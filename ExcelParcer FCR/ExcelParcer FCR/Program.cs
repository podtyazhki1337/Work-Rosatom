using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using OfficeOpenXml;

namespace FCRHybrid
{
    class Program
    {
        static void Main()
        {
            Console.Write("Введите путь к файлу FCR Excel: ");
            string filePath = Console.ReadLine();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Console.WriteLine("Файл не найден или путь пустой. Завершение работы.");
                return;
            }

            Console.WriteLine("Открываем Worksheets[3], снимаем объединения, логируем…");

            // 1) Unmerge + лог
            var dict = UnmergeAndReadAllCells(filePath);

            // 2) Создаем модель
            var fcr = new FcrData();

            // ===== (A) Верхние статические поля (до ~12 строки) =====
            fcr.FieldChangeRequestNo   = GetCell(dict, 2, 3);
            fcr.RegistrationDate       = GetCell(dict, 2, 7);
            fcr.ContractorChangeCoord  = GetCell(dict, 2, 12);

            fcr.ChangeInitiatorOrg     = GetCell(dict, 5, 3);
            fcr.ChangeInitiatorInternal= GetCell(dict, 5, 7);
            fcr.ChangeInitiator        = GetCell(dict, 5, 10);
            fcr.PositionOfChangeInit   = GetCell(dict, 5, 14);

            fcr.TypeOfDocToBeChanged   = GetCell(dict, 6, 4);
            fcr.TypeOfChanges          = GetCell(dict, 6, 7);
            fcr.TypeOfActivity         = GetCell(dict, 6, 11);
            fcr.ConstructionFacility   = GetCell(dict, 6, 15);

            fcr.InitiatorProposalMethod= GetCell(dict, 7, 4);
            fcr.JustificationSimpleMeth= GetCell(dict, 7, 12);

            fcr.ChangeInProjectPosEquip= GetCell(dict, 8, 4);
            fcr.CodeReasonChange       = GetCell(dict, 8, 9);
            fcr.OtherReason            = GetCell(dict, 8, 11);

            // DescriptionEngChange => (r=10,c=1)
            fcr.DescriptionEngChange   = GetCell(dict, 10, 1);

            // ===== (B) KKS (label-based) =====
            fcr.KKSList = ParseKKS(
                dict,
                "List of affected SSC",
                "If the code of the SSC is not specified",
                false
            );

            // ===== (C) Documents =====
            fcr.Documents = ParseDocuments(
                dict,
                "Document Set Code",
                "Supporting and describing documents",
                false
            );

            // ===== (D) Supporting docs =====
            fcr.SupportingDocs = ParseSupportingDocsUntilLabel(
                dict,
                "Supporting and describing documents\nIf the document is attached",
                "The material is equivalent:",
                false
            );

            // ===== (E) Material + CommentsRejectMaterial =====
            ParseMaterialFields(dict, fcr, ignoreCase:false);

            // ===== (F) Link to docs =====
            fcr.LinkToDocs = ParseFileSectionSkipHeaders(
                dict,
                "Link to documents justifying the decision",
                false
            );

            // ===== (G) Impacts =====
            ParseImpacts(
                dict,
                "For the 'Documentation Red Change' type of change",
                out bool nuclear,
                out bool fire,
                out bool industrial,
                out bool environ,
                out bool sched,
                out bool promptDDD,
                out bool promptMDD,
                out bool structural,
                out bool impactDDD,
                out bool licensing,
                out bool cost,
                false
            );
            fcr.NuclearSafety     = nuclear;
            fcr.FireSafety        = fire;
            fcr.IndustrialSafety  = industrial;
            fcr.EnvironmentalSafe = environ;
            fcr.ScheduleImpact2   = sched;
            fcr.PromptReleaseDDD  = promptDDD;
            fcr.PromptReleaseMDD  = promptMDD;
            fcr.StructuralReliab  = structural;
            fcr.ImpactOnOtherDDD  = impactDDD;
            fcr.LicensingDoc      = licensing;
            fcr.CostImpact2       = cost;

            // ===== (H) CommentsRefusalDocs =====
            fcr.CommentsRefusalDocs = GetCell(dict, 81, 1);

            // ===== (I) FinalApproval =====
            fcr.FinalApprovalMethod = GetCell(dict, 85, 3);
            fcr.FinalApprovalJustif = GetCell(dict, 85, 8);

            // ===== (J) Signatures =====
            fcr.Signatures = ParseSignaturesByLabels(
                dict,
                "Project Participant position",
                "* digital signature can be used for signing the FCR",
                false
            );

            // Убираем время в RegistrationDate, Signatures
            fcr.RegistrationDate = TrimDate(fcr.RegistrationDate);
            foreach(var s in fcr.Signatures)
            {
                s.DateVal = TrimDate(s.DateVal);
            }

            // 3) Верификация данных
            Console.WriteLine("\n=== Проверка данных на полноту и соответствие ===");
            bool isValid = VerifyFcrData(fcr);
            
            // 4) Заглушки для записи в БД
            if (isValid)
            {
                Console.WriteLine("\n=== Все данные корректны, выполняется запись в БД ===");
                SaveToDatabase(fcr);
            }
            else
            {
                Console.WriteLine("\n=== Обнаружены ошибки в данных, запись в БД не выполнена ===");
            }

            // 5) Вывод
            PrintFCR(fcr);
        }

        #region ============ (1) Unmerge + GetCell ============

        static Dictionary<CellPos,string> UnmergeAndReadAllCells(string filePath)
        {
            var dict = new Dictionary<CellPos,string>();
            using(var pkg= new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet= pkg.Workbook.Worksheets[3];
                if(sheet==null)
                {
                    Console.WriteLine("Worksheet[3] не найден.");
                    return dict;
                }

                var merges= sheet.MergedCells;
                var mergesCopy= new List<string>(merges);
                foreach(var addr in mergesCopy)
                {
                    sheet.Cells[addr].Merge=false;
                }

                int maxRows=200, maxCols=50;
                Console.WriteLine($"Читаем {maxRows}x{maxCols}, логируем непустые:");

                for(int r=1; r<=maxRows; r++)
                {
                    for(int c=1; c<=maxCols; c++)
                    {
                        var valObj= sheet.Cells[r,c].Value;
                        if(valObj!=null)
                        {
                            string val= valObj.ToString().Trim();
                            if(val.Length>0)
                            {
                                Console.WriteLine($"(r={r}, c={c}) => \"{val}\"");
                                dict[new CellPos(r,c)] = val;
                            }
                        }
                    }
                }
            }
            return dict;
        }

        static string GetCell(Dictionary<CellPos,string> dict, int row, int col)
        {
            var pos= new CellPos(row,col);
            if(dict.TryGetValue(pos, out string val)) return val.Trim();
            return "";
        }

        #endregion

        #region ============ (2) KKS, Docs, SupportingDocs ============

        static List<FcrKksEntry> ParseKKS(
            Dictionary<CellPos,string> dict,
            string startLabel,
            string stopLabel,
            bool ignoreCase
        )
        {
            var list= new List<FcrKksEntry>();
            int rowStart= FindRowByLabel(dict, startLabel, ignoreCase);
            if(rowStart<1) return list;

            for(int r=rowStart+2; r<=9999; r++)
            {
                string valA= GetCell(dict, r,1);
                if(string.IsNullOrEmpty(valA)) break;

                bool stop= ignoreCase
                    ? valA.ToLower().Contains(stopLabel.ToLower())
                    : valA.Contains(stopLabel);
                if(stop) break;

                string building= valA;
                string system=   GetCell(dict, r,2);
                string comp=     GetCell(dict, r,3);

                bool allEmpty= string.IsNullOrEmpty(building)
                             && string.IsNullOrEmpty(system)
                             && string.IsNullOrEmpty(comp);
                if(allEmpty) continue;

                list.Add(new FcrKksEntry
                {
                    BuildingKks= building,
                    SystemKks= system,
                    ComponentKks= comp
                });
            }
            return list;
        }

        static List<FcrDocumentEntry> ParseDocuments(
            Dictionary<CellPos,string> dict,
            string startLabel,
            string stopLabel,
            bool ignoreCase
        )
        {
            var docs= new List<FcrDocumentEntry>();
            int rowStart= FindRowByLabel(dict, startLabel, ignoreCase);
            if(rowStart<1) return docs;

            for(int r=rowStart+1; r<=9999; r++)
            {
                string valA= GetCell(dict, r,1);
                if(string.IsNullOrEmpty(valA)) break;

                bool stop= ignoreCase
                    ? valA.ToLower().Contains(stopLabel.ToLower())
                    : valA.Contains(stopLabel);
                if(stop) break;

                string crDoc= valA;
                string setRev= GetCell(dict, r,3);
                string engDoc= GetCell(dict, r,4);
                string engName=GetCell(dict, r,6);
                string edRev=  GetCell(dict, r,9);
                string sheets= GetCell(dict, r,10);
                string changeA=GetCell(dict, r,11);
                string desc=   GetCell(dict, r,12);

                bool allEmpty= 
                    string.IsNullOrEmpty(crDoc)
                    && string.IsNullOrEmpty(setRev)
                    && string.IsNullOrEmpty(engDoc)
                    && string.IsNullOrEmpty(engName)
                    && string.IsNullOrEmpty(edRev)
                    && string.IsNullOrEmpty(sheets)
                    && string.IsNullOrEmpty(changeA)
                    && string.IsNullOrEmpty(desc);
                if(allEmpty) continue;

                docs.Add(new FcrDocumentEntry
                {
                    CRDocumentSetCode= crDoc,
                    SetRevisionVersion= setRev,
                    EngDocCode= engDoc,
                    EngDocName= engName,
                    EdRevisionVersion= edRev,
                    SheetsOrPageNumbers= sheets,
                    ChangeAMx= changeA,
                    ChangeDescription= desc
                });
            }
            return docs;
        }

        static List<FcrFileNameEntry> ParseSupportingDocsUntilLabel(
            Dictionary<CellPos,string> dict,
            string startLabel,
            string stopLabel,
            bool ignoreCase
        )
        {
            var list= new List<FcrFileNameEntry>();
            int rowStart= FindRowByLabel(dict, startLabel, ignoreCase);
            if(rowStart<1) return list;

            for(int r=rowStart+1; r<=9999; r++)
            {
                string valA= GetCell(dict, r,1);
                if(string.IsNullOrEmpty(valA)) break;

                bool stop= ignoreCase
                    ? valA.ToLower().Contains(stopLabel.ToLower())
                    : valA.Contains(stopLabel);
                if(stop) break;

                string fn= valA;
                string ct= GetCell(dict, r,4);

                bool emptyAll= string.IsNullOrEmpty(fn)
                             && string.IsNullOrEmpty(ct);
                if(emptyAll) continue;

                list.Add(new FcrFileNameEntry
                {
                    FileNameExt= fn,
                    CodeOrTitle= ct
                });
            }
            return list;
        }

        #endregion

        #region ============ (3) ParseMaterialFields ============

        static void ParseMaterialFields(
            Dictionary<CellPos,string> dict,
            FcrData fcr,
            bool ignoreCase
        )
        {
            int matRow= FindRowByLabel(dict, "The material is equivalent:", ignoreCase);
            if(matRow<1) return;

            fcr.MaterialIsEquivalent= GetCell(dict, matRow, 4);
            fcr.ReplaceTypeOfChange = GetCell(dict, matRow, 15);

            int nextRow= matRow+1;
            string labelNext= GetCell(dict, nextRow,1);
            if(labelNext.Contains("Comments and/or reasons to reject approving replacement of material"))
            {
                fcr.CommentsRejectMaterial= GetCell(dict, nextRow+1, 1);
            }
        }

        #endregion

        #region ============ (4) ParseFileSectionSkipHeaders ============

        static List<FcrFileNameEntry> ParseFileSectionSkipHeaders(
            Dictionary<CellPos,string> dict,
            string startLabel,
            bool ignoreCase
        )
        {
            var list= new List<FcrFileNameEntry>();
            int rowStart= FindRowByLabel(dict, startLabel, ignoreCase);
            if(rowStart<1) return list;

            int rData= rowStart+1;
            for(int i=0; i<2; i++)
            {
                var lineA= GetCell(dict, rData,1);
                if(string.IsNullOrEmpty(lineA)) break;
                if(lineA.ToLower().Contains("filename & extension")
                   || lineA.ToLower().Contains("code, title or summary of the document"))
                {
                    rData++;
                }
            }

            for(int r=rData; r<=9999; r++)
            {
                string valA= GetCell(dict, r,1);
                if(string.IsNullOrEmpty(valA)) break;

                string fn= valA;
                string ct= GetCell(dict, r,4);

                bool emptyAll= string.IsNullOrEmpty(fn)
                             && string.IsNullOrEmpty(ct);
                if(emptyAll) continue;

                list.Add(new FcrFileNameEntry
                {
                    FileNameExt= fn,
                    CodeOrTitle= ct
                });
            }
            return list;
        }

        #endregion

        #region ============ (5) ParseImpacts, Signatures, FindRowByLabel ============

        static void ParseImpacts(
            Dictionary<CellPos,string> dict,
            string startLabel,
            out bool nuclear,
            out bool fire,
            out bool industrial,
            out bool environ,
            out bool sched,
            out bool promptDDD,
            out bool promptMDD,
            out bool structural,
            out bool impactDDD,
            out bool licensing,
            out bool cost,
            bool ignoreCase
        )
        {
            nuclear=false; fire=false; industrial=false; environ=false;
            sched=false; promptDDD=false; promptMDD=false;
            structural=false; impactDDD=false; licensing=false; cost=false;

            int rowStart= FindRowByLabel(dict, startLabel, ignoreCase);
            if(rowStart<1) return;

            for(int r=rowStart+1; r<=9999; r++)
            {
                string labelVal= GetCell(dict, r,1);
                if(string.IsNullOrEmpty(labelVal)) break;

                string valB= GetCell(dict, r,2).ToLower();
                bool isYes= (valB=="yes"||valB=="true");

                if(labelVal.Contains("Nuclear Safety")) nuclear= isYes;
                else if(labelVal.Contains("Fire Safety")) fire= isYes;
                else if(labelVal.Contains("Industrial Safety")) industrial= isYes;
                else if(labelVal.Contains("Environmental safety")) environ= isYes;
                else if(labelVal.Contains("Schedule")) sched= isYes;
                else if(labelVal.Contains("Prompt release of a new revision"))
                {
                    promptDDD= isYes;
                    var valNext= GetCell(dict, r+1,2).ToLower();
                    promptMDD= (valNext=="yes"||valNext=="true");
                    r++;
                }
                else if(labelVal.Contains("Structural reliability")) structural= isYes;
                else if(labelVal.Contains("Impact on other")) impactDDD= isYes;
                else if(labelVal.Contains("Licensing Documentation")) licensing= isYes;
                else if(labelVal.Contains("Cost")) cost= isYes;
            }
        }

        static List<FcrSignature> ParseSignaturesByLabels(
            Dictionary<CellPos,string> dict,
            string startLabel,
            string stopLabel,
            bool ignoreCase
        )
        {
            var list= new List<FcrSignature>();
            int rowStart= FindRowByLabel(dict, startLabel, ignoreCase);
            if(rowStart<1) return list;

            for(int r=rowStart+1; r<=9999; r++)
            {
                var valA= GetCell(dict, r,1);
                if(string.IsNullOrEmpty(valA)) break;

                bool stop= ignoreCase
                    ? valA.ToLower().Contains(stopLabel.ToLower())
                    : valA.Contains(stopLabel);
                if(stop) break;

                string pos= valA;
                string name= GetCell(dict, r,7);
                string date= GetCell(dict, r,11);

                bool allEmpty= string.IsNullOrEmpty(pos)
                             && string.IsNullOrEmpty(name)
                             && string.IsNullOrEmpty(date);
                if(allEmpty) continue;

                list.Add(new FcrSignature
                {
                    Position= pos,
                    Name= name,
                    DateVal= date
                });
            }
            return list;
        }

        static int FindRowByLabel(Dictionary<CellPos,string> dict, string label, bool ignoreCase)
        {
            foreach(var kvp in dict)
            {
                if(kvp.Key.Col==1)
                {
                    string cellVal= kvp.Value;
                    bool match= ignoreCase
                        ? cellVal.ToLower().Contains(label.ToLower())
                        : cellVal.Contains(label);
                    if(match) return kvp.Key.Row;
                }
            }
            return 0;
        }

        #endregion

        #region ============ (6) TrimDate, Print ============

        static string TrimDate(string input)
        {
            if(string.IsNullOrEmpty(input)) return "";
            if(DateTime.TryParse(input, out DateTime dt))
            {
                return dt.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
            }
            return input;
        }

        static void PrintFCR(FcrData fcr)
        {
            Console.WriteLine("\n=== Итоговый вывод ===\n");

            Console.WriteLine($"FieldChangeRequestNo:   {fcr.FieldChangeRequestNo}");
            Console.WriteLine($"RegistrationDate:       {fcr.RegistrationDate}");
            Console.WriteLine($"ContractorChangeCoord:  {fcr.ContractorChangeCoord}");
            Console.WriteLine($"ChangeInitiatorOrg:     {fcr.ChangeInitiatorOrg}");
            Console.WriteLine($"ChangeInitiatorInternal:{fcr.ChangeInitiatorInternal}");
            Console.WriteLine($"ChangeInitiator:        {fcr.ChangeInitiator}");
            Console.WriteLine($"PositionOfChangeInit:   {fcr.PositionOfChangeInit}");
            Console.WriteLine($"TypeOfDocToBeChanged:   {fcr.TypeOfDocToBeChanged}");
            Console.WriteLine($"TypeOfChanges:          {fcr.TypeOfChanges}");
            Console.WriteLine($"TypeOfActivity:         {fcr.TypeOfActivity}");
            Console.WriteLine($"ConstructionFacility:   {fcr.ConstructionFacility}");
            Console.WriteLine($"InitiatorProposalMethod:{fcr.InitiatorProposalMethod}");
            Console.WriteLine($"JustificationSimpleMeth:{fcr.JustificationSimpleMeth}");
            Console.WriteLine($"ChangeInProjectPosEquip:{fcr.ChangeInProjectPosEquip}");
            Console.WriteLine($"CodeReasonChange:       {fcr.CodeReasonChange}");
            Console.WriteLine($"OtherReason:            {fcr.OtherReason}");
            Console.WriteLine($"DescriptionEngChange:   {fcr.DescriptionEngChange}");

            Console.WriteLine("\n--- KKS ---");
            if(fcr.KKSList.Count>0)
            {
                int i=1;
                foreach(var k in fcr.KKSList)
                {
                    Console.WriteLine($"  {i++}) B='{k.BuildingKks}', S='{k.SystemKks}', C='{k.ComponentKks}'");
                }
            }
            else Console.WriteLine("  (пусто)");

            Console.WriteLine("\n--- Documents ---");
            if(fcr.Documents.Count>0)
            {
                int i=1;
                foreach(var d in fcr.Documents)
                {
                    Console.WriteLine($"  {i++}) CRDoc='{d.CRDocumentSetCode}', Rev='{d.SetRevisionVersion}', EngDoc='{d.EngDocCode}', EDName='{d.EngDocName}', EDRev='{d.EdRevisionVersion}', Sheets='{d.SheetsOrPageNumbers}', AMx='{d.ChangeAMx}', Desc='{d.ChangeDescription}'");
                }
            }
            else Console.WriteLine("  (пусто)");

            Console.WriteLine("\n--- Supporting and describing documents ---");
            if(fcr.SupportingDocs.Count>0)
            {
                int i=1;
                foreach(var sd in fcr.SupportingDocs)
                {
                    Console.WriteLine($"  {i++}) File='{sd.FileNameExt}', Title='{sd.CodeOrTitle}'");
                }
            }
            else Console.WriteLine("  (пусто)");

            Console.WriteLine($"\nMaterialIsEquivalent:   {fcr.MaterialIsEquivalent}");
            Console.WriteLine($"ReplaceTypeOfChange:    {fcr.ReplaceTypeOfChange}");
            Console.WriteLine($"CommentsRejectMaterial: {fcr.CommentsRejectMaterial}");

            Console.WriteLine("\n--- Link to documents justifying the decision ---");
            if(fcr.LinkToDocs.Count>0)
            {
                int i=1;
                foreach(var doc in fcr.LinkToDocs)
                {
                    Console.WriteLine($"  {i++}) File='{doc.FileNameExt}', Title='{doc.CodeOrTitle}'");
                }
            }
            else Console.WriteLine("  (пусто)");

            Console.WriteLine("\n--- Impacts ---");
            Console.WriteLine($"NuclearSafety:      {fcr.NuclearSafety}");
            Console.WriteLine($"FireSafety:         {fcr.FireSafety}");
            Console.WriteLine($"IndustrialSafety:   {fcr.IndustrialSafety}");
            Console.WriteLine($"EnvironmentalSafe:  {fcr.EnvironmentalSafe}");
            Console.WriteLine($"ScheduleImpact2:    {fcr.ScheduleImpact2}");
            Console.WriteLine($"PromptReleaseDDD:   {fcr.PromptReleaseDDD}");
            Console.WriteLine($"PromptReleaseMDD:   {fcr.PromptReleaseMDD}");
            Console.WriteLine($"StructuralReliab:   {fcr.StructuralReliab}");
            Console.WriteLine($"ImpactOnOtherDDD:   {fcr.ImpactOnOtherDDD}");
            Console.WriteLine($"LicensingDoc:       {fcr.LicensingDoc}");
            Console.WriteLine($"CostImpact2:        {fcr.CostImpact2}");

            Console.WriteLine($"\nCommentsRefusalDocs: {fcr.CommentsRefusalDocs}");

            Console.WriteLine($"FinalApprovalMethod: {fcr.FinalApprovalMethod}");
            Console.WriteLine($"FinalApprovalJustif: {fcr.FinalApprovalJustif}");

            Console.WriteLine("\n--- Project Participants (Signatures) ---");
            if(fcr.Signatures.Count>0)
            {
                int i=1;
                foreach(var s in fcr.Signatures)
                {
                    Console.WriteLine($"  {i++}) Position='{s.Position}', Name='{s.Name}', Date='{s.DateVal}'");
                }
            }
            else Console.WriteLine("  (пусто)");

            Console.WriteLine("\n=== Конец ===");
        }

        #endregion

        #region ============ (7) Верификация и БД ============

        private static bool VerifyFcrData(FcrData fcr)
        {
            bool isValid = true;
            List<string> errors = new List<string>();

            if (string.IsNullOrEmpty(fcr.FieldChangeRequestNo))
                errors.Add("Field Change Request No is required");
            if (string.IsNullOrEmpty(fcr.RegistrationDate))
                errors.Add("Registration date is required");
            if (string.IsNullOrEmpty(fcr.ContractorChangeCoord))
                errors.Add("Contractor’s Change Coordinator is required");
            if (string.IsNullOrEmpty(fcr.ChangeInitiatorOrg))
                errors.Add("Change Initiator's organization is required");
            if (string.IsNullOrEmpty(fcr.ChangeInitiator))
                errors.Add("Change Initiator is required");
            if (string.IsNullOrEmpty(fcr.PositionOfChangeInit))
                errors.Add("Position of the Change Initiator is required");
            if (string.IsNullOrEmpty(fcr.CodeReasonChange))
                errors.Add("Code of reason of engineering change is required");
            if (string.IsNullOrEmpty(fcr.DescriptionEngChange))
                errors.Add("Description of Engineering Change is required");

            string[] validDocTypes = { "DDD", "MDD" };
            if (!string.IsNullOrEmpty(fcr.TypeOfDocToBeChanged) && !Array.Exists(validDocTypes, x => x == fcr.TypeOfDocToBeChanged))
                errors.Add($"Type of documentation to be changed must be one of: {string.Join(", ", validDocTypes)}");

            string[] validChangeTypes = { "Replacing Of Materials", "Documentation Red Change" };
            if (!string.IsNullOrEmpty(fcr.TypeOfChanges) && !Array.Exists(validChangeTypes, x => x == fcr.TypeOfChanges))
                errors.Add($"Type of changes must be one of: {string.Join(", ", validChangeTypes)}");

            string[] validActivityTypes = { "Construction", "Manufacturing" };
            if (!string.IsNullOrEmpty(fcr.TypeOfActivity) && !Array.Exists(validActivityTypes, x => x == fcr.TypeOfActivity))
                errors.Add($"Type of activity must be one of: {string.Join(", ", validActivityTypes)}");

            string[] validFacilities = { "NPP", "CEB" };
            if (!string.IsNullOrEmpty(fcr.ConstructionFacility) && !Array.Exists(validFacilities, x => x == fcr.ConstructionFacility))
                errors.Add($"Construction facility must be one of: {string.Join(", ", validFacilities)}");

            string[] validApprovalMethods = { "Normal", "Simple" };
            if (!string.IsNullOrEmpty(fcr.InitiatorProposalMethod) && !Array.Exists(validApprovalMethods, x => x == fcr.InitiatorProposalMethod))
                errors.Add($"Initiator's proposal for choosing a approval method must be one of: {string.Join(", ", validApprovalMethods)}");

            if (!string.IsNullOrEmpty(fcr.FinalApprovalMethod) && !Array.Exists(validApprovalMethods, x => x == fcr.FinalApprovalMethod))
                errors.Add($"Final approval method must be one of: {string.Join(", ", validApprovalMethods)}");

            if (fcr.InitiatorProposalMethod == "Simple" && string.IsNullOrEmpty(fcr.JustificationSimpleMeth))
                errors.Add("Justification of the 'simple' approval method is required when Initiator's proposal is Simple");

            foreach (var kks in fcr.KKSList)
            {
                if (!string.IsNullOrEmpty(kks.BuildingKks) && (string.IsNullOrEmpty(kks.SystemKks) || string.IsNullOrEmpty(kks.ComponentKks)))
                    errors.Add($"For Building KKS '{kks.BuildingKks}', System KKS and Component KKS are required");
            }

            foreach (var doc in fcr.Documents)
            {
                if (!string.IsNullOrEmpty(doc.CRDocumentSetCode) && 
                    (string.IsNullOrEmpty(doc.SetRevisionVersion) ||
                     string.IsNullOrEmpty(doc.EngDocCode) ||
                     string.IsNullOrEmpty(doc.EngDocName) ||
                     string.IsNullOrEmpty(doc.EdRevisionVersion) ||
                     string.IsNullOrEmpty(doc.SheetsOrPageNumbers) ||
                     string.IsNullOrEmpty(doc.ChangeAMx)))
                    errors.Add($"For Document Set Code '{doc.CRDocumentSetCode}', all fields (Set Revision, Eng Doc Code, Eng Doc Name, ED Revision, Sheets, Change AMx) are required");
            }

            if (fcr.FinalApprovalMethod == "Simple" && string.IsNullOrEmpty(fcr.FinalApprovalJustif))
                errors.Add("Justification is required when Final approval method is Simple");

            if (errors.Count > 0)
            {
                isValid = false;
                Console.WriteLine("Обнаружены следующие ошибки:");
                for (int i = 0; i < errors.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {errors[i]}");
                }
            }
            else
            {
                Console.WriteLine("Все проверки пройдены успешно.");
            }

            return isValid;
        }

        private static void SaveToDatabase(FcrData fcr)
        {
            try
            {
                Console.WriteLine("Подключение к базе данных...");
                Console.WriteLine("Сохранение основных данных FCR...");
                Console.WriteLine($"INSERT INTO FcrMain (FieldChangeRequestNo, RegistrationDate, ContractorChangeCoord, ChangeInitiatorOrg, ChangeInitiator, PositionOfChangeInit, TypeOfDocToBeChanged, TypeOfChanges, TypeOfActivity, ConstructionFacility, InitiatorProposalMethod, JustificationSimpleMeth, CodeReasonChange, DescriptionEngChange, FinalApprovalMethod, FinalApprovalJustif) " +
                    $"VALUES ('{fcr.FieldChangeRequestNo}', '{fcr.RegistrationDate}', '{fcr.ContractorChangeCoord}', '{fcr.ChangeInitiatorOrg}', '{fcr.ChangeInitiator}', '{fcr.PositionOfChangeInit}', '{fcr.TypeOfDocToBeChanged}', '{fcr.TypeOfChanges}', '{fcr.TypeOfActivity}', '{fcr.ConstructionFacility}', '{fcr.InitiatorProposalMethod}', '{fcr.JustificationSimpleMeth}', '{fcr.CodeReasonChange}', '{fcr.DescriptionEngChange}', '{fcr.FinalApprovalMethod}', '{fcr.FinalApprovalJustif}')");

                if (fcr.KKSList.Count > 0)
                {
                    Console.WriteLine("Сохранение KKS данных...");
                    foreach (var kks in fcr.KKSList)
                    {
                        Console.WriteLine($"INSERT INTO FcrKKS (BuildingKks, SystemKks, ComponentKks) VALUES ('{kks.BuildingKks}', '{kks.SystemKks}', '{kks.ComponentKks}')");
                    }
                }

                if (fcr.Documents.Count > 0)
                {
                    Console.WriteLine("Сохранение данных документов...");
                    foreach (var doc in fcr.Documents)
                    {
                        Console.WriteLine($"INSERT INTO FcrDocuments (CRDocumentSetCode, SetRevisionVersion, EngDocCode, EngDocName, EdRevisionVersion, SheetsOrPageNumbers, ChangeAMx, ChangeDescription) " +
                            $"VALUES ('{doc.CRDocumentSetCode}', '{doc.SetRevisionVersion}', '{doc.EngDocCode}', '{doc.EngDocName}', '{doc.EdRevisionVersion}', '{doc.SheetsOrPageNumbers}', '{doc.ChangeAMx}', '{doc.ChangeDescription}')");
                    }
                }

                if (fcr.SupportingDocs.Count > 0)
                {
                    Console.WriteLine("Сохранение SupportingDocs...");
                    foreach (var sd in fcr.SupportingDocs)
                    {
                        Console.WriteLine($"INSERT INTO FcrSupportingDocs (FileNameExt, CodeOrTitle) VALUES ('{sd.FileNameExt}', '{sd.CodeOrTitle}')");
                    }
                }

                if (fcr.LinkToDocs.Count > 0)
                {
                    Console.WriteLine("Сохранение LinkToDocs...");
                    foreach (var ld in fcr.LinkToDocs)
                    {
                        Console.WriteLine($"INSERT INTO FcrLinkToDocs (FileNameExt, CodeOrTitle) VALUES ('{ld.FileNameExt}', '{ld.CodeOrTitle}')");
                    }
                }

                if (fcr.Signatures.Count > 0)
                {
                    Console.WriteLine("Сохранение Signatures...");
                    foreach (var sig in fcr.Signatures)
                    {
                        Console.WriteLine($"INSERT INTO FcrSignatures (Position, Name, DateVal) VALUES ('{sig.Position}', '{sig.Name}', '{sig.DateVal}')");
                    }
                }

                Console.WriteLine("Запись в БД успешно завершена.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сохранении в БД: {ex.Message}");
            }
        }

        #endregion

    }

    #region ============ (8) МОДЕЛИ ============

    public struct CellPos
    {
        public int Row, Col;
        public CellPos(int r,int c){ Row=r; Col=c; }
        public override bool Equals(object obj)
        {
            if(!(obj is CellPos))return false;
            var other=(CellPos)obj;
            return Row==other.Row && Col==other.Col;
        }
        public override int GetHashCode(){ return (Row<<16)^Col; }
    }

    public class FcrData
    {
        public string FieldChangeRequestNo;
        public string RegistrationDate;
        public string ContractorChangeCoord;

        public string ChangeInitiatorOrg;
        public string ChangeInitiatorInternal;
        public string ChangeInitiator;
        public string PositionOfChangeInit;

        public string TypeOfDocToBeChanged;
        public string TypeOfChanges;
        public string TypeOfActivity;
        public string ConstructionFacility;

        public string InitiatorProposalMethod;
        public string JustificationSimpleMeth;

        public string ChangeInProjectPosEquip;
        public string CodeReasonChange;
        public string OtherReason;

        public string DescriptionEngChange;

        public List<FcrKksEntry> KKSList= new List<FcrKksEntry>();
        public List<FcrDocumentEntry> Documents= new List<FcrDocumentEntry>();
        public List<FcrFileNameEntry> SupportingDocs= new List<FcrFileNameEntry>();

        public string MaterialIsEquivalent;
        public string ReplaceTypeOfChange;
        public string CommentsRejectMaterial;

        public List<FcrFileNameEntry> LinkToDocs= new List<FcrFileNameEntry>();

        public bool NuclearSafety;
        public bool FireSafety;
        public bool IndustrialSafety;
        public bool EnvironmentalSafe;
        public bool ScheduleImpact2;
        public bool PromptReleaseDDD;
        public bool PromptReleaseMDD;
        public bool StructuralReliab;
        public bool ImpactOnOtherDDD;
        public bool LicensingDoc;
        public bool CostImpact2;

        public string CommentsRefusalDocs;

        public string FinalApprovalMethod;
        public string FinalApprovalJustif;

        public List<FcrSignature> Signatures= new List<FcrSignature>();
    }

    public class FcrKksEntry
    {
        public string BuildingKks;
        public string SystemKks;
        public string ComponentKks;
    }

    public class FcrDocumentEntry
    {
        public string CRDocumentSetCode;
        public string SetRevisionVersion;
        public string EngDocCode;
        public string EngDocName;
        public string EdRevisionVersion;
        public string SheetsOrPageNumbers;
        public string ChangeAMx;
        public string ChangeDescription;
    }

    public class FcrFileNameEntry
    {
        public string FileNameExt;
        public string CodeOrTitle;
    }

    public class FcrSignature
    {
        public string Position;
        public string Name;
        public string DateVal;
    }

    #endregion
}