using System; 
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace FCRExcelWriterReverse
{
    class Program
    {
        static void Main()
        {
            Console.Write("Введите путь к существующему Excel (шаблону FCR): ");
            string filePath = Console.ReadLine();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Console.WriteLine("Файл не найден. Завершение работы.");
                return;
            }

            // 1) Имитация данных (SQL)
            FcrData fcr = GetFcrDataFromSQLStub();

            // 2) Запись в Excel
            try
            {
                WriteFcrToExcel(filePath, fcr);
                Console.WriteLine("\nДанные успешно записаны в файл: " + filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при записи в Excel: " + ex.Message);
            }

            Console.WriteLine("Нажмите любую клавишу для завершения...");
            Console.ReadKey();
        }

        #region (A) "Заглушка" - тестовые данные

        private static FcrData GetFcrDataFromSQLStub()
        {
            // Пример заполнения FcrData
            var fcr = new FcrData
            {
                FieldChangeRequestNo     = "FCR-2025-REV-TEST",
                RegistrationDate         = "21.03.2025",
                ContractorChangeCoord    = "ContractorCoordName",
                ChangeInitiatorOrg       = "OrgName",
                ChangeInitiatorInternal  = "InternalCoord01",
                ChangeInitiator          = "Initiator Person",
                PositionOfChangeInit     = "Lead Engineer",
                TypeOfDocToBeChanged     = "DDD",
                TypeOfChanges            = "Replacing Of Materials",
                TypeOfActivity           = "Construction",
                ConstructionFacility     = "NPP",
                InitiatorProposalMethod  = "Simple",
                JustificationSimpleMeth  = "Justification text for Simple method",
                ChangeInProjectPosEquip  = true,
                CodeReasonChange         = "R-99",
                OtherReason              = "Some free-text reason",
                DescriptionEngChange     = "Detailed description of the FCR...",

                MaterialIsEquivalent     = true,
                ReplaceTypeOfChange      = true,
                CommentsRejectMaterial   = "No rejection so far.",

                NuclearSafety            = true,
                FireSafety               = false,
                IndustrialSafety         = true,
                EnvironmentalSafe        = false,
                ScheduleImpact2          = false,
                PromptReleaseDDD         = true,
                PromptReleaseMDD         = false,
                StructuralReliab         = false,
                ImpactOnOtherDDD         = true,
                LicensingDoc             = false,
                CostImpact2              = true,

                CommentsRefusalDocs      = "No refusal reasons.",
                FinalApprovalMethod      = "Normal procedure",
                FinalApprovalJustif      = "All participants agreed."
            };

            // KKS (5 строк)
            fcr.KKSList.Add(new FcrKksEntry { BuildingKks="B1", SystemKks="S1", ComponentKks="C1" });
            fcr.KKSList.Add(new FcrKksEntry { BuildingKks="B2", SystemKks="S2", ComponentKks="C2" });
            fcr.KKSList.Add(new FcrKksEntry { BuildingKks="B3", SystemKks="S3", ComponentKks="C3" });
            fcr.KKSList.Add(new FcrKksEntry { BuildingKks="B4", SystemKks="S4", ComponentKks="C4" });
            fcr.KKSList.Add(new FcrKksEntry { BuildingKks="B5", SystemKks="S5", ComponentKks="C5" });

            // Documents (5 строк)
            fcr.Documents.Add(new FcrDocumentEntry {
                CRDocumentSetCode="DOCSET-001",
                SetRevisionVersion="RevA",
                EngDocCode="ENG-001",
                EngDocName="Piping Layout",
                EdRevisionVersion="E1",
                SheetsOrPageNumbers="1-5",
                ChangeAMx="AM1",
                ChangeDescription="Minor text fix"
            });
            fcr.Documents.Add(new FcrDocumentEntry {
                CRDocumentSetCode="DOCSET-002",
                SetRevisionVersion="RevB",
                EngDocCode="ENG-002",
                EngDocName="Electrical Diagram",
                EdRevisionVersion="E2",
                SheetsOrPageNumbers="10-12",
                ChangeAMx="AM2",
                ChangeDescription="Replaced symbol"
            });
            fcr.Documents.Add(new FcrDocumentEntry {
                CRDocumentSetCode="DOCSET-003",
                SetRevisionVersion="RevC",
                EngDocCode="ENG-003",
                EngDocName="Layout of Pump",
                EdRevisionVersion="E3",
                SheetsOrPageNumbers="15-16",
                ChangeAMx="AM3",
                ChangeDescription="Added new note"
            });
            fcr.Documents.Add(new FcrDocumentEntry {
                CRDocumentSetCode="DOCSET-004",
                SetRevisionVersion="RevD",
                EngDocCode="ENG-004",
                EngDocName="Wiring Diagram",
                EdRevisionVersion="E4",
                SheetsOrPageNumbers="20-22",
                ChangeAMx="AM4",
                ChangeDescription="Changed cable specs"
            });
            fcr.Documents.Add(new FcrDocumentEntry {
                CRDocumentSetCode="DOCSET-005",
                SetRevisionVersion="RevE",
                EngDocCode="ENG-005",
                EngDocName="Instrumentation Scheme",
                EdRevisionVersion="E5",
                SheetsOrPageNumbers="30-31",
                ChangeAMx="AM5",
                ChangeDescription="Clarified signals"
            });

            // SupportingDocs (4 строки)
            fcr.SupportingDocs.Add(new FcrFileNameEntry {
                FileNameExt="Attach1.pdf",
                CodeOrTitle="Attachment #1"
            });
            fcr.SupportingDocs.Add(new FcrFileNameEntry {
                FileNameExt="Attach2.pdf",
                CodeOrTitle="Attachment #2"
            });
            fcr.SupportingDocs.Add(new FcrFileNameEntry {
                FileNameExt="Attach3.pdf",
                CodeOrTitle="Attachment #3"
            });
            fcr.SupportingDocs.Add(new FcrFileNameEntry {
                FileNameExt="Attach4.pdf",
                CodeOrTitle="Attachment #4"
            });

            // LinkToDocs (5 штук)
            fcr.LinkToDocs.Add(new FcrFileNameEntry {
                FileNameExt="LinkDoc1.docx",
                CodeOrTitle="Doc #1 for justification"
            });
            fcr.LinkToDocs.Add(new FcrFileNameEntry {
                FileNameExt="LinkCalc2.xlsx",
                CodeOrTitle="Calc #2 for justification"
            });
            fcr.LinkToDocs.Add(new FcrFileNameEntry {
                FileNameExt="LinkDesc3.pdf",
                CodeOrTitle="Description #3"
            });
            fcr.LinkToDocs.Add(new FcrFileNameEntry {
                FileNameExt="LinkDoc4.docx",
                CodeOrTitle="Justification #4"
            });
            fcr.LinkToDocs.Add(new FcrFileNameEntry {
                FileNameExt="LinkTable5.xlsx",
                CodeOrTitle="Table #5 with reasons"
            });

            // 6 подписей
            fcr.Signatures.Add(new FcrSignature {
                Position="Position #1",
                Name="Ivanov I.I.",
                DateVal="2025-03-25"
            });
            fcr.Signatures.Add(new FcrSignature {
                Position="Position #2",
                Name="Petrov P.P.",
                DateVal="2025-03-26"
            });
            fcr.Signatures.Add(new FcrSignature {
                Position="Position #3",
                Name="Sidorov S.S.",
                DateVal="2025-03-27"
            });
            fcr.Signatures.Add(new FcrSignature {
                Position="Position #4",
                Name="Fedorov F.F.",
                DateVal="2025-03-28"
            });
            fcr.Signatures.Add(new FcrSignature {
                Position="Position #5",
                Name="Smirnov S.S.",
                DateVal="2025-03-29"
            });
            fcr.Signatures.Add(new FcrSignature {
                Position="Position #6",
                Name="Ershov E.E.",
                DateVal="2025-03-30"
            });

            return fcr;
        }

        #endregion

        #region (B) Основной метод записи

        private static void WriteFcrToExcel(string filePath, FcrData fcr)
        {
            var fi = new FileInfo(filePath);
            using (var package = new ExcelPackage(fi))
            {
                // Worksheet[3]
                var sheet = package.Workbook.Worksheets[3];
                if (sheet == null)
                {
                    Console.WriteLine("Ошибка: не найден Worksheet[3].");
                    return;
                }

                int totalRows = sheet.Dimension.Rows;
                int totalCols = sheet.Dimension.Columns;
                Console.WriteLine($"Лист [3]: {totalRows} строк, {totalCols} столбцов.");

                // ===== (B.1) Статические ячейки =====
                sheet.Cells[2, 3].Value  = fcr.FieldChangeRequestNo;
                sheet.Cells[2, 7].Value  = fcr.RegistrationDate;
                sheet.Cells[2, 12].Value = fcr.ContractorChangeCoord;

                sheet.Cells[5, 3].Value  = fcr.ChangeInitiatorOrg;
                sheet.Cells[5, 7].Value  = fcr.ChangeInitiatorInternal;
                sheet.Cells[5, 10].Value = fcr.ChangeInitiator;
                sheet.Cells[5, 14].Value = fcr.PositionOfChangeInit;

                sheet.Cells[6, 4].Value  = fcr.TypeOfDocToBeChanged;
                sheet.Cells[6, 7].Value  = fcr.TypeOfChanges;
                sheet.Cells[6, 11].Value = fcr.TypeOfActivity;
                sheet.Cells[6, 15].Value = fcr.ConstructionFacility;

                sheet.Cells[7, 4].Value  = fcr.InitiatorProposalMethod;
                sheet.Cells[7, 12].Value = fcr.JustificationSimpleMeth;

                sheet.Cells[8, 4].Value  = fcr.ChangeInProjectPosEquip ? "Yes" : "No";
                sheet.Cells[8, 9].Value  = fcr.CodeReasonChange;
                sheet.Cells[8, 11].Value = fcr.OtherReason;

                sheet.Cells[10, 1].Value = fcr.DescriptionEngChange;

                // (1) List of affected SSC
                int rowKksHeader = FindRowByLabel(sheet, totalRows, "List of affected SSC");
                if (rowKksHeader > 0)
                {
                    int rowKksDataStart = rowKksHeader + 2;
                    int rowKksEnd = FindRowByLabel(sheet, totalRows, "If the code of the SSC is not specified");
                    if (rowKksEnd == 0) rowKksEnd = totalRows + 1;

                    int reservedKks = rowKksEnd - rowKksDataStart;
                    if (reservedKks < 0) reservedKks = 0;
                    int neededKks = fcr.KKSList.Count;

                    if (neededKks > reservedKks)
                    {
                        int toAdd = neededKks - reservedKks;
                        sheet.InsertRow(rowKksEnd, toAdd);
                        int refRow = rowKksEnd - 1;
                        for (int i = 0; i < toAdd; i++)
                        {
                            int newRow = rowKksEnd + i;
                            CopyRowFormatting(sheet, refRow, newRow, totalCols);
                            UnmergeRow(sheet, newRow);
                            ApplyMergeFromReferenceRow(sheet, refRow, newRow);
                        }
                        totalRows = sheet.Dimension.Rows;
                        rowKksEnd += toAdd;
                    }

                    for (int i = 0; i < neededKks; i++)
                    {
                        int row = rowKksDataStart + i;
                        var k = fcr.KKSList[i];
                        sheet.Cells[row, 1].Value  = k.BuildingKks;
                        sheet.Cells[row, 5].Value  = k.SystemKks;
                        sheet.Cells[row, 11].Value = k.ComponentKks;
                    }
                }

                // (2) List of the relevant documents
                {
                    int rowDocsHeader = FindRowByLabel(sheet, totalRows, "List of the relevant documents");
                    if (rowDocsHeader > 0)
                    {
                        int rowDocsDataStart = rowDocsHeader + 3; 
                        int rowDocsEnd = FindRowByLabel(sheet, totalRows, "Supporting and describing documents");
                        if (rowDocsEnd == 0) rowDocsEnd = totalRows + 1;

                        int reservedDocs = rowDocsEnd - rowDocsDataStart;
                        if (reservedDocs < 0) reservedDocs = 0;
                        int neededDocs = fcr.Documents.Count;

                        if (neededDocs > reservedDocs)
                        {
                            int toAdd = neededDocs - reservedDocs;
                            sheet.InsertRow(rowDocsEnd, toAdd);
                            int refRow = rowDocsEnd - 1;
                            for (int i = 0; i < toAdd; i++)
                            {
                                int newRow = rowDocsEnd + i;
                                CopyRowFormatting(sheet, refRow, newRow, totalCols);
                                UnmergeRow(sheet, newRow);
                                ApplyMergeFromReferenceRow(sheet, refRow, newRow);
                            }
                            totalRows = sheet.Dimension.Rows;
                            rowDocsEnd += toAdd;
                        }

                        for (int i = 0; i < neededDocs; i++)
                        {
                            int row = rowDocsDataStart + i;
                            var doc = fcr.Documents[i];
                            sheet.Cells[row, 1].Value  = doc.CRDocumentSetCode;
                            sheet.Cells[row, 3].Value  = doc.SetRevisionVersion;
                            sheet.Cells[row, 4].Value  = doc.EngDocCode;
                            sheet.Cells[row, 6].Value  = doc.EngDocName;
                            sheet.Cells[row, 9].Value  = doc.EdRevisionVersion;
                            sheet.Cells[row, 10].Value = doc.SheetsOrPageNumbers;
                            sheet.Cells[row, 11].Value = doc.ChangeAMx;
                            sheet.Cells[row, 12].Value = doc.ChangeDescription;
                        }
                    }
                }

                // (3) Supporting and describing documents
                //  --- 3 строки зарезервированы, если >3 => вставить
                {
                    int rowSuppHeader = FindRowByLabel(sheet, totalRows,
                        "Supporting and describing documents\nIf the document is attached");
                    if (rowSuppHeader > 0)
                    {
                        // Допустим, шаблон имеет 3 строки начиная с rowSuppHeader+1
                        int rowSuppDataStart = rowSuppHeader + 2;
                        int reservedSupp = 3;  // ровно 3 строки зарезервированы
                        int neededSupp = fcr.SupportingDocs.Count;

                        // строка после зарезервированных:
                        int insertPos = rowSuppDataStart + reservedSupp;

                        if (neededSupp > reservedSupp)
                        {
                            // вставляем
                            int toAdd = neededSupp - reservedSupp;
                            sheet.InsertRow(insertPos, toAdd);

                            // Копируем стили с последней из 3 резервных (insertPos-1)
                            int refRow = insertPos - 2;
                            for (int i = 0; i < toAdd; i++)
                            {
                                int newRow = insertPos + i;
                                CopyRowFormatting(sheet, refRow, newRow, totalCols);
                                UnmergeRow(sheet, newRow);
                                ApplyMergeFromReferenceRow(sheet, refRow, newRow);
                            }
                            totalRows = sheet.Dimension.Rows;
                        }

                        // Теперь заполняем все neededSupp
                        for (int i = 0; i < neededSupp; i++)
                        {
                            int row = rowSuppDataStart + i;
                            var sdoc = fcr.SupportingDocs[i];
                            sheet.Cells[row, 1].Value = sdoc.FileNameExt;
                            sheet.Cells[row, 4].Value = sdoc.CodeOrTitle;
                        }
                    }
                }

                // (4) The material is equivalent => col D
                {
                    int rowMatEquivalent = FindRowByLabel(sheet, totalRows, "The material is equivalent");
                    if (rowMatEquivalent > 0)
                    {
                        sheet.Cells[rowMatEquivalent, 4].Value = fcr.MaterialIsEquivalent ? "Yes" : "No";
                        sheet.Cells[rowMatEquivalent, 15].Value = fcr.ReplaceTypeOfChange ? "Yes" : "No"; // column O
                        
                    }
                }

                // (5) Comments reasons => A, next line
                {
                    int rowReject = FindRowByLabel(sheet, totalRows,
                        "Comments and/or reasons to reject approving replacement of material");
                    if (rowReject > 0)
                    {
                        sheet.Cells[rowReject + 1, 1].Value = fcr.CommentsRejectMaterial;
                    }
                }

                // (6) Первый Link
                {
                    int rowLinkHeader = FindRowByLabel(sheet, totalRows, "Link to documents justifying the decision");
                    if (rowLinkHeader > 0)
                    {
                        int rowLinkDataStart = rowLinkHeader + 2;
                        int rowNextChange = FindRowByLabel(sheet, totalRows,
                            "For the 'Documentation Red Change' type of change");
                        if (rowNextChange == 0) rowNextChange = totalRows + 1;

                        int reservedLinks = rowNextChange - rowLinkDataStart;
                        if (reservedLinks < 0) reservedLinks = 0;
                        int neededLinks = fcr.LinkToDocs.Count;

                        if (neededLinks > reservedLinks)
                        {
                            int toAdd = neededLinks - reservedLinks;
                            sheet.InsertRow(rowNextChange, toAdd);
                            int refRow = rowNextChange - 2;
                            for (int i = 0; i < toAdd; i++)
                            {
                                int newRow = rowNextChange + i;
                                CopyRowFormatting(sheet, refRow, newRow, totalCols);
                                UnmergeRow(sheet, newRow);
                                ApplyMergeFromReferenceRow(sheet, refRow, newRow);
                            }
                            totalRows = sheet.Dimension.Rows;
                            rowNextChange += toAdd;
                        }

                        for (int i = 0; i < neededLinks; i++)
                        {
                            int row = rowLinkDataStart + i;
                            var ld = fcr.LinkToDocs[i];
                            sheet.Cells[row, 1].Value = ld.FileNameExt;
                            sheet.Cells[row, 4].Value = ld.CodeOrTitle;
                        }
                    }
                }

                // (7) Impacts - ищем "Nuclear Safety:"
                {
                    int rowNucSafety = FindRowByLabel(sheet, totalRows, "Nuclear Safety:");
                    if (rowNucSafety > 0)
                    {
                        sheet.Cells[rowNucSafety, 2].Value  = fcr.NuclearSafety     ? "Yes" : "No";
                        sheet.Cells[rowNucSafety, 4].Value  = fcr.FireSafety        ? "Yes" : "No";
                        sheet.Cells[rowNucSafety, 6].Value  = fcr.IndustrialSafety  ? "Yes" : "No";
                        sheet.Cells[rowNucSafety, 9].Value  = fcr.EnvironmentalSafe ? "Yes" : "No";
                        sheet.Cells[rowNucSafety, 11].Value = fcr.ScheduleImpact2   ? "Yes" : "No";
                        sheet.Cells[rowNucSafety, 15].Value = fcr.PromptReleaseDDD  ? "Yes" : "No";

                        sheet.Cells[rowNucSafety + 1, 3].Value  = fcr.StructuralReliab ? "Yes" : "No";
                        sheet.Cells[rowNucSafety + 1, 6].Value  = fcr.ImpactOnOtherDDD ? "Yes" : "No";
                        sheet.Cells[rowNucSafety + 1, 9].Value  = fcr.LicensingDoc     ? "Yes" : "No";
                        sheet.Cells[rowNucSafety + 1, 11].Value = fcr.CostImpact2      ? "Yes" : "No";
                    }
                }

                // (8) Comments refusal => A, next line
                {
                    int rowRefuse2 = FindRowByLabel(sheet, totalRows,
                        "Comments and/or reasons for refusal to approve changes to the documentation");
                    if (rowRefuse2 > 0)
                    {
                        sheet.Cells[rowRefuse2 + 1, 1].Value = fcr.CommentsRefusalDocs;

                        // (8.1) Второй Link => от rowRefuse2+1 до "CONCURRENCE SHEET"
                        int rowLink2Header = rowRefuse2 + 1; 
                        // Или если есть метка, например "Link to documents justifying the decision (2)"
                        // int rowLink2Header = FindRowByLabel(sheet, totalRows, "Link to documents justifying the decision (2)");

                        if (rowLink2Header > 0)
                        {
                            int rowConcurrence = FindRowByLabel(sheet, totalRows, "CONCURRENCE SHEET");
                            if (rowConcurrence == 0) rowConcurrence = totalRows + 1;

                            int rowLink2DataStart = rowLink2Header + 3;
                            int reservedLink2 = rowConcurrence - rowLink2DataStart;
                            if (reservedLink2 < 0) reservedLink2 = 0;

                            // Берём те же 5 LinkToDocs (или отдельный список, если нужно)
                            int neededLink2 = fcr.LinkToDocs.Count;
                            if (neededLink2 > reservedLink2)
                            {
                                int toAdd = neededLink2 - reservedLink2;
                                sheet.InsertRow(rowConcurrence, toAdd);

                                int refRow = rowConcurrence - 1;
                                for (int i = 0; i < toAdd; i++)
                                {
                                    int newRow = rowConcurrence + i;
                                    CopyRowFormatting(sheet, refRow, newRow, totalCols);
                                    UnmergeRow(sheet, newRow);
                                    ApplyMergeFromReferenceRow(sheet, refRow, newRow);
                                }
                                totalRows = sheet.Dimension.Rows;
                                rowConcurrence += toAdd;
                            }

                            for (int i = 0; i < neededLink2; i++)
                            {
                                int row = rowLink2DataStart + i;
                                var doc2 = fcr.LinkToDocs[i];
                                sheet.Cells[row, 1].Value = doc2.FileNameExt; 
                                sheet.Cells[row, 4].Value = doc2.CodeOrTitle;
                            }
                        }
                    }
                }

                // (9) Final approval method => col C, Justif => col H
                {
                    int rowFinal = FindRowByLabel(sheet, totalRows, "Final approval method");
                    if (rowFinal > 0)
                    {
                        sheet.Cells[rowFinal, 3].Value = fcr.FinalApprovalMethod;
                        sheet.Cells[rowFinal, 8].Value = fcr.FinalApprovalJustif;
                    }
                }

                // (10) Раздел с должностями (6 шт)
                {
                    int rowPositions = FindRowByLabel(sheet, totalRows, "Project Participant position");
                    if (rowPositions > 0 && fcr.Signatures.Count >= 6)
                    {
                        for (int i = 0; i < 6; i++)
                        {
                            int row = (rowPositions + 1) + i;
                            var sig = fcr.Signatures[i];
                            sheet.Cells[row, 1].Value  = sig.Position;
                            sheet.Cells[row, 7].Value  = sig.Name;
                            sheet.Cells[row, 11].Value = sig.DateVal;
                        }
                    }
                }

                // (11) FCR final status => col E
                {
                    int rowStatus = FindRowByLabel(sheet, totalRows, "FCR final status");
                    if (rowStatus > 0)
                    {
                        sheet.Cells[rowStatus, 5].Value = "Approved";
                    }
                }

                // Сохраняем
                package.Save();
                Console.WriteLine("Файл сохранён успешно.");
            }
        }

        #endregion

        #region (C) Методы для поиска меток, копирования стилей и Merge

        private static int FindRowByLabelPartial(ExcelWorksheet sheet, int totalRows, string partial)
        {
            // Частичное совпадение
            for (int r = 1; r <= totalRows; r++)
            {
                string val = (sheet.Cells[r, 1].Text ?? "").Trim();
                if (!string.IsNullOrEmpty(val) && val.Contains(partial))
                {
                    return r;
                }
            }
            return 0;
        }

        private static int FindRowByLabel(ExcelWorksheet sheet, int totalRows, string label)
        {
            // Полное совпадение label -> Contains(label)
            // (Если нужно точное совпадение, можно делать ==, здесь оставляем Contains.)
            for (int r = 1; r <= totalRows; r++)
            {
                string val = (sheet.Cells[r, 1].Text ?? "").Trim();
                if (!string.IsNullOrEmpty(val) && val.Contains(label))
                {
                    return r;
                }
            }
            return 0;
        }

        private static void CopyRowFormatting(ExcelWorksheet sheet, int sourceRow, int targetRow, int cols)
        {
            for (int c = 1; c <= cols; c++)
            {
                var srcCell = sheet.Cells[sourceRow, c];
                var dstCell = sheet.Cells[targetRow, c];
                dstCell.StyleID = srcCell.StyleID;
            }
        }

        private static void UnmergeRow(ExcelWorksheet sheet, int row)
        {
            var mergesToRemove = new List<string>();
            foreach (var rng in sheet.MergedCells)
            {
                var addr = new ExcelAddress(rng);
                if (addr.Start.Row == row && addr.End.Row == row)
                    mergesToRemove.Add(rng);
            }
            foreach (string rem in mergesToRemove)
            {
                sheet.Cells[rem].Merge = false;
            }
        }

        private static void ApplyMergeFromReferenceRow(ExcelWorksheet sheet, int refRow, int newRow)
        {
            var mergesInRef = new List<OfficeOpenXml.ExcelAddressBase>();
            foreach (var rng in sheet.MergedCells)
            {
                var addr = new ExcelAddress(rng);
                if (addr.Start.Row == refRow && addr.End.Row == refRow)
                {
                    mergesInRef.Add(addr);
                }
            }
            foreach (var addr in mergesInRef)
            {
                int c1 = addr.Start.Column;
                int c2 = addr.End.Column;
                sheet.Cells[newRow, c1, newRow, c2].Merge = true;
            }
        }

        #endregion
    }

    #region (D) Классы-модели

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

        public bool ChangeInProjectPosEquip;
        public string CodeReasonChange;
        public string OtherReason;

        public string DescriptionEngChange;

        public List<FcrKksEntry> KKSList = new List<FcrKksEntry>();
        public List<FcrDocumentEntry> Documents = new List<FcrDocumentEntry>();
        public List<FcrFileNameEntry> SupportingDocs = new List<FcrFileNameEntry>();

        public bool MaterialIsEquivalent;
        public bool ReplaceTypeOfChange;
        public string CommentsRejectMaterial;

        public List<FcrFileNameEntry> LinkToDocs = new List<FcrFileNameEntry>();

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

        // Поле для Signatures
        public List<FcrSignature> Signatures = new List<FcrSignature>();
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
