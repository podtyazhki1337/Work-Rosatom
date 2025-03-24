using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ExcelParser35
{
    class Program
    {
        static void Main()
        {
            Console.Write("Введите путь к файлу Excel: ");
            string file = Console.ReadLine();

            if (IsNullOrWhiteSpace(file) || !File.Exists(file))
            {
                Console.WriteLine("Файл не найден или путь пустой. Завершение работы.");
                return;
            }

            Console.WriteLine("Парсим файл: " + file);

            // Парсим данные со второго листа (Worksheets[1])
            ChangeRequest ecr = ParseExcel(file);

            // Выводим всё в консоль
            PrintECR(ecr);

            // Верификация данных
            Console.WriteLine("\n=== Верификация данных ECR ===");
            List<string> validationErrors;
            if (ValidateECR(ecr, out validationErrors))
            {
                Console.WriteLine("Все данные корректны. Сохранение в базу данных...");
                SaveECRToDatabase(ecr);
            }
            else
            {
                Console.WriteLine("Обнаружены ошибки в данных:");
                foreach (var error in validationErrors)
                {
                    Console.WriteLine(" - " + error);
                }
                Console.WriteLine("Сохранение в базу данных отменено из-за ошибок.");
            }

            Console.WriteLine("\nНажмите любую клавишу для завершения...");
            Console.ReadKey();
        }

        /// <summary>
        /// Аналог string.IsNullOrWhiteSpace() для .NET 3.5
        /// </summary>
        private static bool IsNullOrWhiteSpace(string s)
        {
            return s == null || s.Trim().Length == 0;
        }

        /// <summary>
        /// Парсим данные со второго листа:
        ///  - Раздел 1.1 (статические)
        ///  - 2.1, 2.2, 2.3 (динамические)
        ///  - 2.4, 2.5 (статические)
        ///  - 3, 4 (динамические, построчно, начиная с row+2)
        ///    * Раздел 4 заполняется до \"Change Manager of the Owner\" включительно
        /// </summary>
        private static ChangeRequest ParseExcel(string filePath)
        {
            var ecr = new ChangeRequest();

            FileInfo fi = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                // Берём второй лист
                ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                if (sheet.Dimension == null)
                {
                    Console.WriteLine("Ошибка: Лист пустой (Worksheets[1]).");
                    return ecr;
                }

                int rows = sheet.Dimension.Rows;
                int cols = sheet.Dimension.Columns;
                Console.WriteLine("Лист [1] содержит {0} строк и {1} столбцов.", rows, cols);

                // --- (A) Парсим статические поля (1.1, 2.4, 2.5) ---
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        string cellTextRaw = GetMergedCellValue(sheet, r, c);
                        if (IsNullOrWhiteSpace(cellTextRaw)) 
                            continue;

                        string cellText = cellTextRaw.Replace("\r","").Replace("\n","").Trim();

                        // --- 1.1 ---
                        if (cellText.Contains("Change Request No:"))
                            ecr.ChangeRequestNo = GetValueRight(sheet, r, c, cols);
                        else if (cellText.Contains("Contractor Change Coordinator:"))
                            ecr.ContractorChangeCoordinator = GetValueRight(sheet, r, c, cols);
                        else if (cellText.Contains("Initiator's internal CR No:"))
                            ecr.InitiatorInternalCRNo = GetValueRight(sheet, r, c, cols);
                        else if (cellText.Contains("Change Initiator:"))
                            ecr.ChangeInitiator = GetValueRight(sheet, r, c, cols);
                        else if (cellText.Contains("Initiator's organization:"))
                            ecr.InitiatorOrganization = GetValueRight(sheet, r, c, cols);
                        else if (cellText.Contains("Type of documentation, where Engineering Change will be reflected"))
                            ecr.DocumentationType = GetValueRight(sheet, r, c, cols);
                        else if (cellText.Contains("Reason of Engineering Change:"))
                            ecr.Reason = GetValueRight(sheet, r, c, cols);
                        else if (cellText.Contains("description of technical solution"))
                            ecr.TechnicalSolutionDescription = GetValueRight(sheet, r, c, cols);

                        // --- 2.4 ---
                        else if (cellText.Contains("Final NSC category of the Engineering Change:"))
                            ecr.Section24.FinalNSCCategory = GetValueRight(sheet, r, c, cols);
                        else if (cellText.Contains("Presence of direct or indirect impact on equipment of 1,2 and 3 safety classes"))
                        {
                            string val = GetValueRight(sheet, r, c, cols);
                            ecr.Section24.ImpactOnSafetyClasses = (val.ToLower() == "yes");
                        }
                        else if (cellText.Contains("Presence the impact on the results DSA") 
                                 || cellText.Contains("Наличие влияния на результаты ДАБ"))
                        {
                            string val = GetValueRight(sheet, r, c, cols);
                            ecr.Section24.ImpactOnDSA = (val.ToLower() == "yes");
                        }
                        else if (cellText.Contains("Method* of CR:")
                              || cellText.Contains("Вид ЗИ:"))
                            ecr.Section24.CRMethod = GetValueRight(sheet, r, c, cols);
                        else if (cellText.Contains("Comments for engineering evaluation:")
                              || cellText.Contains("Комментарии к инженерной оценке:"))
                            ecr.Section24.EngineeringComments = GetValueRight(sheet, r, c, cols);

                        // --- 2.5 ---
                        else if (cellText.Contains("Contract (its  presence indicates")
                              || cellText.Contains("Контракт (его наличие указывает"))
                        {
                            string val = GetValueRight(sheet, r, c, cols);
                            ecr.Section25.ContractAffected = (val.ToLower() == "yes");
                        }
                        else if (cellText.Contains("Cost impact:") || cellText.Contains("Влияние на стоимость:"))
                        {
                            string val = GetValueRight(sheet, r, c, cols);
                            ecr.Section25.CostImpact = (val.ToLower() == "yes");
                        }
                        else if (cellText.Contains("Schedule:") || cellText.Contains("График:"))
                        {
                            string val = GetValueRight(sheet, r, c, cols);
                            ecr.Section25.ScheduleImpact = (val.ToLower() == "yes");
                        }
                        else if (cellText.Contains("Comments for non-technical assessment:")
                              || cellText.Contains("Комментарии к нетехнической оценке:"))
                            ecr.Section25.NonTechnicalComments = GetValueRight(sheet, r, c, cols);
                    }
                }

                // --- (B) Динамические разделы 2.1, 2.2, 2.3 ---
                int row21 = FindRowByPrefix(sheet, rows, "2.1");
                if (row21 > 0)
                {
                    ecr.InitiatorTDDImpacts = ParseTDDImpacts2_1(sheet, row21 + 2, "2.2", rows);
                }

                int row22 = FindRowByPrefix(sheet, rows, "2.2");
                if (row22 > 0)
                {
                    ecr.OtherTDDImpacts = ParseTDDImpacts2_2(sheet, row22 + 2, "2.3", rows);
                }

                int row23 = FindRowByPrefix(sheet, rows, "2.3");
                if (row23 > 0)
                {
                    ecr.AffectedSSCs = ParseKKSImpacts(sheet, row23 + 2, "2.4", rows);
                }

                // --- (C) Раздел 3 ---
                int row3 = FindRowByPrefix(sheet, rows, "3.");
                if (row3 > 0)
                {
                    ecr.Confirmations = ParseConfirmations(sheet, row3 + 2, "4.", rows);
                }

                // --- (D) Раздел 4 (содержит несколько строк, заканчивая \"Change Manager of the Owner\" включительно)
                int row4 = FindRowByPrefix(sheet, rows, "4.");
                if (row4 > 0)
                {
                    ecr.Approvals = ParseApprovalsUntilOwner(sheet, row4 + 2, rows);
                }
            }

            return ecr;
        }

        #region ==== Методы для поиска строк и чтения ячеек ====

        /// <summary>
        /// Ищет строку, где в первом столбце текст начинается на prefix (\"2.1\", \"3.\", \"4.\" и т.п.)
        /// </summary>
        private static int FindRowByPrefix(ExcelWorksheet sheet, int totalRows, string prefix)
        {
            for (int r = 1; r <= totalRows; r++)
            {
                string cellVal = GetMergedCellValue(sheet, r, 1);
                if (!IsNullOrWhiteSpace(cellVal))
                {
                    string normalized = cellVal.Replace("\r","").Replace("\n","").Trim();
                    if (normalized.StartsWith(prefix))
                        return r;
                }
            }
            return 0;
        }

        /// <summary>
        /// Считываем значение ячейки (row,col) с учётом объединённых ячеек
        /// </summary>
        private static string GetMergedCellValue(ExcelWorksheet sheet, int row, int col)
        {
            var cell = sheet.Cells[row, col];
            if (cell.Merge)
            {
                foreach (var mergedCell in sheet.MergedCells)
                {
                    var range = sheet.Cells[mergedCell];
                    if (range.Start.Row <= row && range.End.Row >= row &&
                        range.Start.Column <= col && range.End.Column >= col)
                    {
                        string mergedText = sheet.Cells[range.Start.Row, range.Start.Column].Text;
                        return mergedText == null ? "" : mergedText.Trim();
                    }
                }
            }

            string text = cell.Text;
            return text == null ? "" : text.Trim();
        }

        /// <summary>
        /// Возвращаем первое непустое значение справа (в той же строке)
        /// </summary>
        private static string GetValueRight(ExcelWorksheet sheet, int row, int startCol, int totalCols)
        {
            for (int c = startCol + 1; c <= totalCols; c++)
            {
                string val = GetMergedCellValue(sheet, row, c);
                if (!IsNullOrWhiteSpace(val))
                    return val.Trim();
            }
            return "";
        }

        #endregion

        #region ========== Разделы 2.1, 2.2, 2.3 (динамические) ==========

        private static List<TDDImpact> ParseTDDImpacts2_1(
            ExcelWorksheet sheet, int startRow, string nextSectionPrefix, int totalRows)
        {
            var list = new List<TDDImpact>();
            for (int r = startRow; r <= totalRows; r++)
            {
                // Остановка, если встретили \"2.2\"
                string firstCell = GetMergedCellValue(sheet, r, 1);
                if (!IsNullOrWhiteSpace(firstCell))
                {
                    string norm = firstCell.Replace("\r","").Replace("\n","").Trim();
                    if (norm.StartsWith(nextSectionPrefix))
                        break;
                }

                if (IsNullOrWhiteSpace(firstCell) || firstCell.ToLower().Contains("tdd code"))
                    continue;

                // A=1(TDD code), D=4(Revision), E=5(NewRev?), F=6(Name), H=8(State), I=9(Desc),
                // K=11..O=15(safety)
                string revision      = GetMergedCellValue(sheet, r, 4);
                string isNewRev      = GetMergedCellValue(sheet, r, 5);
                string tddName       = GetMergedCellValue(sheet, r, 6);
                string tddState      = GetMergedCellValue(sheet, r, 8);
                string changeDesc    = GetMergedCellValue(sheet, r, 9);
                string nuclearVal    = GetMergedCellValue(sheet, r, 11);
                string fireVal       = GetMergedCellValue(sheet, r, 12);
                string industrialVal = GetMergedCellValue(sheet, r, 13);
                string environVal    = GetMergedCellValue(sheet, r, 14);
                string structVal     = GetMergedCellValue(sheet, r, 15);

                var impact = new TDDImpact
                {
                    TDDCode = firstCell,
                    Revision = revision,
                    NewRevisionRequired = (isNewRev.ToLower() == "yes"),
                    TDDName = tddName,
                    TDDState = tddState,
                    ChangeDescription = changeDesc,
                    SafetyImpacts = new SafetyImpacts
                    {
                        NuclearSafety    = (nuclearVal.ToLower()   == "yes"),
                        FireSafety       = (fireVal.ToLower()      == "yes"),
                        IndustrialSafety = (industrialVal.ToLower()== "yes"),
                        Environmental    = (environVal.ToLower()   == "yes"),
                        Structural       = (structVal.ToLower()    == "yes")
                    }
                };
                list.Add(impact);
            }
            return list;
        }

        private static List<TDDImpact> ParseTDDImpacts2_2(
            ExcelWorksheet sheet, int startRow, string nextSectionPrefix, int totalRows)
        {
            var list = new List<TDDImpact>();
            for (int r = startRow; r <= totalRows; r++)
            {
                // Остановка, если \"2.3\"
                string firstCell = GetMergedCellValue(sheet, r, 1);
                if (!IsNullOrWhiteSpace(firstCell))
                {
                    string norm = firstCell.Replace("\r","").Replace("\n","").Trim();
                    if (norm.StartsWith(nextSectionPrefix))
                        break;
                }

                if (IsNullOrWhiteSpace(firstCell) || firstCell.ToLower().Contains("evaluation organization"))
                    continue;

                // A=1(EvalOrg), B=2(TDD code), D=4(Rev), E=5(NewRev?), F=6(Name),
                // H=8(State), I=9(Desc), K=11..O=15(safety)
                string tddCode       = GetMergedCellValue(sheet, r, 2);
                string revision      = GetMergedCellValue(sheet, r, 4);
                string isNewRev      = GetMergedCellValue(sheet, r, 5);
                string tddName       = GetMergedCellValue(sheet, r, 6);
                string tddState      = GetMergedCellValue(sheet, r, 8);
                string changeDesc    = GetMergedCellValue(sheet, r, 9);
                string nuclearVal    = GetMergedCellValue(sheet, r, 11);
                string fireVal       = GetMergedCellValue(sheet, r, 12);
                string industrialVal = GetMergedCellValue(sheet, r, 13);
                string environVal    = GetMergedCellValue(sheet, r, 14);
                string structVal     = GetMergedCellValue(sheet, r, 15);

                var impact = new TDDImpact
                {
                    EvaluationOrganization = firstCell,
                    TDDCode = tddCode,
                    Revision = revision,
                    NewRevisionRequired = (isNewRev.ToLower() == "yes"),
                    TDDName = tddName,
                    TDDState = tddState,
                    ChangeDescription = changeDesc,
                    SafetyImpacts = new SafetyImpacts
                    {
                        NuclearSafety    = (nuclearVal.ToLower()   == "yes"),
                        FireSafety       = (fireVal.ToLower()      == "yes"),
                        IndustrialSafety = (industrialVal.ToLower()== "yes"),
                        Environmental    = (environVal.ToLower()   == "yes"),
                        Structural       = (structVal.ToLower()    == "yes")
                    }
                };
                list.Add(impact);
            }
            return list;
        }

        private static List<SSCImpact> ParseKKSImpacts(
            ExcelWorksheet sheet, int startRow, string nextSectionPrefix, int totalRows)
        {
            var sscList = new List<SSCImpact>();
            for (int r = startRow; r <= totalRows; r++)
            {
                string firstCell = GetMergedCellValue(sheet, r, 1);
                if (!IsNullOrWhiteSpace(firstCell))
                {
                    string norm = firstCell.Replace("\r","").Replace("\n","").Trim();
                    if (norm.StartsWith(nextSectionPrefix)) 
                        break;
                }

                if (!IsNullOrWhiteSpace(firstCell) && firstCell.ToLower().Contains("kks code"))
                    continue;
                if (IsNullOrWhiteSpace(firstCell))
                    continue;

                // KKS code = A=1, SSCName=B=2, Desc=F=6
                string sscName    = GetMergedCellValue(sheet, r, 2);
                string changeDesc = GetMergedCellValue(sheet, r, 6);

                var ssc = new SSCImpact
                {
                    KKSCode = firstCell,
                    SSCName = sscName,
                    ChangeDescription = changeDesc
                };
                sscList.Add(ssc);
            }
            return sscList;
        }

        #endregion

        #region ========== Раздел 3, 4 (динамические) ==========

        /// <summary>
        /// Раздел 3 (Confirmation), начиная с rowStart, 
        /// пока не встретим \"4.\" или конец листа.
        /// A=1(Organization), B=2(Position), E=5(Responsible), K=11(Date)
        /// </summary>
        private static List<Confirmation> ParseConfirmations(
            ExcelWorksheet sheet, int rowStart, string nextSectionPrefix, int totalRows)
        {
            var confirmations = new List<Confirmation>();
            for (int r = rowStart; r <= totalRows; r++)
            {
                string prefixCell = GetMergedCellValue(sheet, r, 1);
                if (!IsNullOrWhiteSpace(prefixCell))
                {
                    string norm = prefixCell.Replace("\r","").Replace("\n","").Trim();
                    // Если встретили \"4.\", выходим
                    if (norm.StartsWith(nextSectionPrefix))
                        break;
                }

                // Читаем колонки (A,B,E,K)
                string orgVal  = GetMergedCellValue(sheet, r, 1);  // A
                string posVal  = GetMergedCellValue(sheet, r, 2);  // B
                string respVal = GetMergedCellValue(sheet, r, 5);  // E
                string dateVal = GetMergedCellValue(sheet, r, 11); // K

                // Проверяем, не пустая ли строка
                if (IsNullOrWhiteSpace(orgVal) && 
                    IsNullOrWhiteSpace(posVal) &&
                    IsNullOrWhiteSpace(respVal) &&
                    IsNullOrWhiteSpace(dateVal))
                {
                    continue; // пустая
                }

                var c = new Confirmation
                {
                    Organization      = orgVal,
                    Position          = posVal,
                    ResponsiblePerson = respVal,
                    Date             = dateVal
                };
                confirmations.Add(c);
            }
            return confirmations;
        }

        /// <summary>
        /// Раздел 4 (Approval), начиная с rowStart, 
        /// до \"Change Manager of the Owner\" включительно или до конца файла.
        /// A=1 (Position), E=5 (Responsible), K=11 (Date)
        /// </summary>
        private static List<Approval> ParseApprovalsUntilOwner(
            ExcelWorksheet sheet, int rowStart, int totalRows)
        {
            var approvals = new List<Approval>();

            for (int r = rowStart; r <= totalRows; r++)
            {
                string firstCell = GetMergedCellValue(sheet, r, 1);
                // Если пусто и все колонки пустые — пропустим
                string posVal  = firstCell; 
                string respVal = GetMergedCellValue(sheet, r, 5);
                string dateVal = GetMergedCellValue(sheet, r, 11);

                if (IsNullOrWhiteSpace(posVal) &&
                    IsNullOrWhiteSpace(respVal) &&
                    IsNullOrWhiteSpace(dateVal))
                {
                    // Пустая строка
                    continue;
                }

                // Создаём запись
                var a = new Approval
                {
                    Position          = posVal,
                    ResponsiblePerson = respVal,
                    Date             = dateVal
                };
                approvals.Add(a);

                // Проверяем, есть ли в firstCell \"Change Manager of the Owner\"?
                if (!IsNullOrWhiteSpace(firstCell))
                {
                    var norm = firstCell.ToLower();
                    if (norm.Contains("change manager of the owner"))
                    {
                        // Нашли искомую строку — останавливаемся
                        break;
                    }
                }
            }

            return approvals;
        }

        #endregion

        #region ========== Вывод в консоль ==========

        private static void PrintECR(ChangeRequest ecr)
        {
            Console.WriteLine("\n=== Change Request (CR) ===");
            // 1.1
            Console.WriteLine("ChangeRequestNo:               " + ecr.ChangeRequestNo);
            Console.WriteLine("ContractorChangeCoordinator:   " + ecr.ContractorChangeCoordinator);
            Console.WriteLine("InitiatorInternalCRNo:         " + ecr.InitiatorInternalCRNo);
            Console.WriteLine("ChangeInitiator:               " + ecr.ChangeInitiator);
            Console.WriteLine("InitiatorOrganization:         " + ecr.InitiatorOrganization);
            Console.WriteLine("DocumentationType:             " + ecr.DocumentationType);
            Console.WriteLine("Reason:                        " + ecr.Reason);
            Console.WriteLine("TechnicalSolutionDescription:  " + ecr.TechnicalSolutionDescription);

            // 2.1
            Console.WriteLine("\n--- 2.1 Initiator TDD Impacts ---");
            if (ecr.InitiatorTDDImpacts.Count > 0)
            {
                foreach (var imp in ecr.InitiatorTDDImpacts)
                {
                    Console.WriteLine("  [2.1] Code={0}, Rev={1}, NewRev={2}, Name={3}, State={4}, Desc={5}",
                        imp.TDDCode, imp.Revision, imp.NewRevisionRequired,
                        imp.TDDName, imp.TDDState, imp.ChangeDescription);
                    if (imp.SafetyImpacts != null)
                    {
                        Console.WriteLine("       Nuclear={0}, Fire={1}, Industrial={2}, Env={3}, Struct={4}",
                            imp.SafetyImpacts.NuclearSafety,
                            imp.SafetyImpacts.FireSafety,
                            imp.SafetyImpacts.IndustrialSafety,
                            imp.SafetyImpacts.Environmental,
                            imp.SafetyImpacts.Structural);
                    }
                }
            }
            else
            {
                Console.WriteLine("  (нет данных)");
            }

            // 2.2
            Console.WriteLine("\n--- 2.2 Other TDD Impacts ---");
            if (ecr.OtherTDDImpacts.Count > 0)
            {
                foreach (var imp in ecr.OtherTDDImpacts)
                {
                    Console.WriteLine("  [2.2] Org={0}, Code={1}, Rev={2}, NewRev={3}, Name={4}, State={5}, Desc={6}",
                        imp.EvaluationOrganization,
                        imp.TDDCode, imp.Revision, imp.NewRevisionRequired,
                        imp.TDDName, imp.TDDState, imp.ChangeDescription);
                    if (imp.SafetyImpacts != null)
                    {
                        Console.WriteLine("       Nuclear={0}, Fire={1}, Industrial={2}, Env={3}, Struct={4}",
                            imp.SafetyImpacts.NuclearSafety,
                            imp.SafetyImpacts.FireSafety,
                            imp.SafetyImpacts.IndustrialSafety,
                            imp.SafetyImpacts.Environmental,
                            imp.SafetyImpacts.Structural);
                    }
                }
            }
            else
            {
                Console.WriteLine("  (нет данных)");
            }

            // 2.3
            Console.WriteLine("\n--- 2.3 Affected SSCs ---");
            if (ecr.AffectedSSCs.Count > 0)
            {
                foreach (var ssc in ecr.AffectedSSCs)
                {
                    Console.WriteLine("  KKSCode={0}; SSCName={1}; Desc={2}",
                        ssc.KKSCode, ssc.SSCName, ssc.ChangeDescription);
                }
            }
            else
            {
                Console.WriteLine("  (нет данных)");
            }

            // 2.4
            Console.WriteLine("\n--- 2.4 Final Evaluation ---");
            Console.WriteLine("FinalNSCCategory:               " + ecr.Section24.FinalNSCCategory);
            Console.WriteLine("ImpactOnSafetyClasses (1,2,3):  " + ecr.Section24.ImpactOnSafetyClasses);
            Console.WriteLine("ImpactOnDSAorHazard:            " + ecr.Section24.ImpactOnDSA);
            Console.WriteLine("CRMethod:                       " + ecr.Section24.CRMethod);
            Console.WriteLine("EngineeringComments:            " + ecr.Section24.EngineeringComments);

            // 2.5
            Console.WriteLine("\n--- 2.5 Non-technical Assessment ---");
            Console.WriteLine("ContractAffected: " + ecr.Section25.ContractAffected);
            Console.WriteLine("CostImpact:       " + ecr.Section25.CostImpact);
            Console.WriteLine("ScheduleImpact:   " + ecr.Section25.ScheduleImpact);
            Console.WriteLine("NonTechComments:  " + ecr.Section25.NonTechnicalComments);

            // 3
            Console.WriteLine("\n--- 3 Confirmation (множество строк) ---");
            if (ecr.Confirmations.Count > 0)
            {
                int idx=1;
                foreach (var conf in ecr.Confirmations)
                {
                    Console.WriteLine("  Confirmation #{0}: Org(A)={1}, Pos(B)={2}, Resp(E)={3}, Date(K)={4}",
                        idx++, conf.Organization, conf.Position, conf.ResponsiblePerson, conf.Date);
                }
            }
            else
            {
                Console.WriteLine("  (нет данных)");
            }

            // 4
            Console.WriteLine("\n--- 4 Approval (множество строк) ---");
            if (ecr.Approvals.Count > 0)
            {
                int idx=1;
                foreach (var app in ecr.Approvals)
                {
                    Console.WriteLine("  Approval #{0}: Pos(A)={1}, Resp(E)={2}, Date(K)={3}",
                        idx++, app.Position, app.ResponsiblePerson, app.Date);
                }
            }
            else
            {
                Console.WriteLine("  (нет данных)");
            }

            Console.WriteLine("\n=== Конец вывода ECR ===");
        }

        #endregion

        #region ========== Верификация и сохранение в БД ==========

        /// <summary>
        /// Проверка полноты и корректности данных ECR согласно заданным правилам
        /// </summary>
        private static bool ValidateECR(ChangeRequest ecr, out List<string> errors)
        {
            errors = new List<string>();
            bool isValid = true;

            // 1. Обязательные текстовые поля
            if (string.IsNullOrEmpty(ecr.ChangeRequestNo))
            {
                errors.Add("Поле 'Change Request No' обязательно для заполнения");
                isValid = false;
            }
            if (string.IsNullOrEmpty(ecr.ContractorChangeCoordinator))
            {
                errors.Add("Поле 'Contractor Change Coordinator' обязательно для заполнения");
                isValid = false;
            }
            if (string.IsNullOrEmpty(ecr.InitiatorOrganization))
            {
                errors.Add("Поле 'Initiator's organization' обязательно для заполнения");
                isValid = false;
            }
            if (string.IsNullOrEmpty(ecr.ChangeInitiator))
            {
                errors.Add("Поле 'Change Initiator' обязательно для заполнения");
                isValid = false;
            }
            if (string.IsNullOrEmpty(ecr.Reason))
            {
                errors.Add("Поле 'Reason of Engineering Change' обязательно для заполнения");
                isValid = false;
            }
            if (string.IsNullOrEmpty(ecr.TechnicalSolutionDescription))
            {
                errors.Add("Поле 'Initiator's description of technical solution' обязательно для заполнения");
                isValid = false;
            }
            // Проверка 2.2 на наличие хотя бы одной записи с обязательным полем EvaluationOrganization
            if (ecr.OtherTDDImpacts.Count > 0)
            {
                foreach (var item in ecr.OtherTDDImpacts)
                {
                    if (string.IsNullOrEmpty(item.EvaluationOrganization))
                    {
                        errors.Add("Поле 'Evaluation organization' в разделе 2.2 обязательно для заполнения");
                        isValid = false;
                        break;
                    }
                }
            }

            // 2. Списочные значения
            string[] validDocTypes = { "MLA", "CLA", "MDD", "DDD", "DD" };
            if (!string.IsNullOrEmpty(ecr.DocumentationType))
            {
                string[] docTypeParts = ecr.DocumentationType.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (docTypeParts.Length == 0)
                {
                    errors.Add("Поле 'Type of documentation' не может содержать только символы '/'");
                    isValid = false;
                }
                else
                {
                    foreach (string part in docTypeParts)
                    {
                        string trimmedPart = part.Trim();
                        if (!Array.Exists(validDocTypes, x => x == trimmedPart))
                        {
                            errors.Add($"Поле 'Type of documentation' содержит недопустимое значение '{trimmedPart}'. Допустимые значения: {string.Join(", ", validDocTypes)}");
                            isValid = false;
                        }
                    }
                }
            }

            string[] validNSCCategories = { "b/bb", "ba", "a/ab", "aa", "b", "a", "bb"};
            if (!string.IsNullOrEmpty(ecr.Section24.FinalNSCCategory) && !Array.Exists(validNSCCategories, x => x == ecr.Section24.FinalNSCCategory))
            {
                errors.Add($"Поле 'Final NSC category' должно быть одним из: {string.Join(", ", validNSCCategories)}. Текущее значение: '{ecr.Section24.FinalNSCCategory}'");
                isValid = false;
            }

            string[] validCRMethods = { "Simple", "Normal", "Complex" };
            if (!string.IsNullOrEmpty(ecr.Section24.CRMethod) && !Array.Exists(validCRMethods, x => x == ecr.Section24.CRMethod))
            {
                errors.Add($"Поле 'Method of CR' должно быть одним из: {string.Join(", ", validCRMethods)}. Текущее значение: '{ecr.Section24.CRMethod}'");
                isValid = false;
            }

            // 3. Проверка зависимости полей в разделах 2.1, 2.2, 2.3
            // Раздел 2.1
            foreach (var item in ecr.InitiatorTDDImpacts)
            {
                if (!string.IsNullOrEmpty(item.TDDCode))
                {
                    if (string.IsNullOrEmpty(item.Revision))
                    {
                        errors.Add($"Для TDDCode '{item.TDDCode}' в разделе 2.1 отсутствует 'Revision'");
                        isValid = false;
                    }
                    if (string.IsNullOrEmpty(item.TDDName))
                    {
                        errors.Add($"Для TDDCode '{item.TDDCode}' в разделе 2.1 отсутствует 'TDD name'");
                        isValid = false;
                    }
                    if (string.IsNullOrEmpty(item.TDDState))
                    {
                        errors.Add($"Для TDDCode '{item.TDDCode}' в разделе 2.1 отсутствует 'TDD state'");
                        isValid = false;
                    }
                    if (string.IsNullOrEmpty(item.ChangeDescription))
                    {
                        errors.Add($"Для TDDCode '{item.TDDCode}' в разделе 2.1 отсутствует 'Engineering Change description'");
                        isValid = false;
                    }
                }
            }

            // Раздел 2.2
            foreach (var item in ecr.OtherTDDImpacts)
            {
                if (!string.IsNullOrEmpty(item.TDDCode))
                {
                    if (string.IsNullOrEmpty(item.Revision))
                    {
                        errors.Add($"Для TDDCode '{item.TDDCode}' в разделе 2.2 отсутствует 'Revision'");
                        isValid = false;
                    }
                    if (string.IsNullOrEmpty(item.TDDName))
                    {
                        errors.Add($"Для TDDCode '{item.TDDCode}' в разделе 2.2 отсутствует 'TDD name'");
                        isValid = false;
                    }
                    if (string.IsNullOrEmpty(item.TDDState))
                    {
                        errors.Add($"Для TDDCode '{item.TDDCode}' в разделе 2.2 отсутствует 'TDD state'");
                        isValid = false;
                    }
                    if (string.IsNullOrEmpty(item.ChangeDescription))
                    {
                        errors.Add($"Для TDDCode '{item.TDDCode}' в разделе 2.2 отсутствует 'Engineering Change description'");
                        isValid = false;
                    }
                }
            }

            // Раздел 2.3
            foreach (var item in ecr.AffectedSSCs)
            {
                if (!string.IsNullOrEmpty(item.KKSCode))
                {
                    if (string.IsNullOrEmpty(item.SSCName))
                    {
                        errors.Add($"Для KKSCode '{item.KKSCode}' в разделе 2.3 отсутствует 'Name of SSC'");
                        isValid = false;
                    }
                    if (string.IsNullOrEmpty(item.ChangeDescription))
                    {
                        errors.Add($"Для KKSCode '{item.KKSCode}' в разделе 2.3 отсутствует 'Engineering Change description'");
                        isValid = false;
                    }
                }
            }

            // 4. Проверка на английский шрифт для Reason и TechnicalSolutionDescription
            if (!string.IsNullOrEmpty(ecr.Reason) && !IsEnglishText(ecr.Reason))
            {
                errors.Add("Поле 'Reason of Engineering Change' должно содержать только английский текст");
                isValid = false;
            }
            if (!string.IsNullOrEmpty(ecr.TechnicalSolutionDescription) && !IsEnglishText(ecr.TechnicalSolutionDescription))
            {
                errors.Add("Поле 'Initiator's description of technical solution' должно содержать только английский текст");
                isValid = false;
            }

            return isValid;
        }

        /// <summary>
        /// Проверка, что строка содержит только английские символы, пробелы и базовую пунктуацию
        /// </summary>
        private static bool IsEnglishText(string text)
        {
            foreach (char c in text)
            {
                if (c > 127) // ASCII выше 127 - не английские символы
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Заглушка для сохранения ECR в базу данных SQL
        /// </summary>
        private static void SaveECRToDatabase(ChangeRequest ecr)
        {
            try
            {
                Console.WriteLine("\n=== Сохранение ECR в базу данных ===");
                Console.WriteLine("Подключение к базе данных...");
                // Здесь должен быть код подключения к SQL (например, SqlConnection)

                // Пример заглушки для сохранения основных полей 1.1
                Console.WriteLine("INSERT INTO ECR_Main (ChangeRequestNo, ContractorChangeCoordinator, InitiatorInternalCRNo) " +
                                 "VALUES ('{0}', '{1}', '{2}')", 
                                 ecr.ChangeRequestNo, ecr.ContractorChangeCoordinator, ecr.InitiatorInternalCRNo);

                // Заглушка для 2.1 InitiatorTDDImpacts
                if (ecr.InitiatorTDDImpacts.Count > 0)
                {
                    Console.WriteLine("Сохранение {0} записей InitiatorTDDImpacts...", ecr.InitiatorTDDImpacts.Count);
                    foreach (var item in ecr.InitiatorTDDImpacts)
                    {
                        Console.WriteLine("INSERT INTO TDD_Impacts (TDDCode, Revision, ChangeDescription) " +
                                        "VALUES ('{0}', '{1}', '{2}')", 
                                        item.TDDCode, item.Revision, item.ChangeDescription);
                    }
                }

                // Заглушка для 2.2 OtherTDDImpacts
                if (ecr.OtherTDDImpacts.Count > 0)
                {
                    Console.WriteLine("Сохранение {0} записей OtherTDDImpacts...", ecr.OtherTDDImpacts.Count);
                    foreach (var item in ecr.OtherTDDImpacts)
                    {
                        Console.WriteLine("INSERT INTO Other_TDD_Impacts (EvaluationOrg, TDDCode, Revision) " +
                                        "VALUES ('{0}', '{1}', '{2}')", 
                                        item.EvaluationOrganization, item.TDDCode, item.Revision);
                    }
                }

                // Заглушка для 2.3 AffectedSSCs
                if (ecr.AffectedSSCs.Count > 0)
                {
                    Console.WriteLine("Сохранение {0} записей AffectedSSCs...", ecr.AffectedSSCs.Count);
                    foreach (var item in ecr.AffectedSSCs)
                    {
                        Console.WriteLine("INSERT INTO Affected_SSCs (KKSCode, SSCName) " +
                                        "VALUES ('{0}', '{1}')", 
                                        item.KKSCode, item.SSCName);
                    }
                }

                // Заглушка для 2.4 и 2.5
                Console.WriteLine("INSERT INTO ECR_Evaluation (FinalNSCCategory, ImpactOnSafetyClasses) " +
                                 "VALUES ('{0}', '{1}')", 
                                 ecr.Section24.FinalNSCCategory, ecr.Section24.ImpactOnSafetyClasses);
                Console.WriteLine("INSERT INTO ECR_NonTechnical (ContractAffected, CostImpact) " +
                                 "VALUES ('{0}', '{1}')", 
                                 ecr.Section25.ContractAffected, ecr.Section25.CostImpact);

                // Заглушка для 3 Confirmations
                if (ecr.Confirmations.Count > 0)
                {
                    Console.WriteLine("Сохранение {0} записей Confirmations...", ecr.Confirmations.Count);
                    foreach (var item in ecr.Confirmations)
                    {
                        Console.WriteLine("INSERT INTO Confirmations (Organization, Position, Date) " +
                                        "VALUES ('{0}', '{1}', '{2}')", 
                                        item.Organization, item.Position, item.Date);
                    }
                }

                // Заглушка для 4 Approvals
                if (ecr.Approvals.Count > 0)
                {
                    Console.WriteLine("Сохранение {0} записей Approvals...", ecr.Approvals.Count);
                    foreach (var item in ecr.Approvals)
                    {
                        Console.WriteLine("INSERT INTO Approvals (Position, ResponsiblePerson, Date) " +
                                        "VALUES ('{0}', '{1}', '{2}')", 
                                        item.Position, item.ResponsiblePerson, item.Date);
                    }
                }

                Console.WriteLine("Данные успешно сохранены в базу данных.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при сохранении в базу данных: " + ex.Message);
            }
        }

        #endregion
    }

    #region ========== Классы-модели ==========
    
    public class ChangeRequest
    {
        // 1.1
        public string ChangeRequestNo;
        public string ContractorChangeCoordinator;
        public string InitiatorInternalCRNo;
        public string ChangeInitiator;
        public string InitiatorOrganization;
        public string DocumentationType;
        public string Reason;
        public string TechnicalSolutionDescription;

        // 2.1
        public List<TDDImpact> InitiatorTDDImpacts = new List<TDDImpact>();
        // 2.2
        public List<TDDImpact> OtherTDDImpacts = new List<TDDImpact>();
        // 2.3
        public List<SSCImpact> AffectedSSCs = new List<SSCImpact>();

        // 2.4
        public Section24 Section24 = new Section24();
        // 2.5
        public Section25 Section25 = new Section25();

        // 3: список подтверждений
        public List<Confirmation> Confirmations = new List<Confirmation>();
        // 4: список согласований
        public List<Approval> Approvals = new List<Approval>();
    }

    public class Section24
    {
        public string FinalNSCCategory;
        public bool ImpactOnSafetyClasses;
        public bool ImpactOnDSA;
        public string CRMethod;
        public string EngineeringComments;
    }

    public class Section25
    {
        public bool ContractAffected;
        public bool CostImpact;
        public bool ScheduleImpact;
        public string NonTechnicalComments;
    }

    /// <summary>
    /// Для Раздела 3 (A=Organization, B=Position, E=Responsible, K=Date)
    /// </summary>
    public class Confirmation
    {
        public string Organization;
        public string Position;
        public string ResponsiblePerson;
        public string Date;
    }

    /// <summary>
    /// Для Раздела 4 (A=Position, E=Responsible, K=Date),
    /// до \"Change Manager of the Owner\" включительно
    /// </summary>
    public class Approval
    {
        public string Position;
        public string ResponsiblePerson;
        public string Date;
    }

    public class TDDImpact
    {
        public string EvaluationOrganization;
        public string TDDCode;
        public string Revision;
        public bool NewRevisionRequired;
        public string TDDName;
        public string TDDState;
        public string ChangeDescription;
        public SafetyImpacts SafetyImpacts;
    }

    public class SSCImpact
    {
        public string KKSCode;
        public string SSCName;
        public string ChangeDescription;
    }

    public class SafetyImpacts
    {
        public bool NuclearSafety;
        public bool FireSafety;
        public bool IndustrialSafety;
        public bool Environmental;
        public bool Structural;
    }

    #endregion
}