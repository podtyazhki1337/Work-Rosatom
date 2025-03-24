using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;           // Для .Any() в "The presence of an impact on docs"
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelWriterReverse
{
    class Program
    {
        static void Main()
        {
            Console.Write("Введите путь к существующему Excel файлу: ");
            string filePath = Console.ReadLine();

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Console.WriteLine("Файл не найден или путь пустой. Завершение работы.");
                return;
            }

            // Тестовые данные (заглушка)
            var ecr = GetECRFromSQLStub();

            // Пишем данные в Excel
            WriteToExcel(filePath, ecr);

            Console.WriteLine("\nДанные успешно записаны в файл: " + filePath);
            Console.WriteLine("Нажмите любую клавишу для завершения...");
            Console.ReadKey();
        }

        #region Пример тестовых данных (заглушка)

        private static ChangeRequest GetECRFromSQLStub()
        {
            var ecr = new ChangeRequest
            {
                ChangeRequestNo = "SQL-CR-001",
                ContractorChangeCoordinator = "SQL Coordinator",
                InitiatorInternalCRNo = "SQL-Int-456",
                ChangeInitiator = "Bob Initiator",
                InitiatorOrganization = "SQL Org",
                DocumentationType = "SQL-MLA",
                Reason = "SQL Test Reason",
                TechnicalSolutionDescription = "SQL Tech Solution (Initiator's long description...)",
            };

            // Пример: 5 SupportingDocuments и т.д.
            for (int i = 1; i <= 5; i++)
            {
                ecr.SupportingDocuments.Add(new SupportingDocument
                {
                    Filename = $"SQL-Doc{i}.pdf",
                    CodeTitleOrSummary = $"SQL Summary {i}"
                });

                ecr.InitiatorTDDImpacts.Add(new TDDImpact
                {
                    TDDCode = $"SQL-TDD-{i}00",
                    Revision = $"Rev{i}",
                    NewRevisionRequired = i % 2 == 0,
                    TDDName = $"SQL TDD {i}",
                    TDDState = "Active",
                    ChangeDescription = $"SQL Change {i}",
                    SafetyImpacts = new SafetyImpacts { NuclearSafety = i % 2 == 1 }
                });

                ecr.OtherTDDImpacts.Add(new TDDImpact
                {
                    EvaluationOrganization = $"SQL-EvalOrg{i}",
                    TDDCode = $"SQL-OtherTDD-{i}00",
                    Revision = $"Rev{i}X",
                    NewRevisionRequired = i % 2 == 1,
                    TDDName = $"SQL Other TDD {i}",
                    TDDState = "Draft",
                    ChangeDescription = $"SQL Other Change {i}",
                    SafetyImpacts = new SafetyImpacts { Environmental = i % 2 == 0 }
                });

                ecr.AffectedSSCs.Add(new SSCImpact
                {
                    KKSCode = $"SQL-KKS-{i}00",
                    SSCName = $"SQL SSC {i}",
                    ChangeDescription = $"SQL SSC Change {i}"
                });

                ecr.Confirmations.Add(new Confirmation
                {
                    Organization = $"SQL-Org{i}",
                    Position = $"SQL-Pos{i}",
                    ResponsiblePerson = $"SQL-Person{i}",
                    Date = $"2025-03-{i:D2}"
                });

                ecr.Approvals.Add(new Approval
                {
                    Position = $"SQL-ApprovalPos{i}",
                    ResponsiblePerson = $"SQL-AppPerson{i}",
                    Date = $"2025-04-{i:D2}"
                });
            }

            ecr.Section24 = new Section24
            {
                FinalNSCCategory = "SQL-aa",
                ImpactOnSafetyClasses = true,
                ImpactOnDSA = true,
                CRMethod = "SQL-Normal",
                EngineeringComments = "SQL Eng Comment"
            };

            ecr.Section25 = new Section25
            {
                ContractAffected = true,
                CostImpact = true,
                ScheduleImpact = false,
                NonTechnicalComments = "SQL Non-Tech Comment"
            };

            return ecr;
        }

        #endregion

        #region Основной метод записи в Excel

        private static void WriteToExcel(string filePath, ChangeRequest ecr)
        {
            FileInfo fi = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                if (sheet == null)
                {
                    Console.WriteLine("Ошибка: не найден лист с индексом 1.");
                    return;
                }

                int rows = sheet.Dimension.Rows;
                int cols = sheet.Dimension.Columns;
                Console.WriteLine("Лист [1] содержит {0} строк и {1} столбцов.", rows, cols);

                //
                // (1) Заполняем поля "Раздел 1" СТРОГО по вашим правилам:
                //     - если надпись в колонке A (c=1) -> ставим значение в колонку D (c=4)
                //     - если надпись в колонке H (c=8) -> ставим значение в колонку J (c=10)
                //     - если надпись в колонке K (c=11) -> ставим значение в колонку O (c=15)
                //

                // Подготовим переменную, если нужно отмечать "The presence of an impact"
                // (допустим, проверяем, есть ли вообще NewRevisionRequired в TDDImpacts)
                bool isAnyRevisionRequired = ecr.InitiatorTDDImpacts.Any(x => x.NewRevisionRequired)
                                          || ecr.OtherTDDImpacts.Any(x => x.NewRevisionRequired);
                string docImpactValue = isAnyRevisionRequired ? "Yes" : "No";

                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        string cellText = sheet.Cells[r, c].Text?.Trim() ?? "";
                        if (IsNullOrWhiteSpace(cellText)) continue;

                        // ---- Логика: смотрим, что за надпись, и куда писать ----

                        // 1) Если в этой ячейке "Change Request No:" / "Initiator's internal CR No:" / ...
                        //    и она в кол. A (c=1) -> пишем в D (col=4)
                        //    (Можно проверять cellText.Contains(...) для русского/англ. вариантов)

                        if (c == 1) // поле A
                        {
                            if (cellText.Contains("Change Request No:"))
                            {
                                SetValueFixedColumn(sheet, r, 4, ecr.ChangeRequestNo);
                            }
                            else if (cellText.Contains("Initiator's internal CR No:"))
                            {
                                SetValueFixedColumn(sheet, r, 4, ecr.InitiatorInternalCRNo);
                            }
                            else if (cellText.Contains("Change Initiator:"))
                            {
                                SetValueFixedColumn(sheet, r, 4, ecr.ChangeInitiator);
                            }
                            else if (cellText.Contains("Reason of Engineering Change:"))
                            {
                                SetValueFixedColumn(sheet, r, 4, ecr.Reason);
                            }
                            else if (cellText.Contains("Initiator's description of technical solution") 
                                     || cellText.Contains("description of technical solution")) 
                            {
                                SetValueFixedColumn(sheet, r, 4, ecr.TechnicalSolutionDescription);
                            }
                            else if (cellText.Contains("The presence of an impact on documents that need to be revised"))
                            {
                                // Пишем "Yes"/"No" в D
                                SetValueFixedColumn(sheet, r, 4, docImpactValue);
                            }
                        }
                        // 2) Если в этой ячейке "Contractor Change Coordinator:" / "Initiator's organization:"
                        //    и она в кол. H (c=8) -> пишем в col=10 (J)
                        else if (c == 8) // поле H
                        {
                            if (cellText.Contains("Contractor Change Coordinator:"))
                            {
                                SetValueFixedColumn(sheet, r, 10, ecr.ContractorChangeCoordinator);
                            }
                            else if (cellText.Contains("Initiator's organization:"))
                            {
                                SetValueFixedColumn(sheet, r, 10, ecr.InitiatorOrganization);
                            }
                        }
                        // 3) Если в кол. K (c=11) -> пишем в col=15 (O)
                        //    "Type of documentation, where Engineering Change will be reflected :"
                        else if (c == 11) // поле K
                        {
                            if (cellText.Contains("Type of documentation, where Engineering Change will be reflected"))
                            {
                                SetValueFixedColumn(sheet, r, 15, ecr.DocumentationType);
                            }
                        }
                    }
                }

                // (2) Динамические разделы
                int rowCount = sheet.Dimension.Rows; 
                
                // 1.2
                int row12 = FindRowByPrefix(sheet, rowCount, "1.2");
                if (row12 > 0)
                {
                    WriteSection12(sheet, row12 + 2, "2.1", ref rowCount, ecr.SupportingDocuments);
                }

                // 2.1
                int row21 = FindRowByPrefix(sheet, rowCount, "2.1");
                if (row21 > 0)
                {
                    WriteSection21(sheet, row21 + 3, "2.2", ref rowCount, ecr.InitiatorTDDImpacts);
                }

                // 2.2
                int row22 = FindRowByPrefix(sheet, rowCount, "2.2");
                if (row22 > 0)
                {
                    WriteSection22(sheet, row22 + 3, "2.3", ref rowCount, ecr.OtherTDDImpacts);
                }

                // 2.3
                int row23 = FindRowByPrefix(sheet, rowCount, "2.3");
                if (row23 > 0)
                {
                    WriteSection23(sheet, row23 + 2, "2.4", ref rowCount, ecr.AffectedSSCs);
                }

                // 3
                int row3 = FindRowByPrefix(sheet, rowCount, "3.");
                if (row3 > 0)
                {
                    WriteSection3(sheet, row3 + 2, "4.", ref rowCount, ecr.Confirmations);
                }

                // 4
                int row4 = FindRowByPrefix(sheet, rowCount, "4.");
                if (row4 > 0)
                {
                    WriteSection4(sheet, row4 + 2, ref rowCount, ecr.Approvals);
                }

                // Сохраняем
                package.Save();
                Console.WriteLine("Файл сохранён принудительно.");
            }
        }

        #endregion

        #region Утилиты для записи статических полей

        private static bool IsNullOrWhiteSpace(string s) => s == null || s.Trim().Length == 0;

        private static void SetValueFixedColumn(ExcelWorksheet sheet, int row, int fixedCol, string value)
        {
            var cell = sheet.Cells[row, fixedCol];
            cell.Value = value;
            Console.WriteLine($"Записано поле (row={row}, col={fixedCol}) -> {value}");
        }

        #endregion

        #region Копирование стилей, снятие Merge, задание Merge

        private static void CopyRowFormatting(ExcelWorksheet sheet, int sourceRow, int targetRow, int cols)
        {
            for (int c = 1; c <= cols; c++)
            {
                var sourceCell = sheet.Cells[sourceRow, c];
                var targetCell = sheet.Cells[targetRow, c];
                targetCell.StyleID = sourceCell.StyleID;
            }
        }

        private static void UnmergeRow(ExcelWorksheet sheet, int row)
        {
            var mergesToRemove = new List<string>();

            foreach (var mergedAddress in sheet.MergedCells)
            {
                if (string.IsNullOrEmpty(mergedAddress)) 
                    continue;

                var addr = new ExcelAddress(mergedAddress);

                // Снимаем только те слияния, которые полностью лежат в ОДНОЙ строке `row`
                // то есть Start.Row == End.Row == row
                if (addr.Start.Row == row && addr.End.Row == row)
                {
                    mergesToRemove.Add(mergedAddress);
                }
            }

            foreach (var addr in mergesToRemove)
            {
                sheet.Cells[addr].Merge = false;
            }
        }


        // 1.2
        private static void ApplyMergeSection12(ExcelWorksheet sheet, int row)
        {
            UnmergeRow(sheet, row);
            sheet.Cells[row, 1, row, 3].Merge = true;   // A..C
            sheet.Cells[row, 4, row, 15].Merge = true;  // D..O
        }

        // 2.1
        private static void ApplyMergeSection21(ExcelWorksheet sheet, int row)
        {
            UnmergeRow(sheet, row);
            sheet.Cells[row, 1, row, 3].Merge = true;   // A..C
            sheet.Cells[row, 6, row, 7].Merge = true;   // F..G
            sheet.Cells[row, 9, row, 10].Merge = true;  // I..J

            // центрируем содержимое по всей строке A..O
            var rng = sheet.Cells[row, 1, row, 15];
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        // 2.2
        private static void ApplyMergeSection22(ExcelWorksheet sheet, int row)
        {
            UnmergeRow(sheet, row);
            sheet.Cells[row, 2, row, 3].Merge = true;  // B..C
            sheet.Cells[row, 6, row, 7].Merge = true;  // F..G
            sheet.Cells[row, 9, row, 10].Merge = true; // I..J
        }

        // 2.3
        private static void ApplyMergeSection23(ExcelWorksheet sheet, int row)
        {
            UnmergeRow(sheet, row);
            sheet.Cells[row, 2, row, 5].Merge = true;   // B..E
            sheet.Cells[row, 6, row, 15].Merge = true;  // F..O
        }

        // 3
        private static void ApplyMergeSection3(ExcelWorksheet sheet, int row)
        {
            UnmergeRow(sheet, row);
            sheet.Cells[row, 2, row, 4].Merge = true;   // B..D
            sheet.Cells[row, 5, row, 9].Merge = true;   // E..I
            sheet.Cells[row, 11, row, 15].Merge = true; // K..O
        }

        // 4
        private static void ApplyMergeSection4(ExcelWorksheet sheet, int row)
        {
            UnmergeRow(sheet, row);
            sheet.Cells[row, 1, row, 4].Merge = true;   // A..D
            sheet.Cells[row, 5, row, 9].Merge = true;   // E..I
            sheet.Cells[row, 11, row, 15].Merge = true; // K..O
        }

        #endregion

        #region Поиск строки (FindRowByPrefix)

        private static int FindRowByPrefix(ExcelWorksheet sheet, int totalRows, string prefix)
        {
            for (int r = 1; r <= totalRows; r++)
            {
                string val = sheet.Cells[r, 1].Text?.Trim() ?? "";
                if (!IsNullOrWhiteSpace(val) && val.StartsWith(prefix))
                {
                    Console.WriteLine($"Найдена строка с префиксом '{prefix}': {r}");
                    return r;
                }
            }
            return 0;
        }

        #endregion

        #region Методы записи динамических разделов (1.2, 2.1, 2.2, 2.3, 3, 4)

        private static void WriteSection12(
            ExcelWorksheet sheet,
            int startRow,
            string nextPrefix,
            ref int totalRows,
            List<SupportingDocument> documents)
        {
            int availableRows = 0;
            int endRow = totalRows;
            for (int r = startRow; r <= totalRows; r++)
            {
                string firstCell = sheet.Cells[r, 1].Text?.Trim() ?? "";
                if (!IsNullOrWhiteSpace(firstCell) && firstCell.StartsWith(nextPrefix))
                {
                    availableRows = r - startRow;
                    endRow = r - 1;
                    break;
                }
                if (r == totalRows)
                    availableRows = totalRows - startRow + 1;
            }

            Console.WriteLine($"Доступно строк для 1.2: {availableRows}, требуется: {documents.Count}");
            if (documents.Count > availableRows)
            {
                int rowsToAdd = documents.Count - availableRows;
                sheet.InsertRow(endRow + 1, rowsToAdd);
                for (int r = endRow + 1; r <= endRow + rowsToAdd; r++)
                {
                    CopyRowFormatting(sheet, startRow, r, sheet.Dimension.Columns);
                    ApplyMergeSection12(sheet, r);
                }
                totalRows = sheet.Dimension.Rows;
            }

            for (int i = 0; i < documents.Count; i++)
            {
                int row = startRow + i;
                if (row <= endRow)
                    ApplyMergeSection12(sheet, row);

                sheet.Cells[row, 1].Value = documents[i].Filename;
                sheet.Cells[row, 4].Value = documents[i].CodeTitleOrSummary;
                Console.WriteLine($"[1.2] Записано: row={row}, Filename={documents[i].Filename}");
            }
        }

        private static void WriteSection21(
            ExcelWorksheet sheet,
            int startRow,
            string nextPrefix,
            ref int totalRows,
            List<TDDImpact> impacts)
        {
            int availableRows = 0;
            int endRow = totalRows;
            for (int r = startRow; r <= totalRows; r++)
            {
                string firstCell = sheet.Cells[r, 1].Text?.Trim() ?? "";
                if (!IsNullOrWhiteSpace(firstCell) && firstCell.StartsWith(nextPrefix))
                {
                    availableRows = r - startRow;
                    endRow = r - 1;
                    break;
                }
                if (r == totalRows)
                    availableRows = totalRows - startRow + 1;
            }

            Console.WriteLine($"Доступно строк для 2.1: {availableRows}, требуется: {impacts.Count}");
            if (impacts.Count > availableRows)
            {
                int rowsToAdd = impacts.Count - availableRows;
                sheet.InsertRow(endRow + 1, rowsToAdd);
                for (int r = endRow + 1; r <= endRow + rowsToAdd; r++)
                {
                    CopyRowFormatting(sheet, startRow, r, sheet.Dimension.Columns);
                    ApplyMergeSection21(sheet, r);
                }
                totalRows = sheet.Dimension.Rows;
            }

            for (int i = 0; i < impacts.Count; i++)
            {
                int row = startRow + i;
                if (row <= endRow)
                    ApplyMergeSection21(sheet, row);

                WriteOneImpact21(sheet, row, impacts[i]);
            }
        }

        private static void WriteOneImpact21(ExcelWorksheet sheet, int row, TDDImpact imp)
        {
            // A..C => TDD code
            sheet.Cells[row, 1].Value = imp.TDDCode;
            // D => Revision
            sheet.Cells[row, 4].Value = imp.Revision;
            // E => Is new revision required?
            sheet.Cells[row, 5].Value = imp.NewRevisionRequired ? "Yes" : "No";
            // F..G => TDD name
            sheet.Cells[row, 6].Value = imp.TDDName;
            // H => TDD state
            sheet.Cells[row, 8].Value = imp.TDDState;
            // I..J => Eng. Change description
            sheet.Cells[row, 9].Value = imp.ChangeDescription;
            // K..O => Potential impacts
            sheet.Cells[row, 11].Value = imp.SafetyImpacts.NuclearSafety ? "Yes" : "No";    
            sheet.Cells[row, 12].Value = imp.SafetyImpacts.FireSafety ? "Yes" : "No";
            sheet.Cells[row, 13].Value = imp.SafetyImpacts.IndustrialSafety ? "Yes" : "No";
            sheet.Cells[row, 14].Value = imp.SafetyImpacts.Environmental ? "Yes" : "No";
            sheet.Cells[row, 15].Value = imp.SafetyImpacts.Structural ? "Yes" : "No";

            Console.WriteLine($"[2.1] Записано: row={row}, TDDCode={imp.TDDCode}");
        }

        private static void WriteSection22(
            ExcelWorksheet sheet,
            int startRow,
            string nextPrefix,
            ref int totalRows,
            List<TDDImpact> impacts)
        {
            int availableRows = 0;
            int endRow = totalRows;
            for (int r = startRow; r <= totalRows; r++)
            {
                string firstCell = sheet.Cells[r, 1].Text?.Trim() ?? "";
                if (!IsNullOrWhiteSpace(firstCell) && firstCell.StartsWith(nextPrefix))
                {
                    availableRows = r - startRow;
                    endRow = r - 1;
                    break;
                }
                if (r == totalRows)
                    availableRows = totalRows - startRow + 1;
            }

            Console.WriteLine($"Доступно строк для 2.2: {availableRows}, требуется: {impacts.Count}");
            if (impacts.Count > availableRows)
            {
                int rowsToAdd = impacts.Count - availableRows;
                sheet.InsertRow(endRow + 1, rowsToAdd);
                for (int r = endRow + 1; r <= endRow + rowsToAdd; r++)
                {
                    CopyRowFormatting(sheet, startRow, r, sheet.Dimension.Columns);
                    ApplyMergeSection22(sheet, r);
                }
                totalRows = sheet.Dimension.Rows;
            }

            for (int i = 0; i < impacts.Count; i++)
            {
                int row = startRow + i;
                if (row <= endRow)
                    ApplyMergeSection22(sheet, row);

                WriteOneImpact22(sheet, row, impacts[i]);
            }
        }

        private static void WriteOneImpact22(ExcelWorksheet sheet, int row, TDDImpact imp)
        {
            sheet.Cells[row, 1].Value = imp.EvaluationOrganization;  
            sheet.Cells[row, 2].Value = imp.TDDCode;                 
            sheet.Cells[row, 4].Value = imp.Revision;                
            sheet.Cells[row, 5].Value = imp.NewRevisionRequired ? "Yes" : "No"; 
            sheet.Cells[row, 6].Value = imp.TDDName;                 
            sheet.Cells[row, 8].Value = imp.TDDState;                
            sheet.Cells[row, 9].Value = imp.ChangeDescription;       
            sheet.Cells[row, 11].Value = imp.SafetyImpacts.NuclearSafety ? "Yes" : "No";
            sheet.Cells[row, 12].Value = imp.SafetyImpacts.FireSafety ? "Yes" : "No";
            sheet.Cells[row, 13].Value = imp.SafetyImpacts.IndustrialSafety ? "Yes" : "No";
            sheet.Cells[row, 14].Value = imp.SafetyImpacts.Environmental ? "Yes" : "No";
            sheet.Cells[row, 15].Value = imp.SafetyImpacts.Structural ? "Yes" : "No";

            Console.WriteLine($"[2.2] Записано: row={row}, TDDCode={imp.TDDCode}");
        }

        private static void WriteSection23(
            ExcelWorksheet sheet,
            int startRow,
            string nextPrefix,
            ref int totalRows,
            List<SSCImpact> sscList)
        {
            int availableRows = 0;
            int endRow = totalRows;
            for (int r = startRow; r <= totalRows; r++)
            {
                string firstCell = sheet.Cells[r, 1].Text?.Trim() ?? "";
                if (!IsNullOrWhiteSpace(firstCell) && firstCell.StartsWith(nextPrefix))
                {
                    availableRows = r - startRow;
                    endRow = r - 1;
                    break;
                }
                if (r == totalRows)
                    availableRows = totalRows - startRow + 1;
            }

            Console.WriteLine($"Доступно строк для 2.3: {availableRows}, требуется: {sscList.Count}");
            if (sscList.Count > availableRows)
            {
                int rowsToAdd = sscList.Count - availableRows;
                sheet.InsertRow(endRow + 1, rowsToAdd);
                for (int r = endRow + 1; r <= endRow + rowsToAdd; r++)
                {
                    CopyRowFormatting(sheet, startRow, r, sheet.Dimension.Columns);
                    ApplyMergeSection23(sheet, r);
                }
                totalRows = sheet.Dimension.Rows;
            }

            for (int i = 0; i < sscList.Count; i++)
            {
                int row = startRow + i;
                if (row <= endRow)
                    ApplyMergeSection23(sheet, row);

                WriteOneSSC23(sheet, row, sscList[i]);
            }
        }

        private static void WriteOneSSC23(ExcelWorksheet sheet, int row, SSCImpact ssc)
        {
            sheet.Cells[row, 1].Value = ssc.KKSCode;
            sheet.Cells[row, 2].Value = ssc.SSCName;
            sheet.Cells[row, 6].Value = ssc.ChangeDescription;
            Console.WriteLine($"[2.3] Записано: row={row}, KKSCode={ssc.KKSCode}");
        }

        private static void WriteSection3(
            ExcelWorksheet sheet,
            int startRow,
            string nextPrefix,
            ref int totalRows,
            List<Confirmation> list)
        {
            int idx = 0;
            for (int r = startRow; r <= totalRows; r++)
            {
                string prefixCell = sheet.Cells[r, 1].Text?.Trim() ?? "";
                if (!IsNullOrWhiteSpace(prefixCell) && prefixCell.StartsWith(nextPrefix))
                {
                    int remain = list.Count - idx;
                    if (remain > 0)
                    {
                        sheet.InsertRow(r, remain);
                        for (int i = 0; i < remain; i++)
                        {
                            CopyRowFormatting(sheet, startRow, r + i, sheet.Dimension.Columns);
                            ApplyMergeSection3(sheet, r + i);
                            WriteOneConfirm3(sheet, r + i, list[idx + i]);
                        }
                        totalRows = sheet.Dimension.Rows;
                    }
                    break;
                }

                if (idx < list.Count)
                {
                    ApplyMergeSection3(sheet, r);
                    WriteOneConfirm3(sheet, r, list[idx++]);
                }
                else
                {
                    break;
                }
            }
        }

        private static void WriteOneConfirm3(ExcelWorksheet sheet, int row, Confirmation conf)
        {
            sheet.Cells[row, 1].Value = conf.Organization;
            sheet.Cells[row, 2].Value = conf.Position;
            sheet.Cells[row, 5].Value = conf.ResponsiblePerson;
            sheet.Cells[row, 11].Value = conf.Date;

            Console.WriteLine($"[3] Записано: row={row}, Org={conf.Organization}");
        }

        private static void WriteSection4(
            ExcelWorksheet sheet,
            int startRow,
            ref int totalRows,
            List<Approval> list)
        {
            int idx = 0;
            for (int r = startRow; r <= totalRows; r++)
            {
                string firstCell = sheet.Cells[r, 1].Text?.Trim() ?? "";
                if (!IsNullOrWhiteSpace(firstCell) &&
                    firstCell.ToLower().Contains("change manager of the owner"))
                {
                    int remain = list.Count - idx;
                    if (remain > 0)
                    {
                        sheet.InsertRow(r, remain);
                        for (int i = 0; i < remain; i++)
                        {
                            CopyRowFormatting(sheet, startRow, r + i, sheet.Dimension.Columns);
                            ApplyMergeSection4(sheet, r + i);
                            WriteOneApproval4(sheet, r + i, list[idx + i]);
                        }
                        totalRows = sheet.Dimension.Rows;
                    }
                    break;
                }

                if (idx < list.Count)
                {
                    if (!string.IsNullOrEmpty(list[idx].Position) &&
                        list[idx].Position.ToLower().Contains("change manager of the owner"))
                    {
                        Console.WriteLine($"[4] Пропущена запись: '{list[idx].Position}'");
                        idx++;
                        continue;
                    }

                    ApplyMergeSection4(sheet, r);
                    WriteOneApproval4(sheet, r, list[idx++]);
                }
                else
                {
                    break;
                }
            }
        }

        private static void WriteOneApproval4(ExcelWorksheet sheet, int row, Approval app)
        {
            sheet.Cells[row, 1].Value = app.Position;
            sheet.Cells[row, 5].Value = app.ResponsiblePerson;
            sheet.Cells[row, 11].Value = app.Date;

            Console.WriteLine($"[4] Записано: row={row}, Pos={app.Position}");
        }

        #endregion

        #region Классы-модели (без изменений)

        public class ChangeRequest
        {
            public string ChangeRequestNo;
            public string ContractorChangeCoordinator;
            public string InitiatorInternalCRNo;
            public string ChangeInitiator;
            public string InitiatorOrganization;
            public string DocumentationType;
            public string Reason;
            public string TechnicalSolutionDescription;

            public List<SupportingDocument> SupportingDocuments = new List<SupportingDocument>();
            public List<TDDImpact> InitiatorTDDImpacts = new List<TDDImpact>();
            public List<TDDImpact> OtherTDDImpacts = new List<TDDImpact>();
            public List<SSCImpact> AffectedSSCs = new List<SSCImpact>();

            public Section24 Section24 = new Section24();
            public Section25 Section25 = new Section25();

            public List<Confirmation> Confirmations = new List<Confirmation>();
            public List<Approval> Approvals = new List<Approval>();
        }

        public class SupportingDocument
        {
            public string Filename;
            public string CodeTitleOrSummary;
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

        public class Confirmation
        {
            public string Organization;
            public string Position;
            public string ResponsiblePerson;
            public string Date;
        }

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
            public SafetyImpacts SafetyImpacts = new SafetyImpacts();
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
}
