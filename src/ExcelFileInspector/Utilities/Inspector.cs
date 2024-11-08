﻿using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;

namespace ExcelFileInspector.Utilities
{
    /// <summary>
    /// 検査クラス
    /// </summary>
    public class Inspector
    {
        /// <summary>
        /// 検査結果クラス
        /// </summary>
        public class InspectionResult
        {
            /// <summary>
            /// ファイル名
            /// </summary>
            public string FileName { get; set; } = string.Empty;

            /// <summary>
            /// セル
            /// </summary>
            public string Cell { get; set; } = string.Empty;

            /// <summary>
            /// 結果
            /// </summary>
            public string ResultMessage { get; set; } = string.Empty;
        }

        /// <summary>
        /// ファイル検査処理
        /// </summary>
        /// <param name="filename">検査対象ファイル名</param>
        /// <param name="methods">検査方法</param>
        /// <returns></returns>
        public static List<InspectionResult> InspectionFile(string filename, List<SettingManager.InspectionMethod> methods)
        {
            List<InspectionResult> results = [];
            InspectionResult result;

            // Bookを開く(読み取り専用で開く)
            using (FileStream fs = new(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using XLWorkbook workbook = new(fs);
                // 全シート調査
                foreach (IXLWorksheet worksheet in workbook.Worksheets)
                {
                    // 検査方法で指定された条件を確認
                    foreach (SettingManager.InspectionMethod method in methods)
                    {
                        switch (method.Condition)
                        {
                            case "Equal":
                                // 指定されたセルが指定された値になっている場合はNG
                                if (worksheet.Name.Equals(method.SheetName) && (worksheet.Cell(method.Cell).Value.ToString() == method.Value))
                                {
                                    result = new()
                                    {
                                        FileName = filename,
                                        Cell = method.Cell,
                                        ResultMessage = string.Format(Resources.Strings.MessageResultInspectionNGEqual, method.Value)
                                    };
                                    results.Add(result);
                                }
                                break;
                            case "NotEqual":
                                // 指定されたセルが指定された値以外になっている場合はNG
                                if (worksheet.Name.Equals(method.SheetName) && (worksheet.Cell(method.Cell).Value.ToString() != method.Value))
                                {
                                    result = new()
                                    {
                                        FileName = filename,
                                        Cell = method.Cell,
                                        ResultMessage = string.Format(Resources.Strings.MessageResultInspectionNGNotEqual, method.Value)
                                    };
                                    results.Add(result);
                                }
                                break;
                            case "Empty":
                                // 指定されたセルが空である場合はNG
                                if (worksheet.Name.Equals(method.SheetName) && (worksheet.Cell(method.Cell).Value.ToString() == string.Empty))
                                {
                                    result = new()
                                    {
                                        FileName = filename,
                                        Cell = method.Cell,
                                        ResultMessage = Resources.Strings.MessageResultInspectionNGEmpty
                                    };
                                    results.Add(result);
                                }
                                break;
                            case "NotEmpty":
                                // 指定されたセルが空ではない場合はNG
                                if (worksheet.Name.Equals(method.SheetName) && (worksheet.Cell(method.Cell).Value.ToString() != string.Empty))
                                {
                                    result = new()
                                    {
                                        FileName = filename,
                                        Cell = method.Cell,
                                        ResultMessage = Resources.Strings.MessageResultInspectionNGNotEmpty
                                    };
                                    results.Add(result);
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            // NGが1つもない場合は問題なしとする
            if (results.Count == 0)
            {
                result = new()
                {
                    FileName = filename,
                    Cell = "―",
                    ResultMessage = Resources.Strings.MessageResultInspectionOK
                };
                results.Add(result);
            }

            return results;
        }
    }
}
