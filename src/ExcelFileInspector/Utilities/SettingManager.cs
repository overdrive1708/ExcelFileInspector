namespace ExcelFileInspector.Utilities
{
    /// <summary>
    /// 設定管理クラス
    /// </summary>
    public class SettingManager
    {
        /// <summary>
        /// 検査方法クラス
        /// </summary>
        public class InspectionMethod
        {
            /// <summary>
            /// シート名
            /// </summary>
            public string SheetName { get; set; } = string.Empty;

            /// <summary>
            /// セル
            /// </summary>
            public string Cell { get; set; } = string.Empty;

            /// <summary>
            /// 条件
            /// </summary>
            public string Condition { get; set; } = string.Empty;

            /// <summary>
            /// 値
            /// </summary>
            public string Value { get; set; } = string.Empty;
        }
    }
}
