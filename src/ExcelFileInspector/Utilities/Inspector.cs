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
    }
}
