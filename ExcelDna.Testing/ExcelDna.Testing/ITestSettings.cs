namespace ExcelDna.Testing
{
    internal interface ITestSettings
    {
        /// <summary>
        /// Whether tests are executed outside of the Excel process.
        /// </summary>
        bool OutOfProcess { get; set; }

        /// <summary>
        /// A relative path to workbook file to open. Use empty path to create a new workbook.
        /// </summary>
        string Workbook { get; set; }

        /// <summary>
        /// A relative path to .xll add-in to load. Without bitness and .xll extension.
        /// </summary>
        string AddIn { get; set; }

        /// <summary>
        /// Whether to start Excel in safe mode using the /safe command-line switch.
        /// </summary>
        bool SafeMode { get; set; }
    }
}
