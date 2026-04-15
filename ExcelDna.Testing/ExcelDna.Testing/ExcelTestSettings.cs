namespace ExcelDna.Testing
{
    public class ExcelTestSettings : ITestSettings
    {
        public ExcelTestSettings(bool outOfProcess, string workbook, string addin, bool safeMode = false)
        {
            OutOfProcess = outOfProcess;
            Workbook = workbook;
            AddIn = addin;
            SafeMode = safeMode;
        }

        public bool OutOfProcess { get; set; }

        public string Workbook { get; set; }

        public string AddIn { get; set; }

        public bool SafeMode { get; set; }
    }
}
