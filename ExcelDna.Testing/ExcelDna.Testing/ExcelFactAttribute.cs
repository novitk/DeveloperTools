using System;
using Xunit.Sdk;
using ExcelDna.Testing;

namespace Xunit
{
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
    [XunitTestCaseDiscoverer("ExcelDna.Testing." + nameof(ExcelFactDiscoverer), "ExcelDna.Testing")]
    public class ExcelFactAttribute : FactAttribute, ITestSettings
    {
        /// <inheritdoc />
        public bool OutOfProcess { get; set; }

        /// <inheritdoc />
        public string Workbook { get; set; }

        /// <inheritdoc />
        public string AddIn { get; set; }

        /// <inheritdoc />
        public bool SafeMode { get; set; }
    }
}
