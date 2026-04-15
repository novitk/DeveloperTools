using ExcelDna.Testing;
using System;

namespace Xunit
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelTestSettingsAttribute : Attribute, ITestSettings
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
