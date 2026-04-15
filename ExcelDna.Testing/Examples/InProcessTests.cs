using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Xunit;

namespace Examples
{
    public class InProcessTests
    {
        [ExcelFact]
        public void GetExcelVersion()
        {
            Assert.Equal("16.0", Utils.GetVersion());
        }

        [ExcelFact(Workbook = "")]
        public void SetCellValue()
        {
            Range targetRange = (ExcelDna.Testing.Util.Workbook.Sheets[1] as Worksheet).Range["A1:C2"];

            object[,] newValues = new object[,] { { "One", 2, "Three" }, { true, System.DateTime.Now, "" } };
            targetRange.Value = newValues;

            var cell = new ExcelReference(0, 0);
            Assert.Equal("One", cell.GetValue().ToString());
        }

        [ExcelFact]
        public void ApplicationVersion()
        {
            Assert.Equal("16.0", ExcelDna.Testing.Util.Application.Version);
        }

        [ExcelFact]
        public void TestAssemblyDirectory()
        {
            Assert.False(string.IsNullOrEmpty(ExcelDna.Testing.Util.TestAssemblyDirectory));
        }

        [ExcelFact(Workbook = "MrExcel.xlsx")]
        public void PreCreatedWorkbook()
        {
            Range cell = (ExcelDna.Testing.Util.Workbook.Sheets[1] as Worksheet).Range["A2:A2"];
            Assert.Equal("Red Ford Truck", cell.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExampleAddin\bin\Debug\net472\ExampleAddin-AddIn")]
        public void ClickRibbonButton()
        {
            ExcelDna.Testing.Automation.ClickRibbonButton("ExampleAddin", "My Button");

            var cell = new ExcelReference(0, 0);
            Assert.Equal("One", cell.GetValue().ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExampleAddin\bin\Debug\net472\ExampleAddin-AddIn")]
        public void AsyncFunctionTest()
        {
            Range targetRange = (ExcelDna.Testing.Util.Workbook.Sheets[1] as Worksheet).Range["A1"];
            targetRange.Formula = "=MyAsyncFunction()";

            ExcelDna.Testing.Automation.WaitFor(() => targetRange.Value?.ToString() == "Completed", 1000);

            Assert.Equal("Completed", targetRange.Value);
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExampleAddinNET6\bin\Debug\net6.0-windows\ExampleAddinNET6-AddIn", SafeMode = true)]
        public void FunctionTestNET6()
        {
            Range targetRange = (ExcelDna.Testing.Util.Workbook.Sheets[1] as Worksheet).Range["A1"];
            targetRange.Formula = "=SayHelloNET6(\"world\")";

            Range cell = (ExcelDna.Testing.Util.Workbook.Sheets[1] as Worksheet).Range["A1:A1"];
            Assert.Equal("Hello .NET6 world", cell.Value.ToString());
        }

        [ExcelFact(Workbook = "", AddIn = @"..\..\..\..\ExampleAddinNET8\bin\Debug\net8.0-windows\ExampleAddinNET8-AddIn")]
        public void FunctionTestNET8()
        {
            Range targetRange = (ExcelDna.Testing.Util.Workbook.Sheets[1] as Worksheet).Range["A1"];
            targetRange.Formula = "=SayHelloNET8(\"world\")";

            Range cell = (ExcelDna.Testing.Util.Workbook.Sheets[1] as Worksheet).Range["A1:A1"];
            Assert.Equal("Hello .NET8 world", cell.Value.ToString());
        }
    }
}
