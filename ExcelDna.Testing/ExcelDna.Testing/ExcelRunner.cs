using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace ExcelDna.Testing
{
    internal class ExcelRunner
    {
        public ExcelRunner()
        {
            ExcelDetector excelDetector = new ExcelDetector();
            excelDetected = excelDetector.TryFindLatestExcel(out excelExePath) && excelDetector.TryFindExcelBitness(excelExePath, out bitness);
        }

        public Process Start(string addinAssemblyPath, IEnumerable<string> addins, bool safeMode = false)
        {
            if (!excelDetected)
                throw new ApplicationException("Can't find an installed version of Excel.");

            string addinAssemblyDirectory = Path.GetDirectoryName(addinAssemblyPath);
            string arguments = "";
            
            if (safeMode)
            {
                arguments += "/safe ";
            }
            
            foreach (string externalAddinRelativePath in addins)
            {
                arguments += Quote(GetXllPath(addinAssemblyDirectory, externalAddinRelativePath, bitness)) + " ";
            }

            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = excelExePath;
            info.Arguments = arguments + Quote(GetExcelAgentXllPath(addinAssemblyDirectory, bitness));
            return Process.Start(info);
        }

        public static string GetXllPath(string addinAssemblyDirectory, string externalXllRelativePath, Bitness bitness)
        {
            return Path.Combine(addinAssemblyDirectory, externalXllRelativePath + (bitness == Bitness.Bit64 ? "64" : "") + ".xll");
        }

        private static string GetExcelAgentXllPath(string addinAssemblyDirectory, Bitness bitness)
        {
            return GetXllPath(addinAssemblyDirectory, "ExcelAgent", bitness);
        }

        private string Quote(string s)
        {
            return "\"" + s + "\"";
        }

        private string excelExePath;
        private Bitness bitness;
        private bool excelDetected;
    }
}
