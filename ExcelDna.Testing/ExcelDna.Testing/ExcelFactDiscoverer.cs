using System.Collections.Generic;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;
using System.Linq;

namespace ExcelDna.Testing
{
    class ExcelFactDiscoverer : FactDiscoverer
    {
        public ExcelFactDiscoverer(IMessageSink diagnosticMessageSink) : base(diagnosticMessageSink)
        {
        }

        public override IEnumerable<IXunitTestCase> Discover(ITestFrameworkDiscoveryOptions discoveryOptions, ITestMethod testMethod, IAttributeInfo factAttribute)
        {
            var results = new List<IXunitTestCase>();
            results.Add(new ExcelTestCase(GetSettings(testMethod), DiagnosticMessageSink, discoveryOptions.MethodDisplayOrDefault(), testMethod, null));
            return results;
        }

        private static ExcelTestSettings GetSettings(ITestMethod testMethod)
        {
            return new ExcelTestSettings(
                GetSetting<bool>(testMethod, nameof(ExcelFactAttribute.OutOfProcess)),
                GetSetting<string>(testMethod, nameof(ExcelFactAttribute.Workbook)),
                GetSetting<string>(testMethod, nameof(ExcelFactAttribute.AddIn)),
                GetSetting<bool>(testMethod, nameof(ExcelFactAttribute.SafeMode)));
        }

        private static T GetSetting<T>(ITestMethod testMethod, string name)
        {
            object arg = GetNamedArg(testMethod.Method.GetCustomAttributes(typeof(ExcelFactAttribute)).FirstOrDefault(), name);
            if (arg == null)
                arg = GetNamedArg(testMethod.TestClass.Class.GetCustomAttributes(typeof(ExcelTestSettingsAttribute)).FirstOrDefault(), name);

            return arg is T ? (T)arg : default(T);
        }

        private static object GetNamedArg(IAttributeInfo attributeInfo, string argumentName)
        {
            if (!IsNamedArg(attributeInfo, argumentName))
                return null;

            return attributeInfo.GetNamedArgument<object>(argumentName);
        }

        private static bool IsNamedArg(IAttributeInfo attributeInfo, string argumentName)
        {
            if (attributeInfo is ReflectionAttributeInfo reflectionAttributeInfo)
                return reflectionAttributeInfo.AttributeData.NamedArguments.Any(arg => arg.MemberName == argumentName);
            return false;
        }
    }
}
