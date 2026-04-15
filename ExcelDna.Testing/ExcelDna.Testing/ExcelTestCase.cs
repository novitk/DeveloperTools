using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Xunit.Abstractions;
using Xunit.Internal;
using Xunit.Sdk;

namespace ExcelDna.Testing
{
    [Newtonsoft.Json.JsonConverter(typeof(Remote.ExcelTestCaseConverter))]
    public class ExcelTestCase : XunitTestCase
    {
        [Obsolete("Called by the de-serializer; should only be called by deriving classes for de-serialization purposes")]
        public ExcelTestCase()
        {
        }

#pragma warning disable CS0618
        public ExcelTestCase(ExcelTestSettings testSettings, IMessageSink diagnosticMessageSink, TestMethodDisplay defaultMethodDisplay, ITestMethod testMethod, object[] testMethodArguments = null)
            : base(diagnosticMessageSink, defaultMethodDisplay, testMethod, testMethodArguments)
        {
            this.testSettings = testSettings;
        }
#pragma warning restore CS0618

        public ExcelTestSettings Settings => testSettings;

        public override Task<RunSummary> RunAsync(IMessageSink diagnosticMessageSink, IMessageBus messageBus, object[] constructorArguments, ExceptionAggregator aggregator, CancellationTokenSource cancellationTokenSource)
    => new ExcelTestCaseRunner(this, DisplayName, SkipReason, constructorArguments, TestMethodArguments, messageBus, aggregator, cancellationTokenSource).RunAsync();

        public override void Serialize(IXunitSerializationInfo info)
        {
            base.Serialize(info);
            info.AddValue(nameof(testSettings.OutOfProcess), testSettings.OutOfProcess);
            info.AddValue(nameof(testSettings.Workbook), testSettings.Workbook);
            info.AddValue(nameof(testSettings.AddIn), testSettings.AddIn);
            info.AddValue(nameof(testSettings.SafeMode), testSettings.SafeMode);
        }

        public override void Deserialize(IXunitSerializationInfo info)
        {
            base.Deserialize(info);
            testSettings = new ExcelTestSettings(
                info.GetValue<bool>(nameof(testSettings.OutOfProcess)),
                info.GetValue<string>(nameof(testSettings.Workbook)),
                info.GetValue<string>(nameof(testSettings.AddIn)),
                info.GetValue<bool>(nameof(testSettings.SafeMode)));
        }

        public string SerializeToString()
        {
            var triple = new XunitSerializationTriple(nameof(ExcelTestCase), this, GetType());
            return XunitSerializationInfo.SerializeTriple(triple);
        }

        public static ExcelTestCase DeserializeFromString(string value)
        {
            var triple = XunitSerializationInfo.DeserializeTriple(value);
            return (ExcelTestCase)triple.Value;
        }

        private ExcelTestSettings testSettings;
    }
}
