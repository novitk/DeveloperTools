using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace ExcelDna.Testing
{
    internal class ExcelTestAssemblyRunner : XunitTestAssemblyRunner
    {
        private readonly ITestAssembly testAssembly;
        private readonly IMessageSink diagnosticMessageSink;
        private readonly IMessageSink executionMessageSink;
        private readonly ITestFrameworkExecutionOptions executionOptions;
        private readonly ExcelRunner excelRunner;

        public ExcelTestAssemblyRunner(ITestAssembly testAssembly, IEnumerable<IXunitTestCase> testCases, IMessageSink diagnosticMessageSink, IMessageSink executionMessageSink, ITestFrameworkExecutionOptions executionOptions)
            : base(testAssembly, testCases, diagnosticMessageSink, executionMessageSink, executionOptions)
        {
            this.testAssembly = testAssembly;
            this.diagnosticMessageSink = diagnosticMessageSink;
            this.executionMessageSink = executionMessageSink;
            this.executionOptions = executionOptions;
            excelRunner = new ExcelRunner();
        }

        protected override async Task<RunSummary> RunTestCollectionsAsync(IMessageBus messageBus, CancellationTokenSource cancellationTokenSource)
        {
            IEnumerable<IXunitTestCase> localTestCases = TestCases.Except(TestCases.OfType<ExcelTestCase>());
            IEnumerable<ExcelTestCase> excelInProcessTestCases = TestCases.OfType<ExcelTestCase>().Where(i => !i.Settings.OutOfProcess && !i.Settings.SafeMode);
            IEnumerable<ExcelTestCase> excelInProcessSafeTestCases = TestCases.OfType<ExcelTestCase>().Where(i => !i.Settings.OutOfProcess && i.Settings.SafeMode);
            IEnumerable<ExcelTestCase> excelOutOfProcessTestCases = TestCases.OfType<ExcelTestCase>().Where(i => i.Settings.OutOfProcess);

            var result = await LocalRunTestCasesAsync(localTestCases, messageBus, cancellationTokenSource);
            if (excelOutOfProcessTestCases.Count() > 0)
                result.Aggregate(await COMRunTestCasesAsync(excelOutOfProcessTestCases, messageBus, cancellationTokenSource));
            if (excelInProcessTestCases.Count() > 0)
                result.Aggregate(await RemoteRunTestCasesAsync(excelInProcessTestCases, messageBus, cancellationTokenSource, false));
            if (excelInProcessSafeTestCases.Count() > 0)
                result.Aggregate(await RemoteRunTestCasesAsync(excelInProcessSafeTestCases, messageBus, cancellationTokenSource, true));

            CleanupReferences();

            return result;
        }

        private async Task<RunSummary> LocalRunTestCasesAsync(IEnumerable<IXunitTestCase> testCases, IMessageBus messageBus, CancellationTokenSource cancellationTokenSource)
        {
            var allTestCases = TestCases;
            TestCases = testCases;
            var result = await base.RunTestCollectionsAsync(messageBus, cancellationTokenSource);
            TestCases = allTestCases;
            return result;
        }

        private async Task<RunSummary> RemoteRunTestCasesAsync(IEnumerable<ExcelTestCase> testCases, IMessageBus messageBus, CancellationTokenSource cancellationTokenSource, bool safeMode)
        {
            RunSummary result = new RunSummary();
            try
            {
                ExcelStartupEvent.Create();
                Process excelProcess = excelRunner.Start(testAssembly.Assembly.AssemblyPath, GetAddins(testCases), safeMode);
                if (!ExcelStartupEvent.Wait(30000))
                    throw new System.ApplicationException("Excel startup failed.");

                if (Debugger.IsAttached)
                    VS.VisualStudioInstance.AttachDebugger(excelProcess);

                var bus = new DynMessageBus(messageBus, message =>
                {
                    if (cancellationTokenSource.Token.IsCancellationRequested)
                        return false;

                    switch (message)
                    {
                        case ITestAssemblyStarting assemblyStarting:
                        case ITestAssemblyFinished assemblyFinished:
                            return true;
                        case ITestCaseStarting testCaseStarting:
                            break;
                        case ITestCaseFinished testCaseFinished:
                            break;
                    }

                    return messageBus.QueueMessage(message);
                });

                using (var stream = new NamedPipeClientStream(".", "ExcelDna.Testing", PipeDirection.InOut, PipeOptions.Asynchronous))
                {
                    await stream.ConnectAsync();
                    Remote.IRemoteExcel remoteObject = StreamJsonRpc.JsonRpc.Attach<Remote.IRemoteExcel>(stream);

                    remoteObject.BusMessage += (_, args) =>
                    {
                        var msg = args.GetMessage();
                        if (msg != null)
                            bus.QueueMessage(msg);
                    };

                    EventWaitHandle finalMessageEvent = new EventWaitHandle(false, EventResetMode.ManualReset);
                    remoteObject.FinalMessage += (_, args) => finalMessageEvent.Set();

                    result = await remoteObject.RunTestsAsync(testAssembly.Assembly.AssemblyPath, testAssembly.ConfigFileName, testCases.Select(i => i.SerializeToString()).ToArray());
                    await remoteObject.SendFinalMessageAsync();
                    finalMessageEvent.WaitOne(30000);
                    await remoteObject.CloseHostAsync();
                }
            }
            catch (System.Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
            }

            return result;
        }

        private async Task<RunSummary> COMRunTestCasesAsync(IEnumerable<ExcelTestCase> testCases, IMessageBus messageBus, CancellationTokenSource cancellationTokenSource)
        {
            try
            {
                Util.Application = new Microsoft.Office.Interop.Excel.Application();
                Util.TestAssemblyDirectory = RunnerUtil.TestAssemblyDirectory(testAssembly, testCases);
                Bitness bitness = Marshal.SizeOf(Util.Application.HinstancePtr) == 8 ? Bitness.Bit64 : Bitness.Bit32;
                foreach (string addin in GetAddins(testCases))
                    Util.Application.RegisterXLL(ExcelRunner.GetXllPath(Util.TestAssemblyDirectory, addin, bitness));
            }
            catch (System.Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
                return new RunSummary();
            }

            TestCases = testCases;
            try
            {
                return await base.RunTestCollectionsAsync(messageBus, cancellationTokenSource);
            }
            finally
            {
                Util.Application = null;
            }
        }

        private static List<string> GetAddins(IEnumerable<ExcelTestCase> testCases)
        {
            return testCases.Select(i => i.Settings.AddIn).
                Where(i => i != null).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList();
        }

        private class DynMessageBus : IMessageBus
        {
            private readonly IMessageBus messageBus;

            public DynMessageBus(IMessageBus messageSink, Func<IMessageSinkMessage, bool> onMessage)
            {
                this.messageBus = messageSink;
                OnMessageCallback = onMessage;
            }

            public Func<IMessageSinkMessage, bool> OnMessageCallback { get; }

            public void Dispose()
            {
            }

            public bool QueueMessage(IMessageSinkMessage message)
                => OnMessageCallback(message);
        }

        private static void CleanupReferences()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}