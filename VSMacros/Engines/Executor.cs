//-----------------------------------------------------------------------
// <copyright file="Executor.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using EnvDTE80;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using VSMacros.ExecutionEngine.Pipes.Shared;
using VSMacros.Helpers;
using VSMacros.Interfaces;
using VSMacros.Pipes;

namespace VSMacros.Engines
{
    /// <summary>
    /// Implements the execution engine.
    /// </summary>   
    internal sealed class Executor : IExecutor
    {
        /// <summary>
        /// The execution engine.
        /// </summary>
        internal static Process executionEngine;
        internal static JobHandle Job;

        /// <summary>
        /// Informs subscribers of success or error during execution.
        /// </summary>
        public event EventHandler<CompletionReachedEventArgs> Complete;
        public string CurrentlyExecutingMacro { get; set; }
        public bool IsEngineRunning { get; set; }

        #region Helpers
        internal void SendCompletionMessage(bool isError, string errorMessage)
        {
            if (Complete != null)
            {
                var eventArgs = new CompletionReachedEventArgs(isError, errorMessage);
                Complete(this, eventArgs);
            }
        }

        internal void ResetMessages()
        {
            Complete = null;
        }

        private string ProvidePipeArguments(Guid guid, string version)
        {
            int pid = Process.GetCurrentProcess().Id;
            return string.Format("{0}{1}{2}{1}{3}", guid, SharedVariables.Delimiter, pid, version);
        }

        private System.Runtime.InteropServices.ComTypes.IRunningObjectTable GetRunningObjectTable()
        {
            System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot;
            ErrorHandler.ThrowOnFailure(NativeMethods.GetRunningObjectTable(0, out rot));

            return rot;
        }

        private static string GetProgID(Type type)
        {
            Guid guid = Marshal.GenerateGuidForType(type);
            string progID = String.Format(CultureInfo.InvariantCulture, "{0:B}", guid);
            return progID;
        }

        private void RegisterCommandDispatcherinROT()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            int hResult;
            System.Runtime.InteropServices.ComTypes.IBindCtx bc;
            hResult = NativeMethods.CreateBindCtx(0, out bc);

            if (hResult == NativeMethods.S_OK && bc != null)
            {
                System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot = GetRunningObjectTable();
                Validate.IsNotNull(rot, "rot");

                string delim = "!";
                string progID = GetProgID(typeof(SUIHostCommandDispatcher));

                System.Runtime.InteropServices.ComTypes.IMoniker moniker;
                hResult = NativeMethods.CreateItemMoniker(delim, progID, out moniker);
                Validate.IsNotNull(moniker, "moniker");

                if (hResult == NativeMethods.S_OK)
                {
                    var serviceProvider = ServiceProvider.GlobalProvider;

                    var suiHost = serviceProvider.GetService(typeof(SUIHostCommandDispatcher));
                    hResult = rot.Register(NativeMethods.ROTFLAGS_REGISTRATIONKEEPSALIVE, suiHost, moniker);
                }
            }
        }

        private void RegisterCmdNameMappinginROT()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            int hResult;
            System.Runtime.InteropServices.ComTypes.IBindCtx bc;
            hResult = NativeMethods.CreateBindCtx(0, out bc);

            if (hResult == NativeMethods.S_OK && bc != null)
            {
                System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot = GetRunningObjectTable();
                Validate.IsNotNull(rot, "rot");

                string delim = "!";
                string progID = GetProgID(typeof(SVsCmdNameMapping));

                System.Runtime.InteropServices.ComTypes.IMoniker moniker;
                hResult = NativeMethods.CreateItemMoniker(delim, progID, out moniker);
                Validate.IsNotNull(moniker, "moniker");

                if (hResult == NativeMethods.S_OK)
                {
                    var serviceProvider = ServiceProvider.GlobalProvider;

                    var svsCmdName = serviceProvider.GetService(typeof(SVsCmdNameMapping));
                    hResult = rot.Register(NativeMethods.ROTFLAGS_REGISTRATIONKEEPSALIVE, svsCmdName, moniker);
                }
            }
        }
        #endregion

        /// <summary>
        /// Initializes the engine.
        /// </summary>
        /// 
        public void InitializeEngine()
        {
            RegisterCmdNameMappinginROT();
            RegisterCommandDispatcherinROT();

            Server.InitializeServer();
            Server.serverWait = new Thread(Server.WaitForMessage);
            Server.serverWait.Start();

            DTE2 dte = ((IServiceProvider)VSMacrosPackage.Current).GetService(typeof(SDTE)) as DTE2;
            string version = dte.Version;
            executionEngine = new Process();
            string exeFileName = "VisualStudio.Macros.ExecutionEngine.exe";
            string processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), exeFileName);
            executionEngine.StartInfo.FileName = processName;
            executionEngine.StartInfo.UseShellExecute = false;
            executionEngine.StartInfo.CreateNoWindow = true;
            executionEngine.StartInfo.Arguments = ProvidePipeArguments(Server.Guid, version);
            executionEngine.Start();

            Job = JobHandle.CreateNewJob();
            Job.AddProcess(executionEngine);
        }

        private static bool IsExecutorReady()
        {
            return executionEngine != null && !executionEngine.HasExited;
        }

        private static bool IsServerReady()
        {
            return Server.ServerStream != null && Server.ServerStream.IsConnected;
        }

        /// <summary>
        /// If engine is initialized, runs the engine.  Otherwise, initializes and runs the engine.
        /// </summary>
        /// 
        public void RunEngine(int iterations, string path)
        {
            if (IsServerReady() && IsExecutorReady())
            {
                Thread sendPath = new Thread(() =>
                {
                    Server.SendFilePath(iterations, path);
                }
                );
                sendPath.Start();
            }
            else
            {
                InitializeEngine();
                Thread waitForConnection = new Thread(() =>
                    {
                        WaitForConnection();
                        Server.SendFilePath(iterations, path);
                    }
                );
                waitForConnection.Start();
            }

            IsEngineRunning = true;
        }

        private static void WaitForConnection()
        {
            while (!Server.ServerStream.IsConnected) { }
        }

        public void StopEngine()
        {
            Job.Close();
            IsEngineRunning = false;
        }
    }
}
