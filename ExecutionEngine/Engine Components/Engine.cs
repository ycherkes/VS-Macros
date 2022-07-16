//-----------------------------------------------------------------------
// <copyright file="Engine.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using EnvDTE80;
using ExecutionEngine.Enums;
using ExecutionEngine.Helpers;
using ExecutionEngine.Interfaces;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using VSMacros.ExecutionEngine;

namespace ExecutionEngine
{
    public sealed class Engine : IDisposable
    {
        private IActiveScript engine;
        private readonly Parser parser;
        private readonly Site scriptSite;
        private object _dispatch;

        public static DTE2 DteObject { get; private set; }
        public static object CommandHelper { get; private set; }

        private static IMoniker GetItemMoniker(int pid, string version)
        {
            ErrorHandler.ThrowOnFailure(NativeMethods.CreateItemMoniker("!", string.Format(CultureInfo.InvariantCulture, "VisualStudio.DTE.{0}:{1}", version, pid), out var moniker));

            return moniker;
        }

        private static IRunningObjectTable GetRunningObjectTable()
        {
            int hr = NativeMethods.GetRunningObjectTable(0, out var rot);
            if (ErrorHandler.Failed(hr))
            {
                ErrorHandler.ThrowOnFailure(hr, null);
            }

            return rot;
        }

        private static DTE2 GetDteObject(IRunningObjectTable rot, IMoniker moniker)
        {
            int hr = rot.GetObject(moniker, out var dteObject);
            if (ErrorHandler.Failed(hr))
            {
                ErrorHandler.ThrowOnFailure(hr, null);
            }

            return dteObject as DTE2;
        }

        private static void InitializeDteObject(int pid, string version)
        {
            IMoniker moniker = GetItemMoniker(pid, version);
            IRunningObjectTable rot = GetRunningObjectTable();
            DteObject = GetDteObject(rot, moniker);

            if (DteObject == null)
            {
                throw new InvalidOperationException();
            }
        }

        private static void InitializeCommandHelper()
        {
            var globalProvider = ServiceProvider.GlobalProvider;
            Validate.IsNotNull(globalProvider, "globalProvider");

            CommandHelper = new CommandHelper(globalProvider);
            Validate.IsNotNull(CommandHelper, "Engine.CommandHelper");
        }

        private static IActiveScript CreateEngine()
        {
            Type engine = Type.GetTypeFromProgID("jscript", true);
            return Activator.CreateInstance(engine) as IActiveScript;
        }

        public Engine(int pid, string version)
        {
            engine = CreateEngine();
            scriptSite = new Site();
            parser = new Parser(engine);
            engine.SetScriptSite(scriptSite);

            InformEngineOfNewObjects(pid, version);
        }

        private void InformEngineOfNewObjects(int pid, string version)
        {
            const string dte = "dte";
            InitializeDteObject(pid, version);
            engine.AddNamedItem(dte, ScriptItem.CodeOnly | ScriptItem.IsVisible);

            const string cmdHelper = "cmdHelper";
            InitializeCommandHelper();
            engine.AddNamedItem(cmdHelper, ScriptItem.CodeOnly | ScriptItem.IsVisible);
        }

        public void Dispose()
        {
            engine = null;
        }

        public bool CallMethod(string methodName, params object[] arguments)
        {
            try
            {
                _dispatch.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, _dispatch, arguments);
                return true;
            }
            catch (Exception e)
            {
                if (Site.RuntimeError) return false;

                Site.InternalVsException = new InternalVsException(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());

                return false;
            }
        }

        internal void Parse(string unparsed)
        {
            try
            {
                engine.SetScriptState(ScriptState.Connected);
                parser.Parse(unparsed);
                engine.GetScriptDispatch(null, out var dispatch);
                _dispatch = Marshal.GetObjectForIUnknown(dispatch);
            }
            catch (Exception e)
            {
                Site.InternalVsException = new InternalVsException(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
            }
        }
    }
}
