//-----------------------------------------------------------------------
// <copyright file="Site.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using ExecutionEngine.Enums;
using ExecutionEngine.Interfaces;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using VisualStudio.Macros.ExecutionEngine;
using VSMacros.ExecutionEngine;

namespace ExecutionEngine
{
    internal sealed class Site : IActiveScriptSite
    {
        private const int TypeEElementNotFound = unchecked((int)0x8002802B);
        private const string Dte = "dte";
        private const string CmdHelper = "cmdHelper";
        private Dictionary<string, object> objectsEngineKnowsAbout;
        private Dictionary<string, string> errorMessages;
        internal static bool RuntimeError;
        internal static RuntimeException RuntimeException;
        internal static InternalVsException InternalVsException;

        public void GetLCID(out int lcid)
        {
            lcid = Thread.CurrentThread.CurrentCulture.LCID;
        }

        private static void CreateInternalVsException(string message)
        {
            string source = "VSMacros.ExecutionEngine";
            string stackTrace = string.Empty;
            string targetSite = string.Empty;
            InternalVsException = new InternalVsException(message, source, stackTrace, targetSite);
        }

        public void GetItemInfo(string name, ScriptInfo returnMask, out IntPtr item, IntPtr typeInfo)
        {
            EnsureDictionariesAreInitialized();

            if ((returnMask & ScriptInfo.ITypeInfo) == ScriptInfo.ITypeInfo)
            {
                throw new NotImplementedException();
            }

            if (objectsEngineKnowsAbout.TryGetValue(name, out var objectEngineKnowsAbout))
            {
                item = Marshal.GetIUnknownForObject(objectEngineKnowsAbout);
            }
            else if (errorMessages.TryGetValue(name, out var errorMessage))
            {
                CreateInternalVsException(errorMessage);
                item = IntPtr.Zero;
            }
            else
            {
                throw new COMException(null, TypeEElementNotFound);
            }
        }

        private void EnsureDictionariesAreInitialized()
        {
            if (objectsEngineKnowsAbout == null)
            {
                InitializedAddedObjects();
            }

            if (errorMessages == null)
            {
                InitializeErrorMessages();
            }
        }

        private void InitializeErrorMessages()
        {
            errorMessages = new Dictionary<string, string>
            {
                {Dte, Resources.NullDte},
                {CmdHelper, Resources.NullCommandHelper}
            };
        }

        private void InitializedAddedObjects()
        {
            objectsEngineKnowsAbout = new Dictionary<string, object>
            {
                {Dte, Engine.DteObject},
                {CmdHelper, Engine.CommandHelper}
            };
        }

        public void GetDocVersionString(out string version)
        {
            version = null;
        }

        public void OnScriptTerminate(object result, System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo)
        {
        }

        public void OnStateChange(ScriptState scriptState)
        {
        }

        public static void ResetError()
        {
            RuntimeError = false;
            RuntimeException = null;

            InternalVsException = null;
        }

        private static string GetErrorDescription(System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo)
        {
            string description = exceptionInfo.bstrDescription;
            if (description.Equals("Object required"))
            {
                description = Resources.NoActiveDocumentErrorMessage;
            }
            else if (string.IsNullOrEmpty(description))
            {
                description = Resources.CommandNotValid;
            }
            return description;
        }

        public void OnScriptError(IActiveScriptError scriptError)
        {
            scriptError.GetSourcePosition(out _, out var lineNumber, out var column);
            scriptError.GetExceptionInfo(out var exceptionInfo);
            string source = exceptionInfo.bstrSource;
            string description = GetErrorDescription(exceptionInfo);

            RuntimeError = true;
            RuntimeException = new RuntimeException(description, source, lineNumber, column);
        }

        public void OnEnterScript()
        {
        }

        public void OnLeaveScript()
        {
        }
    }
}