//-----------------------------------------------------------------------
// <copyright file="Parser.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using ExecutionEngine.Enums;
using ExecutionEngine.Interfaces;
using System;

namespace ExecutionEngine
{
    internal sealed class Parser
    {
        private IActiveScriptParse32 parse32;
        private IActiveScriptParse64 parse64;
        private bool isParse32;

        public Parser(IActiveScript engine)
        {
            isParse32 = Is32BitEnvironment();
            InitializeParsers(engine);
        }

        private void InitializeParsers(IActiveScript engine)
        {
            if (isParse32)
            {
                parse32 = (IActiveScriptParse32)engine;
                parse32.InitNew();
            }
            else
            {
                parse64 = (IActiveScriptParse64)engine;
                parse64.InitNew();
            }
        }

        private static bool Is32BitEnvironment()
        {
            return IntPtr.Size == 4;
        }

        internal void Parse(string unparsed)
        {
            ScriptText flags = ScriptText.None;

            if (isParse32)
            {
                parse32.ParseScriptText(unparsed,
                    itemName: null,
                    context: null,
                    delimiter: null,
                    sourceContextCookie: IntPtr.Zero,
                    startingLineNumber: 0,
                    flags: flags,
                    result: out _,
                    exceptionInfo: out _);
            }
            else
            {
                parse64.ParseScriptText(unparsed,
                    itemName: null,
                    context: null,
                    delimiter: null,
                    sourceContextCookie: IntPtr.Zero,
                    startingLineNumber: 0,
                    flags: flags,
                    result: out _,
                    exceptionInfo: out _);
            }
        }
    }
}
