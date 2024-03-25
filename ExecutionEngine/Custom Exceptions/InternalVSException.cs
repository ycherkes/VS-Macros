//-----------------------------------------------------------------------
// <copyright file="InternalVSException.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace VSMacros.ExecutionEngine
{
    public class InternalVsException : Exception
    {
        public InternalVsException(string message, string source, string stackTrace, string targetSite)
        {
            Description = message;
            Source = source;
            StackTrace = stackTrace;
            TargetSite = targetSite;
        }

        public string Description { get; }
        public override string Source { get; set; }
        public new string StackTrace { get; }
        public new string TargetSite { get; }
    }
}
