//-----------------------------------------------------------------------
// <copyright file="BinderHelper.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Reflection;
using System.Runtime.Serialization;

namespace VSMacros.ExecutionEngine.Pipes
{
    public class BinderHelper : SerializationBinder
    {
        public override Type BindToType(string assemblyName, string typeName)
        {
            string currentAssembly = Assembly.GetExecutingAssembly().FullName;
            return Type.GetType(string.Format("{0}, {1}", typeName, currentAssembly));
        }
    }
}