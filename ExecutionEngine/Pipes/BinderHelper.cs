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
            Type toDeserialize = Type.GetType(string.Format("{0}, {1}", typeName, currentAssembly));

            return toDeserialize;
        }
    }
}
