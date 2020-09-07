using System.Reflection;

namespace ExcelToPdf
{
    public static partial class AssemblyLoader
    {
        public static void LoadAssembly()
        {
            var assPath = AssemblyPath.GetAssemblyPath();
            foreach (var assemblyName in AssemblyNames)
                Assembly.LoadFrom(assPath + "\\" + assemblyName);
        }
    }
}
