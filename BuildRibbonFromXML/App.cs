#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.Versioning;
using System.Windows.Markup;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

#endregion

namespace BuildRibbonFromXML
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            //// Create ribbon
            //string assembly_Dll_File = @"C:\Directory\Path\Assembly.dll";
            //string xml_Ribbon_File_Path = @"C:\Directory\Path";

            //or

            // Get the path of the executing assembly (DLL)
            string assembly_Dll_File = Assembly.GetExecutingAssembly().Location;

            // Get the directory path of the assembly
            string xml_Ribbon_File_Path = Path.GetDirectoryName(assembly_Dll_File);

            // call the RibbonBuilder Class and pass it the path of the dll and *.ribbon path.
            RibbonBuilder.build_ribbon(application, assembly_Dll_File, xml_Ribbon_File_Path);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }


    }
}
