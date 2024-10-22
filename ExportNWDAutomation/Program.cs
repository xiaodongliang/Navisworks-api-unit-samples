using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Navisworks.Api.Automation;
using Autodesk.Navisworks.Api.Resolver;

namespace ExportNWDAutomation
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // The "Navisworks.Autodesk.Resolver" assembly should be included in your installation.
            // Other Navisworks.Autodesk.*" assemblies will be resolved at runtime,
            // and thus should not be included and thus should not be included.
            // Autodesk.Navisworks.Controls controls usage inside the VS Designer is deprecated, and instead they should be created dynamically.
            String runtimeName = Resolver.TryBindToRuntime(RuntimeNames.Any);
            if (String.IsNullOrEmpty(runtimeName))
            {
                throw new Exception("Failed to bind to Navisworks runtime");
            }
            XMain(args);
        }

        static void XMain(string[] args)
        {
            NavisworksApplication navisworksApplication = null;
            try
            {
                //create NavisworksApplication automation object
                navisworksApplication = new NavisworksApplication();

                string workingDir = System.IO.Directory.GetCurrentDirectory().TrimEnd('\\');

                //disable progress whilst we do this procedure
                navisworksApplication.DisableProgress();

                //invisible model
                navisworksApplication.Visible = false;

                //Open two AutoCAD files
                navisworksApplication.OpenFile(workingDir + "\\hello.nwc", workingDir + "\\world.nwc");

                //Save the combination into a Navisworks file
                navisworksApplication.SaveFile(workingDir + "\\hello_world.nwd");


                //Call your plugin of NW for advanced features
                //such as manipulating models, properties, timeliner, clash etc..
                //navisworksApplication.ExecuteAddInPlugin("MyPlugin.ADSK", args);

                //Re-enable progress
                navisworksApplication.EnableProgress();



            }
            catch (InvalidOperationException e)
            {
                Console.WriteLine("InvalidOperationException: " + e.Message);
            }
            catch (ArgumentException e)
            {
                Console.WriteLine("ArgumentException: " + e.Message);
            }
            finally
            {
                //Close Navisworks
                if (navisworksApplication != null)
                {
                    navisworksApplication.Dispose();
                }
            }
        }
    }
}
