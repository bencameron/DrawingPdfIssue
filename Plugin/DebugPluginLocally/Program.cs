using Inventor;
using System;
using System.IO;
using System.IO.Compression;
using Path = System.IO.Path;

namespace DebugPluginLocally
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (var inv = new InventorConnector())
            {
                InventorServer server = inv.GetInventorServer();

                try
                {
                    Console.WriteLine("Running locally...");
                    // run the plugin
                    DebugSamplePlugin(server);
                }
                catch (Exception e)
                {
                    string message = $"Exception: {e.Message}";
                    if (e.InnerException != null)
                        message += $"{System.Environment.NewLine}    Inner exception: {e.InnerException.Message}";

                    Console.WriteLine(message);
                }
                finally
                {
                    if (System.Diagnostics.Debugger.IsAttached)
                    {
                        Console.WriteLine("Press any key to exit. All documents will be closed.");
                        Console.ReadKey();
                    }
                }
            }
        }

        /// <summary>
        /// Opens box.ipt and runs SamplePlugin
        /// </summary>
        /// <param name="app"></param>
        private static void DebugSamplePlugin(InventorServer app)
        {
            var repoRootDir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName;
            var zipDir = Path.Combine(repoRootDir, "ZipFiles");
            var workDir = Path.Combine(repoRootDir, "work");
            var drawingsDir = Path.Combine(workDir, "Drawings");
            //var modelName = "-127-162765-1-Default Color Option";
            var modelName = "-142-163904-1-Default Color Option";

            // Clear contents of work directory
            try
            {
                if (Directory.Exists(workDir))
                {
                    Directory.Delete(workDir, true);
                }

                Directory.CreateDirectory(workDir);

                Directory.CreateDirectory(Path.Combine(workDir, "Pdfs"));
                System.Environment.CurrentDirectory = workDir;
            }
            catch (IOException)
            {
                Console.WriteLine("The specified file is in use. It might be open by Inventor");
                return;
            }

            // Extract model to work directory
            using (var zipArchive = ZipFile.Open(Path.Combine(zipDir, $"{modelName}_Models.zip"), ZipArchiveMode.Read))
            {
                zipArchive.ExtractToDirectory(workDir);
            }

            // Note that the second part of this path is the value used for "PathInZip"
            var modelPath = Path.Combine(workDir, $"Designs\\{modelName}.iam");

            // Extract drawings to Drawings subdirectory
            using (var zipArchive = ZipFile.Open(Path.Combine(zipDir, $"{modelName}_drawings.zip"), ZipArchiveMode.Read))
            {
                zipArchive.ExtractToDirectory(drawingsDir);
            }

            Inventor.NameValueMap map = app.TransientObjects.CreateNameValueMap();
            map.Add("Model", modelPath);

            // create an instance of PluginPlugin
            PluginPlugin.SampleAutomation plugin = new PluginPlugin.SampleAutomation(app);

            // run the plugin
            plugin.RunWithArguments(null, map);

        }
    }
}
