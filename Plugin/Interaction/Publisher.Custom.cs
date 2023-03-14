using Autodesk.Forge.DesignAutomation.Model;
using System;
using System.Collections.Generic;
using System.IO;

namespace Interaction
{
    /// <summary>
    /// Customizable part of Publisher class.
    /// </summary>
    internal partial class Publisher
    {
        /// <summary>
        /// Constants.
        /// </summary>
        private static class Constants
        {
            // Constants containing the name of the specific Inventor engines. Please note
            // that potentially not all engines are listed - new egines may have been added meanwhile.
            // Please use the interaction tools function - list all engines in order to find name of engine that is not listed here
            //private const string InventorEngine2018 = "Autodesk.Inventor+22"; // Inventor 2018
            //private const string InventorEngine2019 = "Autodesk.Inventor+23"; // Inventor 2019
            //private const string InventorEngine2020 = "Autodesk.Inventor+24"; // Inventor 2020
            //private const string InventorEngine2021 = "Autodesk.Inventor+2021"; // Inventor 2021
            //private const string InventorEngine2022 = "Autodesk.Inventor+2022"; // Inventor 2022
            private const string InventorEngine2023 = "Autodesk.Inventor+2023"; // Inventor 2022

            public static readonly string Engine = InventorEngine2023;

            public const string Description = "PUT DESCRIPTION HERE";

            internal static class Bundle
            {
                public static readonly string Id = "Plugin";
                public const string Label = "alpha";

                public static readonly AppBundle Definition = new AppBundle
                {
                    Engine = Engine,
                    Id = Id,
                    Description = Description
                };
            }

            internal static class Activity
            {
                public static readonly string Id = Bundle.Id;
                public const string Label = Bundle.Label;
            }

            internal static class Parameters
            {
                public const string Model = nameof(Model);
                public const string DrawingsZip = nameof(DrawingsZip);
                public const string PdfsZip = nameof(PdfsZip);
            }
        }


        /// <summary>
        /// Get command line for activity.
        /// </summary>
        private static List<string> GetActivityCommandLine()
        {
            //return new List<string> { $"$(engine.path)\\InventorCoreConsole.exe /al \"$(appbundles[{Constants.Activity.Id}].path)\" /i \"$(args[{Constants.Parameters.DrawingsZip}].path)\"" };
            return new List<string> { $"$(engine.path)\\InventorCoreConsole.exe /al \"$(appbundles[{Constants.Activity.Id}].path)\" /Model \"$(args[Model].path)\"" };
        }

        /// <summary>
        /// Get activity parameters.
        /// </summary>
        private static Dictionary<string, Parameter> GetActivityParams()
        {
            return new Dictionary<string, Parameter>
                    {
                        {
                            Constants.Parameters.Model,
                            new Parameter
                            {
                                Zip = true,
                                Verb = Verb.Get,
                                Description = "Input model"
                            }
                        },
                        {
                            Constants.Parameters.DrawingsZip,
                            new Parameter
                            {
                                Zip = true,
                                Verb = Verb.Get,
                                Description = "Zip full of drawings"
                            }
                        },
                        {
                            Constants.Parameters.PdfsZip,
                            new Parameter
                            {
                                Verb = Verb.Put,
                                LocalName = "Pdfs",
                                Zip = true,
                                Description = "Zip full of pdfs",
                                Ondemand = false,
                                Required = true
                            }
                        }
                    };
        }

        /// <summary>
        /// Get arguments for workitem.
        /// </summary>
        private static Dictionary<string, IArgument> GetWorkItemArgs(string bucketName, string token)
        {
            var ossPath = $"https://developer.api.autodesk.com/oss/v2/buckets/{bucketName}/objects/";
            //var modelName = "-127-162765-1-Default Color Option";
            var modelName = "-142-163904-1-Default Color Option";
            var modelInput = $"{modelName}_Models.zip";
            var dwgInput = $"{modelName}_drawings.zip";
            var pdfOutput = "Pdfs.zip";

            return new Dictionary<string, IArgument>
                    {
                        {
                            Constants.Parameters.Model,
                            new XrefTreeArgument
                            {
                                PathInZip = $"Designs\\{modelName}.iam",
                                Url = $"{ossPath}{modelInput}",
                                Headers = new Dictionary<string, string>
                                {
                                    {
                                        "Authorization",
                                        "Bearer " + token
                                    }
                                }
                            }
                        },
                        {
                            Constants.Parameters.DrawingsZip,
                            new XrefTreeArgument
                            {
                                LocalName = "Drawings",
                                Url = $"{ossPath}{dwgInput}",
                                Headers = new Dictionary<string, string>
                                {
                                    {
                                        "Authorization",
                                        "Bearer " + token
                                    }
                                }
                            }
                        },
                        {
                            Constants.Parameters.PdfsZip,
                            new XrefTreeArgument
                            {
                                LocalName = "Pdfs",
                                Verb = Verb.Put,
                                Url = $"{ossPath}{pdfOutput}",
                                Headers = new Dictionary<string, string>
                                {
                                    {
                                        "Authorization",
                                        "Bearer " + token
                                    }
                                }
                            }
                        }
                    };
        }
    }
}
