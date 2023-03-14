/////////////////////////////////////////////////////////////////////
// Copyright (c) Autodesk, Inc. All rights reserved
// Written by Forge Partner Development
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
/////////////////////////////////////////////////////////////////////

using Autodesk.Forge.DesignAutomation.Inventor.Utils;
using Autodesk.Forge.DesignAutomation.Inventor.Utils.Helpers;
using Inventor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace PluginPlugin
{
    [ComVisible(true)]
    public class SampleAutomation
    {
        private readonly InventorServer inventorApplication;

        public SampleAutomation(InventorServer inventorApp)
        {
            inventorApplication = inventorApp;
        }

        public void Run(Document doc)
        {
            RunWithArguments(doc, inventorApplication.TransientObjects.CreateNameValueMap());
        }

        public void RunWithArguments(Document doc, NameValueMap map)
        {
            try
            {
                using (new HeartBeat())
                {
                    LogTrace("Looking for model");
                    var modelPath = map.AsString("Model");

                    LogTrace($"Model path: {modelPath}");
                    var modelDir = System.IO.Path.GetDirectoryName(modelPath);
                    if (System.IO.File.Exists(modelPath))
                    {
                        LogTrace("Model exists");
                    }
                    else
                    {
                        LogError("Model DOES NOT exist");

                        LogTrace($"{modelDir} contents:");
                        foreach (var filePath in System.IO.Directory.EnumerateFiles(modelDir))
                        {
                            LogTrace(filePath);
                        }
                        return;
                    }

                    var dwgDir = ResolveLocalName("Drawings");
                    LogTrace($"Checking for drawings folder: {dwgDir}");
                    if (System.IO.Directory.Exists(dwgDir))
                    {
                        LogTrace("Found it");
                    }
                    else
                    {
                        LogError("It's not there");
                        return;
                    }
                 
                    foreach (var dwgPath in System.IO.Directory.EnumerateFiles(dwgDir))
                    {
                        var updatedDwgPath = System.IO.Path.Combine(modelDir, System.IO.Path.GetFileName(dwgPath));

                        LogTrace($"Moving {dwgPath} to {modelDir}");
                        System.IO.File.Move(dwgPath, updatedDwgPath);

                        LogTrace($"Opening {updatedDwgPath}");
                        var dwgDoc = inventorApplication.Documents.Open(updatedDwgPath);

                        LogTrace($"{updatedDwgPath} opened. Creating PDFs");
                        SaveAsPdf((DrawingDocument)dwgDoc);
                    }
                }
            }
            catch (Exception e)
            {
                LogError("Processing failed. " + e.ToString());
            }
        }

        private void SaveAsPdf(DrawingDocument doc)
        {
            var pdfDir = ResolveLocalName("Pdfs");
            var pdfAddIn = GetPdfAddIn();
            var context = inventorApplication.TransientObjects.CreateTranslationContext();
            var translatorOptions = inventorApplication.TransientObjects.CreateNameValueMap();
            var dataMedium = inventorApplication.TransientObjects.CreateDataMedium();

            context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

            // Check whether the translator has 'SaveCopyAs' options
            if (pdfAddIn.HasSaveCopyAsOptions[doc, context, translatorOptions])
            {
                for (int i = 1; i <= doc.Sheets.Count; i++)
                {
                    var pdfPath = System.IO.Path.Combine(pdfDir, $"{System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName)}_{i}.pdf");

                    //translatorOptions.Value["Sheet_Range"] = PrintRangeEnum.kPrintAllSheets;
                    translatorOptions.Value["Sheet_Range"] = PrintRangeEnum.kPrintSheetRange;
                    translatorOptions.Value["Custom_Begin_Sheet"] = i;
                    translatorOptions.Value["Custom_End_Sheet"] = i;
                    LogTrace("Set saveas options");

                    // Set the destination file name
                    dataMedium.FileName = pdfPath;

                    // Publish document
                    pdfAddIn.SaveCopyAs(doc, context, translatorOptions, dataMedium);
                }
            }
        }

        private TranslatorAddIn GetPdfAddIn()
        {
            var pdfAddinId = "{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}";
            var addIn = (TranslatorAddIn)inventorApplication.ApplicationAddIns.ItemById[pdfAddinId];

            if (addIn != null && (!addIn.Activated))
            {
                addIn.Activate();
            }

            return addIn;
        }

        #region Logging utilities

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string format, params object[] args)
        {
            Trace.TraceInformation(format, args);
        }

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string message)
        {
            Trace.TraceInformation(message);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string format, params object[] args)
        {
            Trace.TraceError(format, args);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string message)
        {
            Trace.TraceError(message);
        }

        private static string ResolveLocalName(string localName)
        {
            return System.IO.Path.Combine(System.Environment.CurrentDirectory, localName);
        }

        #endregion
    }
}