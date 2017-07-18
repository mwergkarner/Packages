//
// Copyright © 2017 Ranorex All rights reserved
//

using System;
using System.IO;
using System.Linq;
using System.Text;
using Ranorex.Core.Testing;

namespace Ranorex.AutomationHelpers.UserCodeCollections
{
    /// <summary>
    /// A collection of useful file methods.
    /// </summary>
    [UserCodeCollection]
    public class FileLibrary
    {
        /// <summary>
        /// Creates a log file containing a custom test in the output folder.
        /// </summary>
        /// <param name="text">Text that the log file should contain</param>
        /// <param name="filenamePrefix">Prefix used for the log filename</param>
        /// <param name="fileExtension">Extension of the log file</param>
        [UserCodeMethod]
        public static void WriteToFile(string text, string filenamePrefix, string fileExtension)
        {
            System.DateTime now = System.DateTime.Now;
            string strTimestamp = now.ToString("yyyyMMdd_HHmmss");
            string filename = filenamePrefix + "_" + strTimestamp + "." + fileExtension;
            Report.Info(filename);

            try
            {
                //Create the File
                using (FileStream fs = File.Create(filename))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes(text);
                    fs.Write(info, 0, info.Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            filenamePrefix = filename;
        }

        /// <summary>
        /// Checks if files in a directory exist.
        /// </summary>
        /// <param name="path">The relative or absolute path to search for the files</param>
        /// <param name="pattern">The pattern to search for in the filename</param>
        /// <param name="expectedCount">Number of expected files to be found</param>
        /// <param name="timeout">Defines the search timeout in seconds</param>
        [UserCodeMethod]
        public static void CheckFilesExist(string path, string pattern, int expectedCount, int timeout)
        {
            string[] listofFiles = System.IO.Directory.GetFiles(path, pattern);
            System.DateTime start = System.DateTime.Now;

            while (listofFiles.Length != expectedCount && System.DateTime.Now < start.AddSeconds(timeout))
            {
                listofFiles = System.IO.Directory.GetFiles(path, pattern);
            }

            Ranorex.Report.Info("Check if '" + expectedCount + "' file(s) with pattern '" + pattern + "' exist in the directory '" + path + "'. Search time " + timeout + " seconds.");
            Validate.AreEqual(listofFiles.Length, expectedCount);
        }

        /// <summary>
        /// Deletes files.
        /// </summary>
        /// <param name="path">The relative or absolute path to search for the files</param>
        /// <param name="pattern">The pattern to search for in the filename</param>
        [UserCodeMethod]
        public static void DeleteFiles(string path, string pattern)
        {
            string[] listofFiles = System.IO.Directory.GetFiles(path, pattern);

            if (listofFiles.Length == 0)
            {
                Ranorex.Report.Warn("No files have been found in '" + path + "' with the pattern '" + pattern + "'.");
            }

            foreach (string file in listofFiles)
            {
                try
                {
                    System.IO.File.Delete(file);
                    Ranorex.Report.Info("File has been deleted: " + file.ToString());
                }
                catch (Exception ex)
                {
                    Ranorex.Report.Error(ex.Message);
                }
            }
        }

        /// <summary>
        /// Repeatedly checks if files in a directory exist.
        /// </summary>
        /// <param name="path">>The relative or absolute path to search for the files</param>
        /// <param name="pattern">The pattern to search for in the filename</param>
        /// <param name="duration">Defines the search timeout in milliseconds</param>
        /// <param name="intervall">Sets the interval in ms at which the files are checked for the pattern</param>
        [UserCodeMethod]
        public static void WaitForFile(string path, string pattern, int duration, int intervall)
        {
            path = getPathForFile(path);
            bool bFound = Directory.GetFiles(path, pattern).Length > 0;
            System.DateTime start = System.DateTime.Now;

            while (!bFound && (System.DateTime.Now < start + new Duration(duration)))
            {

                bFound = Directory.GetFiles(path, pattern).Length > 0;

                if (bFound)
                {
                    break;
                }

                Delay.Duration(intervall, false);
            }

            if (bFound)
            {
                Ranorex.Report.Success("Validation", "File with pattern '" + pattern + "' was found in directory '" + path + "'.");
            }
            else
            {
                Ranorex.Report.Failure("Validation", "File with pattern '" + pattern + "' wasn't found in directory '" + path + "'.");
            }
        }

        /// <summary>
        /// Performs the playback of actions in this module.
        /// </summary>
        [UserCodeMethod]
        public static void WordCompare()
        {
            // TODO: Implement/specify
            /*
public static void FILE_compareWithWord(string file1, string file2, string unmask, string StartFrom, string CompareTo) {
    iDeletions=0;
    iInsertions=0;

    file1 = getPathForFile(file1);
    file2 = getPathForFile(file2);


    string sFile1Temp=Path.GetDirectoryName(file1)+@"\"+Path.GetFileNameWithoutExtension(file1)+".temp."+Path.GetExtension(file1);
    //string sFile2Temp=Path.GetDirectoryName(file2)+@"\"+Path.GetFileNameWithoutExtension(file2)+".temp."+Path.GetExtension(file2);


    string sFileContent = File.ReadAllText(file1);
    if (unmask.Length>0) {
        sFileContent = Regex.Replace(sFileContent, unmask, "XXXX");
    }
    File.WriteAllText(sFile1Temp, sFileContent);

    Application app = new Application();
    app.Visible = false;

    Document objDocument1 = app.Documents.Open(sFile1Temp);
    Document objDocument2 = app.Documents.Open(file2);

    app.CompareDocuments(objDocument1,objDocument2 , WdCompareDestination.wdCompareDestinationNew, WdGranularity.wdGranularityWordLevel, false, true, true, true, true, true, true, true, true, true, "Ranorex compare", false);

    objDocument1.Close();
    objDocument2.Close();


    int iStartFrom;
    if (!Int32.TryParse(StartFrom, out iStartFrom)) {
        iStartFrom=-1;
    }


    int pageNumber=9999;

    if (CompareTo.Length>0) {
        Range range= app.ActiveDocument.Content;
        range.Find.Execute(CompareTo,false, true, false, false, false, true, 0, false, "",0);
        if (range.Text.Equals(CompareTo)) {

            object oPageNumber=range.Information[WdInformation.wdActiveEndAdjustedPageNumber];
            pageNumber= Int32.Parse(oPageNumber.ToString());
        }
    }


    FILE_count (app.ActiveDocument.Revisions, unmask,iStartFrom, pageNumber, true);
    Section section = app.ActiveDocument.Sections.First;

    foreach (Microsoft.Office.Interop.Word.HeaderFooter aHeaderFooter in section.Headers) {
        FILE_count(aHeaderFooter.Range.Revisions, unmask, iStartFrom, pageNumber, false);
    }

    foreach (Microsoft.Office.Interop.Word.HeaderFooter aHeaderFooter in section.Headers) {
        FILE_count(aHeaderFooter.Range.Revisions, unmask, iStartFrom, pageNumber, false);
    }

    if (File.Exists(sFile1Temp)) {
        File.Delete(sFile1Temp);
    }

    if (iInsertions>0 || iDeletions>0) {
        string sResultFile=Path.GetDirectoryName(file1)+@"\"+Path.GetFileNameWithoutExtension(file1)+".difference.docx";
        Ranorex.Report.Failure("Differences found in comparison between '"+file1+"' and '"+file2+"'. Found "+iDeletions+" deletions and "+iInsertions+" insertions. Diff file: '"+sResultFile+"'");
        app.ActiveDocument.TrackRevisions=false;
        app.ActiveDocument.ShowRevisions=false;
        app.ActiveDocument.PrintRevisions=false;

        app.ActiveDocument.Protect(WdProtectionType.wdAllowOnlyReading);
        app.ActiveDocument.SaveAs2(sResultFile);
    } else {
        Ranorex.Report.Success("No differences found in comparison between '"+file1+"' and '"+file2+"'.");
    }

    app.ActiveDocument.Close(false);
    app.Quit();
}


private static void FILE_count (Revisions revisions, string unmask, int startFrom, int pagenumber, bool content) {

    foreach (Revision r in revisions) {
        bool accepted=false;
        string sRangeText=r.Range.Text;
        if (r.Type== WdRevisionType.wdRevisionDelete) {
            object oPageNumber=r.Range.Information[WdInformation.wdActiveEndAdjustedPageNumber];
            int iActualPageNumber= Int32.Parse(oPageNumber.ToString());



            if ((iActualPageNumber<=startFrom || iActualPageNumber>=pagenumber) && content) {
                r.Range.HighlightColorIndex=WdColorIndex.wdGray50;
                r.Accept();
                accepted=true;
            } else if ((!Regex.IsMatch(sRangeText, ".*XXXX.*")) || unmask.Length==0) {
                iDeletions++;
                Ranorex.Report.Info("Deletion", sRangeText);
            } else {
                r.Accept();
                accepted=true;
            }
        }
        if (!accepted) {
            if (r.Type== WdRevisionType.wdRevisionInsert) {
                object oPageNumber=r.Range.Information[WdInformation.wdActiveEndAdjustedPageNumber];
                int iActualPageNumber= Int32.Parse(oPageNumber.ToString());

                if ((iActualPageNumber<=startFrom || iActualPageNumber>=pagenumber) && content) {
                    r.Range.HighlightColorIndex=WdColorIndex.wdGray50;
                    r.Accept();
                    accepted=true;
                } else if (!Regex.IsMatch(sRangeText, unmask) || unmask.Length==0) {
                    iInsertions++;
                    Ranorex.Report.Info("Insertion", sRangeText);
                } else {
                    r.Range.HighlightColorIndex=WdColorIndex.wdGray50;
                    r.Accept();
                    accepted=true;
                }
            }
        }
    }
}
*/
            throw new NotImplementedException();
        }

        private static string getPathForFile(string path)
        {
            return path.StartsWith(".") ? Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, path)) : path;
        }
    }
}
