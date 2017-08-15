using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.IO;
using Marshal = System.Runtime.InteropServices.Marshal;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelResaver
{
    public class Controler
    {
        #region Private Values

        private string folderPath;
        private string oldName;
        private string newName;
        private string extension;
        private string statusString = "OK";
        private Exception level2Exception = null;
        private List<string> errorsMessages = new List<string>();
            
        #endregion

        /// <summary>
        /// Resave files with old name, changing it to new name
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="oldName"></param>
        /// <param name="newName"></param>
        public void ResaveFilesInFolder(string[] args)
        {
            Helper.ExstractDataFromArgs(args, out folderPath, out oldName, out newName, out extension);
            PlusReplaceCompleted();
            byte excelExtension = Helper.ExtensionToByte(extension);
            FileWorker fileWorker = new FileWorker();
            fileWorker.Logs("", "1");
            List<string> fullPath = new List<string>();
            fullPath = fileWorker.GetAllFullFilePath(fileWorker.SortFilesByNameAndExt(fileWorker.GetAllFiesInDirectory(folderPath), oldName));

            List<string> fileNames = new List<string>();
            fileNames = fileWorker.GetAllFileNames(fileWorker.SortFilesByNameAndExt(fileWorker.GetAllFiesInDirectory(folderPath), oldName));

            List<string> filesForDeletion = new List<string>();
            List<string> sameFiles = new List<string>();

            if (fullPath.Count != 0)
            {
                int q = 0;
                bool flag = false;
               
                foreach (string s in fullPath) 
                {
                    fileWorker.Logs(s);
                    ExcelWorker excelWorker = new ExcelWorker();
                    try
                    {                
                        if (!excelWorker.ResaveExcelFile(folderPath, s, Helper.Replace(fileNames[q], oldName, newName), excelExtension, oldName))
                        {
                            q++;
                            flag = true;
                            filesForDeletion.Add(s);
                            excelWorker.GetLastErrors(ref errorsMessages);
                            if (errorsMessages.Count != 0)
                            {
                                fileWorker.Logs(errorsMessages);
                            }
                            GColectorStart();
                        }
                        else
                        {
                            excelWorker.GetLastErrors(ref errorsMessages);
                            if (errorsMessages.Count != 0)
                            {
                                fileWorker.Logs(errorsMessages);
                            }
                            GColectorStart();
                        }

                        GColectorStart();
                    }
                    catch (Exception ex)
                    {
                        level2Exception = ex;
                        fileWorker.Logs("Level 2 Error" + level2Exception.Message);
                        GColectorStart();
                    }
                    if (flag != true)
                    {
                        q++;
                        GColectorStart();
                    }
                    flag = false;
                    GColectorStart();
                }

                DeleteFiles(filesForDeletion);
                fileWorker.Logs(statusString);
                fileWorker.Logs("\n");
            }
        }

        /// <summary>
        /// Replace +++ by space bar if we have +++ in file name
        /// </summary>
        /// <returns></returns>
        private void PlusReplaceCompleted()
        {
            string plus = "+++";
            string space = " ";
            if (Helper.IsStringContains(oldName, plus))
            {
                oldName = Helper.ReplaceSubStringByString(oldName, plus, space);
                if (Helper.IsStringContains(newName, plus))
                {
                    newName = Helper.ReplaceSubStringByString(newName, plus, space);
                }
            }
            else
            {
                if (Helper.IsStringContains(newName, plus))
                {
                    newName = Helper.ReplaceSubStringByString(newName, plus, space);
                }
            }
        }

        /// <summary>
        /// Delete files in queue 
        /// </summary>
        /// <param name="fileList"></param>
        private void DeleteFiles(List<string> fileList)
        {
            Queue myQueue = new Queue();
            int maxTryCount = 200;
            int currentTryCount = 0;
            FileWorker fileWorker = new FileWorker();
            
            foreach (string s in fileList)
            {
                myQueue.Enqueue(s);
            }

            while (myQueue.Count != 0 && maxTryCount > currentTryCount)
            {
                string s = myQueue.Dequeue().ToString();
                currentTryCount++;
                try
                {
                    System.Threading.Thread.Sleep(20);
                    fileWorker.DeleteFiles(s);
                    System.Threading.Thread.Sleep(20);
                }
                catch (Exception ex)
                {
                    myQueue.Enqueue(s);
                    fileWorker.Logs("Level 2 Error" + ex.Message);
                }
            }
        }

        /// <summary>
        /// Collect COM objects
        /// </summary>
        private void GColectorStart()
        {
            int collectorTryCount = 0;
            do
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                collectorTryCount++;
            }
            while (Marshal.AreComObjectsAvailableForCleanup() && collectorTryCount < 20);           
        }
    }
}
