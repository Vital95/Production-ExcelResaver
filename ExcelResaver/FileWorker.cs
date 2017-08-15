using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelResaver
{
    public class FileWorker
    {
        #region File worker methods

        private List<string> extensions = new List<string>();

        public FileWorker()
        {
            extensions.Add(".xls");
            extensions.Add(".xlsx");
            extensions.Add(".xlsb");
        }

        /// <summary>
        /// Log to log.txt (No file - no logs)
        /// </summary>
        /// <param name="statusString"></param>
        public void Logs(string statusString, string startLog = "notFirstTime")
        {
            string time = DateTime.Now.ToString("MM/dd/yyyy h:mm tt");
            if (File.Exists(Directory.GetCurrentDirectory() + "\\log.txt"))
            {
                try
                {
                    TextWriter tw = new StreamWriter(Directory.GetCurrentDirectory() + "\\log.txt", true);
                    if (startLog != "notFirstTime")
                    {
                        tw.WriteLine(time);
                    }
                    tw.WriteLine(statusString);
                    tw.Close();
                }
                catch (Exception ex)
                {
                }
             }
         }

        /// <summary>
        /// Log to log.txt list of strings (No file - no logs)
        /// </summary>
        /// <param name="statusString"></param>
        public void Logs(List<string> statusString, string startLog = "notFirstTime")
        {
            string time = DateTime.Now.ToString("MM/dd/yyyy h:mm tt");
            if (File.Exists(Directory.GetCurrentDirectory() + "\\log.txt"))
            {
                try
                {
                    TextWriter tw = new StreamWriter(Directory.GetCurrentDirectory() + "\\log.txt", true);
                    if (startLog != "notFirstTime")
                    {
                        tw.WriteLine(time);
                    }
                    foreach (string s in statusString)
                    {
                        tw.WriteLine(s);
                    }
                    tw.Close();
                }
                catch (Exception ex)
                {
                }
            }
        }

        /// <summary>
        /// Gets all files in directory
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns></returns>
        public FileInfo[] GetAllFiesInDirectory(string folderPath)
        {
            DirectoryInfo d = new DirectoryInfo(folderPath);
            FileInfo[] infos = d.GetFiles();        
            return infos;
        }

        /// <summary>
        /// Return List of files with old file names
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        public List<FileInfo> SortFilesByNameAndExt(FileInfo[] files,string oldName)
        {
            FileWorker fw = new FileWorker();
            bool goodExt = false;
            List<FileInfo> newFiles = new List<FileInfo>();
            foreach (FileInfo f in files)
            {
                fw.Logs(f.Name);
                foreach (string s in extensions) 
                {
                    string tmp = Helper.GetExtension(f.Name);
                    if (tmp == s)
                    {
                        goodExt = true;
                        break;
                    }
                }
                if (Helper.HasFirstName(f.Name,oldName) && goodExt)
                {
                    newFiles.Add(f);
                }
                else
                {
                    goodExt = false;
                    continue;
                }

                goodExt = false;
            }
            return newFiles;
        }

        /// <summary>
        /// Return file names from list of files
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        public List<string> GetAllFileNames(List<FileInfo> files)
        {
            List<string> fileNames = new List<string>();
            foreach(FileInfo fileInfo in files){
                fileNames.Add(fileInfo.Name);
            }
            return fileNames;
        }

        /// <summary>
        /// Return all files paths from list of files
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        public List<string> GetAllFullFilePath(List<FileInfo> files)
        {
            List<string> filePath = new List<string>();
            foreach (FileInfo fileInfo in files)
            {
                filePath.Add(fileInfo.FullName.ToString());
            }
            return filePath;
        }

        /// <summary>
        /// Delete files by path
        /// </summary>
        /// <param name="filePath"></param>
        public void DeleteFiles(string filePath)
        {
            try
            {
                File.Delete(filePath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion
    }
}
