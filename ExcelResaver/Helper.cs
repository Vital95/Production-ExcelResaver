using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace ExcelResaver
{
    public static class Helper
    {
        #region Helper methods

        /// <summary>
        /// Replace old name by new name
        /// </summary>
        /// <param name="mainValue"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public static string Replace(string mainValue, string oldValue, string newValue)
        {
            Regex MyRegEx = new Regex(oldValue);
            string result = MyRegEx.Replace(mainValue, newValue, 1);
            return result;
        }

        /// <summary>
        /// Return file name without extansion
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string SplitExstansionByDotAndTakeName(string input)
        {
            int fileExtPos = input.LastIndexOf(".");
            string output = input.Substring(0,fileExtPos);
            return output;  
        }

        /// <summary>
        /// Return extension of a file
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string GetExtension(string fileName)
        {
            return Path.GetExtension(fileName);
        }

        /// <summary>
        /// Exstract data from console args
        /// </summary>
        /// <param name="args"></param>
        /// <param name="folderPath"></param>
        /// <param name="oldName"></param>
        /// <param name="newName"></param>
        public static void ExstractDataFromArgs(string[] args, out string folderPath, out string oldName, out string newName, out string exstansion)
        {
            folderPath = args[0].Remove(args[0].Length - 1);
            oldName = args[1];
            newName = args[2];
            exstansion = args[3];
        }

        /// <summary>
        /// Convert extension to byte value
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static byte ExtensionToByte(string input)
        {
            if (input == "xls")
            {
                return 1;
            }
            else
            {
                return 2;
            }
        }

        /// <summary>
        /// Check if file need to be resaved
        /// </summary>
        /// <param name="oldName"></param>
        /// <returns></returns>
        public static bool HasFirstName(string fileName, string oldName)
        {
            if (fileName.Length >= oldName.Length)
            {
                string subString = fileName.Substring(0, oldName.Length);

                if (subString == oldName)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Shifts files for deleting by files wich already exists
        /// </summary>
        /// <returns></returns>
        public static List<string> SiftSameFileNames(List<string> deletionList, List<string> sameFilesList)
        {
            List<string> output = new List<string>();
            foreach(string s in sameFilesList)
            {
                foreach (string fileForDeletion in deletionList)
                {
                    if (fileForDeletion == s)
                    {
                        output.Add(s);
                        break;
                    }
                }
            }
            if (output.Count != 0)
            {
                foreach (string outputFiles in output)
                {
                    deletionList.Remove(outputFiles);
                }
            }
            return deletionList;
        }

        /// <summary>
        /// Splits input string by space 
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string[] SplitBySpaceBar(string input)
        {
            char[] whitespace = new char[] { ' ' };
            return input.Split(whitespace);
        }

        /// <summary>
        /// Check if input string contain input sub string
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="newFileName"></param>
        /// <returns></returns>
        public static bool IsStringContains(string inputString, string subString)
        {
            return inputString.Contains(subString);
        }

        /// <summary>
        /// Replace all oldValues by newValues
        /// </summary>
        /// <param name="mainString"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        /// <returns></returns>
        public static string ReplaceSubStringByString(string mainString, string oldValue, string newValue )
        {
            return mainString.Replace(oldValue,newValue);
        }

        /// <summary>
        /// Validate files after convert and delete lost ones
        /// </summary>
        /// <param name="listOfFiles"></param>
        /// <param name="extension"></param>
        public static void PosValidateExpectedFiles(List<string> listOfFiles, string extension = "xlsb")
        {
            FileWorker fw = new FileWorker();
            List<string> expectedFiles = new List<string>();

            foreach(string filePath in listOfFiles)
            {
                string fileWithoutExt = Helper.SplitExstansionByDotAndTakeName(filePath);
                expectedFiles.Add(fileWithoutExt + "."+extension);
            }

            int isExistsCount = 0;
            foreach(string newFilePath in expectedFiles)
            {
                if (File.Exists(newFilePath))
                {
                    isExistsCount++;
                }
            }
            if (isExistsCount == listOfFiles.Count)
            {
                foreach (string oldFilePath in listOfFiles)
                {
                    if (File.Exists(oldFilePath))
                    {
                        try
                        {
                            fw.DeleteFiles(oldFilePath);
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                    else
                    {
                        continue;
                    }
                }
            }

        }

        #endregion

        #region code not in use

        /// <summary>
        /// Puts txt file to .exe directory
        /// </summary>
        /// <param name="dirPath"></param>
        /// <param name="input"></param>
        public static void WriteDataToFile(string input, string fileName)
        {
            System.IO.StreamWriter textFile = new System.IO.StreamWriter(Directory.GetCurrentDirectory() + "\\" + fileName + ".txt");
            textFile.WriteLine(input);
            textFile.Close();
        }

        /// <summary>
        /// Checks if input data is valid
        /// </summary>
        /// <returns></returns>
        public static bool CheckArgumentCount(string[] args, int neededCount)
        {
            int counter = 0;
            foreach (string s in args)
            {
                counter++;
            }
            if (counter == neededCount)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        #endregion
    }
}
