using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelResaver
{
    public static class IsFileLocked
    {
        /// <summary>
        /// Critical section logic
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool IsFileReallyLocked(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
