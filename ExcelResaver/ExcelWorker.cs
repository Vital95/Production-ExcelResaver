using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Marshal = System.Runtime.InteropServices.Marshal;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace ExcelResaver
{
    public class ExcelWorker
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        #region private fields

        private Microsoft.Office.Interop.Excel.Application xlApp;
        private Workbook xlWorkbook;
        private Exception currentExeption = null;
        private byte currentTryCount;
        private Exception lastException = null;
        private int hWnd = 0;
        private uint pid = 0;
        private List<string> errorsMessages = new List<string>(); 

        #endregion

        #region Excel worker methods

        /// <summary>
        /// Resave in correct extension and return true if file was saved successfully
        /// </summary>
        public bool ResaveExcelFile(string folderPath, string filePath, string newFileName, byte extension, string oldName)
        { 
            bool flag = false;
            currentTryCount = 0;
            FileWorker fileWorker = new FileWorker();

            bool retry = true;
            byte maxTryCount = 20;
            while (retry && maxTryCount > currentTryCount)
            {
                try
                {
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    if (hWnd == 0)
                    {
                        try
                        {
                            hWnd = xlApp.Hwnd;
                        }
                        catch (Exception ex)
                        {
                            errorsMessages.Add(ex.Message);
                            hWnd = 0;
                        }
                    }
                    retry = false;
                    currentExeption = null;
                }
                catch (Exception ex)
                {
                    currentTryCount++;
                    currentExeption = ex;
                    errorsMessages.Add(ex.Message);
                    if (hWnd == 0)
                    {
                        try
                        {
                            hWnd = xlApp.Hwnd;
                        }
                        catch (Exception ex1)
                        {
                            errorsMessages.Add(ex1.Message);
                            hWnd = 0;
                        }
                    }
                    lastException = ex;
                }
            }

            currentTryCount = 0;
            retry = true;

            if (currentExeption == null)
            {
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;
                xlApp.UserControl = false;
                
                while (retry && maxTryCount > currentTryCount)
                {
                    try
                    {
                        xlWorkbook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", true, XlPlatform.xlWindows, "", false, false, 0, false, 1, 0);
                        retry = false;
                        currentExeption = null;
                    }
                    catch (Exception ex)
                    {
                        errorsMessages.Add(ex.Message);
                        currentTryCount++;
                        currentExeption = ex;
                        lastException = ex;
                    }
                }

                currentTryCount = 0;
                retry = true;

                if (currentExeption == null)
                {
                    newFileName = Helper.SplitExstansionByDotAndTakeName(newFileName);
                    Microsoft.Office.Interop.Excel.XlFileFormat excelExstansion = PickExstansion(extension);
                    string oldExtensionInString = Helper.GetExtension(filePath);
                    string fullFilePathAndName = folderPath + @"\" + newFileName;

                    while (retry && maxTryCount > currentTryCount)
                    {
                        try
                        {
                            xlWorkbook.SaveAs(fullFilePathAndName, excelExstansion, Type.Missing, Type.Missing,
                                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            retry = false;
                            currentExeption = null;
                        }
                        catch(Exception ex)
                        {
                            errorsMessages.Add(ex.Message);
                            currentTryCount++;
                            currentExeption = ex;
                            lastException = ex;
                        }
                    }

                    currentTryCount = 0;
                    retry = true;

                    if (currentExeption == null)
                    {
                        DisposeCOMObjects();
                        if (IsCurrentExtensionCoincidence(oldExtensionInString, excelExstansion) && Helper.IsStringContains(filePath, newFileName))
                        {
                            flag = true;
                        }
                    }
                    else
                    {
                        if (hWnd == 0)
                        {
                            
                            try
                            {
                                hWnd = xlApp.Hwnd;
                            }
                            catch (Exception ex)
                            {
                                hWnd = 0;
                            }
                        }
                        DisposeCOMObjects();
                        return flag;
                    }
                }
                else
                {
                    if (hWnd == 0)
                    {
                        try
                        {
                            hWnd = xlApp.Hwnd;
                        }
                        catch (Exception ex)
                        {
                            errorsMessages.Add(ex.Message);
                            hWnd = 0;
                        }
                    }
                    DisposeCOMObjects();
                    return flag;
                }
            }
            else
            {
                if (hWnd == 0)
                {
                    try
                    {
                        hWnd = xlApp.Hwnd;
                    }
                    catch (Exception ex)
                    {
                        errorsMessages.Add(ex.Message);
                        hWnd = 0;
                    }
                }
                DisposeCOMObjects();
                return flag;
            }
            if (hWnd == 0)
            {
                try
                {
                    hWnd = xlApp.Hwnd;
                }
                catch (Exception ex)
                {
                    errorsMessages.Add(ex.Message);
                    hWnd = 0;
                }
            }
            DisposeCOMObjects();

            return flag;
        }

        /// <summary>
        /// Return needed extension
        /// </summary>
        /// <param name="extension"></param>
        /// <returns></returns>
        private Microsoft.Office.Interop.Excel.XlFileFormat PickExstansion(byte extension)
        {
            if (extension == 1)
            {
                return XlFileFormat.xlWorkbookNormal;
            }
            else
            {
                return XlFileFormat.xlOpenXMLWorkbook;
            }
        }

        /// <summary>
        /// Return true if extension of saved files is the same
        /// </summary>
        /// <param name="oldExtensionInString"></param>
        /// <returns></returns>
        private bool IsCurrentExtensionCoincidence(string oldExtensionInString, Microsoft.Office.Interop.Excel.XlFileFormat extension)
        {
            if ((oldExtensionInString == ".xlsx" && extension == XlFileFormat.xlOpenXMLWorkbook) || (oldExtensionInString == ".xls" && extension == XlFileFormat.xlWorkbookNormal))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Dispose COM objects
        /// </summary>
        private void DisposeCOMObjects()
        {
            FileWorker fileWorker = new FileWorker();
            try
            {
                xlWorkbook.Close(false);
            }
            catch (Exception ex)
            {
                errorsMessages.Add(ex.Message);
            }

            try
            {
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                errorsMessages.Add(ex.Message);
                if (ex.Message == "Exception from HRESULT: 0x800AC472")
                {
                    GetWindowThreadProcessId((IntPtr)hWnd, out pid);
                    if (hWnd != 0)
                    {
                        try
                        {
                            Process[] procs = Process.GetProcessesByName("EXCEL");
                            if (pid != 0)
                            {
                                foreach (Process p in procs)
                                {
                                    if (p.Id == pid)
                                    {
                                        fileWorker.Logs(p.Id.ToString());
                                        p.Kill();
                                    }
                                }
                            }
                        }
                        catch (Exception ex1)
                        {
                            errorsMessages.Add(ex1.Message);
                        }
                    }
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            try
            {
                Marshal.ReleaseComObject(xlWorkbook);
            }
            catch (Exception ex)
            {
                errorsMessages.Add(ex.Message);
            }
            try
            {
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                errorsMessages.Add(ex.Message);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Indicate last error message
        /// </summary>
        /// <param name="error"></param>
        /// <returns></returns>
        public bool GetLastErrors(ref List<string> error)
        {
            error = errorsMessages;
            return true;
        }
        #endregion  
    }
}
