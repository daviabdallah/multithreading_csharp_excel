using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;


//Author: Davi Abdallah
//© Copyright 2014 Davi Abdallah

//Code Disclaimer Information

//This document contains programming examples.
//You are granted a nonexclusive copyright license to use all programming code examples from which you can generate similar function tailored to your own specific needs.

//All sample code is provided for illustrative purposes only. These examples have not been thoroughly tested under all conditions. Therefore there are not guarantees or implicit reliability, serviceability, or function of these programs.

//All programs contained here are provided to you "AS IS" without any warranties of any kind.
//The implied warranties of non-infringement, merchantability and fitness for a particular purpose are expressly disclaimed.

namespace MultithreadingSampleApp
{
    [ComVisible(true)]
    public class MultiThreadingRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private static Excel.Application _excelApp = null;
        private static Excel.Worksheet _activeSheet = null;
        private static Mutex _mut = new Mutex();

        public static void SetExcelApp(ref Excel.Application excelApp)
        {
            _excelApp = excelApp;
        }

        private static void SetActiveSheet()
        {
            if (_excelApp != null)
            {
            _activeSheet = (Excel.Worksheet)_excelApp.ActiveSheet;
            }
            else
            {
                throw new ArgumentException("ActiveSheet references could not be set (null app reference).");
            }
            if (_activeSheet == null)
            {
                throw new ArgumentException("ActiveSheet references could not be set. (null ActiveSheet reference)");
            }
        }

        private void RunSampleCode()
        {
            SetActiveSheet();

            int noOfRuns = 50;
            int noOfThreads = 5;
            int noOfIteractionsPerThread = 10;
            System.Windows.Forms.MessageBox.Show("Starting MultiThreaded write job, writing " + noOfRuns + " cells.");
            _activeSheet.Cells[1, 1] = "MultiThread Write BatchID";
            _activeSheet.Cells[1, 2] = "MultiThread Write IndexID";
            _activeSheet.Cells[1, 3] = "MultiThread Write IndexCalculation";
            _activeSheet.Cells[1, 4] = "MultiThread Write Search State";
            _activeSheet.Cells[1, 5] = "MultiThread Write BatchTime (Start/End)";

            _activeSheet.Cells[1, 6] = "SingleThread Read BatchID";
            _activeSheet.Cells[1, 7] = "SingleThread Read IndexID";
            _activeSheet.Cells[1, 8] = "SingleThread Write IndexCalculation";
            _activeSheet.Cells[1, 9] = "SingleThread Read Search State";
            _activeSheet.Cells[1, 10] = "SingleThread Read BatchTime (Start/End)";
            MultiThreadWriteJob(noOfRuns, noOfThreads, noOfIteractionsPerThread);
            long jobExecCount = 0;
            while (jobExecCount < noOfRuns)
            {
                if (_mut.WaitOne())
                {
                    // System.Windows.Forms.MessageBox.Show("Read job, jobExecCount:" + jobExecCount.ToString());
                    jobExecCount = ThreadData.GetJobExecWaitCount();
                    _mut.ReleaseMutex();
                }
            }
            if (_mut.WaitOne())
            {
                System.Windows.Forms.MessageBox.Show("Starting SingleThreaded read job, reading " + noOfRuns + " cells.");
                _mut.ReleaseMutex();
            }
            //reset and start read job
            jobExecCount = 0;
            ThreadData.ResetJobExecWaitCount();
            SingleThreadWriteJob(noOfRuns, noOfThreads, noOfIteractionsPerThread);
            _activeSheet.Cells[noOfRuns + 2, 5] = "=MIN(E2:E" + (noOfRuns + 1).ToString() + ")";
            _activeSheet.Cells[noOfRuns + 3, 5] = "=MAX(E2:E" + (noOfRuns + 1).ToString() + ")";
            _activeSheet.Cells[noOfRuns + 4, 4] = "Run Time in milliseconds:";
            _activeSheet.Range[_activeSheet.Cells[noOfRuns + 4, 4], _activeSheet.Cells[noOfRuns + 4, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            _activeSheet.Cells[noOfRuns + 4, 5] = "=(E" + (noOfRuns + 3).ToString() + "-E" + (noOfRuns + 2).ToString() + ")*24*60*60*1000";
            _activeSheet.Cells[noOfRuns + 5, 4] = "Run Time in seconds:";
            _activeSheet.Range[_activeSheet.Cells[noOfRuns + 5, 4], _activeSheet.Cells[noOfRuns + 5, 4]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            _activeSheet.Cells[noOfRuns + 5, 5] = "=(E" + (noOfRuns + 3).ToString() + "-E" + (noOfRuns + 2).ToString() + ")*24*60*60";

            _activeSheet.Cells[noOfRuns + 2, 10] = "=MIN(J2:J" + (noOfRuns + 1).ToString() + ")";
            _activeSheet.Cells[noOfRuns + 3, 10] = "=MAX(J2:J" + (noOfRuns + 1).ToString() + ")";
            _activeSheet.Cells[noOfRuns + 4, 9] = "Run Time in milliseconds:";
            _activeSheet.Range[_activeSheet.Cells[noOfRuns + 4, 9], _activeSheet.Cells[noOfRuns + 4, 9]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            _activeSheet.Cells[noOfRuns + 4, 10] = "=(J" + (noOfRuns + 3).ToString() + "-J" + (noOfRuns + 2).ToString() + ")*24*60*60*1000";
            _activeSheet.Cells[noOfRuns + 5, 9] = "Run Time in seconds:";
            _activeSheet.Range[_activeSheet.Cells[noOfRuns + 5, 9], _activeSheet.Cells[noOfRuns + 5, 9]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            _activeSheet.Cells[noOfRuns + 5, 10] = "=(J" + (noOfRuns + 3).ToString() + "-J" + (noOfRuns + 2).ToString() + ")*24*60*60";

            _activeSheet.Cells.Columns.AutoFit();
            _activeSheet.Cells.Rows.AutoFit();
            _activeSheet.Range[_activeSheet.Cells[1, 1], _activeSheet.Cells[1, 10]].EntireRow.Font.Bold = true;
            System.Windows.Forms.MessageBox.Show("All jobs completed.");
        }

        private void MultiThreadWriteJob(int runs, int noOfThreads, int noOfIteractionsPerThread)
        {
            ThreadData tData;
            ThreadPool.SetMaxThreads(noOfThreads, noOfThreads);
            int j = 0;
            for (int i = 0; i < runs; i = i + noOfIteractionsPerThread)
            {
                if (i + noOfIteractionsPerThread < runs)
                {
                    tData = new ThreadData(ref _mut, ref _activeSheet, i + 2, (long)i + 2 + noOfIteractionsPerThread, j);
                }
                else
                {
                    tData = new ThreadData(ref _mut, ref _activeSheet, i + 2, (long)2 + runs, j);
                }
                ThreadPool.QueueUserWorkItem(new WaitCallback(Worker1), tData);
                j++;
            }
        }

        private void SingleThreadWriteJob(int runs, int noOfThreads, int noOfIteractionsPerThread)
        {
            ThreadData tData;
            for (int i = 0; i < runs; i++)
            {
                if (i != runs - 1 && noOfIteractionsPerThread > 2)
                {
                    tData = new ThreadData(ref _mut, ref _activeSheet, i + 2, ((i + 2) % noOfIteractionsPerThread == 2) || ((i + 2) % noOfIteractionsPerThread == 1), 0);
                }
                else
                {
                    tData = new ThreadData(ref _mut, ref _activeSheet, i + 2, true, 0);
                }
                tData.SetIndex(i + 2);
                Worker0(tData);
            }
        }

        private void Worker0(object data)
        {
            ThreadData tData = (ThreadData)data;
            try
            {
                //for (int i = 0; i < 10000; i++)
                //{
                //    int j = (i * 100) / 10;
                //}
                tData.MutexWaitOne();
                string searchValue = tData.SetSingleValue();
                tData.Release();
                try
                {
                    System.Net.WebClient wb = new System.Net.WebClient();
                    wb.DownloadFile("http://www.google.com/search?q=" + searchValue, searchValue + "_SingleThread.html");
                }
                catch
                {
                }
            }
            catch (Exception e)
            {
                tData.Release();
                //System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }

        private void Worker1(object data)
        {
            ThreadData tData = (ThreadData)data;
            try
            {
                //for (int i = 0; i < 10000; i++)
                //{
                //    int j = (i * 100) / 10;
                //}
                tData.MutexWaitOne();
                string[] searchValues = tData.SetRangeValue();
                tData.Release();
                foreach (string s in searchValues)
                {
                    try
                    {
                        System.Net.WebClient wb = new System.Net.WebClient();
                        wb.DownloadFile("http://www.google.com/search?q=" + s, s + "_MultiThread.html");
                    }
                    catch
                    {
                    }
                }
            }
            catch (Exception e)
            {
                tData.Release();
                //System.Windows.Forms.MessageBox.Show(e.Message);
            }
        }

        public MultiThreadingRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MultithreadingSampleApp.MultiThreadingRibbon.xml");
        }

        #endregion

        public void MultiThreadingSampleApp(Office.IRibbonControl control)
        {
            //_activeSheet.Cells[1, 1] = "MultiThread Write BatchID";
            RunSampleCode();
        }

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
