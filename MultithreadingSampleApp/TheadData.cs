using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace MultithreadingSampleApp
{
    public class ThreadData
    {
        private static Mutex _mut = null;
        private static Excel.Worksheet _activeWorksheet = null;
        private static int _jobExecWaitCount = 0;
        private long _index;
        private long _startIndex;
        private long _endIndex;
        private bool _setTS;
        private int _threadId;

        public ThreadData(ref Mutex mut, ref Excel.Worksheet activeWorksheet, long index, bool setTS, int threadId)
        {
            _mut = mut;
            _activeWorksheet = activeWorksheet;
            _index = index;
            _setTS = setTS;
            _threadId = threadId;
        }

        public ThreadData(ref Mutex mut, ref Excel.Worksheet activeWorksheet, long startIndex, long endIndex, int threadId)
        {
            _mut = mut;
            _activeWorksheet = activeWorksheet;
            _startIndex = startIndex;
            _endIndex = endIndex;
            _threadId = threadId;
        }

        public static long GetJobExecWaitCount()
        {
            return _jobExecWaitCount;
        }

        public static void ResetJobExecWaitCount()
        {
            _jobExecWaitCount = 0;
        }

        public void SetIndex(int index)
        {
            _index = index;
        }

        public string SetSingleValue()
        {
            _activeWorksheet.Cells[_index, 6] = "threadId " + _threadId.ToString();
            _activeWorksheet.Cells[_index, 7] = "indexId " + _index.ToString();
            _activeWorksheet.Cells[_index, 8] = "indexId calculation " + (1024 * _index).ToString();
            _activeWorksheet.Cells[_index, 9] = "Seach Results File Name " + (1024 * _index).ToString() + "_SingleThread.html";
            if (_setTS)
            {
                _activeWorksheet.Cells[_index, 10] = DateTime.Now.ToString();
            }
            Interlocked.Increment(ref _jobExecWaitCount);
            return (1024 * _index).ToString();
        }

        public string[] SetRangeValue()
        {
            List<string> results = new List<string>();
            for (long i = _startIndex; i < _endIndex; i++)
            {
                _activeWorksheet.Cells[i, 1] = "threadId " + _threadId.ToString();
                _activeWorksheet.Cells[i, 2] = "indexId " + i.ToString();
                _activeWorksheet.Cells[i, 3] = "indexId calculation " + (1024 * 1024 * i).ToString();
                _activeWorksheet.Cells[i, 4] = "Seach Results File Name " + (1024 * 1024 * i).ToString() + "_MultiThread.html";
                if (i == _startIndex || i == _endIndex - 1)
                {
                    _activeWorksheet.Cells[i, 5] = DateTime.Now.ToString();
                }
                results.Add((1024 * 1024 * i).ToString());
                Interlocked.Increment(ref _jobExecWaitCount);
            }
            return results.ToArray();
        }

        public bool MutexWaitOne()
        {
            while (true)
            {
                if (_mut.WaitOne())
                {
                    return true;
                }
            }
        }
        public void Release()
        {
            try
            {
                _mut.ReleaseMutex();
            }
            catch
            {
            }
        }
    }
}
