using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace MyExcelHelper
{
    public static class MyExcel
    {
        /// <summary>
        /// 根据完整路径返回一个Excel.Workbook
        /// </summary>
        /// <param name="path">excel文件的完整路径</param>
        /// <returns>返回Excel.Workbook</returns>
        public static Excel.Workbook GetWorkBook(string path)
        {
            Excel.Workbook wb;

            Excel.Application oXL;
            try
            {
                oXL = (Excel.Application)Marshal.GetActiveObject("Excel.Application");//检查内存，看看是否已有excel进程打开
            }
            catch (Exception)
            {
                oXL = new Excel.Application();//无已经打开的Excel进程则新建Excel进程
            }
            if (IsOpen(oXL.Workbooks, path.Split('\\').Last()))
            {

                wb = oXL.Workbooks.get_Item(path.Split('\\').Last());
                //若excel文件已打开则直接拉Item生成副本
                //多用户打开怎么办？弹窗提示？还是直接调试失败？记得复制office.interop.excel
                //不需要打开，作为模板直接复制项目
            }
            else
            {
                wb = oXL.Workbooks.Open(path,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                //未打开excel文件则用Open打开
            }

            return wb;
        }

        /// <summary>
        /// 返回单个单元格
        /// </summary>
        /// <param name="ws">Worksheet</param><param name="column">字符串列名，如A,AA,AAA</param><param name="row">整数行号</param>
        /// <returns>需改返回值可以Return.Value</returns>
        public static Excel.Range GetCell(this Excel.Worksheet ws, string column, int row)
        {
            return (Excel.Range)ws.Cells[row, ToIndex(column)];//让字符的列名转化为对应的列号
            
        }

        /// <summary>
        /// 检查workbook是否被打开，当excel进程打开workbook时检查文件是否被占用
        /// </summary>
        /// <param name="wbs">字母列名称</param><param name="book">workbook的路径</param>
        /// <returns></returns>
        static bool IsOpen(Excel.Workbooks wbs, string book)
        {
            foreach (Excel._Workbook wb in wbs)
            {
                if (wb.Name.Contains(book))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 将excel中字符列转换为列index
        /// </summary>
        /// <param name="columnName">字母列名称，如A，AA，AAA</param>
        /// <returns>返回列号</returns>
        public static int ToIndex(string columnName)
        {
            if (!Regex.IsMatch(columnName.ToUpper(), @"[A-Z]+")) { throw new Exception("invalid excel column parameter"); }
            int index = 0;
            char[] chars = columnName.ToUpper().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
            }
            return index;
        }

    }


    //用messagefiter锁住线程中的excel。可以解释之前为什么那么缓慢了，如果没有锁，多线程同时访问excel资源会卡死
    public class MessageFilter : IOleMessageFilter
    {
        // Class containing the IOleMessageFilter
        // thread error-handling functions.

        // Start the filter.
        public static void Register()
        {
            IOleMessageFilter newFilter = new MessageFilter();
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(newFilter, out oldFilter);
        }

        // Done with the filter, close it.
        public static void Revoke()
        {
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(null, out oldFilter);
        }

        //
        // IOleMessageFilter functions.
        // Handle incoming thread requests.
        int IOleMessageFilter.HandleInComingCall(int dwCallType,
          System.IntPtr hTaskCaller, int dwTickCount, System.IntPtr
          lpInterfaceInfo)
        {
            //Return the flag SERVERCALL_ISHANDLED.
            return 0;
        }

        // Thread call was rejected, so try again.
        int IOleMessageFilter.RetryRejectedCall(System.IntPtr
          hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (dwRejectType == 2)
            // flag = SERVERCALL_RETRYLATER.
            {
                // Retry the thread call immediately if return >=0 & 
                // <100.
                return 99;
            }
            // Too busy; cancel call.
            return -1;
        }

        int IOleMessageFilter.MessagePending(System.IntPtr hTaskCallee,
          int dwTickCount, int dwPendingType)
        {
            //Return the flag PENDINGMSG_WAITDEFPROCESS.
            return 2;
        }

        // Implement the IOleMessageFilter interface.
        [DllImport("Ole32.dll")]
        private static extern int
          CoRegisterMessageFilter(IOleMessageFilter newFilter, out
          IOleMessageFilter oldFilter);
    }

    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(
            int dwCallType,
            IntPtr hTaskCaller,
            int dwTickCount,
            IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwRejectType);

        [PreserveSig]
        int MessagePending(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwPendingType);
    }
}
