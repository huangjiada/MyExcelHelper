using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyExcelHelper;
using Excel = Microsoft.Office.Interop.Excel;

namespace MyExcelDemo
{
    class Program
    {
        static void Main(string[] args)
        {

            Random r = new Random();
            Excel.Workbook myWb = MyExcel.GetWorkBook(System.AppDomain.CurrentDomain.BaseDirectory + "demo.xlsx");
            try
            {
                MessageFilter.Register();
                #region 写excel

                myWb.Application.ScreenUpdating = true;//控制excel刷新暂停，false之后等到true才刷新，不暂停将会导致写一个格刷新一次
                                                       //这里写入excel
                Excel.Worksheet ws = (Excel.Worksheet)myWb.Sheets["Sheet1"];//想要搞的workbook的名字
                string sample = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD";
                string[] samples = sample.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var item in samples)
                {
                    for (int i = 1; i < 10; i++)
                    {
                        ws.GetCell(item, i).Value = r.Next(1, 100);
                    }
                }

                myWb.Save();
                myWb.Application.ScreenUpdating = true;

                #endregion
                MessageFilter.Revoke();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                myWb.Application.Quit();
            }
            

        }
    }
}
