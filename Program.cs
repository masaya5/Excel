using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test_console
{
    class Program
    {
        static void Main(string[] args)
        {
            //DBConnect dc = new DBConnect();
            
            //Console.WriteLine(dc.試験ID);


            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            //エクセルを非表示
            ExcelApp.Visible = false;

            //エクセルファイルのオープン
            Microsoft.Office.Interop.Excel.Workbook WorkBook = ExcelApp.Workbooks.Open(@"C:\Users\maeda.masaya\Desktop\test.xlsx");

            //1シート目の選択
            Microsoft.Office.Interop.Excel.Worksheet sheet = WorkBook.Sheets[1];
            sheet.Select();

            //A1セルのデータの取得
            Microsoft.Office.Interop.Excel.Range range = sheet.get_Range("A1");
            if (string.IsNullOrEmpty(range.Value))
            {
                Console.WriteLine("ブランクです。");

                
            }
            else if (range != null)
            {

                Console.WriteLine(range.Value);
            }

            //A1セルから見たら下への連続データ数
            int row_count = sheet.get_Range("A1").End[Microsoft.Office.Interop.Excel.XlDirection.xlDown].Row;

            //A1セルから見たら右への連続データ数
            int column_count = sheet.get_Range("A1").End[Microsoft.Office.Interop.Excel.XlDirection.xlToRight].Column;

            //A1セルのデータの書き込み
            range = sheet.get_Range("A1");
            if (range != null)
            {
                range.Value = 10;
            }
            
            //データの保存
            WorkBook.Save();

            //workbookを閉じる
            WorkBook.Close();
            //エクセルを閉じる
            ExcelApp.Quit();



            // Keep the console window open in debug mode.
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
