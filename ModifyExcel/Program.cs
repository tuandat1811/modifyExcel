#define demo                    //Sinh cấu trúc dữ liệu đầu vào mặc định  demo.json

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
//using Excel = Microsoft.Office.Interop.Excel;       //Khong su dung nua, do da lien ket động


namespace ModifyExcel
{
    class Program
    {

        /// <summary>
        ///    Đường dẫn tuyệt đối tới file excel
        /// </summary>
        static string ExcelDBPath;
  
        static public bool IsExcelSupport()
        {
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            return (officeType != null);
        }

        static int Main(string[] args)
        {
            /// Triệu gọi thư viện Interop Excel kiểu động.    
            Type typeExcel = Type.GetTypeFromProgID("Excel.Application");
            dynamic Excel = Activator.CreateInstance(typeExcel);

            string errMsg = null;

            /// Tạo các biến thuộc kiểu dữ liệu liên kết động ở thư viện  Interop Excel. Thế là xong. Mọi việc lại diễn ra bình thường
            dynamic MyBook = null;
            dynamic MyApp = null;
            dynamic MySheet = null;

            string ExcelTemplateFileName  ;
            CellPosition CellMSCH;
            CellPosition CellTuan;
            CellPosition CellPTID;
            CellPosition CellNgaySinh;
            String SelectedSheetName ;  /// Sheet được chỉ định

            string MSCH;
            string Tuan;
            string PTID;
            string NgaySinh;

            Boolean ExcelIsVisible;
            Boolean PrintNow;

            if (!IsExcelSupport())
            {
                errMsg = "Chưa cài đặt MS Excel.";
                goto _END_;
            }

            if (args.Length == 3)
            {
                if (args[0] == "-json")
                {
                    string InputFile = args[1];
                    ExcelTemplateFileName = args[2];
                    string json;

                    if (!File.Exists(InputFile))
                    {
                        errMsg = "File " + InputFile + " không tôn tại \n. Thư mục hiện thời: " + Directory.GetCurrentDirectory() + "\n. Xem file demo.json để biết cấu trúc đầu vào";
                        json = "";

                    }
                    else
                    {
                        json = File.ReadAllText(InputFile);
                    }
                    
                    Modification2(json, ExcelTemplateFileName) ;
                    goto _END_;
                }

            }
                if (args.Length < 11)
            {
                errMsg = "Không đủ tham số dòng lệnh.\n   <file name> <CellMSCH> <MSCH> <CellTuan> <Tuan> <CellPTID> <PTID> <CellNgaySinh> <NgaySinh>  <Visible> <PrintNow> [<Sheet Name>]";
                errMsg += "  \n   -json <JsonFileName> <TemplateFileName>";
                goto _END_;

            }

            try
            {
                ExcelTemplateFileName = args[0];
                CellMSCH = new CellPosition(args[1]);
                MSCH = args[2];
                CellTuan = new CellPosition(args[3]);
                Tuan = args[4];
                CellPTID = new CellPosition(args[5]);
                PTID = args[6];
                CellNgaySinh = new CellPosition(args[7]);
                NgaySinh = args[8];

                ExcelIsVisible = Convert.ToBoolean(args[9]);
            }
            catch (Exception e3)
            {
                errMsg = "Tham số vào không hợp lệ." + e3.Message;
                goto _END_;
            }

            PrintNow = Convert.ToBoolean(args[10]);

            if (args.Length >= 12)
            {
                SelectedSheetName = args[11];
            }
            else
            {
                SelectedSheetName = null;
            }

            ExcelDBPath = ExcelTemplateFileName;
            if (!File.Exists(ExcelDBPath))
            {
                errMsg = "File " + ExcelDBPath + " không tồn tại";
                if (Directory.Exists(Path.GetDirectoryName(ExcelDBPath)))
                {
                    errMsg += " (có thư mục).";
                }   else
                {
                    errMsg += " (không có thư mục).";
                }
                goto _END_;
            }


            try
            {
                MyApp = Excel.Application();
                if (MyApp == null)
                {
                    errMsg = "Excel không khởi động được, do phiên bản Excel không phù hợp với phần mềm SUCe. Hãy sử dụng định dạng CSV thay thế.\r\nVui lòng liên hệ với công ty Techlink để được tư vấn và hỗ trợ.";
                    goto _END_;
                }

                // Yêu cầu ứng dụng excel không hiển thị ra màn hình, --> chạy ngầm. //
                /* 
                 * Nếu file excel DB đã được mở, không hiển thị file nữa
                 * và đóng tiến trình chạy ngầm sau khi load xong dữ liệu
                 */
                MyApp.Visible = ExcelIsVisible;
                // Mở file excel theo đường dẫn đã cho
                MyBook = MyApp.Workbooks.Open(ExcelDBPath, ReadOnly: true);
                if (MyBook == null)
                {
                    errMsg = "Không mở được Workbook do phiên bản Excel không phù hợp với phần mềm SUCe. Hãy sử dụng định dạng CSV thay thế.\r\nVui lòng liên hệ với công ty Techlink để được tư vấn và hỗ trợ.";
                    goto _END_;
                }
            }
            catch
            {
                errMsg = "Phiên bản Excel không phù hợp với phần mềm. Hãy sử dụng định dạng CSV thay thế.\r\nVui lòng liên hệ với công ty Techlink để được tư vấn và hỗ trợ.";
                goto _END_;
            }

            {
                //Mở Sheet có tên như chỉ định, hoặc sử dụng số thứ tự
                try
                {
                    if (SelectedSheetName != null)
                    {
                        MySheet = MyBook.Worksheets[SelectedSheetName];
                    }
                    else
                    {
                        MySheet = MyBook.Sheets[1];
                    }
                }
                catch
                {
                    errMsg = "File dữ liệu Excel không có worksheet có tên qui định là " + SelectedSheetName;
                    goto _END_;
                }

                try
                {
                    MySheet.Cells[CellMSCH.RowIndex, CellMSCH.ColumnIndex].Value = MSCH;
                    MySheet.Cells[CellTuan.RowIndex, CellTuan.ColumnIndex].Value = Tuan;
                    MySheet.Cells[CellPTID.RowIndex, CellPTID.ColumnIndex].Value = PTID;
                    MySheet.Cells[CellNgaySinh.RowIndex, CellNgaySinh.ColumnIndex].Value = NgaySinh;
                }
                catch (Exception e)
                {
                    errMsg = e.Message;
                    goto _END_;
                }

                // Thực hiện in luôn ra máy in nếu có yêu cầu
                if (PrintNow)
                {
                    try
                    {
                        MySheet.PrintOut(
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            
                    }
                    catch (Exception e2)
                    {
                        errMsg = e2.Message;
                        goto _END_;
                    }
                }
            }

            try
            {
                //MyBook.SaveAs(ExcelTemplateFileName, Excel.XlFileFormat.xlWorkbookDefault);//, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
                //MyBook.Save();
            }
            catch (Exception e)
            {
                errMsg = e.Message;
                goto _END_;
            }

            _END_:
            // Đóng file, kết thúc
            if (MyBook != null)
            {
                MyBook.Close( SaveChanges: false);
            }
            if (MyApp != null)
            {
                MyApp.Quit();
            }
            Console.WriteLine("Version 2.0");
            Console.WriteLine(errMsg);
            Debug.WriteLine(errMsg);
            return 0;
        }

        /// <summary>
        ///     Nhồi dữ liệu vào file excel
        /// </summary>
        /// <param name="jsontext"> Chứa thông tin cần nhồi vào file excel </param>
        /// <param name="excelTemplateFileName" File Excel template gốc </param> 
        private static void Modification2(string jsontext, string excelTemplateFileName)
        {
            //-----------------------------------JSON PARSER to  ORM ----------------------------
            /// Đọc và phân tích nội dung json đầu vào
            JsonTextReader reader = new JsonTextReader(new StringReader(jsontext));


            WorkbookType MyWorkbookData;
            
            /// Nếu dữ liệu đầu vào không có, thì sẽ hiển thị cấu trúc dữ liệu demo 
            if (jsontext == "")
            {
                /// Đối tượng json handler
                JsonSerializer serializer = new JsonSerializer();

                MyWorkbookData = new WorkbookType();

                ///    - Bổ sung cấu hình
                MyWorkbookData.config.printername = "pdf";
                MyWorkbookData.config.visible = true;
                MyWorkbookData.config.saveas = "demo.xlsx";

                ///    - Bổ sung sheet mới để có dữ liệu minh họa
                SheetType MySheetData = new SheetType(new List<CellType>(), "Programable Sheet");
                MyWorkbookData.sheets.Add(MySheetData);

                ///    - Bổ sung các cell mới vào sheet nói trên, để có dữ liệu minh họa
                MySheetData.cells.Add(new CellType("A5", true));
                MySheetData.cells.Add(new CellType("B1", "Text"));
                MySheetData.cells.Add(new CellType("C3", "20/11/2012"));

                ////    - Ghi file demo
                using (StreamWriter sw = new StreamWriter(@"demo.json"))
                using (JsonWriter writer = new JsonTextWriter(sw))
                {
                    serializer.Serialize(writer, MyWorkbookData);
                }
                return;
            }

            try
            {
                //ORM hóa nội dung json vào class
                MyWorkbookData = JsonConvert.DeserializeObject<WorkbookType>(jsontext);
            }catch (Exception e)
            {
                Console.WriteLine("Json Error: " + e.Message);
                Debug.WriteLine("Json Error: " + e.Message);
                return;
            }

            //----------------------------------- ORM To WORKBOOK  ----------------------------

            /// Triệu gọi thư viện Interop Excel kiểu động.
            Type typeExcel = Type.GetTypeFromProgID("Excel.Application");
            dynamic Excel = Activator.CreateInstance(typeExcel);

            string errMsg = null;

            /// Tạo các biến thuộc kiểu dữ liệu liên kết động ở thư viện  Interop Excel. Thế là xong. Mọi việc lại diễn ra bình thường
            dynamic MyBook = null;
            dynamic MyApp = null;
            dynamic MySheet = null;

            try
            {
                MyApp = Excel.Application();
                if (MyApp == null)
                {
                    MyWorkbookData.errMessage = "Excel không khởi động được.";
                    goto _END_;
                }

                // Yêu cầu ứng dụng excel không hiển thị ra màn hình, --> chạy ngầm. //
                /* 
                 * Nếu file excel DB đã được mở, không hiển thị file nữa
                 * và đóng tiến trình chạy ngầm sau khi load xong dữ liệu
                 */
                MyApp.Visible = MyWorkbookData.config.visible;

                // Mở file excel theo đường dẫn đã cho
                MyBook = MyApp.Workbooks.Open(excelTemplateFileName, ReadOnly: true);
                if (MyBook == null)
                {
                    MyWorkbookData.errMessage = "Không mở được Workbook do phiên bản Excel không phù hợp.";
                    goto _END_;
                }
            }
            catch (Exception e)
            {
                MyWorkbookData.errMessage = e.Message;
                goto _END_;
            }
            //----------------------------------- ORM To SHEETS  ----------------------------
            foreach (SheetType MySheetData in MyWorkbookData.sheets)
            {
                //Mở Sheet có tên như chỉ định, hoặc sử dụng số thứ tự
                try
                {
                    if (MySheetData.name != null)
                    {
                        MySheet = MyBook.Worksheets[MySheetData.name];
                        if (MyWorkbookData.config.activatesheet)
                        {
                            MySheet.Activate();
                        }
                    }
                    else
                    {
                        MySheetData.errMessage += "Tên Sheet trống, không hợp lệ.\n";
                        continue;
                    }
                }
                catch
                {
                    MySheetData.errMessage = "File dữ liệu Excel không có worksheet có tên qui định là " + MySheetData.name;
                    Debug.WriteLine(MySheetData.errMessage);
                    continue;
                }

                /// Ghi nội dung vào các cell
                MySheetData.errMessage = "";
                foreach (CellType MyCellData in MySheetData.cells)
                {
                    try
                    {
                        MySheet.Cells[MyCellData.RowIndex, MyCellData.ColumnIndex].Value = MyCellData.value;
                    }
                    catch (Exception e)
                    {
                        MySheetData.errMessage += (MyCellData.pos + " " + e.Message);
                        Debug.WriteLine(MySheetData.errMessage);
                        continue;
                    }
                }
            }

            //----------------------------------- PRINT / SAVE  ----------------------------

            /// Lưu lại file đã có kết quả
            if (MyWorkbookData.config.saveas != string.Empty)
            {
                if (!System.IO.Path.IsPathRooted(MyWorkbookData.config.saveas))
                {// Đường dẫn mặc định là thư mục Documents. Cần đổi về đường dẫn tương đối
                    MyWorkbookData.config.saveas = Directory.GetCurrentDirectory() + @"\" + MyWorkbookData.config.saveas;
                }
                try
                {
                    MyApp.DisplayAlerts = false;
                    //Lưu ý: Do muốn overwritten, mà hằng Excel.XlSaveAsAccessMode.xlNoChange lại không có khi triệu gọi lb động --> sử dụng luôn hằng số 1.
                    MyBook.SaveAs(MyWorkbookData.config.saveas, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MyApp.DisplayAlerts = true;
                }
                catch (Exception e)
                {
                    MyWorkbookData.errMessage += (e.Message + "\n");
                    Debug.WriteLine(e.Message);
                    goto _END_;
                }
            }

            /// Thực hiện in luôn ra máy in tất cả các sheet nếu có yêu cầu
            if (MyWorkbookData.config.printnow) {
                /// Mặc định là máy in mặc định của hệ thống
                var printer = Type.Missing;

                /// Trường hợp có chỉ định tên máy in, thì sẽ tìm tên 
                if (MyWorkbookData.config.printername != string.Empty)
                {
                    var printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters;

                    foreach (String s in printers)
                    {
                        if (s.IndexOf(MyWorkbookData.config.printername,StringComparison.OrdinalIgnoreCase)>=0)
                        {
                            printer = s;
                            break;
                        }
                    }
                }

                foreach (SheetType MySheetData in MyWorkbookData.sheets)
                {
                    //Mở lại sheet đó
                    if (MySheetData.name != null)
                    {
                        MySheet = MyBook.Worksheets[MySheetData.name];
                    }
                    //và in đúng máy in chỉ định
                    try
                    {
                        MySheet.PrintOut(
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            printer, Type.Missing, Type.Missing, Type.Missing);
                    }
                    catch (Exception e)
                    {
                        MyWorkbookData.errMessage += (e.Message + "\n");
                        continue;
                    }
                }
            }    

            _END_:


            // Đóng file, kết thúc
            if (MyWorkbookData.config.terminate)
            {
                if (MyBook != null)
                {
                    MyBook.Close(SaveChanges: false);
                }
                if (MyApp != null)
                {
                    MyApp.Quit();
                }
            }
            Console.WriteLine("Version 2.0");

            /// Hiển thị tất cả các lỗi
            Console.WriteLine(MyWorkbookData.errMessage);
            foreach (SheetType MySheetData in MyWorkbookData.sheets)
            {
                Console.WriteLine(MySheetData.errMessage);
            }
        }
    }
}
