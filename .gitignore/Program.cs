using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Voucher
{
    public static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
    public static class Run
    {
        static int counter = 1;
        static bool breakflag = false;
        static int carryIndex;
        static int carryTranIndex;
        static int carryEntries;
        static double carrySum;
        static int sheet = 1;
        public static void Generate(string path, string companyName, string filepath)
        {
            List<Voucher> ll = new List<Voucher>();
            FileStream createStream = new FileStream(filepath + "\\Voucher.xlsx", FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
            ExcelPackage ep = new ExcelPackage(createStream);
            Excel.GetList(ll, path);
            Console.WriteLine(Excel.noOfRecords(ll));
            Console.WriteLine(Excel.record(3,ll));
            for (int i = 0; i < Excel.noOfVouchers(ll); i++)
            {
                Excel.CreateSheet(ep, ((i + 1).ToString()));
            }
            while (sheet <= Excel.noOfVouchers(ll))
            {
                double sum = 0;
                for (int j = 0, row = 6, tranIndex = Excel.record(counter, ll), index = Excel.record(counter, ll); j < Excel.noOfEntries(counter, ll); j++)
                {
                    ep.Workbook.Worksheets[sheet].Cells["B2"].Value = companyName;
                    ep.Workbook.Worksheets[sheet].Cells["D2"].Value = ll[tranIndex].tranID;
                    ep.Workbook.Worksheets[sheet].Cells["B4"].Value = "Date: " + ll[tranIndex].Date;
                    ep.Workbook.Worksheets[sheet].Cells[row, 1].Value = ll[index].accountID;
                    ep.Workbook.Worksheets[sheet].Cells[row, 2].Value = ll[index].Description;
                    if (ll[index].debit == 0)
                    {
                        ep.Workbook.Worksheets[sheet].Cells[row, 4].Value = "";
                    }
                    else
                    {
                        ep.Workbook.Worksheets[sheet].Cells[row, 4].Value = ll[index].debit;
                    }
                    if (ll[index].credit == 0)
                    {
                        ep.Workbook.Worksheets[sheet].Cells[row, 5].Value = "";
                    }
                    else
                    {
                        ep.Workbook.Worksheets[sheet].Cells[row, 5].Value = ll[index].credit;
                    }
                    sum += ll[index].debit;
                    index++;
                    row++;
                    if (row == 18)
                    {
                        if (sheet < Excel.noOfVouchers(ll))
                        {
                            sheet++;
                            row = 6;
                        }
                        else
                        {
                            breakflag = true;
                            carryIndex = index;
                            carryTranIndex = tranIndex;
                            carryEntries = Excel.noOfEntries(counter, ll) - j;
                            carrySum = sum;
                            break;
                        }
                    }
                }
                if (breakflag == true)
                {
                    break;
                }
                else
                {
                    ep.Workbook.Worksheets[sheet].Cells["D18"].Value = sum;
                    ep.Workbook.Worksheets[sheet].Cells["E18"].Value = sum;
                    sheet++;
                    counter++;
                }
            }
            if ((Excel.noOfRecords(ll) > 1) || (breakflag == true))
            {
                double sum = 0;
                int index;
                int tranIndex;
                int entries;
                sheet = 1;
                if (breakflag == true)
                {
                    index = carryIndex;
                    tranIndex = carryTranIndex;
                    entries = carryEntries;
                    sum = carrySum;
                    breakflag = false;
                }
                else
                {
                    index = Excel.record(counter, ll);
                    tranIndex = Excel.record(counter, ll);
                    entries = Excel.noOfEntries(counter, ll);
                }
                while ((sheet <= Excel.noOfVouchers(ll)) && (index <= ll.Count - 1) && (breakflag == false))
                {

                    for (int j = 0, row = 28; j < entries; j++)
                    {
                        ep.Workbook.Worksheets[sheet].Cells["B24"].Value = companyName;
                        ep.Workbook.Worksheets[sheet].Cells["D24"].Value = ll[tranIndex].tranID;
                        ep.Workbook.Worksheets[sheet].Cells["B26"].Value = "Date: " + ll[tranIndex].Date;
                        ep.Workbook.Worksheets[sheet].Cells[row, 1].Value = ll[index].accountID;
                        ep.Workbook.Worksheets[sheet].Cells[row, 2].Value = ll[index].Description;
                        if (ll[index].debit == 0)
                        {
                            ep.Workbook.Worksheets[sheet].Cells[row, 4].Value = "";
                        }
                        else
                        {
                            ep.Workbook.Worksheets[sheet].Cells[row, 4].Value = ll[index].debit;
                        }
                        if (ll[index].credit == 0)
                        {
                            ep.Workbook.Worksheets[sheet].Cells[row, 5].Value = "";
                        }
                        else
                        {
                            ep.Workbook.Worksheets[sheet].Cells[row, 5].Value = ll[index].credit;
                        }
                        sum += ll[index].debit;
                        if (index != ll.Count - 1)
                        {
                            index++;
                            row++;
                        }
                        else
                        {
                            breakflag = true;
                            break;
                        }
                        if (row == 40)
                        {
                            sheet++;
                            row = 28;
                        }
                    }
                    ep.Workbook.Worksheets[sheet].Cells["D40"].Value = sum;
                    ep.Workbook.Worksheets[sheet].Cells["E40"].Value = sum;
                    sheet++;
                    counter++;
                    entries = Excel.noOfEntries(counter, ll);
                    tranIndex = Excel.record(counter, ll);
                    index = Excel.record(counter, ll);
                    sum = 0;
                }
            }
            ep.Save();
            counter = 1;
            breakflag = false;
            sheet = 1;
        }
    }
    public class Voucher
    {
        public string Date { get; set; }
        public object accountID { get; set; }
        public string tranID { get; set; }
        public string Description { get; set; }
        public double debit { get; set; }
        public double credit { get; set; }

    }
    public class Excel
    {
        public static void GetList(List<Voucher> lists, string path)
        {
            using (FileStream readStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (ExcelPackage ep2 = new ExcelPackage(readStream))
                {

                    ExcelWorksheet sheet = ep2.Workbook.Worksheets[1];//取得Sheet1
                    int startRowNumber = 2;//起始列編號，從1算起
                    int endRowNumber = sheet.Dimension.End.Row - 3;//結束列編號，從1算起
                    int startColumn = sheet.Dimension.Start.Column;//開始欄編號，從1算起
                    int endColumn = sheet.Dimension.End.Column;
                    for (int currentRow = startRowNumber; currentRow <= endRowNumber; currentRow++)
                    {
                        ExcelRange range = sheet.Cells[currentRow, startColumn, currentRow, endColumn];//抓出目前的Excel列
                        Voucher obj = new Voucher();
                        obj.Date = "20" + Convert.ToString(sheet.Cells[currentRow, 1].Value);
                        obj.Date = obj.Date.Replace("/", "-");
                        obj.accountID = sheet.Cells[currentRow, 2].Value;
                        obj.tranID = sheet.Cells[currentRow, 3].Text;
                        obj.Description = sheet.Cells[currentRow, 4].Text;
                        obj.debit = Convert.ToDouble(sheet.Cells[currentRow, 5].Value);
                        obj.credit = Convert.ToDouble(sheet.Cells[currentRow, 6].Value);
                        lists.Add(obj);
                    }
                }
            }
        }
        public static int noOfEntries(int no, List<Voucher> list)
        {
            int answer;
            int count = 1;
            if (no == noOfRecords(list))
            {
                for (int i = list.Count - 1; i > record(noOfRecords(list), list); i--)
                {
                    count++;
                }
                return count;
            }
            else
            {
                answer = record(no + 1, list) - record(no, list) - 1;
                return answer;
            }
        }
        public static int noOfRecords(List<Voucher> list)
        {
            int counter = 1;
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].accountID == null)
                {
                    counter++;
                }
            }
            return counter;
        }
        public static int record(int record, List<Voucher> list)
        {

            int index = 0;
            if (record == 1)
            {
                return 0;
            }
            else
            {
                int count = 1;
                for (int i = 0; i < list.Count(); i++)
                {
                    if (list[i].accountID == null)
                    {
                        count++;
                        if (count == record)
                        {
                            index = i + 1;
                        }
                    }
                }
                return index;
            }

        }
        public static int noOfVouchers(List<Voucher> list)
        {
            int count = 0;
            for (int i = 0; i < noOfRecords(list); i++)
            {
                int cast = noOfEntries(i + 1, list) / 12;
                if (cast > 0 && (cast % 12 != 0))
                {
                    count += cast + 1;
                }
                else if (cast > 0 && (cast % 12 == 0))
                {
                    count += cast;
                }
                else
                {
                    count += 1;
                }
            }
            if (count % 2 == 0)
                return (count / 2);
            else
                return ((count / 2) + 1);
        }
        public static ExcelWorksheet CreateSheet(ExcelPackage p, string sheetName)
        {
            p.Workbook.Worksheets.Add(sheetName);
            ExcelWorksheet ws = p.Workbook.Worksheets[sheetName];
            ws.Name = sheetName; //Setting Sheet's name
            ws.Cells.Style.Font.Size = 12; //Default font size for whole sheet
            ws.Cells.Style.Font.Name = "Arial"; //Default Font name for whole sheet
            ws.PrinterSettings.PaperSize = ePaperSize.A4;
            ws.PrinterSettings.Orientation = eOrientation.Portrait;
            ws.PrinterSettings.FitToPage = true;
            //ws.PrinterSettings.Scale = 100;
            //ws.PrinterSettings.PrintArea = ws.Cells[1, 1, 43, 5];
            ws.PrinterSettings.FitToHeight = 1;
            ws.PrinterSettings.FitToWidth = 1;
            for (int i = 1; i < 42; i++)
            {
                ws.Row(i).Height = 19.7;
            }
            ws.Column(1).Width = 15.14;
            ws.Column(2).Width = 53.43;
            ws.Column(3).Width = 1.71;
            ws.Column(4).Width = 13.57;
            ws.Column(5).Width = 13.57;
            ws.PrinterSettings.TopMargin = 0;
            ws.PrinterSettings.BottomMargin = 0;
            ws.PrinterSettings.RightMargin = 0;
            ws.PrinterSettings.LeftMargin = 0;
            ws.PrinterSettings.VerticalCentered = true;
            ws.PrinterSettings.HorizontalCentered = true;
            ws.PrinterSettings.FooterMargin = 0;
            ws.PrinterSettings.HeaderMargin = 0;
            ws.PrinterSettings.BlackAndWhite = true;
            ws.PrinterSettings.Draft = false;
            ws.Cells[3, 2].Value = "VOUCHER";
            ws.Cells["B2:B4"].Style.Font.Bold = true;
            ws.Cells["B2:B4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["B4"].Style.Numberformat.Format = "yyyy-mm-dd";
            ws.Cells["B4"].Style.Font.Size = 14;
            ws.Cells[3, 2].Style.Font.Size = 16;
            ws.Cells[2, 2].Style.Font.Size = 18;
            var border = ws.Cells[3, 2].Style.Border;
            border.Bottom.Style = border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells[2, 4, 2, 5].Style.Font.Bold = true;
            ws.Cells[2, 4, 2, 5].Merge = true;
            ws.Cells[18, 1, 18, 2].Merge = true;
            border = ws.Cells[2, 4, 2, 5].Style.Border;
            border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            border = ws.Cells[6, 1, 18, 5].Style.Border;
            border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            border = ws.Cells[5, 1, 5, 5].Style.Border;
            border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            border.Bottom.Style = ExcelBorderStyle.Medium;
            ws.Cells[5, 1, 5, 5].Style.Font.Size = 14;
            ws.Cells[5, 1, 5, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[5, 1].Value = "ACCOUNT";
            ws.Cells[5, 2].Value = "PARTICULARS";
            ws.Cells[5, 4].Value = "Debit";
            ws.Cells[5, 5].Value = "Credit";
            ws.Cells[18, 1].Value = "Bills Attached           Sheets                                                       TOTAL";
            ws.Cells[19, 1, 20, 5].Merge = true;
            border = ws.Cells[19, 1, 20, 5].Style.Border;
            border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells["D6:E18"].Style.Numberformat.Format = "#,##0.00";
            ws.Cells["A6:A17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A2:E18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["B2"].Style.ShrinkToFit = true;
            ws.Cells["D2"].Style.ShrinkToFit = true;
            ws.Cells["A6:E17"].Style.ShrinkToFit = true;
            ws.Cells["D18:E18"].Style.ShrinkToFit = true;
            ws.Cells[19, 1].Value = "Approved by                Accountant                    Checked by                    Made by";
            //BOTTOM SHEET
            ws.Cells[25, 2].Value = "VOUCHER";
            ws.Cells["B24:B26"].Style.Font.Bold = true;
            ws.Cells["B24:B26"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[25, 2].Style.Font.Size = 16;
            ws.Cells[24, 2].Style.Font.Size = 18;
            border = ws.Cells[25, 2].Style.Border;
            border.Bottom.Style = border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells[24, 4, 24, 5].Style.Font.Bold = true;
            ws.Cells[24, 4, 24, 5].Merge = true;
            ws.Cells[40, 1, 40, 2].Merge = true;
            border = ws.Cells[24, 4, 24, 5].Style.Border;
            border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            border = ws.Cells[27, 1, 42, 5].Style.Border;
            border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            border = ws.Cells[27, 1, 27, 5].Style.Border;
            border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            border.Bottom.Style = ExcelBorderStyle.Medium;
            ws.Cells[27, 1, 27, 5].Style.Font.Size = 14;
            ws.Cells[27, 1, 27, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[27, 1].Value = "ACCOUNT";
            ws.Cells[27, 2].Value = "PARTICULARS";
            ws.Cells[27, 4].Value = "Debit";
            ws.Cells[27, 5].Value = "Credit";
            ws.Cells[40, 1].Value = "Bills Attached           Sheets                                                       TOTAL";
            ws.Cells[41, 1, 42, 5].Merge = true;
            border = ws.Cells[41, 1, 42, 5].Style.Border;
            border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells["D28:E40"].Style.Numberformat.Format = "#,##0.00";
            ws.Cells["A28:A39"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A24:E40"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["B39"].Style.Numberformat.Format = "yyyy-mm-dd";
            ws.Cells["B39"].Style.Font.Size = 14;
            ws.Cells["D24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["B24"].Style.ShrinkToFit = true;
            ws.Cells["D24"].Style.ShrinkToFit = true;
            ws.Cells["A28:E39"].Style.ShrinkToFit = true;
            ws.Cells["D40:E40"].Style.ShrinkToFit = true;
            ws.Cells[41, 1].Value = "Approved by                Accountant                    Checked by                    Made by";
            return ws;
        }
    }
}

