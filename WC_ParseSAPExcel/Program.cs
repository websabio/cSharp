using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;



namespace WC_ParseSAPExcel
{
    class Program
    {
        #region Global Variables
        static private string _FILEPATH = @"testouput1.txt";
        static private int _NUMEMPLOYEES = 0;
        static private DataSet _DS = null;
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        static private string _YTD_WCOMP_2_XLS_PATH = @"C:\Worker's com\SAP\YTD Wcomp 2.xls";
        #endregion

        static void Main(string[] args)
        {

            switch (args[0])
            {
                case "1":
                    Console.WriteLine("Parse SAP F1 Report");
                    ParseSAPF1_Report();
                    break;
                case "2":
                    Console.WriteLine("Read Excel File YTD Wcomp 2.xls");
                    break;
                default:
                    Console.WriteLine("Default case");
                    break;
            }


            //Console.ReadLine();
        }

        static private void ParseSAPF1_Report()
        {
            initialiseDataSet();
            ReadText();
            writeDataSetToXML();
        }

        static private void ReadText()
        {
            string resultTextBlock = "";
            //string name;
            //bool writeFlag = false;
            string[] resultArray = null;
            string[] stringSeparators = new string[] { "Pay period" };
            if (File.Exists(_FILEPATH))
            {
                File.Delete(_FILEPATH);
            }
            if (!File.Exists(_FILEPATH))
            {
                using (FileStream f = new FileStream(_FILEPATH, FileMode.Append, FileAccess.Write))
                using (StreamWriter s = new StreamWriter(f))
                    s.WriteLine("EmployeeID;EmployeeName;WageType;WageTypeDescription;ThisPayHours;ThisPayAmount;MTDHours;MTDAmount;YTDHours;YTDAmount");
            }
            try
            {
                using (StreamReader sr = new StreamReader(@"C:\Worker's com\SAP\F1 20 2013_2.txt"))
                {
                    resultTextBlock = sr.ReadToEnd();
                }
                resultArray = resultTextBlock.Split(stringSeparators, StringSplitOptions.None);
            }
            catch (Exception e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }

            for (int i = 1; i < resultArray.Length; i++)
            {
                processTextBlock(resultArray[i].ToString());
            }
            using (StreamWriter sw = File.AppendText(_FILEPATH))
            {
                sw.WriteLine("Total Employees in Report = " + _NUMEMPLOYEES.ToString());
            }


        }

        static private void initialiseDataSet()
        {
            _DS = new DataSet("ResultDS");
            DataColumn dc1 = new DataColumn("EmployeeID", typeof(string));
            DataColumn dc2 = new DataColumn("EmployeeName", typeof(string));
            DataColumn dc3 = new DataColumn("WageType", typeof(string));
            DataColumn dc4 = new DataColumn("WageTypeDescription", typeof(string));
            DataColumn dc5 = new DataColumn("ThisPayHours", typeof(decimal));
            DataColumn dc6 = new DataColumn("ThisPayAmount", typeof(decimal));
            DataColumn dc7 = new DataColumn("MTDHours", typeof(decimal));
            DataColumn dc8 = new DataColumn("MTDAmount", typeof(decimal));
            DataColumn dc9 = new DataColumn("YTDHours", typeof(decimal));
            DataColumn dc10 = new DataColumn("YTDAmount", typeof(decimal));

            DataTable dt = new DataTable("Wages");
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            dt.Columns.Add(dc3);
            dt.Columns.Add(dc4);
            dt.Columns.Add(dc5);
            dt.Columns.Add(dc6);
            dt.Columns.Add(dc7);
            dt.Columns.Add(dc8);
            dt.Columns.Add(dc9);
            dt.Columns.Add(dc10);
            _DS.Tables.Add(dt);

        }

        static private void writeDataSetToXML()
        {
            System.IO.StreamWriter xmlSW = new System.IO.StreamWriter("ResultDS.xml");
            _DS.WriteXml(xmlSW, XmlWriteMode.WriteSchema);
            xmlSW.Close();
        }

        static private void processTextBlock(string input)
        {
            string[] resultArray = null;
            string[] stringSeparators = null;
            //resultArray = input.Split(stringSeparators, StringSplitOptions.None);
            string Result = "";
            // bool flagInitialHeaderRow = false;
            bool flagWages = false;
            string employeeName = "";
            string employeeId = "";

            //input is a report for each employee
            if (input.Contains("Report: RPLDETQ1"))
            {
                stringSeparators = new string[] { "\n" };
                resultArray = input.Split(stringSeparators, StringSplitOptions.None);

                foreach (string str in resultArray)
                {
                    if (Regex.IsMatch(str, @"\d{8} "))
                    {
                        Result += ("EmployeeID & Name are: " + str);
                        employeeId = str.Substring(0, 8);
                        employeeName = str.Substring(9).Replace("\r", "");
                        //initial = true;
                        _NUMEMPLOYEES++;
                    }
                    if (str.Contains("Retro Indicator") && flagWages == false)
                    {
                        flagWages = true;
                        string text = str.Replace("\t\t\t\t\t", "\t");
                        text = text.Replace("\t\t\t\t", "\t");
                        text = text.Replace("\t\t\t", "\t");
                        text = text.Replace("\t\t", "\t");
                        text = text.Replace("\t", ";");
                        text = text.Replace("  ", "");
                        Result += text;

                    }
                    if (flagWages == true)//&& !str.Contains("Retro Indicator"))
                    {
                        string text = str.Replace("\t\t\t\t\t", "\t");
                        text = text.Replace("\t\t\t\t", "\t");
                        text = text.Replace("\t\t\t", "\t");
                        text = text.Replace("\t\t", "\t");
                        text = text.Replace("\t", ";");
                        text = text.Replace("  ", "");
                        int columns = text.Count(f => f == ';');
                        if (columns == 8)
                        {
                            Result += text;
                            //load the data into a data set
                            text = text.Replace("\r", "");
                            string[] textArr = text.Split(';');
                            DataRow dr = _DS.Tables["Wages"].NewRow();

                            dr["EmployeeID"] = employeeId;
                            dr["EmployeeName"] = employeeName;
                            dr["WageType"] = textArr[1];
                            dr["WageTypeDescription"] = textArr[2];
                            dr["ThisPayHours"] = Decimal.Parse(textArr[3]);
                            dr["ThisPayAmount"] = Decimal.Parse(textArr[4]);
                            dr["MTDHours"] = Decimal.Parse(textArr[5]);
                            dr["MTDAmount"] = Decimal.Parse(textArr[6]);
                            dr["YTDHours"] = Decimal.Parse(textArr[7]);
                            dr["YTDAmount"] = Decimal.Parse(textArr[8]);
                            _DS.Tables["Wages"].Rows.Add(dr);

                            using (StreamWriter sw = File.AppendText(_FILEPATH))
                            {
                                sw.WriteLine(employeeId + ";" + employeeName + ";" + textArr[1] + ";" + textArr[2] + ";" + textArr[3] + ";" + textArr[4] + ";" + textArr[5] + ";" + textArr[6] + ";" + textArr[7] + ";" + textArr[8] + ";");
                            }
                        }
                    }
                }

            }
            //Get the Data in the tables




        }

        /*
        static private void ReadExcel_YTD_Wcomp()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(_YTD_WCOMP_2_XLS_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explicit cast is not required here
            lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

        }*/
    }
}




/*
 static private void ReadExcel()
    {
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Worker's com\SAP\F1 20 2013.xls", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
        Excel.Range xlRange = xlWorksheet.UsedRange;

        int rowCount = xlRange.Rows.Count;
        int colCount = xlRange.Columns.Count;

        for (int i = 1; i <= 50; i++)
        {
            for (int j = 1; j <= colCount; j++)
            {
                try
                {

                    if (!xlWorksheet.Cells[i, j].Equals(null))
                    {
                        Console.WriteLine(xlWorksheet.Cells[i, j].Value.ToString());
                    }
                }
                catch (Exception e)
                {
                    //throw new Exception(e.ToString());
                }


            }
        }
    
    }
*/