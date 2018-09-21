using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace sisysmtracker
{
    class siSysData
    {
        public int ODB { get; set; }
        public int EWB { get; set; }
        public int EDB { get; set; }
        public int FANAVF { get; set; }
        public double TotalCAP { get; set; }

        public double SenseCAP { get; set; }

        public double InputPWR { get; set; }

        public override string ToString()
        {

            return ODB.ToString() + EWB.ToString() + EDB.ToString() + FANAVF.ToString() + TotalCAP.ToString() + SenseCAP.ToString()
                + InputPWR.ToString();
        }
    }

    class HPData
    {
        public int ODB { get; set; }

        public string EDB { get; set; }
        public int FANAVF { get; set; }
        public double TotalCAP { get; set; }



        public double InputPWR { get; set; }

        public override string ToString()
        {

            return ODB.ToString() + EDB.ToString() + FANAVF.ToString() + TotalCAP.ToString()
                + InputPWR.ToString();
        }
    }
    class Program
    {
        public static void ReadExistingExcel(Dictionary<string, siSysData> siDic, string siSysmID, int workSheetNo)
        {
            String fileName = @"C:\Users\rxn14\Downloads\sis.xlsx";


            Application xlApp = new Application();
            Workbook xlWorkBook = default(Workbook);
            Worksheet xlWorkSheet = default(Worksheet);
            object misValue = System.Reflection.Missing.Value;

            try
            {


                if (xlApp == null)
                {
                    Console.WriteLine("Excel is not installed in the system...");
                    return;
                }


                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", false,
                XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(workSheetNo);


                siSysmID = siSysmID.Replace(";", "");

                // Range xlRange = xlWorkSheet.UsedRange;

                int rowCount = 14;//xlRange.Rows.Count + 1;
                foreach (KeyValuePair<string, siSysData> kvp in siDic)
                {




                    xlWorkSheet.Cells[rowCount, 1] = siSysmID.Replace(";", "");
                    xlWorkSheet.Cells[rowCount, 3] = kvp.Value.ODB;
                    xlWorkSheet.Cells[rowCount, 4] = kvp.Value.EWB;
                    xlWorkSheet.Cells[rowCount, 5] = kvp.Value.EDB;
                    xlWorkSheet.Cells[rowCount, 6] = kvp.Value.FANAVF;
                    xlWorkSheet.Cells[rowCount, 7] = kvp.Value.TotalCAP;
                    xlWorkSheet.Cells[rowCount, 8] = kvp.Value.SenseCAP;
                    xlWorkSheet.Cells[rowCount, 9] = kvp.Value.InputPWR;



                    rowCount++;
                }
                xlApp.DisplayAlerts = false;




            }
            catch (Exception ex)
            {

            }
            finally
            {
                xlWorkBook.SaveAs(fileName, XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue,
                XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        public static void ReadExistingExcel(Dictionary<string, HPData> siDic, string siSysmID, int workSheetNo)
        {
            String fileName = @"C:\Users\rxn14\Downloads\sis.xlsx";


            Application xlApp = new Application();
            Workbook xlWorkBook = default(Workbook);
            Worksheet xlWorkSheet = default(Worksheet);
            object misValue = System.Reflection.Missing.Value;

            try
            {


                if (xlApp == null)
                {
                    Console.WriteLine("Excel is not installed in the system...");
                    return;
                }


                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", false,
                XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(workSheetNo);


                siSysmID = siSysmID.Replace(";", "");

                // Range xlRange = xlWorkSheet.UsedRange;

                int rowCount = 14;//xlRange.Rows.Count + 1;
                foreach (KeyValuePair<string, HPData> kvp in siDic)
                {




                    xlWorkSheet.Cells[rowCount, 1] = siSysmID.Replace(";", "");
                    xlWorkSheet.Cells[rowCount, 3] = kvp.Value.ODB;

                    xlWorkSheet.Cells[rowCount, 4] = kvp.Value.EDB;
                    xlWorkSheet.Cells[rowCount, 5] = kvp.Value.FANAVF;
                    xlWorkSheet.Cells[rowCount, 6] = kvp.Value.TotalCAP;

                    xlWorkSheet.Cells[rowCount, 7] = kvp.Value.InputPWR;



                    rowCount++;
                }
                xlApp.DisplayAlerts = false;




            }
            catch (Exception ex)
            {

            }
            finally
            {
                xlWorkBook.SaveAs(fileName, XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue,
                XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                xlWorkBook.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        public static void UnitConfig(string connectionString, int workSheetNo, string family, string baseunit)
        {
            string queryString =
              "SELECT top 1 * FROM UnitConf WHERE ( UnitConf.family = '@family') " +
               " AND ( UnitConf.baseunit = '@baseunit') and Stage in (3,2,1) ORDER BY createdate DESC, Stage desc";

            string[] rtLines = new string[12];
            string siSysmID = "";

            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {

                // create a SqlCommand object for this connection
                queryString = queryString.Replace("@family", family)
                    .Replace("@baseunit", baseunit);
                SqlCommand command = new SqlCommand(queryString, connection);

                //   command.Parameters.AddWithValue("@family", "XC25-036-230-01");
                //  command.Parameters.AddWithValue("@baseunit", "XC25-036-230-01 - CR33-48B-F + SL280DF090V48B-3");
                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        Console.WriteLine("\t{0}\t{1}\t{2}",
                            reader[0], reader[1], reader[2]);
                        for (int rtLineData = 1; rtLineData <= 12; rtLineData++)
                        {
                            rtLines[rtLineData - 1] = reader["RTLine" + rtLineData.ToString()].ToString();
                        }
                        siSysmID = reader["sysimInputsID"].ToString();
                    }
                    reader.Close();
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            Dictionary<string, siSysData> siDic = new Dictionary<string, siSysData>();
            foreach (string rtLine in rtLines)
            {
                if (rtLine == null || rtLine == "")
                    continue;

                string[] rtArray = rtLine.Split(' ');
                int[] sensible = { 75, 80, 85 };
                int[] totalValue = { 85, 95, 105, 115, 125 };
                int arrCount = 2;
                foreach (int totalIndex in totalValue)
                {

                    int arrSensible = 0;
                    foreach (int sensibleIndex in sensible)
                    {
                        siSysData sd = new siSysData();
                        sd.ODB = totalIndex;
                        if (rtArray[0] != null && rtArray[0] != "")
                            sd.EWB = int.Parse(rtArray[0]);

                        sd.EDB = sensibleIndex;
                        if (rtArray[1] != null && rtArray[1] != "")
                            sd.FANAVF = int.Parse(rtArray[1]);
                        sd.TotalCAP = (double.Parse(rtArray[arrCount]) * 1000);
                        sd.SenseCAP = (double.Parse(rtArray[arrCount + arrSensible + 2]) * sd.TotalCAP);
                        sd.InputPWR = double.Parse(rtArray[arrCount + 1]);

                        if (!siDic.ContainsKey(sd.ToString()))
                            siDic.Add(sd.ToString(), sd);

                        arrSensible += 1;
                    }
                    arrCount += 5;
                }
            }
            if (siDic.Count > 0)
                ReadExistingExcel(siDic, siSysmID, workSheetNo);
        }

        public static void UnitConfigHPDate(string connectionString, int workSheetNo, string family, string baseunit)
        {
            string queryString =
            "select top 1 ho.*, hd.SysimInputsID from hsout ho inner join hpdata hd on ho.HeatID = hd.HeatID " +
              "and hd.family = '@family' and hd.baseunit = '@baseunit'  order by HeatID desc ";
            string siSysmID = "";
            Dictionary<string, HPData> siDic = new Dictionary<string, HPData>();

            using (SqlConnection connection =
                new SqlConnection(connectionString))
            {

                // create a SqlCommand object for this connection
                queryString = queryString.Replace("@family", family)
                    .Replace("@baseunit", baseunit);
                SqlCommand command = new SqlCommand(queryString, connection);

                //   command.Parameters.AddWithValue("@family", "XC25-036-230-01");
                //  command.Parameters.AddWithValue("@baseunit", "XC25-036-230-01 - CR33-48B-F + SL280DF090V48B-3");
                try
                {
                    connection.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    string[] oddryblubs = { "65", "45", "25", "05", "15" };
                    string[] cfms = { "hcfm", "lcfm", "mcfm" };

                    while (reader.Read())
                    {
                        Console.WriteLine("\t{0}\t{1}\t{2}",
                            reader[0], reader[1], reader[2]);
                        foreach (string cfm in cfms)
                        {

                            foreach (string oddryblub in oddryblubs)
                            {
                                HPData sd = new HPData();
                                sd.ODB = 70;
                                sd.EDB = oddryblub;
                                sd.FANAVF = int.Parse(reader[cfm].ToString());
                                sd.TotalCAP = (double.Parse(reader[cfm.Substring(0, 1) + oddryblub + "btu"].ToString()) * 1000);

                                sd.InputPWR = double.Parse(reader[cfm.Substring(0, 1) + oddryblub + "watt"].ToString());

                                if (!siDic.ContainsKey(sd.ToString()))
                                    siDic.Add(sd.ToString(), sd);
                            }
                        }
                        siSysmID = reader["sysimInputsID"].ToString();
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            //string rtLine = "59 875 31.4 1.80 .89 1.00 1.00 30.2 2.09 .91 1.00 1.00 29.0 2.39 .93 1.00 1.00 27.8 2.73 .95 1.00 1.00 26.4 3.08 .98 1.00 1.00";
            ReadExistingExcel(siDic, siSysmID, workSheetNo);
        }

        static void Main(string[] args)
        {
            string connectionString =
           "Data Source=RCHSQL8P1,1488;Persist Security Info=True;Initial Catalog=CoolingAndHeatPumpRatings;User Id=CoolingAndHeat;PASSWORD=C00l1ng@Heat;";


            UnitConfig(connectionString, 2, "XP21-060-230-06", "XP21-060-230-06 - CHX35-60D-6F - TXV");
            UnitConfigHPDate(connectionString, 3, "XP21-060-230-06", "XP21-060-230-06 - CHX35-60D-6F - TXV");

        }
    }
}
