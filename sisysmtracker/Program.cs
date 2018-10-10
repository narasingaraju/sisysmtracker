using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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

    class sysimID
    {
        public string SyssimValue { get; set; }
    }
    class Program
    {
        private static sysimID sid = new sysimID();
        public static Dictionary<string, string> Update_ASHP(string inPutFile, string siSysmID, int workSheetNo)
        {
            String fileName = @"C:\Users\rxn14\Downloads\HP_Data1.xlsx";
            Dictionary<string, string> ASHP_SrcOut = new Dictionary<string, string>();

            Dictionary<string, string> ASHP_Src = new Dictionary<string, string>();
            ASHP_Src.Add("AHRI Certified Reference Number", "AHRIRefNo");
            ASHP_Src.Add("Outdoor Unit Model Number  (Condenser or Single Package)", "CondenserModel");
            ASHP_Src.Add("Indoor Unit Model Number (Evaporator and/or Air Handler)", "CoilModel");
            ASHP_Src.Add("Furnace Model Number", "FurnModel");
            ASHP_Src.Add("Cooling Capacity (A2) - Single or High Stage (95F),btuh", "Cap95");
            ASHP_Src.Add("EER (A2) - Single or High Stage (95F)", "EER95");
            ASHP_Src.Add("SEER", "SEER");
            ASHP_Src.Add("Heating Capacity (H12) - Single or High Stage (47F),btuh", "Cap47");
            ASHP_Src.Add("Heating COP (H12) - Single or High Stage (47F)", "COP47");
            ASHP_Src.Add("HSPF (Region IV)", "HSPF");
            ASHP_Src.Add("calc1", "HtgPower");
            ASHP_Src.Add("Indoor Full-Load Heating Air Volume Rate (H12 SCFM)", "AVF");
            ASHP_Src.Add("calc2", "PwrRtd");

            ASHP_Src.Add("calc3", "AccCode");
            ASHP_Src.Add("calc4", "UnitType");
            ASHP_Src.Add("calc5", "Voltage");
            ASHP_Src.Add("calc6", "Phase");
            ASHP_Src.Add("calc7", "BaseModelID");


            Application xlApp = new Application();
            Workbook xlWorkBook = default(Workbook);
            Worksheet xlWorkSheet = default(Worksheet);
            object misValue = System.Reflection.Missing.Value;

            try
            {


                if (xlApp == null)
                {
                    Console.WriteLine("Excel is not installed in the system...");
                    return null;
                }


                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, false, 5, "", "", false,
                XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(workSheetNo);


                siSysmID = siSysmID.Replace(";", "");

                // Range xlRange = xlWorkSheet.UsedRange;

                int targetRow = 1;//xlRange.Rows.Count + 1;
                Range xlRange = xlWorkSheet.UsedRange;

                for (int i = targetRow; i <= targetRow/*xlRange.Rows.Count*/; i++)
                {
                    for (int j = 1; j <= xlRange.Columns.Count; j++)
                    {


                        Range currentRange = (Range)xlWorkSheet.Cells[i, j];
                        if (currentRange.Value2 != null)
                        {

                            string curVal = currentRange.Value2.ToString();
                            if (ASHP_Src.ContainsKey(curVal))
                            {
                                currentRange = (Range)xlWorkSheet.Cells[i + 1, j];
                                ASHP_SrcOut.Add(ASHP_Src[curVal], currentRange.Value2.ToString());
                            }
                        }


                    }


                }

                foreach (KeyValuePair<string, string> kvp in ASHP_Src)
                {
                    if (kvp.Key.StartsWith("calc"))
                    {
                        switch (kvp.Key)
                        {
                            case "calc1":
                                ASHP_SrcOut.Add("HtgPower", ((double.Parse(ASHP_SrcOut["Cap47"]) / 1000) /
                                                double.Parse(ASHP_SrcOut["COP47"])).ToString());
                                break;
                            case "calc2":
                                ASHP_SrcOut.Add("PwrRtd", ((double.Parse(ASHP_SrcOut["Cap95"]) / 1000) /
                                                double.Parse(ASHP_SrcOut["EER95"])).ToString());
                                break;
                            case "calc3":
                                ASHP_SrcOut.Add("AccCode", System.Configuration.ConfigurationSettings.AppSettings["AccCode"]);
                                break;
                            case "calc4":
                                ASHP_SrcOut.Add("UnitType", System.Configuration.ConfigurationSettings.AppSettings["UnitType"]);
                                break;
                            case "calc5":
                                ASHP_SrcOut.Add("Voltage", System.Configuration.ConfigurationSettings.AppSettings["Voltage"]);
                                break;
                            case "calc6":
                                ASHP_SrcOut.Add("Phase", System.Configuration.ConfigurationSettings.AppSettings["Phase"]);
                                break;
                            case "calc7":
                                ASHP_SrcOut.Add("BaseModelID", sid.SyssimValue);
                                break;
                                

                        }
                    }
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
            return ASHP_SrcOut;
        }

        public static void Update_ASHP_Excel(Dictionary<string, string> siDic, string siSysmID, int workSheetNo)
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

                int targetRow = 29;//xlRange.Rows.Count + 1;
                Range xlRange = xlWorkSheet.UsedRange;

                for (int i = targetRow; i <= targetRow/*xlRange.Rows.Count*/; i++)
                {
                    for (int j = 1; j <= xlRange.Columns.Count; j++)
                    {

                        Range currentRange = (Range)xlWorkSheet.Cells[i, j];
                        if (currentRange.Value2 != null)
                        {

                            string curVal = currentRange.Value2.ToString();
                            if (siDic.ContainsKey(curVal))
                            {
                                currentRange = (Range)xlWorkSheet.Cells[i + 1, j];
                                // ASHP_SrcOut.Add(ASHP_Src[curVal], 
                                currentRange.Value2 = siDic[curVal];
                            }
                        }


                    }


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
                sid.SyssimValue = siSysmID;
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
                sid.SyssimValue = siSysmID;
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
            Update_ASHP_Excel(Update_ASHP("", "", 1), "systemid", 5);

        }
    }
}
