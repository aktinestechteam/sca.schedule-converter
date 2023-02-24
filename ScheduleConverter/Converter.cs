using ScheduleConverter.FileExtractor;
using System;
using System.Collections;
using System.Data;
using System.IO;
namespace ScheduleConverter
{
    public class Converter
    {
        private Hashtable _hashFlightes { get; set; }

        public void ConvertSchedule(string filename, string sheetName)
        {
            try
            {
                string col = "Q";
                var dtresult = new DataTable();
                var tables = ExcelExtract.DataReader(filename, sheetName, col);
                if (tables.Count > 0)
                {
                    _hashFlightes = new Hashtable();
                    var colinx = 16;
                    var dt = tables[0];
                    if (dt.Rows.Count > 0 && dt.Columns.Count > 0)
                    {
                        for (int i = 2; i < dt.Rows.Count; i++)
                        {
                            for (int j = 0; j <= colinx; j++)
                            {
                                var value = Convert.ToString(dt.Rows[i][j]).Trim();
                                if (string.IsNullOrEmpty(value))
                                {
                                    dt.Rows[i][j] = dt.Rows[i - 1][j];
                                }
                            }
                        }

                        for (int j = 0; j <= colinx; j++)
                        {
                            dtresult.Columns.Add(Convert.ToString(dt.Rows[0][j]));
                        }

                        dtresult.Columns.Add("FlightDate");
                        dtresult.Columns.Add("Day");
                        dtresult.Columns.Add("StartDate");
                        dtresult.Columns.Add("EndDate");
                        dtresult.Columns.Add("ActStartDate");
                        dtresult.Columns.Add("ActEndDate");
                        dtresult.Columns.Add("Isduplicate");


                        var sheetYear = DateTime.Now.Year;
                        var stindex = colinx + 1;
                        var edindex = dt.Columns.Count - 1;
                        for (int i = 1; i < dt.Rows.Count; i++)
                        {
                            for (int j = stindex; j <= edindex; j++)
                            {
                                var value = Convert.ToString(dt.Rows[i][j]).Trim();
                                if (value == "1")
                                {
                                    sheetYear = insertRow(dtresult, dt, i, colinx, j, sheetYear);
                                }
                            }
                        }
                        ProcessData(dtresult);

                    }


                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nError:" + ex.Message);
                Console.ResetColor();
            }
        }

        private int insertRow(DataTable dtresult, DataTable dt, int i, int colinx, int datacol, int sheetYear)
        {
            dtresult.Rows.Add();
            for (int j = 0; j <= colinx; j++)
            {
                dtresult.Rows[dtresult.Rows.Count - 1][j] = dt.Rows[i][j];
            }
            int stinx = Convert.ToString(dt.Rows[0][datacol]).IndexOf('(') + 1;
            int length = Convert.ToString(dt.Rows[0][datacol]).IndexOf(')') - stinx;

            string[] datesplit1 = Convert.ToString(dt.Rows[0][datacol]).Substring(0, stinx - 2).Split('/');
            int dd1 = Convert.ToInt32(datesplit1[0]);
            int mm1 = Convert.ToInt32(datesplit1[1]);
            if (datacol > colinx + 1)
            {
                string[] datesplit2 = Convert.ToString(dt.Rows[0][datacol - 1]).Substring(0, stinx - 2).Split('/');
                int mm2 = Convert.ToInt32(datesplit2[1]);
                if (mm2 == 12 && mm1 >= 1)
                    sheetYear = sheetYear + 1;
            }
            DateTime date = new DateTime(sheetYear, mm1, dd1);
            int weekDay = Convert.ToInt32(Convert.ToString(dt.Rows[0][datacol]).Substring(stinx, length));
            dtresult.Rows[dtresult.Rows.Count - 1]["FlightDate"] = Convert.ToDateTime(date);
            dtresult.Rows[dtresult.Rows.Count - 1]["Day"] = weekDay;
            dtresult.Rows[dtresult.Rows.Count - 1]["StartDate"] = date.AddDays(4 - (weekDay >= 4 ? weekDay : 7 + weekDay));
            dtresult.Rows[dtresult.Rows.Count - 1]["EndDate"] = date.AddDays(4 - (weekDay >= 4 ? weekDay : 7 + weekDay)).AddDays(6);
            dtresult.Rows[dtresult.Rows.Count - 1]["IsDuplicate"] = 0;
            UpdateFlightsHash(Convert.ToString(dtresult.Rows[dtresult.Rows.Count - 1]["Flt"]), Convert.ToString(dtresult.Rows[dtresult.Rows.Count - 1]["Dep. Time"]), Convert.ToString(dtresult.Rows[dtresult.Rows.Count - 1]["Arr. Time"]), Convert.ToString(dtresult.Rows[dtresult.Rows.Count - 1]["AC Type"]), dtresult.Rows.Count - 1, dtresult);

            return sheetYear;
        }

        private void UpdateFlightsHash(string flight, string deptime, string arrtime, string acttype, int rowindex, DataTable dt)
        {
            string unkey = flight + "$" + deptime + "$" + arrtime + "$" + acttype;
            if (!_hashFlightes.Contains(unkey))
            {
                _hashFlightes.Add(unkey, new ArrayList() { rowindex });
                dt.Rows[rowindex]["ActStartDate"] = dt.Rows[rowindex]["StartDate"];
                dt.Rows[rowindex]["ActEndDate"] = dt.Rows[rowindex]["EndDate"];
            }
            else
            {
                ArrayList hashrowinxarray = _hashFlightes[unkey] as ArrayList;
                hashrowinxarray.Add(rowindex);
                _hashFlightes[unkey] = hashrowinxarray;
                bool flag = false;
                foreach (Int32 arowindex in hashrowinxarray)
                {
                    var dateDiff1 = (Convert.ToDateTime(dt.Rows[rowindex]["StartDate"]) - Convert.ToDateTime(dt.Rows[arowindex]["EndDate"])).TotalDays;
                    var dateDiff2 = (Convert.ToDateTime(dt.Rows[rowindex]["StartDate"]) - Convert.ToDateTime(dt.Rows[arowindex]["ActEndDate"])).TotalDays;

                    if ((dateDiff1 == 1 || dateDiff2 == 1) && Convert.ToString(dt.Rows[rowindex]["Day"]) == Convert.ToString(dt.Rows[arowindex]["Day"]))
                    {
                        var startDate = Convert.ToDateTime(dt.Rows[arowindex]["ActStartDate"]);
                        var endDate = Convert.ToDateTime(dt.Rows[rowindex]["EndDate"]);
                        dt.Rows[arowindex]["ActStartDate"] = startDate;
                        dt.Rows[arowindex]["ActEndDate"] = endDate;
                        dt.Rows[rowindex]["ActStartDate"] = startDate;
                        dt.Rows[rowindex]["ActEndDate"] = endDate;
                        flag = true;
                    }
                    else if (!flag)
                    {
                        dt.Rows[rowindex]["ActStartDate"] = dt.Rows[rowindex]["StartDate"];
                        dt.Rows[rowindex]["ActEndDate"] = dt.Rows[rowindex]["EndDate"];
                    }
                }

            }
        }

        private void ProcessData(DataTable dtResult)
        {
            Hashtable hashKey = new Hashtable();
            for (int i = 0; i < dtResult.Rows.Count; i++)
            {
                string unkey = Convert.ToString(dtResult.Rows[i]["Flt"]) + "$" + Convert.ToString(dtResult.Rows[i]["Dep. Time"]) + "$" + Convert.ToString(dtResult.Rows[i]["Arr. Time"]) + "$" + Convert.ToString(dtResult.Rows[i]["AC Type"]) + "$" + Convert.ToString(dtResult.Rows[i]["ActStartDate"]) + "$" + Convert.ToString(dtResult.Rows[i]["ActEndDate"]);
                if (!hashKey.Contains(unkey))
                {
                    hashKey.Add(unkey, i);
                }
                else
                {
                    var rwindex = Convert.ToInt32(hashKey[unkey]);
                    if (!Convert.ToString(dtResult.Rows[rwindex]["Day"]).Contains(Convert.ToString(dtResult.Rows[i]["Day"])))
                    {
                        dtResult.Rows[rwindex]["Day"] = Convert.ToString(dtResult.Rows[rwindex]["Day"]) + " " + Convert.ToString(dtResult.Rows[i]["Day"]);
                    }
                    dtResult.Rows[i]["IsDuplicate"] = 1;
                }
            }

            DataTable dtClone = dtResult.Clone();
            for (int i = 0; i < dtResult.Rows.Count; i++)
            {
                if (Convert.ToInt32(dtResult.Rows[i]["IsDuplicate"]) == 0)
                {
                    dtClone.Rows.Add(dtResult.Rows[i].ItemArray);
                }
            }
            dtResult = dtClone;
            ExportToCSV(dtResult, @"C:\POC\export.csv");
        }

        public void ExportToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers    
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
    

    }
}
