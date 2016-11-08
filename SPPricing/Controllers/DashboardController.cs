using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data.Objects;
using DotNet.Highcharts;
using DotNet.Highcharts.Enums;
using DotNet.Highcharts.Options;
using DotNet.Highcharts.Helpers;
using System.Data;
using System.Data.SqlClient;


namespace SPPricing.Controllers
{
    public class DashboardController : Controller
    {
        //
        // GET: /Dashboard/
        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        public ActionResult Index(DashboardModel objDashboardModel)
        {
            LoginController objLoginController = new LoginController();
            try
            {

                if (ValidateSession())
                {
                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "DI");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    objDashboardModel.IsPricerCountChart = true;
                    objDashboardModel.IsPricerBarChart = true;

                    #region Pie Chart For Blacksholes Pricers

                    ObjectResult<PricerCountPieResult> objPricerCountPieResult = objSP_PRICINGEntities.FETCH_PRICE_COUNT_PIE_CHART();
                    List<PricerCountPieResult> PricerCountPieResultList = objPricerCountPieResult.ToList();

                    List<KeyValuePair<string, string>> pricerCounts = new List<KeyValuePair<string, string>>();
                    foreach (var a in PricerCountPieResultList)
                    {
                        pricerCounts.Add(new KeyValuePair<string, string>(a.PricerType.ToString(), a.Count.ToString()));
                    }

                    List<object> ChartValues1 = new List<object>();
                    foreach (var item in pricerCounts)
                    {
                        ChartValues1.Add(new object[] { item.Key, item.Value });
                    }

                    //instanciate an object of the Highcharts type
                    Highcharts pricerChart = new Highcharts("pricer")
                        //define the type of chart 
                        .InitChart(new Chart { DefaultSeriesType = ChartTypes.Pie, PlotShadow = false })

                        //overall Title of the chart 
                     .SetTitle(new Title { Text = "Blackscholes Pricers " })
                        //small label below the main Title
                     .SetSubtitle(new Subtitle { Text = "Count" })
                     .SetTooltip(new Tooltip
                     {
                         PointFormat = "{series.name}: <b>{point.percentage:.1f}%</b>",
                         Enabled = true
                         //Formatter = @"function() { return '<b>'+ this.series.name +'</b><br/>'+ this.x +': '+ this.y; }"
                     })
                     .SetPlotOptions(new PlotOptions
                     {
                         Pie = new PlotOptionsPie
                         {
                             AllowPointSelect = true,
                             DataLabels = new PlotOptionsPieDataLabels
                             {
                                 Enabled = true
                             },
                             EnableMouseTracking = true,
                             //ShowInLegend = true
                         }
                     })

                      .SetSeries(new Series
                      {
                          //Type = ChartTypes.Line,
                          Name = "Count",
                          Data = new Data(ChartValues1.ToArray())

                      });



                    objDashboardModel.PricerCountChart = pricerChart;
                    #endregion


                    #region Pie Chart For Blacksholes Pricers


                    DataSet dsResult = new DataSet();
                    dsResult = General.ExecuteDataSet("PRICER_STATUS_FOR_BAR_CHART");

                    //dsResult.DataSetName = "Test1";

                    List<KeyValuePair<string, string>> pricerCounts1 = new List<KeyValuePair<string, string>>();
                    for (int i = 0; i < dsResult.Tables[0].Rows.Count; i++)
                    {
                        pricerCounts1.Add(new KeyValuePair<string, string>(dsResult.Tables[0].Rows[i]["STATUS"].ToString(), dsResult.Tables[0].Rows[i]["COUNT"].ToString()));
                    }

                    List<object> Pending = new List<object>();
                    List<object> Confirmed = new List<object>();
                    List<object> Expired = new List<object>();
                    List<object> Cancel = new List<object>();
                    foreach (var item in pricerCounts1)
                    {

                        if (item.Key == "Pending For Approval")
                            Pending.Add(new object[] { item.Value });

                        if (item.Key == "Confirmed")
                            Confirmed.Add(new object[] { item.Value });

                        if (item.Key == "Cancelled")
                            Cancel.Add(new object[] { item.Value });

                        if (item.Key == "Expired")
                            Expired.Add(new object[] { item.Value });
                    }


                    var transactionCounts = new List<Graph>();

                    if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[1].Rows.Count > 0)
                    {
                        foreach (DataRow dr in dsResult.Tables[1].Rows)
                        {
                            transactionCounts.Add(new Graph() { PricerType = Convert.ToString(dr["PRICERTYPE"]) });
                        }
                    }

                    var xPricer = transactionCounts.Select(i => i.PricerType).ToArray();

                    //instanciate an object of the Highcharts type
                    Highcharts BarChart = new Highcharts("pricerbar")
                        //define the type of chart 
                        .InitChart(new Chart { DefaultSeriesType = ChartTypes.Column })
                        //overall Title of the chart 
                     .SetTitle(new Title { Text = "Blackscholes Pricers " })
                        //small label below the main Title
                     .SetSubtitle(new Subtitle { Text = "Count" })
                      .SetXAxis(new XAxis { Title = new XAxisTitle { Text = "Pricer Types" }, Categories = xPricer })

                     //.SetTooltip(new Tooltip
                        //{
                        //    HeaderFormat = "<span style=font-size:11px>{series.name}</span><br>",
                        //    ////PointFormat = "<span style=color:{point.color}>{point.name}</span>: <b>{point.y:.2f}</b> of total<br/>",
                        //    //PointFormat = "{series.name}: <b>{point.percentage:.1f}%</b>",
                        //    Enabled = true

                     //})
                     .SetPlotOptions(new PlotOptions
                     {
                         Column = new PlotOptionsColumn
                         {
                             AllowPointSelect = true,
                             DataLabels = new PlotOptionsColumnDataLabels
                             {
                                 Enabled = false
                             }
                             // EnableMouseTracking = false,
                         }
                     })

                      .SetSeries(new[]
              {                     
        
                  new Series {Name = "Pending For Approval",Color=System.Drawing.Color.Orange ,Data = new Data(Pending.ToArray())},
                      //you can add more y data to create a second line
                  new Series { Name ="Confirmed",Color=System.Drawing.Color.Green, Data = new Data(Confirmed.ToArray()) },

                  new Series { Name ="Cancelled",Color=System.Drawing.Color.Red, Data = new Data(Cancel.ToArray()) },

                  new Series { Name ="Expired", Data = new Data(Expired.ToArray()) }
              });

                    objDashboardModel.PricerBarChart = BarChart;
                    #endregion

                    #region Fetch UnderlyingList

                    List<Underlying> UnderlyingList = new List<Underlying>();

                    DataSet dsResult2 = new DataSet();
                    dsResult2 = General.ExecuteDataSet("FETCH_UNDERLYING_CREATION_TICKER");

                    if (dsResult2 != null && dsResult2.Tables.Count > 0 && dsResult2.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dr in dsResult2.Tables[0].Rows)
                        {
                            Underlying obj = new Underlying();

                            obj.UnderlyingName = Convert.ToString(dr["UnderlyingName"]);
                            obj.UnderlyingType = Convert.ToString(dr["UnderlyingType"]);
                            obj.PriceDiff = Convert.ToString(dr["PriceDiff"]);
                            obj.PriceDiffInPercentage = Convert.ToString(dr["PriceDiffInPercentage"]);
                            obj.NewDate = Convert.ToString(dr["NewDate"]);
                            obj.NewPrice = Convert.ToString(dr["NewPrice"]);

                            UnderlyingList.Add(obj);
                        }
                    }
                    objDashboardModel.UnderlyingList = UnderlyingList;
                    #endregion

                    return View(objDashboardModel);
                }
                else
                {
                    return RedirectToAction("Login", "Login");
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "Dashboard Index Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchPricerPendingList(string SearchFlag)
        {
            try
            {
                //string SearchFlag = "";

                List<DashboardModel> DashboardModelList = new List<DashboardModel>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("PRICER_STATUS_FOR_LINE_CHART", Convert.ToInt32(SearchFlag));

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        DashboardModel obj = new DashboardModel();

                        obj.PricerType = Convert.ToString(dr["PRICERTYPE"]);
                        obj.Count = Convert.ToInt32(dr["COUNT"]);

                        DashboardModelList.Add(obj);
                    }
                }

                var DashboardModelListData = DashboardModelList.ToList();
                return Json(DashboardModelList, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json("");
            }
        }

        public void LogError(string strErrorDescription, string strStackTrace, string strClassName, string strMethodName, Int32 intUserId)
        {
            SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
            var Count = objSP_PRICINGEntities.SP_ERROR_LOG(strErrorDescription, strStackTrace, strClassName, strMethodName, intUserId);
        }

        public bool ValidateSession()
        {
            LoginController objLoginController = new LoginController();

            try
            {
                if (Session["LoggedInUser"] != null)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                objLoginController.LogError(ex.Message, ex.StackTrace, "DashboardController", "ValidateSession", objUserMaster.UserID);
                return false;
            }
        }

        ////public void DataSetsToExcel(List<DataSet> dataSets, string fileName)
        //public void DataSetsToExcel(DataSet dataSets, string fileName)
        //{
        //    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
        //    Microsoft.Office.Interop.Excel.Sheets xlSheets = null;
        //    Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = null;

        //    Int32 intMaxRowCount = 30;

        //    //for (int cnt = 1; cnt <= intMaxRowCount; cnt++)
        //    //{
        //    //    DataRow dataRow;
        //    //    dataRow = dataSets[0].Tables[0].NewRow();
        //    //    dataRow["ROW_NUM"] = cnt;
        //    //}

        //    var a = dataSets.Tables[0].Rows.Count;
        //    var b = dataSets.Tables[0].Rows.Count;

        //    a = a / intMaxRowCount;
        //    b = b % intMaxRowCount;

        //    if (b != 0)
        //        a = a + 1;

        //    DataTable dtFirst = null;

        //    if (dataSets.Tables[0].Rows.Count > intMaxRowCount)
        //    {
        //        for (int i = 0; i < a; i++)
        //        {
        //            DataView dvFirst = new DataView(dataSets.Tables[0], "ROW_NUM>" + (i * intMaxRowCount) + " AND ROW_NUM<=" + ((i + 1) * intMaxRowCount), "ROW_NUM", DataViewRowState.CurrentRows);
        //            dtFirst = dvFirst.ToTable();

        //            System.Data.DataTable dataTable = dtFirst;
        //            dtFirst.Columns.Remove("ROW_NUM");

        //            int rowNo = dataTable.Rows.Count;
        //            int columnNo = dataTable.Columns.Count;
        //            int colIndex = 0;

        //            //Create Excel Sheets
        //            xlSheets = xlWorkbook.Sheets;
        //            xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
        //            xlWorksheet.Name = "AutocallSheet" + (i + 1).ToString();

        //            //Generate Field Names
        //            foreach (DataColumn dataColumn in dataTable.Columns)
        //            {
        //                colIndex++;
        //                xlApp.Cells[1, colIndex] = dataColumn.ColumnName;
        //            }

        //            object[,] objData = new object[rowNo, columnNo];

        //            //Convert DataSet to Cell Data
        //            for (int row = 0; row < rowNo; row++)
        //            {
        //                for (int col = 0; col < columnNo; col++)
        //                {
        //                    objData[row, col] = dataTable.Rows[row][col];
        //                }
        //            }

        //            //Add the Data
        //            Microsoft.Office.Interop.Excel.Range range = xlWorksheet.Range[xlApp.Cells[2, 1], xlApp.Cells[rowNo + 1, columnNo]];
        //            range.Value2 = objData;

        //            //Format Data Type of Columns 
        //            colIndex = 0;
        //            foreach (DataColumn dataColumn in dataTable.Columns)
        //            {
        //                colIndex++;
        //                string format = "@";
        //                switch (dataColumn.DataType.Name)
        //                {
        //                    case "Boolean":
        //                        break;
        //                    case "Byte":
        //                        break;
        //                    case "Char":
        //                        break;
        //                    case "DateTime":
        //                        format = "dd/mm/yyyy";
        //                        break;
        //                    case "Decimal":
        //                        format = "$* #,##0.00;[Red]-$* #,##0.00";
        //                        break;
        //                    case "Double":
        //                        break;
        //                    case "Int16":
        //                        format = "0";
        //                        break;
        //                    case "Int32":
        //                        format = "0";
        //                        break;
        //                    case "Int64":
        //                        format = "0";
        //                        break;
        //                    case "SByte":
        //                        break;
        //                    case "Single":
        //                        break;
        //                    case "TimeSpan":
        //                        break;
        //                    case "UInt16":
        //                        break;
        //                    case "UInt32":
        //                        break;
        //                    case "UInt64":
        //                        break;
        //                    default: //String
        //                        break;
        //                }
        //                //Format the Column accodring to Data Type
        //                xlWorksheet.Range[xlApp.Cells[2, colIndex], xlApp.Cells[rowNo + 1, colIndex]].NumberFormat = format;
        //            }
        //        }
        //    }

        //    //Remove the Default Worksheet
        //    ((Microsoft.Office.Interop.Excel.Worksheet)xlApp.ActiveWorkbook.Sheets[xlApp.ActiveWorkbook.Sheets.Count]).Delete();

        //    //Save
        //    xlWorkbook.SaveAs(fileName,
        //        System.Reflection.Missing.Value,
        //        System.Reflection.Missing.Value,
        //        System.Reflection.Missing.Value,
        //        System.Reflection.Missing.Value,
        //        System.Reflection.Missing.Value,
        //        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
        //        System.Reflection.Missing.Value,
        //        System.Reflection.Missing.Value,
        //        System.Reflection.Missing.Value,
        //        System.Reflection.Missing.Value,
        //        System.Reflection.Missing.Value);

        //    xlWorkbook.Close();
        //    xlApp.Quit();
        //    GC.Collect();
        //}

    }
}



