using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WagnerCable.Models;

namespace WagnerCable.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home
        public ActionResult ImportExcel()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ImportExcel(HttpPostedFileBase PostedFile)
        {
            try
            {

                if (PostedFile.ContentLength > 0)
                {
                    string extension = System.IO.Path.GetExtension(PostedFile.FileName).ToLower();
                    string query = null;
                    string connString = "";
                    string fileName = Guid.NewGuid().ToString();
                    List<ExcelViewModel> list = new List<ExcelViewModel>();
                    List<ExcelViewModel> newList = new List<ExcelViewModel>();


                    string[] validFileTypes = { ".xls", ".xlsx", ".csv" };

                    string path1 = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), fileName);
                    if (!Directory.Exists(path1))
                    {
                        Directory.CreateDirectory(Server.MapPath("~/Content/Uploads"));
                    }
                    if (validFileTypes.Contains(extension))
                    {
                        if (System.IO.File.Exists(path1))
                        {
                            System.IO.File.Delete(path1);
                        }
                        PostedFile.SaveAs(path1);
                        string data = "";
                        //Create COM Objects. Create a COM object for everything that is referenced
                        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path1);
                        Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                        Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                        int rowCount = xlRange.Rows.Count;
                        int colCount = xlRange.Columns.Count;

                        //iterate over the rows and columns and print to the console as it appears in the file
                        //excel is not zero based!!
                        for (int i = 2; i <= rowCount; i++)
                        {
                            ExcelViewModel model = new ExcelViewModel();
                            //either collect data cell by cell or DO you job like insert to DB 
                            if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                            {
                                model.ID = xlRange.Cells[i, 1].Value2.ToString();
                                string a = (xlRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                                double date = double.Parse(a);
                                var dateTime = DateTime.FromOADate(date).ToString("dd.MM.yyyy");
                                model.DATE = dateTime;
                                model.MOQ = Convert.ToInt32(xlRange.Cells[i, 3].Value2.ToString());
                                model.AMOUNT = Convert.ToInt32(xlRange.Cells[i, 4].Value2.ToString());
                            }
                            list.Add(model);
                        }

                        int amount = 0;
                        int rest = 0;
                        bool firstStep = false;
                        bool firstId = false;
                        ExcelViewModel newModel = new ExcelViewModel();

                        for (int i = 0; i < list.Count; i++)
                        {

                            if (i > 0)
                            {

                                if (list[i].ID == list[i - 1].ID)
                                {
                                    if (!firstId && newModel.ID == null)
                                    {
                                        newModel.ID = list[i].ID;
                                        newModel.DATE = list[i].DATE;
                                        newModel.MOQ = list[i].MOQ;
                                        newModel.AMOUNT = list[i].AMOUNT;
                                        firstId = true;
                                    }
                                    if (!firstStep)
                                    {
                                        amount += list[i - 1].AMOUNT;
                                    }
                                    if (list[i - 1].MOQ > amount)
                                    {
                                        amount += list[i].AMOUNT;
                                        if (rest > 0)
                                        {
                                            amount += rest;
                                            rest = 0;
                                        }
                                        if (list[i - 1].MOQ > amount)
                                        {
                                            firstStep = true;
                                            continue;
                                        }
                                        else
                                        {
                                            rest = amount - list[i - 1].MOQ;
                                            newModel.AMOUNT = amount - rest;
                                            newList.Add(newModel);
                                            newModel = new ExcelViewModel();
                                            firstId = false;
                                            amount = 0;
                                            continue;
                                        }
                                    }
                                }
                                else
                                {
                                    if (!firstId)
                                    {
                                        newList[newList.Count - 1].AMOUNT += rest;
                                        
                                    }
                                    else
                                    {
                                        newModel.AMOUNT = amount;
                                        newList.Add(newModel);
                                    }
                                    newModel = new ExcelViewModel();
                                    firstStep = false;
                                    firstId = false;
                                    rest = 0;
                                    amount = 0;
                                    newModel.ID = list[i].ID;
                                    newModel.DATE = list[i].DATE;
                                    newModel.MOQ = list[i].MOQ;
                                    newModel.AMOUNT = list[i].AMOUNT;
                                    continue;
                                }

                            }
                            else
                            {
                                newModel.ID = list[i].ID;
                                newModel.DATE = list[i].DATE;
                                newModel.MOQ = list[i].MOQ;
                                newModel.AMOUNT = list[i].AMOUNT;
                            }
                        }
                        var result = newList;
                        DownloadExcel(result);

                    }
                    else
                    {
                        ViewBag.Error = "Please Upload Files in .xls, .xlsx or .csv format";

                    }

                }

                return View();
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Failed to Upload File";
                return View();
            }
        }

        public void DownloadExcel(List<ExcelViewModel> list)
        {

            ExcelPackage Ep = new ExcelPackage();
            ExcelWorksheet Sheet = Ep.Workbook.Worksheets.Add("Report");
            Sheet.Cells["A1"].Value = "ID";
            Sheet.Cells["B1"].Value = "DATE";
            Sheet.Cells["C1"].Value = "MOQ";
            Sheet.Cells["D1"].Value = "AMOUNT";
            int row = 2;
            foreach (var item in list)
            {

                Sheet.Cells[string.Format("A{0}", row)].Value = item.ID;
                Sheet.Cells[string.Format("B{0}", row)].Value = item.DATE;
                Sheet.Cells[string.Format("C{0}", row)].Value = item.MOQ;
                Sheet.Cells[string.Format("D{0}", row)].Value = item.AMOUNT;
                row++;
            }


            Sheet.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment: filename=" + "Report.xlsx");
            Response.BinaryWrite(Ep.GetAsByteArray());
            Response.End();
        }


    }
}