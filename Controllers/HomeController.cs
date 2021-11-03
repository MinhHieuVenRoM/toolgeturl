using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using ToolGetThumb.Models;

namespace ToolGetThumb.Controllers
{
    public class HomeController : Controller
    {
        private FileStream stream;
        private static List<String> DSIdNewsError = new List<string>();
        public ActionResult Index()
        {
            ViewData["dataerror"] = DSIdNewsError;
            return View();

        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase FileUpload)
        {
           
            ImportExcel_quickly(FileUpload);
            ViewData["dataerror"] = DSIdNewsError;
            return View();
        }
        public int ImportExcel_quickly(HttpPostedFileBase FileUpload  )
        {
            int rowexecl = 8;
            try
            {
                string filePath = "";
                if (FileUpload != null && FileUpload.ContentLength > 0)
                    try
                    {
                        filePath = Path.Combine(Server.MapPath("~/Content/templates/"),
                                                   Path.GetFileName(FileUpload.FileName));
                        FileUpload.SaveAs(filePath);
                        var Path_Excel = Path.GetFileName(FileUpload.FileName);
                        // Debug.WriteLine("\n +++++++++EXCEL: " + selloutaeon_.Excel);
                        var Excel_Path = "/Content/templates/" + Path.GetFileName(FileUpload.FileName);
                        // Debug.WriteLine("\n +++++++++EXCEL_PATH: " + selloutaeon_.Excel_Path);
                    }
                    catch (Exception ex)
                    {

                        ViewBag.msg = ex.Message.ToString();
                        return 2;
                    }
                else
                {
                    ViewBag.msg = "Không có file Upload ";
                    return 2;
                }

                stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader;

                //1. Reading Excel file
                if (Path.GetExtension(filePath).ToUpper() == ".XLS")
                {
                    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else
                {
                    //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                //2. DataSet - The result of each spreadsheet will be created in the result.Tables
                DataSet result = excelReader.AsDataSet();

                //3. DataSet - Create column names from first row
                // excelReader.IsFirstRowAsColumnNames = true; 

                DataTable dataRecords = new DataTable();
                dataRecords = result.Tables[0];
                dataRecords.Rows[0].Delete();
                dataRecords.AcceptChanges();
                List<PostModel> reqData = new List<PostModel>();
                if (dataRecords.Rows.Count > 0)
                {

                    foreach (DataRow row in dataRecords.Rows)
                    {
                        if (row[0] != null && row[0].ToString().Trim().Length > 0)
                        {
                            PostModel exceldata = new PostModel();
                            rowexecl++;
                            exceldata.NEWSID = row[0].ToString();
                            exceldata.THUMBNAILIMAGE = row[1].ToString();
                            exceldata.DETAILIMAGE = row[2].ToString();
                            exceldata.TITLE = row[3].ToString();
                            exceldata.URL = row[4].ToString();
                            exceldata.METATITLE = row[5].ToString();
                            exceldata.METAKEYWORD = row[6].ToString();
                            exceldata.METADESCRIPTION = row[7].ToString();
                            exceldata.CREATEDDATE = Convert.ToDateTime(row[8]);
                            exceldata.CREATEDUSER = row[9].ToString();
                            exceldata.CREATEDCUSTOMERID = row[10].ToString();
                            exceldata.TAGS = row[11].ToString();
                            exceldata.LISTCATEGORYID = row[12].ToString();
                            exceldata.LISTCATEGORYNAME = row[13].ToString();
                            reqData.Add(exceldata);

                            String Url = "https://cdn.tgdd.vn/Files/";

                            string Year = exceldata.CREATEDDATE.Year.ToString();
                            string month = exceldata.CREATEDDATE.Month.ToString();
                            if (month.Length == 1)
                            {
                                month = "0" + month;

                            }


                            string day = exceldata.CREATEDDATE.Day.ToString();
                            if (day.Length == 1)
                            {
                                day = "0" + day;

                            }

                            Url += Year + "/"+month + "/"+day + "/"+exceldata.NEWSID+"/";
                            if (!exceldata.DETAILIMAGE.Equals("null"))
                            {
                                Url += exceldata.DETAILIMAGE;
                                GetPage(Url, exceldata.NEWSID);

                            }
                            if(exceldata.DETAILIMAGE.Equals("null") && !exceldata.THUMBNAILIMAGE.Equals("null"))

                            {
                                Url += exceldata.THUMBNAILIMAGE+ "_300x300";
                                GetPage(Url, exceldata.NEWSID);
                            }




                        }
                    }

                }
                stream.Close();
                stream.Dispose();
                ViewBag.msg = DSIdNewsError;
                return 1;

            }
            catch (Exception e)
            {
                stream.Close();
                stream.Dispose();
                ViewBag.msg = "File bị lỗi, vui lòng kiểm tra lại file , chú ý dòng " + rowexecl;
                return 2;
            }
        }

        public static void GetPage(String url,string idNews)
        {
            try
            {
                // Creates an HttpWebRequest for the specified URL.
                System.Net.HttpWebRequest myHttpWebRequest = (System.Net.HttpWebRequest)WebRequest.Create(url);
                // Sends the HttpWebRequest and waits for a response.
                HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();
                if (myHttpWebResponse.StatusCode == HttpStatusCode.OK)
                    System.Diagnostics.Debug.WriteLine("\r\nStatus Code is OK"+url);
                else
                {
                    DSIdNewsError.Add(idNews);
                }

                // Releases the resources of the response.
                myHttpWebResponse.Close();
            }
            catch (WebException e)
            {
                
                DSIdNewsError.Add(idNews);
                System.Diagnostics.Debug.WriteLine("\r\n Error", url);
            }
            catch (Exception e)
            {
                DSIdNewsError.Add(idNews);
                System.Diagnostics.Debug.WriteLine("\r\n Error", url);
            }
        }
    }
}