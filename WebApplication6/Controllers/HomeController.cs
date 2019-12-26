using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;
using System.Web.Hosting;
using System.Web.Mvc;
using ExcelDataReader;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using WebApplication6.Models;

namespace WebApplication6.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";

            return View();
        }


        public ActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Upload(HttpPostedFileBase upload)
        {
            if (ModelState.IsValid)
            {

                if (upload != null && upload.ContentLength > 0)
                {
                    // ExcelDataReader works with the binary Excel file, so it needs a FileStream
                    // to get started. This is how we avoid dependencies on ACE or Interop:
                    Stream stream = upload.InputStream;

                    // We return the interface, so that
                    IExcelDataReader reader = null;


                    if (upload.FileName.EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (upload.FileName.EndsWith(".xlsx"))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        ModelState.AddModelError("File", "This file format is not supported");
                        return View();
                    }

                    reader.AsDataSet();

                    DataSet result = reader.AsDataSet();
                    reader.Close();


                    return View(result.Tables[0]);
                }
                else
                {
                    ModelState.AddModelError("File", "Please Upload Your file");
                }
            }
            return View();
        }



        public ActionResult GetExcellSheetRows()
        {
            return View();
        }

        [HttpPost]
        public List<IRow> GetExcellSheetRows(bool skipFirstRow = true)
        {
            string filePath = "C://Users//Aliaa Yasser//Downloads//DSS Midterm Grades_ToStudents";
            List<IRow> ExcellSheetRowList = new List<IRow>();
            try
            {
                FileStream FS = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                List<ISheet> Sheetlist = new List<ISheet>();
                int totalRowCount = new int();
                #region Check Type Of Excel
                if (filePath.EndsWith(".xls"))
                {
                    #region old excel sheet
                    HSSFWorkbook Workbook = new HSSFWorkbook(FS);
                    for (int i = 0; i < Workbook.Count; i++)
                    {
                        ISheet sheet = Workbook.GetSheetAt(i);
                        Sheetlist.Add(sheet);
                        totalRowCount += sheet.PhysicalNumberOfRows;
                        if (skipFirstRow && sheet.PhysicalNumberOfRows > 1)
                            totalRowCount--;
                    }
                    for (int j = 0; j < Sheetlist.Count; j++)
                    {
                        if (Sheetlist[j].IsActive)
                        {
                            System.Collections.IEnumerator rows = Sheetlist[j].GetRowEnumerator();
                            if (skipFirstRow)
                            {
                                rows.MoveNext();
                            }
                            while (rows.MoveNext())
                            {
                                IRow row = (XSSFRow)rows.Current;
                                ExcellSheetRowList.Add(row);
                            }
                        }
                    }
                    #endregion
                }
                else
                {
                    #region excel 2007 and later
                    FS.Position = 0;
                    XSSFWorkbook Workbook = new XSSFWorkbook(FS);
                    for (int i = 0; i < Workbook.Count; i++)
                    {
                        Sheetlist.Add(Workbook.GetSheetAt(i));
                    }
                    for (int j = 0; j < Sheetlist.Count; j++)
                    {
                        if (Sheetlist[j].IsActive)
                        {
                            System.Collections.IEnumerator rows = Sheetlist[j].GetRowEnumerator();
                            // skip first row if required... it may be the header 
                            if (skipFirstRow)
                            {
                                rows.MoveNext();
                            }
                            while (rows.MoveNext())
                            {
                                IRow row = (XSSFRow)rows.Current;
                                ExcellSheetRowList.Add(row);
                            }
                        }
                    }
                    #endregion
                }
                #endregion
            }
            catch (Exception ex)
            {

            }
            return ExcellSheetRowList;

        }
       public ActionResult Createfile()
        {
            return RedirectToAction("Create_file");
        }
        public ActionResult Create_file(HttpPostedFileBase upload)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");
            sheet1.CreateRow(0).CreateCell(0).SetCellValue("This is a Sample");
            int x = 1;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet1.CreateRow(i);
                for (int j = 0; j < 3; j++)
                {
                    if (j == 0)
                    {
                        row.CreateCell(j).SetCellValue("aliaa"+""+ i);
                    }
                    else if (j == 1)
                    {
                        row.CreateCell(j).SetCellValue(i);
                    }
                    else
                    {
                        row.CreateCell(j).SetCellValue(i + 87);
                    }


                }
            }

            ISheet sheet2 = workbook.CreateSheet("Sheet2");
            sheet2.CreateRow(0).CreateCell(0).SetCellValue("secound");
            int n = 1;
            for (int i = 1; i <= 15; i++)
            {
                IRow row = sheet2.CreateRow(i);
                for (int j = 0; j < 3; j++)
                {
                    if (j == 0)
                    {
                        row.CreateCell(j).SetCellValue("aliaa" + " " + i);
                    }
                    else if (j == 1)
                    {
                        row.CreateCell(j).SetCellValue(i);
                    }
                    else
                    {
                        row.CreateCell(j).SetCellValue("Yes");
                    }


                }
            }


            FileStream sw = System.IO.File.Create(HostingEnvironment.MapPath("~/Content/aliaa.xlsx"));
            workbook.Write(sw);
            sw.Close();
            ViewBag.msg = "succusse";
            return View();
        }

        public ActionResult sum()
        {
            return View();
        }

        [HttpPost]
        public ActionResult sum(int row1, int row2)
        {
            XSSFWorkbook book = new XSSFWorkbook(new FileStream(HostingEnvironment.MapPath("~/Content/aliaa.xlsx"), FileMode.Open));
            XSSFSheet sheet1 = book.GetSheetAt(0) as XSSFSheet;
            double y1 = sheet1.GetRow(row1).GetCell(1).NumericCellValue;
            double y2 = sheet1.GetRow(row2).GetCell(1).NumericCellValue;
            double ppm_sum = y1 + y2;
            double y3 = sheet1.GetRow(row1).GetCell(2).NumericCellValue;
            double y4 = sheet1.GetRow(row2).GetCell(2).NumericCellValue;
            double met = y3 + y4;


            ViewBag.ppm = ppm_sum;
            ViewBag.met = met;
            return View();

        }

        [HttpPost]
        public ActionResult read_n(int nrow)
        {
            List<product> products = new List<product>();

            XSSFWorkbook book = new XSSFWorkbook(new FileStream(HostingEnvironment.MapPath("~/Content/aliaa.xlsx"), FileMode.Open));
            XSSFSheet sheet1 = book.GetSheetAt(1) as XSSFSheet;

            for (int i = 1; i <= nrow; i++)
            {
                product product1 = new product();
                product1.description = sheet1.GetRow(i).GetCell(0).StringCellValue;
                product1.ppm = sheet1.GetRow(i).GetCell(1).NumericCellValue;
                product1.met = sheet1.GetRow(i).GetCell(2).StringCellValue;
                products.Add(product1);


            }
            
            return View(products);
        }
        public ActionResult read_n()
        {

            return View();
        }
        [HttpGet]
        public ActionResult setYEStoTrue()
        {


            XSSFWorkbook book1 = new XSSFWorkbook(new FileStream(HostingEnvironment.MapPath("~/Content/aliaa.xlsx"), FileMode.Open));

       
            XSSFSheet sheet1 = book1.GetSheet("sheet2") as XSSFSheet;
          
          
            int x = 1;
            for (int i = 1; i <= sheet1.LastRowNum; i++)
            {


                if (sheet1.GetRow(i).GetCell(2).ToString() == "Yes")
                    sheet1.GetRow(i).CreateCell(2).SetCellValue(1);


                else {
                    sheet1.GetRow(i).CreateCell(2).SetCellValue(0);

                }
            }
            string filepath = HostingEnvironment.MapPath("~/Content/aliaa.xlsx");
             FileStream sw = System.IO.File.Create(filepath);
            book1.Write(sw);
            sw.Close();
            ViewBag.msg = "scucess";
            ViewBag.path = filepath;
            return View();
        }

        [HttpGet]
        
        public ActionResult Download(string file)
        {
         //   string file = "C:\\Users\\Aliaa Yasser\\Desktop\\aliaa1.xlsx";
            //get the temp folder and file path in server
            string fullPath = Path.Combine(Server.MapPath("~/temp"), file);

            //return the file for download, this is an Excel 
            //so I set the file content type to "application/vnd.ms-excel"
            return File(fullPath, "application/vnd.ms-excel", file);
        }
    }
}
