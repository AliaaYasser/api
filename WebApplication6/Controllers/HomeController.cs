using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelDataReader;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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




        public List<IRow> GetExcellSheetRows(string filePath, long parsedFileId, bool skipFirstRow = true)
        {
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
                #region Log Exception
                Log(LogEnum.LogWriteType.Both, LogEnum.LogMethodType.Method, MethodBase.GetCurrentMethod().Name, ex, LogEnum.View.Web);
                #endregion
            }
            return ExcellSheetRowList;
        }


        public bool GenerateExcelSheetWithoutDownload(DataTable dataTable, string exportingSheetPath, out string exportingFileName)
        {
            #region Validate the parameters and Generate the excel sheet
            bool returnValue = false;
            exportingFileName = Guid.NewGuid().ToString() + ".xlsx";

            try
            {
                string excelSheetPath = string.Empty;
                #region Check If The directory is exist
                if (!Directory.Exists(exportingSheetPath))
                {
                    Directory.CreateDirectory(exportingSheetPath);
                }

                excelSheetPath = exportingSheetPath + exportingFileName;
                FileInfo fileInfo = new FileInfo(excelSheetPath);
                #endregion

                #region Write stream to the file
                MemoryStream ms = DataToExcel(dataTable);
                byte[] blob = ms.ToArray();
                if (blob != null)
                {
                    using (MemoryStream inStream = new MemoryStream(blob))
                    {
                        FileStream fs = new FileStream(excelSheetPath, FileMode.Create);
                        inStream.WriteTo(fs);
                        fs.Close();
                    }
                }
                ms.Close();
                returnValue = true;
                #endregion
            }
            catch (Exception ex)
            {

            }
            return returnValue;
            #endregion
        }
    }
}
