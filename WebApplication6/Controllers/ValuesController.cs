using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using WebApplication6.Models;




namespace WebApplication6.Controllers

{
    public class ValuesController : ApiController
    {

         [HttpGet]
 public string Create_file()
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
                     row.CreateCell(j).SetCellValue("aliaa" + i);
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
         FileStream sw = System.IO.File.Create(System.Web.HttpContext.Current.Server.MapPath("~/App_Data/aliaa.xlsx"));
         workbook.Write(sw);
         sw.Close();
         return "created succesfully";
     }
     
        [HttpGet]
        public  string sum(int row1, int row2)
        {
            XSSFWorkbook book = new XSSFWorkbook(new FileStream(System.Web.HttpContext.Current.Server.MapPath("~/Content/aliaa.xlsx"), FileMode.Open));
            XSSFSheet sheet1 = book.GetSheetAt(0) as XSSFSheet;
            double y1 = sheet1.GetRow(row1).GetCell(1).NumericCellValue;
            double y2 = sheet1.GetRow(row2).GetCell(1).NumericCellValue;
            double ppm_sum = y1 + y2;
            double y3 = sheet1.GetRow(row1).GetCell(2).NumericCellValue;
            double y4 = sheet1.GetRow(row2).GetCell(2).NumericCellValue;
            double met = y3 + y4;



            return "ppm sum=" + ppm_sum + "--------- met sum=" + met;

        }

        [HttpGet]
        public IEnumerable<product> read_n(int nrow)
        {
            List<product> products = new List<product>();

            XSSFWorkbook book = new XSSFWorkbook(new FileStream(System.Web.HttpContext.Current.Server.MapPath("~/Content/aliaa.xlsx"), FileMode.Open));
            XSSFSheet sheet1 = book.GetSheetAt(1) as XSSFSheet;

            for (int i = 1; i <= nrow; i++)
            {
                product product1 = new product();
                product1.description = sheet1.GetRow(i).GetCell(0).StringCellValue;
                product1.ppm = sheet1.GetRow(i).GetCell(1).NumericCellValue;
                product1.met = sheet1.GetRow(i).GetCell(2).StringCellValue;
                products.Add(product1);


            }

         return products;
        }
        [HttpGet]
        public HttpResponseMessage setYEStoTrue(int n)
        {

        
            XSSFWorkbook book1 = new XSSFWorkbook(new FileStream(System.Web.HttpContext.Current.Server.MapPath("~/Content/aliaa.xlsx"), FileMode.Open));


            XSSFSheet sheet1 = book1.GetSheet("sheet2") as XSSFSheet;


            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheetx = workbook.CreateSheet("Sheet1");
           
            
            int x = 1;
            for (int i = 1; i <= n; i++)
            {


                if (sheet1.GetRow(i).GetCell(2).ToString() == "Yes")
                    sheet1.GetRow(i).CreateCell(2).SetCellValue(1);


                else if (sheet1.GetRow(i).GetCell(2).ToString() == "No")
                {
                    sheet1.GetRow(i).CreateCell(2).SetCellValue(0);

                }
               
            }
            string filepath = System.Web.HttpContext.Current.Server.MapPath("~/Content/aliaa.xlsx");
            FileStream sw = System.IO.File.Create(filepath);
            book1.Write(sw);
            sw.Close();

            
            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
            response.Content = new StreamContent(new FileStream(filepath, FileMode.Open, FileAccess.Read));
            response.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = "report.xlsx";
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");

            return response;
        }



     


        // POST api/values
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }


       
    }
}
