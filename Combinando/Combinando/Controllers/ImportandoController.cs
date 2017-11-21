using Combinando.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Combinando.Controllers
{
    public class ImportandoController : Controller
    {
        private PruebaBDEntities db = new PruebaBDEntities();
        private string nombreArchivo;
        // GET: Importando
        public ActionResult Index()
        {
            return View("");
        }
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase excelFile)
        {
            if(excelFile == null || excelFile.ContentLength == 0)
            {
                ViewBag.Error = "Seleccione un archivo valido";
                return View();
            }
            else
            {
                if(excelFile.FileName.EndsWith("xls")|| excelFile.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/import/" + excelFile.FileName);
                    if (System.IO.File.Exists(path))
                    {
                        System.IO.File.Delete(path);
                        excelFile.SaveAs(path);
                    }
                    else
                    {
                        excelFile.SaveAs(path);
                    }
                    
                    nombreArchivo = excelFile.FileName;
                    int count = 0;
                    ImportarExcel(out count);
                    return View("");
                }
                else
                {
                    ViewBag.Error = "Tipo de archivo incorrecto";
                    return View();
                }
            }
        }

        public ActionResult Importando()
        {
            var tb_prueba = db.tb_Prueba;

            return View(tb_prueba.ToList());
        }


        private bool ImportarExcel(out int count)
        {
            var result = false;
            count = 0;
            try
            {
                String path = Server.MapPath("/") + "\\import\\"+nombreArchivo;
                var package = new ExcelPackage(new System.IO.FileInfo(path));
                int startColumn = 1;
                int startRow = 5;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                object data = null;
                PruebaBDEntities db = new PruebaBDEntities();
                do
                {
                    data = worksheet.Cells[startRow, startColumn].Value;
                    object className = worksheet.Cells[startRow, startColumn + 1].Value.ToString();
                    if (data != null && className != null)
                    {
                        //importdb
                        var isSuccess = saveClass(className.ToString(), db);
                        if (isSuccess)
                        {
                            count++;
                        }
                    }
                    startRow++;
                }
                while (data != null);
            }
            catch (Exception ex)
            {

            }


            return result;
        }

        private bool saveClass(String className, PruebaBDEntities db)
        {
            var result = false;
            try
            {
                //Checa si existe
                if(db.tb_Prueba.Where(t=>t.ClassName.Equals(className)).Count() == 0)
                {
                    var item = new tb_Prueba();
                    item.ClassName = className;
                    db.tb_Prueba.Add(item);
                    db.SaveChanges();
                    result = true;
                }
            }
            catch
            {

            }
            return result;
        }
    }
}