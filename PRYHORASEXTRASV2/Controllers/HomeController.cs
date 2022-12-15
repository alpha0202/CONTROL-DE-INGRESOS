using CapaDatos;
using CONTROLDEINGRESOS.Models;
using PRYHORASEXTRASV2.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Web;
using System.Web.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using ActionResult = System.Web.Mvc.ActionResult;
using HttpPostAttribute = System.Web.Mvc.HttpPostAttribute;
using NonActionAttribute = System.Web.Mvc.NonActionAttribute;
using ExcelDataReader;
using System.Data.OleDb;



namespace PRYHORASEXTRASV2.Controllers
{
    public class HomeController : Controller
    {



        public ActionResult Reporte()
        {
            try
            {
                if (Request.Cookies["CIuser"] == null)
                {
                    return RedirectToAction("Index", "Login");
                }


                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                if (user == null)
                {
                    throw new ArgumentException("NO SE ENCONTRO REGISTRADO EL USUARIO");
                }

                if (user.estado != 'A')
                {
                    throw new ArgumentException("EL USUARIO " + user.nombre + " ESTA DESACTIVADO");
                }

                ViewBag.NombreUsuario = user.nombre;
                ViewBag.titulo = user.porteria + "-" + user.sede;

                ViewBag.fechaIni = DateTime.Now.ToString("yyyy-MM-dd");
                ViewBag.fechaFin = DateTime.Now.ToString("yyyy-MM-dd");

            }
            catch (Exception ex)
            {
                return RedirectToAction("Error", new { Error = ex.Message });


            }

            return View();


        }

        public ActionResult ReporteVisitantes()
        {
            try
            {
                if (Request.Cookies["CIuser"] == null)
                {
                    return RedirectToAction("Index", "Login");
                }


                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                if (user == null)
                {
                    throw new ArgumentException("NO SE ENCONTRO REGISTRADO EL USUARIO");
                }

                if (user.estado != 'A')
                {
                    throw new ArgumentException("EL USUARIO " + user.nombre + " ESTA DESACTIVADO");
                }

                ViewBag.NombreUsuario = user.nombre;
                ViewBag.titulo = user.porteria + "-" + user.sede;

                ViewBag.fechaIni = DateTime.Now.ToString("yyyy-MM-dd");
                ViewBag.fechaFin = DateTime.Now.ToString("yyyy-MM-dd");

            }
            catch (Exception ex)
            {
                return RedirectToAction("Error", new { Error = ex.Message });


            }

            return View();


        }


        public ActionResult ReporteVisitanteFrecuente()
        {
            try
            {
                if (Request.Cookies["CIuser"] == null)
                {
                    return RedirectToAction("Index", "Login");
                }


                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                if (user == null)
                {
                    throw new ArgumentException("NO SE ENCONTRO REGISTRADO EL USUARIO");
                }

                if (user.estado != 'A')
                {
                    throw new ArgumentException("EL USUARIO " + user.nombre + " ESTA DESACTIVADO");
                }

                ViewBag.NombreUsuario = user.nombre;
                ViewBag.titulo = user.porteria + "-" + user.sede;

                ViewBag.fechaIni = DateTime.Now.ToString("yyyy-MM-dd");
                ViewBag.fechaFin = DateTime.Now.ToString("yyyy-MM-dd");

            }
            catch (Exception ex)
            {
                return RedirectToAction("Error", new { Error = ex.Message });


            }

            return View();


        }


        //Obtener reporte (tabla) con los visitantes frecuentes.
        [HttpPost]
        public ActionResult GetReporteVisitantesFrecuentes()
        {
            Usuario user = new Usuario();
            user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

            List<VisitanteFrecuente> respuesta = new List<VisitanteFrecuente>();

            DataTable dt = new DataTable();
            //List<Parametros> LstParametros = new List<Parametros>();
            //LstParametros.Add(new Parametros("@fechaIni", fechaIni, System.Data.SqlDbType.Date));
            //LstParametros.Add(new Parametros("@fechaFin", fechaFin, System.Data.SqlDbType.Date));
            //LstParametros.Add(new Parametros("@filtro", filtro, System.Data.SqlDbType.Int));
            //LstParametros.Add(new Parametros("@sede", sede, System.Data.SqlDbType.Int));
            dt = Datos.SPObtenerDataTable("SP_ReporteVisitantesFrecuentes");

            foreach (DataRow dr in dt.Rows)
            {
                VisitanteFrecuente res = new VisitanteFrecuente();
                res.cedula = int.Parse(dr["cedulaEmpleado"].ToString());
                res.nombre = dr["nombreEmpleado"].ToString();
                res.arl = dr["arl"].ToString();
                res.empleadoAutoriza = dr["empleadoAutoriza"].ToString();
                res.motivoVisita = dr["motivoVisita"].ToString();
                res.placa = dr["placa"].ToString();
                res.empresa = dr["empresa"].ToString();
                res.fechaIniFrecuente= DateTime.Parse(dr["fechaIniFrecuente"].ToString()).ToString("dd/MM/yyyy");
                res.fechaFinFrecuente= DateTime.Parse(dr["fechaFinFrecuente"].ToString()).ToString("dd/MM/yyyy");
                      
                respuesta.Add(res);

            }

            return Json(respuesta, JsonRequestBehavior.AllowGet);
        }



        #region CARGUE DE ARCHIVO EXCEL 
        public ActionResult ImportExcel()
        {


            return View();
        }
        [System.Web.Mvc.ActionName("Importexcel")]
        [HttpPost]
        public ActionResult Importexcel1()
        {
           
            if (Request.Files["FileUpload1"].ContentLength > 0)
            {
                string extension = System.IO.Path.GetExtension(Request.Files["FileUpload1"].FileName).ToLower();
                
                string connString = "";
                

                string[] validFileTypes = { ".xls", ".xlsx", ".csv" };

                string path1 = string.Format("{0}/{1}", Server.MapPath("~/Recursos/Uploads"), Request.Files["FileUpload1"].FileName);
                if (!Directory.Exists(path1))
                {
                    Directory.CreateDirectory(Server.MapPath("~/Recursos/Uploads"));
                }
                if (validFileTypes.Contains(extension))
                {
                    if (System.IO.File.Exists(path1))
                    {
                        System.IO.File.Delete(path1);
                    }
                    Request.Files["FileUpload1"].SaveAs(path1);
                    if (extension == ".csv")
                    {
                        DataTable dt = Utility.ConvertCSVtoDataTable(path1);
                        ViewBag.Data = dt;
                    }
                    //Connection String to Excel Workbook  
                    else if (extension.Trim() == ".xls")
                    {
                        connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path1 + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                        DataTable dt = Utility.ConvertXSLXtoDataTable(path1, connString);
                        ViewBag.Data = dt;
                    }
                    else if (extension.Trim() == ".xlsx")
                    {
                        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                        DataTable dt = Utility.ConvertXSLXtoDataTable(path1, connString);
                        ViewBag.Data = dt;
                        
                    }

                }
                else
                {
                    ViewBag.Error = "Por favor cargar archivos con extensión .xls, .xlsx or .csv.";

                }

            }

            return View();

        }

        #endregion







        public class Combox
        {
            public string valor { get; set; }
            public string descripcion { get; set; }
        }


        public class ClsRegistro
        {
            public string cedula { get; set; }
            public string nombre { get; set; }
            public string fecha { get; set; }
            public string tipoRegistro { get; set; }
            public string tipoInsert { get; set; }
            public string usuario { get; set; }
            public string porteria { get; set; }
            public string sede { get; set; }
            public string registro { get; set; }
            public string insert { get; set; }

            public string arl { get; set; }
            public string empleadoAutoriza { get; set; }
            public string motivo { get; set; }
            public string placa { get; set; }
            public string empresa { get; set; }


        }

        [HttpPost]
        public ActionResult GetReporteEmpleados(string sede, string filtro, string fechaIni, string fechaFin)
        {
            Usuario user = new Usuario();
            user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

            List<ClsRegistro> respuesta = new List<ClsRegistro>();

            DataTable dt = new DataTable();
            List<Parametros> LstParametros = new List<Parametros>();
            LstParametros.Add(new Parametros("@fechaIni", fechaIni, System.Data.SqlDbType.Date));
            LstParametros.Add(new Parametros("@fechaFin", fechaFin, System.Data.SqlDbType.Date));
            LstParametros.Add(new Parametros("@filtro", filtro, System.Data.SqlDbType.Int));
            LstParametros.Add(new Parametros("@sede", sede, System.Data.SqlDbType.Int));
            dt = Datos.SPObtenerDataTable("SP_ReporteEmpleados", LstParametros);

            foreach (DataRow dr in dt.Rows)
            {
                ClsRegistro res = new ClsRegistro();
                res.cedula = dr["cedulaEmpleado"].ToString();
                res.nombre = dr["nombreEmpleado"].ToString();
                res.fecha = DateTime.Parse(dr["fechaHora"].ToString()).ToString("dd/MM/yyyy HH:mm");
                res.tipoRegistro = dr["tipoRegistro"].ToString();
                res.usuario = dr["usuarioRegistra"].ToString();
                res.tipoInsert = dr["tipoInsert"].ToString();
                res.porteria = dr["descripcionPorteria"].ToString();
                res.sede = dr["descripcionSede"].ToString();
                res.insert = dr["Ingreso"].ToString();
                res.registro = dr["Registro"].ToString();
                respuesta.Add(res);

            }

            return Json(respuesta, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public ActionResult GetReporteVisitantes(string sede, string filtro, string fechaIni, string fechaFin)
        {
            Usuario user = new Usuario();
            user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

            List<ClsRegistro> respuesta = new List<ClsRegistro>();

            DataTable dt = new DataTable();
            List<Parametros> LstParametros = new List<Parametros>();
            LstParametros.Add(new Parametros("@fechaIni", fechaIni, System.Data.SqlDbType.Date));
            LstParametros.Add(new Parametros("@fechaFin", fechaFin, System.Data.SqlDbType.Date));
            LstParametros.Add(new Parametros("@filtro", filtro, System.Data.SqlDbType.Int));
            LstParametros.Add(new Parametros("@sede", sede, System.Data.SqlDbType.Int));
            dt = Datos.SPObtenerDataTable("SP_ReporteVisitantes", LstParametros);

            foreach (DataRow dr in dt.Rows)
            {
                ClsRegistro res = new ClsRegistro();
                res.cedula = dr["cedulaEmpleado"].ToString();
                res.nombre = dr["nombreEmpleado"].ToString();
                res.fecha = DateTime.Parse(dr["fechaHora"].ToString()).ToString("dd/MM/yyyy HH:mm");
                res.tipoRegistro = dr["tipoRegistro"].ToString();
                res.usuario = dr["usuarioRegistra"].ToString();
                res.tipoInsert = dr["tipoInsert"].ToString();
                res.porteria = dr["descripcionPorteria"].ToString();
                res.sede = dr["descripcionSede"].ToString();
                res.insert = dr["Ingreso"].ToString();
                res.registro = dr["Registro"].ToString();


                res.arl = dr["arl"].ToString();
                res.empleadoAutoriza = dr["empleadoAutoriza"].ToString();
                res.motivo = dr["motivoVisita"].ToString();
                res.placa = dr["placa"].ToString();
                res.empresa = dr["empresa"].ToString();


                respuesta.Add(res);

            }

            return Json(respuesta, JsonRequestBehavior.AllowGet);
        }





        [HttpPost]
        public ActionResult GetSedes()
        {
            Usuario user = new Usuario();
            user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

            List<Combox> respuesta = new List<Combox>();

            DataTable dt = new DataTable();
            dt = CapaDatos.Datos.ObtenerDataTable("select S.RowIDSede, S.descripcionSede from Sedes S INNER JOIN UsuarioSedes U ON U.usuario = '" + user.usuario + "' AND U.RowIDSede = s.RowIDSede where estado = 'A'");

            foreach (DataRow dr in dt.Rows)
            {
                Combox res = new Combox();
                res.valor = dr[0].ToString();
                res.descripcion = dr[1].ToString();
                respuesta.Add(res);

            }

            return Json(respuesta, JsonRequestBehavior.AllowGet);
        }






        public ActionResult Index()
        {
            try
            {
                if (Request.Cookies["CIuser"] == null)
                {
                    return RedirectToAction("Index", "Login");
                }


                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                if (user == null)
                {
                    throw new ArgumentException("NO SE ENCONTRO REGISTRADO EL USUARIO");
                }

                if (user.estado != 'A')
                {
                    throw new ArgumentException("EL USUARIO " + user.nombre + " ESTA DESACTIVADO");
                }

                ViewBag.NombreUsuario = user.nombre;
                ViewBag.titulo = user.porteria + "-" + user.sede;
            }
            catch (Exception ex)
            {
                return RedirectToAction("Error", new { Error = ex.Message });


            }

            return View();


        }



        public string ImageToBase64()
        {
            string base64String = null;
            string path = "D:\\FOTOYESID.jpg";
            using (System.Drawing.Image image = System.Drawing.Image.FromFile(path))
            {
                using (MemoryStream m = new MemoryStream())
                {
                    image.Save(m, image.RawFormat);
                    byte[] imageBytes = m.ToArray();
                    base64String = Convert.ToBase64String(imageBytes);
                    return base64String;
                }
            }
        }
        static string imageNull = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAIBAQIBAQICAgICAgICAwUDAwMDAwYEBAMFBwYHBwcGBwcICQsJCAgKCAcHCg0KCgsMDAwMBwkODw0MDgsMDAz/2wBDAQICAgMDAwYDAwYMCAcIDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAz/wAARCALQAtADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9mKKKK0MwooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoozikDA0ALRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUU2SRYYXkdlSOMbndiFVB6kngUAOoHJx39BXnHjT9pnQvDrPDpqya3crxmJtlup95D97/gIP1ryvxZ8evE/i3ehvv7Otm48ixHlZHoX++fz/AAqlFslyR9EeI/GWk+EI92qalZWH+zLKA5+iDLH8q4nXP2pPDmnErZxalqbdjHEIYz/wJyD/AOO18+HmQv1duSx5Y/U9aKvkXUjnZ61qn7WuoSsfsOiWMA7G4neVvyXaKxLz9pjxZcsSlxpttntFYqcfixauAop8qFzM7CT4/eMZf+Y5KvslvCP/AGSmJ8d/GKH/AJD90frFEf8A2SuSop2Qrs7a2/aK8YW551SKYf8ATWyhb+SitbT/ANqrxDbYFxZ6PdjufKeEn/vliP0rzOilyod2e36R+1tZS4GoaJeQer21wsyj8GCn9a7Dw/8AHPwr4jdUi1eG2mfgRXim3bP1b5f1r5goPIx29DS5EPnZ9lqd8SupDI3KspyrfQjg0tfIvhrxjqvg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r9qXxv8A2X4ct9ChfE2qHzrjHVYFPAP+84/JDXqv1IUdyew9a+T/AIleMD488c6hqYz5Msnl2wP8MKfKn5gZ+rGqgtSZvQw6KKK2MQooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKANz4b+NG+H3jWx1T5jDE3l3Kj+OFuHH4D5h7qK+r0kWWNWRhIjgMrr0dTyCPqK+NK+jP2bvGP/CTfDxLSRs3Oiv8AZWyeWjPMR/LK/wDAKzmuppB9D0CiiiszQKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA4748+Kz4T+GGoPG2y4vwLGE9wZMhiPogavmQDA46dq9a/ax8Rfadf0rSVPy2kDXcgH9+Q7V/JVP/fVeS1rDYym9QoooqyAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACu//Zt8Unw/8SorVjiDWYzaMO3mD54z+YI/4FXAVJZ38ulXkN1Ads9rIs8Z9GUhh+opPVDPsgHNFQaZqcet6ZbXsP8AqbyFLhPo6hh/Op6wNwooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACgDccDqeBRUGpagNJ025u2+7aQvOf+AKW/pQB8vfF/Xf+Ej+J+uXIO5BdNBH/uRYjH/oJ/OubpFlaceY5y8nzsfUnk/qaWt0YBRRRTEFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUA4oooA+k/2c9a/tj4TWCE5fTpJbI+wVty/wDjrj8q7mvHP2R9V3W2vWBP3Xhu0H1DI38lr2OsJbm8dgooopDCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigArl/jTfnTfhN4gkBwWszCPq7Kn/sxrqK4L9pa5MHwku1/wCe11bR/h5m4/8AoNNbiex84Hrx0ooorcwCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD0j9lm/+y/EueHtd6fKMepRkcf1r6Er5m/Z7ufs/xi0fniXzoj/wKF/8BX0yDmsZ7msNgoooqSwooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAK82/aok2fDKEf3tSh/HCyGvSa80/arGfhran+7qUX6pJTW4pbHz9RRRW5gFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFAHUfBN/L+Lvhw+t6F/NGFfUafcH0r5a+DCb/i14bH/AE/of/HWr6lT7g+lZT3NYbC0UUVBYUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFeeftPwed8Kmb/AJ439u/5ll/9mr0OuN/aBsvtvwe1r1gWKf8A74lQ/wAiaa3E9j5looIwaK3MAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA634DwfaPjDoI/uTvJ/3zE5r6eUYFfOX7NFj9r+LNu+P+PW0uJT/wB8hB/6HX0dWU9zWGwUUUVBYUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFZXjjSf7e8FaxZYybqxmjX67CR+oFatKjbXBPTPPuO9AHxhG/mRK394A06r/ivRT4b8Vanp5GPsN3LCPoGO3/x3FUK3OcKKKKYBRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFAHrP7JOmeb4i1u9x/x72sduD7u5Y/ole515n+yvo/2H4fXV4R82o3zkH1SNQg/XdXplYy3No7BRRRUlBRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAfOn7TWgf2T8T2ugMR6tbR3Gexdf3b/+gqfxrz6vfv2p/DZ1PwNbaki5k0m4G846RS4U/gGCH8a8BraOxjLcKKKKokKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigApHfy0Legzj1pa6H4UeGP+Ew+I2k2LLuhM4nnGP+WUfzt+eAP+BUhn0j8OvDv/CJ+A9I08jD21qgk/66MN7/APjzGtqhnLsSepOTRWBuFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAVNf0SHxNoV5p1x/qL+F7dz6BhjP4HB/CvkTUNOn0bUJ7O6XZc2krQyqezqcH+VfY1eB/tR+CzpHiyDWol/0fV12TED7s6AD/AMeTB+qmrg9bETWlzy+iiitTIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAr2r9k/wl5VrqeuyL/riLG2J/uqQ0hH47R/wE141Y2E2q30NrbIZbm5kWGJB/E7HAH5mvrbwj4Zh8GeF7HSoCGjsYRGWH/LRurN+LEn8aib0sXBa3NGiiisjUKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKw/iP4KT4heDbzS2IWWZQ9u5/wCWUy8ofpng+zGtyigD41ngktLiSGaNopoXMciN1RgcEH6EU2vWP2nvh1/ZuqJ4jtY/3F8wivQo4jm6K/0cDB/2h715PW6d0YNW0CiiimIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKueHtAuvFWu2mm2Sb7q9kEceei+rH2AyT7CgD0r9lzwF/aeuTeILhP3GnEwWuR9+cj5m/4Ap/N/avdqz/Cvhm18G+HbTTLMf6PZx7AxHMh6s592JJP1rQrBu7N0rIKKKKQwooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAq6zo9t4h0m5sbyMTWt3GYpU9VPp6EdQexAr5W8f+CLr4eeKZ9Musvs+eCbHFxEfuuP5EdiCK+s65T4vfDGL4neGvJXZFqVpmSymboGPVGP91uM+hwe1VF2Jkrny9RUl7ZTabezW1zE8FxbuY5YnGGjYcEGo62MQooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigBCdoyeAOSfSvoD9nH4XHwrop1q+i26jqcYEKMPmtoDyPoz8E+gAHrXE/s+/CP/hM9TGsahHnSLGT93Gw4vZR290U9fU4HrX0GzFjk8k8ms5voaQXUKKKKzNAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigDzn47fBj/hOrQ6ppkYGtW6YZBx9uQdFP8Atj+E9+h7V88spjdlYMrKSrKwwVI4II7GvsyvM/jd8CV8aeZq2josesAZmhztS/x+gk9+jdDzzVxl0ZEo9UfP9FOmie3meORHjkjYo6Ou1kYdQQeQabWpkFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABXYfCD4TXHxO1n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AooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKANxwBk+grC8V/EvQfBGRqWp20Ew/5YIfNnP/AABcn88UAbtAG5sDk+grxzxL+1milk0bSGf0mvpNo+vlpz+bCvPfEvxn8T+KwVudWuIYD/ywtP8AR4//AB3k/iTVcjJ50fSfiHxhpPhNN2p6lY2P+zNMA5+i/eP5VxGuftR+HNOytnFqWpuOhjiEMZ/4E/P/AI7Xz3j5y38R5LHqfqetLV8iI52ep65+1frF5kafpunWCno0pa5f/wBlX9K5PWPjT4r1wETa7fRoeqWxFuv/AI4Af1rmKKqyFdjruV7+XzLh5Lhz1aZzIT+LZpq/KMDgegoopkhRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAZ4xSRoIpAyDYw6Mnyn8xzS0UAb+jfFXxL4fAFprupog6JJN5yD8H3Cur0T9qXxFp5AvLfTNSTuWiNu5/FDj/AMdrzWilZDuz3nRP2rtGvMDUNO1LT2PVo9tyg/La36V2/hz4l+H/ABaQNP1ixnkPSJpPLl/74fB/Svk+kdBIPmAP1GankRXOz7OZSh5BH1GKSvk7w18Stf8AB5H9navewRj/AJYtJ5sJ/wCAPkflXoHhv9rG8t9qaxpUF0vQzWb+S/12NlT+BFS4MpTR7jRXKeE/jZ4a8YlUg1KO2uX6W94PIkJ9Bn5T+DGurKleoIzyPeoLCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKq6zrVn4d09ru/uoLK2TrLM4VfoPU+w5ryjxz+1TFDvg8O2nnN0+2XaEIPdI+p+rEfSmk2JtI9cvb6HTLN7i5mitreMZeWVwiL9SeK838X/tRaLo26LSoZtZnHAkGYbcH/eI3N+A/GvEfE/i7U/Gl55+q31xfOD8okb5I/8AdQYVfwFZ1WodyHPsdb4v+OPiTxlvSW/Njatx9nsswpj3bO9vxP4VyKqEzjvyfeloqyAooopiCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooARlDjBAI9CK3/CXxP17wOQNO1KdIB/y7ynzoD/AMAbIH4YrBopDPb/AAh+1baXW2LXdPks36G4tMyxfUofmH4Fq9Q0DxHp/iuw+1abe219B3eF923/AHh1U+xAr5AqfStVutCv1urK5ns7lOksEhR/zHUex4qXBdClN9T7ForwvwP+1NfaeVg1+2GoQ9PtVuojnX/eXhX/AA2n617B4T8a6V45sjPpV7Ddqoy6L8ssX+8h+Zfyx71m00aJpmpRRRSGFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFcf8S/jTpPw2VoJD9u1TGVsoW+ZPQyN0QfqewoA625uY7O2kmmkjhhiXdJJIwVEHqSeBXlPj/9qOz07fbeHoV1Cbp9smUi3T3VeGf6nA+teWePfihrHxHuc6hcYtVOY7OH5YI/w/iPu2T9K5+tFDuZufYveJPE+o+MNS+16peTXtx/C0h4jHoqjhR9BVGiitDMKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACpbC+n0q9jubWaa2uYjlJYnKOh9iOaiooA9d8AftST2my28RwG6j6C9tkAlX/fj6N9VwfY17JoWv2XifTEvdOuoby1fgSRNkA+hHUH2ODXx9Wj4W8Xal4J1P7Xpd3JaTHh9vKSj0dTww+tQ4di1Nn13RXm/wz/aN0/xc0dnqwi0nUmwquW/0a4Psx+4fZuPQ16QRtODWTVjRO+wUUUUDCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACory8h06zluLiWK3t4F3ySyMFSNfUk9KzPG3jrTfh9oxvdSn8tTkRRJzLcN/dRe/ueg7mvnP4mfFnU/idff6Qfs2nxNmCyjbKJ6Mx/jf3PTsBVKNyXKx2XxS/aWm1IyWPhppLa3+69+y7ZZf+uYP3B/tH5j2xXkrMXdmJLMxLMxOSxPUk9z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QAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFecftMeNj4c8DrpsD7brW2MTYPKwLjzD+OVX8TXo4BZgB1JwK+XfjT4y/4Tf4i39wjbrS0b7Ha+nloSC3/Am3H8RVRV2TJ2RytFFFbGIUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFAHrn7Kvjb7Lql54fmf8Ad3QN3aA9pFH7xR9Vw3/ADXuFfHuh63P4a1m01G1OLixlWeP3Knp9CMg/WvrrSNVg13Sra9tjut7yJJ4j/ssAR+WcfhWU1rc1g9LFiiiioLCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAOc+LPiw+Cvh5qd/G224EfkW/8A11k+RT+GSf8AgNfKqLsQAdhivaP2tPEmI9H0ZD94vfTD2HyR/wA3NeMVrDYym9QoooqyAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACvoD9lzxUdX8DT6Y5zLo821M/88ZMsv5NvH5V8/13/wCzX4j/ALD+J8NszYi1eFrQg9N4+eP9VI/4FUy2KjufRtFAORRWJsFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUE4FFKjiNwzcKp3N9ByaAPmT4/63/bfxZ1XDbo7EpZJ/wBs1+b/AMeLVxtTanqJ1fVLq7blrueScn13sW/rUNbrYwCiiimIKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAqzomrv4f1uyv0zvsbiO4GP9hgx/QGq1BXeMHoeDQB9liV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KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP/Z";
        public ActionResult Registro()
        {
            try
            {
                if (Request.Cookies["CIuser"] == null)
                {
                    return RedirectToAction("Index", "Login");
                }

                ViewBag.imagen = imageNull;// "data:image/jpeg;base64," +  ImageToBase64();

                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                if (user == null)
                {
                    throw new ArgumentException("NO SE ENCONTRO REGISTRADO EL USUARIO");
                }

                if (user.estado != 'A')
                {
                    throw new ArgumentException("EL USUARIO " + user.nombre + " ESTA DESACTIVADO");
                }

                ViewBag.NombreUsuario = user.nombre;
                ViewBag.titulo = user.porteria + "-" + user.sede;

            }
            catch (Exception ex)
            {
                return RedirectToAction("Error", new { Error = ex.Message });


            }

            return View();


        }

        private string guardarEntrada(string cedula, string usuario)
        {
            try
            {
                List<Parametros> LstParametros = new List<Parametros>();
                LstParametros.Add(new Parametros("@Cedula", cedula, System.Data.SqlDbType.Decimal));
                LstParametros.Add(new Parametros("@Usuario", usuario, System.Data.SqlDbType.VarChar));
                string respuesta = CapaDatos.Datos.SPGetEscalar("SP_GuardarRegistroEntrada", LstParametros).ToString();
                if (!string.IsNullOrEmpty(respuesta))
                {

                    throw new ArgumentException(respuesta);

                }

                return respuesta;


            }
            catch (Exception ex)
            {
                throw new ArgumentException(ex.Message);
            }
        }

        private string guardarSalida(string cedula, string usuario)
        {
            try
            {
                List<Parametros> LstParametros = new List<Parametros>();
                LstParametros.Add(new Parametros("@Cedula", cedula, System.Data.SqlDbType.Decimal));
                LstParametros.Add(new Parametros("@Usuario", usuario, System.Data.SqlDbType.VarChar));
                string respuesta = CapaDatos.Datos.SPGetEscalar("SP_GuardarRegistroSalida", LstParametros).ToString();
                if (!string.IsNullOrEmpty(respuesta))
                {

                    throw new ArgumentException(respuesta);

                }

                return respuesta;


            }
            catch (Exception ex)
            {
                throw new ArgumentException(ex.Message);
            }
        }

        [HttpPost]
        public ActionResult GuardarEmpleadoEntrada(string cedula)
        {
            if (Request.Cookies["CIuser"] == null)
            {
                return RedirectToAction("Index", "Login");
            }

            try
            {
                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                string respuesta = guardarEntrada(cedula, user.usuario);
                if (!string.IsNullOrEmpty(respuesta))
                {
                    throw new ArgumentException(respuesta);
                }
                return Json("", JsonRequestBehavior.AllowGet);



            }

            catch (Exception ex)
            {
                return Json(ex.Message, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public ActionResult GuardarEmpleadoSalida(string cedula)
        {
            if (Request.Cookies["CIuser"] == null)
            {
                return RedirectToAction("Index", "Login");
            }

            try
            {
                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                string respuesta = guardarSalida(cedula, user.usuario);
                if (!string.IsNullOrEmpty(respuesta))
                {
                    throw new ArgumentException(respuesta);
                }
                return Json("", JsonRequestBehavior.AllowGet);



            }

            catch (Exception ex)
            {
                return Json(ex.Message, JsonRequestBehavior.AllowGet);
            }
        }



        [HttpPost]
        public ActionResult GuardarImagenEmpleado(string cedula, string imagen)
        {
            try
            {
                string respuesta = "";
                List<Parametros> LstParametros = new List<Parametros>();
                LstParametros.Add(new Parametros("@Cedula", cedula, System.Data.SqlDbType.Decimal));
                LstParametros.Add(new Parametros("@imagen", imagen, System.Data.SqlDbType.VarChar));
                respuesta = CapaDatos.Datos.SPGetEscalar("SP_GuardarImagenEmpleado", LstParametros).ToString();
                if (!string.IsNullOrEmpty(respuesta))
                {

                    throw new ArgumentException(respuesta);

                }


                return Json(respuesta, JsonRequestBehavior.AllowGet);


            }
            catch (Exception ex)
            {
                return Json(ex.Message, JsonRequestBehavior.AllowGet);
            }

        }

        [HttpPost]
        public ActionResult GuardarImagenVisitante(string cedula, string imagen)
        {
            try
            {
                string respuesta = "";
                List<Parametros> LstParametros = new List<Parametros>();
                LstParametros.Add(new Parametros("@Cedula", cedula, System.Data.SqlDbType.Decimal));
                LstParametros.Add(new Parametros("@imagen", imagen, System.Data.SqlDbType.VarChar));
                respuesta = CapaDatos.Datos.SPGetEscalar("SP_GuardarImagenVisitante", LstParametros).ToString();
                if (!string.IsNullOrEmpty(respuesta))
                {

                    throw new ArgumentException(respuesta);

                }


                return Json(respuesta, JsonRequestBehavior.AllowGet);


            }
            catch (Exception ex)
            {
                return Json(ex.Message, JsonRequestBehavior.AllowGet);
            }

        }





        [HttpPost]
        public ActionResult RecuperarEmpleado(string cedula)
        {
            if (Request.Cookies["CIuser"] == null)
            {
                return RedirectToAction("Index", "Login");
            }


            Empleado empleado = new Empleado();
            try
            {
                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                DataSet ds = new DataSet();
                List<Parametros> LstParametros = new List<Parametros>();
                LstParametros.Add(new Parametros("@Cedula", cedula, System.Data.SqlDbType.Decimal));
                ds = Datos.GetDataSet("SP_GetEmpleadoRegistro", LstParametros);
                DataRow dr = ds.Tables[0].Rows[0];

                string mensaje = dr[0].ToString();

                if (string.IsNullOrEmpty(mensaje))
                {
                    empleado.mensaje = "";
                    empleado.cedula = decimal.Parse(dr["cedulaEmpleado"].ToString());
                    empleado.nombre = dr["nombreEmpleado"].ToString();
                    empleado.sedeID = int.Parse(dr["RowIdSede"].ToString());
                    empleado.sede = dr["descripcionSede"].ToString();
                    empleado.imagen = dr["imagenEmpleado"].ToString();
                    if (string.IsNullOrEmpty(empleado.imagen))
                    {
                        empleado.imagen = imageNull;
                    }
                    else
                    {
                        empleado.imagen = empleado.imagen;
                    }

                    if (ds.Tables[1].Rows.Count > 0)
                    {
                        DataRow dr2 = ds.Tables[1].Rows[0];
                        empleado.fechaUltimo = DateTime.Parse(dr2[0].ToString()).ToString("dd/MM/yyyy HH:mm");
                        empleado.tipo = dr2[1].ToString();
                    }

                    if (user.sedeId != empleado.sedeID)
                    {
                        empleado.mensaje = "SEDE";
                    }

                    if (empleado.tipo == "E")
                    {
                        empleado.mensaje = "ENTRADA";
                    }

                    if (string.IsNullOrEmpty(empleado.mensaje))
                    {
                        //GUARDAR
                        string respuesta = guardarEntrada(cedula, user.usuario);
                        if (!string.IsNullOrEmpty(respuesta))
                        {
                            throw new ArgumentException(respuesta);
                        }
                    }


                }
                else
                {
                    empleado.imagen = imageNull;
                    empleado.mensaje = mensaje;
                }


                return Json(empleado, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                empleado.imagen = imageNull;
                empleado.mensaje = ex.Message;
                return Json(empleado, JsonRequestBehavior.AllowGet);
            }


        }


        public class Arl
        {
            public string nombre { get; set; }
        }

        [HttpPost]
        public ActionResult GetArl()
        {
            if (Request.Cookies["CIuser"] == null)
            {
                return RedirectToAction("Index", "Login");
            }


            List<Arl> listArl = new List<Arl>();
            try
            {
                string sql = "SELECT Descripcion FROM Arl";
                DataTable dt = new DataTable();
                dt = Datos.ObtenerDataTable(sql);
                foreach (DataRow dr in dt.Rows)
                {
                    Arl a = new Arl();
                    a.nombre = dr[0].ToString();
                    listArl.Add(a);
                }


                return Json(listArl, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {

                return Json(listArl, JsonRequestBehavior.AllowGet);
            }


        }


        [HttpPost]
        public ActionResult GetEmpleados()
        {
            if (Request.Cookies["CIuser"] == null)
            {
                return RedirectToAction("Index", "Login");
            }

            Usuario user = new Usuario();
            user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

            List<Empleado> listEmpleado = new List<Empleado>();
            try
            {
                string sql = " select cedula,nombre  from EmpleadosAutoizadores where sedeID  = " + user.sedeId + " order by nombre";
                DataTable dt = new DataTable();
                dt = Datos.ObtenerDataTable(sql);
                foreach (DataRow dr in dt.Rows)
                {
                    Empleado a = new Empleado();
                    a.cedula = decimal.Parse(dr[0].ToString());
                    a.nombre = dr[1].ToString();
                    listEmpleado.Add(a);
                }


                return Json(listEmpleado, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {

                return Json(listEmpleado, JsonRequestBehavior.AllowGet);
            }


        }

        [HttpPost]
        public ActionResult GuardarVisitanteEntrada(string cedula, string nombre, string arl, string empledo, string motivo, string placa, string empresa, string cedulaEmpleado)
        {
            if (Request.Cookies["CIuser"] == null)
            {
                return RedirectToAction("Index", "Login");
            }

            try
            {
                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                string respuesta = guardarEntradav(cedula, user.usuario, nombre, arl, empledo, motivo, placa, empresa);
                if (!string.IsNullOrEmpty(respuesta))
                {
                    throw new ArgumentException(respuesta);
                }

                //enviar correo
                try
                {
                    string sql = "select correo  from EmpleadosAutoizadores where cedula = '" + cedulaEmpleado + "'";
                    string correo = Datos.GetEscalar(sql).ToString();
                    if (!string.IsNullOrEmpty(correo))
                    {
                        string html = "<p>Cordial saludo,</p>" +
                            " <h1>Autorización Ingreso Visitante</h1> " +
                             "    <p>Se autorizo el ingreso al visitante: "+cedula+"-"+nombre+" </p>  " +
                             "    <p>Empresa: "+empresa+" </p>  " +
                              "   <p>Motivo Visita: "+motivo+" </p>  " +
                             "<div class='gmail_default'></div>" +
                             "<div class='gmail_default'>Atentamente,</div><br />" +
                             "<div class='gmail_default'></div>" +
                             "<div class='gmail_default'><b>TECNOLOG&Iacute;A - ALIAR S.A.</b></div>";
                                AlternateView htmlView =
                                    AlternateView.CreateAlternateViewFromString(html,
                                                            Encoding.UTF8,
                                                            MediaTypeNames.Text.Html);
                        EnviarCorreoAlter(htmlView, correo, "Ingreso Visitante");
                    }

                }
                catch
                {

                }


                return Json("", JsonRequestBehavior.AllowGet);



            }

            catch (Exception ex)
            {
                return Json(ex.Message, JsonRequestBehavior.AllowGet);
            }
        }


        [NonAction]
        public void EnviarCorreoAlter(AlternateView html, string email, string asunto)
        {
            try
            {

                bool bolUseDefaultCredential = true;
                string userName = "noresponder@aliar.com.co";
                string password = "NRF4z3nd4*2021";
                string senderName = "La Fazenda";
                string emailFrom = "noresponder@aliar.com.co";
                string smtpServer = "smtp.gmail.com";
                int portNumber = 25;
                var Auth = new NetworkCredential(userName, password);
                var From = new MailAddress(emailFrom, senderName);
                var SC = new SmtpClient(smtpServer, portNumber);
                SC.EnableSsl = true;
                SC.UseDefaultCredentials = bolUseDefaultCredential;


                // Dim [To] As New MailAddress(email)
                using (var message = new MailMessage())
                {
                    message.AlternateViews.Add(html);
                    message.From = From;
                    foreach (string mail in email.Split(new char[] { ',' }))
                        message.To.Add(new MailAddress(mail));
                    message.Subject = asunto;
                    //message.Body = body;
                    message.IsBodyHtml = true;
                    if (SC.UseDefaultCredentials)
                    {
                        SC.Credentials = Auth;
                    }

                    SC.DeliveryMethod = SmtpDeliveryMethod.Network;
                    SC.Timeout = 100000;
                    SC.Send(message);
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException(ex.Message);
            }
        }


        [HttpPost]
        public ActionResult GuardarVisitanteSalida(string cedula)
        {
            if (Request.Cookies["CIuser"] == null)
            {
                return RedirectToAction("Index", "Login");
            }

            try
            {
                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                string respuesta = guardarSalidaV(cedula, user.usuario);
                if (!string.IsNullOrEmpty(respuesta))
                {
                    throw new ArgumentException(respuesta);
                }
                return Json("", JsonRequestBehavior.AllowGet);



            }

            catch (Exception ex)
            {
                return Json(ex.Message, JsonRequestBehavior.AllowGet);
            }
        }


        private string guardarEntradav(string cedula, string usuario, string nombre, string arl, string empledo, string motivo, string placa, string empresa)
        {
            try
            {
                List<Parametros> LstParametros = new List<Parametros>();
                LstParametros.Add(new Parametros("@Cedula", cedula, System.Data.SqlDbType.Decimal));
                LstParametros.Add(new Parametros("@Usuario", usuario, System.Data.SqlDbType.VarChar));
                LstParametros.Add(new Parametros("@nombre", nombre, System.Data.SqlDbType.VarChar));
                LstParametros.Add(new Parametros("@arl", arl, System.Data.SqlDbType.VarChar));
                LstParametros.Add(new Parametros("@empleado", empledo, System.Data.SqlDbType.VarChar));
                LstParametros.Add(new Parametros("@motivo", motivo, System.Data.SqlDbType.VarChar));
                LstParametros.Add(new Parametros("@empresa", empresa, System.Data.SqlDbType.VarChar));
                LstParametros.Add(new Parametros("@placa", placa, System.Data.SqlDbType.VarChar));

                string respuesta = CapaDatos.Datos.SPGetEscalar("SP_GuardarRegistroEntradaVisitante", LstParametros).ToString();
                if (!string.IsNullOrEmpty(respuesta))
                {

                    throw new ArgumentException(respuesta);

                }

                return respuesta;


            }
            catch (Exception ex)
            {
                throw new ArgumentException(ex.Message);
            }
        }


        private string guardarSalidaV(string cedula, string usuario)
        {
            try
            {
                List<Parametros> LstParametros = new List<Parametros>();
                LstParametros.Add(new Parametros("@Cedula", cedula, System.Data.SqlDbType.Decimal));
                LstParametros.Add(new Parametros("@Usuario", usuario, System.Data.SqlDbType.VarChar));
                string respuesta = CapaDatos.Datos.SPGetEscalar("SP_GuardarRegistroSalidaVisitante", LstParametros).ToString();
                if (!string.IsNullOrEmpty(respuesta))
                {

                    throw new ArgumentException(respuesta);

                }

                return respuesta;


            }
            catch (Exception ex)
            {
                throw new ArgumentException(ex.Message);
            }
        }

        [HttpPost]
        public ActionResult GetVisitantes(string cedula)
        {

            Usuario user = new Usuario();
            user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);


            Visitante visitante = new Visitante();
            try
            {

                DataSet ds = new DataSet();
                List<Parametros> LstParametros = new List<Parametros>();
                LstParametros.Add(new Parametros("@Cedula", cedula, System.Data.SqlDbType.Decimal));
                ds = Datos.GetDataSet("SP_GetVisitanteRegistro", LstParametros);
                DataRow dr = ds.Tables[0].Rows[0];

                string mensaje = dr[0].ToString();

                if (string.IsNullOrEmpty(mensaje))
                {
                    visitante.mensaje = "";
                    visitante.cedula = decimal.Parse(dr["cedula"].ToString());
                    visitante.nombre = dr["nombre"].ToString();
                    visitante.imagen = dr["imagenVistante"].ToString();
                    visitante.arl = dr["arl"].ToString();
                    visitante.usuarioCreacion = dr["usuarioCreacion"].ToString();
                    visitante.empleadoAutoriza = dr["empleadoAutoriza"].ToString();
                    visitante.motivoVisita = dr["motivoVisita"].ToString();
                    visitante.placa = dr["placa"].ToString();
                    visitante.empresa = dr["empresa"].ToString();
                    visitante.frecuente = bool.Parse(dr["frecuente"].ToString());

                    if (!string.IsNullOrEmpty(dr["fechaCreacion"].ToString()))
                    {
                        visitante.fechaCreacion = DateTime.Parse(dr["fechaCreacion"].ToString()).ToString("dd/MM/yyyy HH:mm");
                    }

                    if (!string.IsNullOrEmpty(dr["fechaCreacion"].ToString()))
                    {
                        visitante.fechaCreacion = DateTime.Parse(dr["fechaCreacion"].ToString()).ToString("dd/MM/yyyy HH:mm");
                    }


                    if (!string.IsNullOrEmpty(dr["fechaIniFrecuente"].ToString()))
                    {
                        visitante.fechaIniFrecuente = DateTime.Parse(dr["fechaIniFrecuente"].ToString()).ToString("dd/MM/yyyy");
                    }


                    if (!string.IsNullOrEmpty(dr["fechaFinFrecuente"].ToString()))
                    {
                        visitante.fechaFinFrecuente = DateTime.Parse(dr["fechaFinFrecuente"].ToString()).ToString("dd/MM/yyyy");
                    }

                    if (string.IsNullOrEmpty(visitante.imagen))
                    {
                        visitante.imagen = imageNull;
                    }

                    if (ds.Tables[1].Rows.Count > 0)
                    {
                        DataRow dr2 = ds.Tables[1].Rows[0];
                        visitante.fechaUltimo = DateTime.Parse(dr2[0].ToString()).ToString("dd/MM/yyyy HH:mm");
                        visitante.tipo = dr2[1].ToString();
                    }

                    if (visitante.tipo == "E")
                    {
                        visitante.mensaje = "ENTRADA";
                    }
                    if (visitante.tipo == "S")
                    {
                        visitante.mensaje = "SALIDA";
                    }

                    if (visitante.mensaje == "")
                    {
                        if (visitante.frecuente == true & dr["fechaIniFrecuente"].ToString() != "" & dr["fechaFinFrecuente"].ToString() != "")
                        {
                            if (DateTime.Now <= DateTime.Parse(dr["fechaFinFrecuente"].ToString()) & DateTime.Now >= DateTime.Parse(dr["fechaIniFrecuente"].ToString()))
                            {
                                string respuesta = guardarEntradav(visitante.cedula.ToString(), user.usuario, visitante.nombre, visitante.arl, visitante.empleadoAutoriza, visitante.motivoVisita, visitante.placa, visitante.empresa);
                                if (!string.IsNullOrEmpty(respuesta))
                                {
                                    throw new ArgumentException(respuesta);
                                }


                            }
                            else
                            {
                                visitante.frecuente = false;
                            }
                        }
                        else
                        {
                            visitante.frecuente = false;
                        }
                    }


                }
                else
                {
                    visitante.imagen = imageNull;
                    visitante.mensaje = mensaje;
                }


                return Json(visitante, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                visitante.imagen = imageNull;
                visitante.mensaje = ex.Message;
                return Json(visitante, JsonRequestBehavior.AllowGet);
            }


        }


        [HttpPost]
        public ActionResult RecuperarEmpleadoSalida(string cedula)
        {
            if (Request.Cookies["CIuser"] == null)
            {
                return RedirectToAction("Index", "Login");
            }


            Empleado empleado = new Empleado();
            try
            {
                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(Request.Cookies["CIuser"].Value);

                DataSet ds = new DataSet();
                List<Parametros> LstParametros = new List<Parametros>();
                LstParametros.Add(new Parametros("@Cedula", cedula, System.Data.SqlDbType.Decimal));
                ds = Datos.GetDataSet("SP_GetEmpleadoRegistro", LstParametros);
                DataRow dr = ds.Tables[0].Rows[0];

                string mensaje = dr[0].ToString();

                if (string.IsNullOrEmpty(mensaje))
                {
                    empleado.mensaje = "";
                    empleado.cedula = decimal.Parse(dr["cedulaEmpleado"].ToString());
                    empleado.nombre = dr["nombreEmpleado"].ToString();
                    empleado.sedeID = int.Parse(dr["RowIdSede"].ToString());
                    empleado.sede = dr["descripcionSede"].ToString();
                    empleado.imagen = dr["imagenEmpleado"].ToString();
                    if (string.IsNullOrEmpty(empleado.imagen))
                    {
                        empleado.imagen = imageNull;
                    }
                    else
                    {
                        empleado.imagen = empleado.imagen;
                    }

                    if (ds.Tables[1].Rows.Count > 0)
                    {
                        DataRow dr2 = ds.Tables[1].Rows[0];
                        empleado.fechaUltimo = DateTime.Parse(dr2[0].ToString()).ToString("dd/MM/yyyy HH:mm");
                        empleado.tipo = dr2[1].ToString();
                    }

                    if (user.sedeId != empleado.sedeID)
                    {
                        empleado.mensaje = "";
                    }

                    if (empleado.tipo == "S")
                    {
                        empleado.mensaje = "SALIDA";
                    }

                    if (string.IsNullOrEmpty(empleado.mensaje))
                    {
                        //GUARDAR
                        string respuesta = guardarSalida(cedula, user.usuario);
                        if (!string.IsNullOrEmpty(respuesta))
                        {
                            throw new ArgumentException(respuesta);
                        }
                    }


                }
                else
                {
                    empleado.imagen = imageNull;
                    empleado.mensaje = mensaje;
                }


                return Json(empleado, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                empleado.imagen = imageNull;
                empleado.mensaje = ex.Message;
                return Json(empleado, JsonRequestBehavior.AllowGet);
            }


        }




        public string Error(string Error)
        {

            return Error;
        }


    }
}