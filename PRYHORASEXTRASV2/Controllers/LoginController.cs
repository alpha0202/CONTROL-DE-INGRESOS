using CapaDatos;
using PRYHORASEXTRASV2.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace PRYHORASEXTRASV2.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        public ActionResult Index()
        {
            return View();
        }

        //[NonAction]
        //private Boolean email_bien_escrito(String email)
        //{
        //    String expresion;
        //    expresion = "\\w+([-+.']\\w+)*@\\w+([-.]\\w+)*\\.\\w+([-.]\\w+)*";
        //    if (Regex.IsMatch(email, expresion))
        //    {
        //        if (Regex.Replace(email, expresion, String.Empty).Length == 0)
        //        {
        //            return true;
        //        }
        //        else
        //        {
        //            return false;
        //        }
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //}


        //[NonAction]
        //public string CrearPassword(int longitud)
        //{
        //    string caracteres = "1234567890";
        //    StringBuilder res = new StringBuilder();
        //    Random rnd = new Random();
        //    while (0 < longitud--)
        //    {
        //        res.Append(caracteres[rnd.Next(caracteres.Length)]);
        //    }
        //    return res.ToString();
        //}


        //[HttpPost]
        //public ActionResult Acceder(string usuario, string codigo, string codigoDigitado)
        //{
        //    try
        //    {
        //        if (codigo != codigoDigitado)
        //        {
        //            throw new ArgumentException("Código Incorrecto");
        //        }

        //        HttpCookie cookie = new HttpCookie("HEuser", usuario);
        //        cookie.Expires = DateTime.Now.AddDays(1d);
        //        Response.Cookies.Add(cookie);

        //        return Json("", JsonRequestBehavior.AllowGet);
        //    }
        //    catch (Exception ex)
        //    {
        //        return Json(ex.Message, JsonRequestBehavior.AllowGet);
        //    }

        //}


        [HttpPost]
        public ActionResult validarCorreo(string usuario, string password)
        {
            try
            {

                Usuario user = new Usuario();
                user = Usuario.RecuperarUsuario(usuario);

                if (user == null)
                {
                    throw new ArgumentException("NO SE ENCONTRO REGISTRADO EL USUARIO");
                }

                if (user.estado != 'A')
                {
                    throw new ArgumentException("EL USUARIO " + user.nombre + " ESTA DESACTIVADO");
                }

                if (user.password != password)
                {
                    throw new ArgumentException("CONTRASEÑA INCORRECTA");
                }

                HttpCookie cookie = new HttpCookie("CIuser", usuario);
                cookie.Expires = DateTime.Now.AddDays(1d);
                Response.Cookies.Add(cookie);

                return Json("OK", JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                return Json(ex.Message, JsonRequestBehavior.AllowGet);
            }

        }

        //[NonAction]
        //public void EnviarCorreoAlter(AlternateView html, string email, string asunto)
        //{
        //    try
        //    {

        //        bool bolUseDefaultCredential = true;
        //        string userName = "noresponder@aliar.com.co";
        //        string password = "NRF4z3nd4*2021";
        //        string senderName = "La Fazenda";
        //        string emailFrom = "noresponder@aliar.com.co";
        //        string smtpServer = "smtp.gmail.com";
        //        int portNumber = 25;
        //        var Auth = new NetworkCredential(userName, password);
        //        var From = new MailAddress(emailFrom, senderName);
        //        var SC = new SmtpClient(smtpServer, portNumber);
        //        SC.EnableSsl = true;
        //        SC.UseDefaultCredentials = bolUseDefaultCredential;


        //        // Dim [To] As New MailAddress(email)
        //        using (var message = new MailMessage())
        //        {
        //            message.AlternateViews.Add(html);
        //            message.From = From;
        //            foreach (string mail in email.Split(new char[] { ',' }))
        //                message.To.Add(new MailAddress(mail));
        //            message.Subject = asunto;
        //            //message.Body = body;
        //            message.IsBodyHtml = true;
        //            if (SC.UseDefaultCredentials)
        //            {
        //                SC.Credentials = Auth;
        //            }

        //            SC.DeliveryMethod = SmtpDeliveryMethod.Network;
        //            SC.Timeout = 100000;
        //            SC.Send(message);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new ArgumentException(ex.Message);
        //    }
        //}

    }
}