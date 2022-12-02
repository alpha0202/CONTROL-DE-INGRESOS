using CapaDatos;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace PRYHORASEXTRASV2.Models
{
    public class Usuario
    {
        public string usuario { get; set; }
        public string nombre { get; set; }
        public string password { get; set; }
        public char estado { get; set; }
        public string sede { get; set; }
        public int sedeId { get; set; }
        public string porteria { get; set; }
        public int porteriaID { get; set; }


        public static Usuario RecuperarUsuario(string usuario)
        {
            Usuario user = new Usuario();
            DataTable dt = new DataTable();
            List<Parametros> LstParametros = new List<Parametros>();
            LstParametros.Add(new Parametros("@usuario", usuario, System.Data.SqlDbType.VarChar));
            dt = CapaDatos.Datos.SPObtenerDataTable("SP_GetUsuario", LstParametros);
            if (dt.Rows.Count == 0)
            {
                user = null;
            }
            else
            {
                DataRow dr;
                dr = dt.Rows[0];
                user.usuario = dr["usuario"].ToString();
                user.nombre = dr["nombre"].ToString();
                user.estado = char.Parse(dr["estado"].ToString());
                user.password = dr["password"].ToString();

                user.porteria = dr["descripcionPorteria"].ToString();
                user.porteriaID = int.Parse(dr["RowIdPorteria"].ToString());
                user.sede = dr["descripcionSede"].ToString();
                user.sedeId = int.Parse( dr["RowIdSede"].ToString());

            }



            return user;
        }

      

    }




}