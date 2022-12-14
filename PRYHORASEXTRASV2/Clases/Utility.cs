using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Data.OleDb;
using CapaDatos;

namespace CONTROLDEINGRESOS.Models
{
    public class Utility
    {
       
        //método que lee los archivos csv y pasa el contenido a un dataset.
        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            DataTable dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }

                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    if (rows.Length > 1)
                    {
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i].Trim();
                        }
                        dt.Rows.Add(dr);
                    }
                }

            }


            return dt;
        }

        //método que leer las hojas del archivo excel, saca los campos, las filas y las incluye en un dataset.
        public static DataTable ConvertXSLXtoDataTable(string strFilePath, string connString)
        {
            OleDbConnection oledbConn = new OleDbConnection(connString);
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            try
            {
               
                oledbConn.Open();
                using (DataTable Sheets = oledbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null))
                {

                    for (int i = 0; i < Sheets.Rows.Count; i++)
                    {
                        string worksheets = Sheets.Rows[i]["TABLE_NAME"].ToString();
                        OleDbCommand cmd = new OleDbCommand(String.Format("SELECT * FROM [{0}]", worksheets), oledbConn);
                        OleDbDataAdapter oleda = new OleDbDataAdapter();
                        oleda.SelectCommand = cmd;

                        oleda.Fill(ds);
                  
                    }

                    dt = ds.Tables[0];

                   

                    List<VisitanteFrecuente> listVisitanteFrecuente = new List<VisitanteFrecuente>();

                    foreach (DataRow row in dt.Rows)
                    {
                        VisitanteFrecuente visitanteFre = new VisitanteFrecuente();

                        listVisitanteFrecuente.Add(new VisitanteFrecuente()
                        {
                            cedula = int.Parse(row[0].ToString()),
                            nombre = row[1].ToString(),
                            arl = row[2].ToString(),
                            empleadoAutoriza = row[3].ToString(),
                            motivoVisita = row[4].ToString(),
                            empresa = row[5].ToString(),
                            placa = row[6].ToString(),
                            fechaIniFrecuente = row[7].ToString(),
                            fechaFinFrecuente = row[8].ToString()
                        });

                        decimal cedula = int.Parse(row[0].ToString());
                        string nombre = row[1].ToString();
                        string arl = row[2].ToString();
                        string empleadoAutoriza = row[3].ToString();
                        string motivo = row[4].ToString();
                        string empresa = row[5].ToString();
                        string placa = row[6].ToString();
                        string fechaIni = DateTime.Parse(row[7].ToString()).ToString("dd/MM/yyyy");
                        string fechafin = DateTime.Parse(row[8].ToString()).ToString("dd/MM/yyyy");
                        //string fechaIni = row[7].ToString();
                        //string fechafin = row[8].ToString();


                        List<Parametros> LstParametros = new List<Parametros>();
                        LstParametros.Add(new Parametros("@cedula", cedula, SqlDbType.VarChar));
                        LstParametros.Add(new Parametros("@nombre", nombre, SqlDbType.VarChar));
                        LstParametros.Add(new Parametros("@arl", arl, SqlDbType.VarChar));
                        LstParametros.Add(new Parametros("@empleadoAutoriza", empleadoAutoriza, SqlDbType.VarChar));
                        LstParametros.Add(new Parametros("@motivo", motivo, SqlDbType.VarChar));
                        LstParametros.Add(new Parametros("@empresa", empresa, SqlDbType.VarChar));
                        LstParametros.Add(new Parametros("@placa", placa, SqlDbType.VarChar));
                        LstParametros.Add(new Parametros("@fechaIni", fechaIni, SqlDbType.Date));
                        LstParametros.Add(new Parametros("@fechaFin", fechafin, SqlDbType.Date));
                        
                        string respuesta = Datos.SPGetEscalar("SP_GuardarExcelVisitanteFrecuente", LstParametros).ToString();


                    }

                }


            }
            catch(Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {

                oledbConn.Close();
            }

            return dt;

        }




        //public static List<VisitanteFrecuente> lstVisitanteFrecuente(DataTable dt)

        //{
        //    var listExcel = (from row in dt.AsEnumerable()
        //                     select new VisitanteFrecuente()
        //                     {
        //                       cedula = int.Parse(row["cedula"].ToString()),
        //                       nombre = row["nombre"].ToString(),
        //                       arl = row["arl"].ToString(),
        //                       empleadoAutoriza = row["empleadoAutoriza"].ToString(),
        //                       motivoVisita = row["motivo"].ToString(),
        //                       empresa = row["empresa"].ToString(),
        //                       placa = row["placa"].ToString(),
        //                       fechaIniFrecuente = row["fechaIni"].ToString(),
        //                       fechaFinFrecuente = row["fechaFin"].ToString() DateTime.Parse(dr["fechaFinFrecuente"].ToString()).ToString("dd/MM/yyyy");
        //                     }
        //                     ).ToList();


        //    return listExcel;
        //}




    }
}