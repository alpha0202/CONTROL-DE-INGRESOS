using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Security.Cryptography;
using System.IO;
using System.Web;
using System.Net.Http;

namespace CapaDatos
{
    public class Parametros
    {
        #region Variables

        string _Nombre;
        SqlDbType _Tipo;
        object _Valor;
        Int32 _Longitud;
        ParameterDirection _Direccion = ParameterDirection.Input;

        #endregion

        #region Constructores

        public Parametros(string Nombre, object Valor, SqlDbType Tipo)
        {
            this._Nombre = Nombre;
            this._Valor = Valor;
            this._Tipo = Tipo;
        }

        public Parametros(string Nombre, object Valor, SqlDbType Tipo, Int32 Longitud)
        {
            this._Nombre = Nombre;
            this._Valor = Valor;
            this._Tipo = Tipo;
            this._Longitud = Longitud;
        }

        public Parametros(string Nombre, object Valor, SqlDbType Tipo, ParameterDirection Direccion = ParameterDirection.Input)
        {
            this._Nombre = Nombre;
            this._Valor = Valor;
            this._Tipo = Tipo;
            this._Direccion = Direccion;
        }

        public Parametros(string Nombre, object Valor, SqlDbType Tipo, Int32 Longitud, ParameterDirection Direccion = ParameterDirection.Input)
        {
            this._Nombre = Nombre;
            this._Valor = Valor;
            this._Tipo = Tipo;
            this._Longitud = Longitud;
            this._Direccion = Direccion;
        }

        #endregion

        #region Propiedades

        public string Nombre
        {
            get { return this._Nombre; }
            set { this._Nombre = value; }
        }

        public object Valor
        {
            get { return this._Valor; }
            set { this._Valor = value; }
        }

        public Int32 Longitud
        {
            get { return this._Longitud; }
            set { this._Longitud = value; }
        }

        public SqlDbType Tipo
        {
            get { return this._Tipo; }
            set { this._Tipo = value; }
        }

        public ParameterDirection Direccion
        {
            get { return this._Direccion; }
            set { this._Direccion = value; }
        }

        #endregion
    }

    public class Datos
    {
        static String SqlConnectionString = @"Data Source= 10.252.0.158;Initial Catalog=CONTROLDEINGRESOS_NET;User Id=sa;Password=0bleaconMora";
        //static String SqlConnectionString = @"Data Source= 10.252.0.159;Initial Catalog=HORASEXTRAS_V2;User Id=sa;Password=0bleaconMora";
        //public static String SqlConnectionString = SqlConnectionString;// @"Data Source= 10.252.0.158\ALIAR;Initial Catalog=CEDIALIAR;User Id=SWAliar;Password=Con@x1on";
        #region Variables



        #endregion

        #region Eventos Estáticos

        #region Registro Errores

        public static void logErrores(string procedure, string parameters, string Funcion, string Error)
        {

        }

        #endregion

        #region Eventos ObtenerDataTable

        public static DataTable SPObtenerDataTable(string Procedimiento, List<Parametros> Parametros)
        {
            SqlConnection Conexion = new SqlConnection();
            Conexion.ConnectionString = SqlConnectionString;
            StringBuilder CadenaSQL = new StringBuilder();
            SqlCommand Comando = new SqlCommand();

            try
            {
                Comando.Connection = Conexion;
                Comando.CommandType = CommandType.StoredProcedure;
                Comando.CommandText = Procedimiento;
                Comando.CommandTimeout = 60000;
                SqlParameter ObjSqlParametro;

                foreach (Parametros ObjSpParametro in Parametros)
                {
                    ObjSqlParametro = new SqlParameter();
                    ObjSqlParametro.ParameterName = ObjSpParametro.Nombre;
                    ObjSqlParametro.SqlDbType = ObjSpParametro.Tipo;
                    ObjSqlParametro.Value = ObjSpParametro.Valor;

                    if (ObjSpParametro.Tipo == SqlDbType.VarChar)
                    {
                        ObjSqlParametro.Size = ObjSpParametro.Longitud;
                    }

                    CadenaSQL.Append(ObjSpParametro.Nombre);
                    CadenaSQL.Append(" = ");
                    CadenaSQL.Append(ObjSpParametro.Valor.ToString());
                    CadenaSQL.Append(" | ");
                    Comando.Parameters.Add(ObjSqlParametro);
                }

                if (Conexion.State != ConnectionState.Open)
                {
                    Conexion.Open();
                }

                DataTable DtResultado = new DataTable();
                SqlDataAdapter ObjDataAdapter = new SqlDataAdapter(Comando);
                ObjDataAdapter.Fill(DtResultado);

                return DtResultado;
            }
            catch (Exception ex)
            {
                logErrores(Procedimiento, CadenaSQL.ToString(), "GetDataTable", ex.Message.ToString());
                throw ex;
            }
            finally
            {
                if (Conexion.State != ConnectionState.Closed)
                {
                    Conexion.Close();
                }

                Comando.Parameters.Clear();
                Comando.Dispose();
            }
        }

        public static DataTable SPObtenerDataTable(string Procedimiento)
        {
            SqlConnection Conexion = new SqlConnection();
            Conexion.ConnectionString = SqlConnectionString;
            SqlCommand Comando = new SqlCommand();

            try
            {
                Comando.Connection = Conexion;
                Comando.CommandType = CommandType.StoredProcedure;
                Comando.CommandText = Procedimiento;
                Comando.CommandTimeout = 60000;
                DataTable DtResultado = new DataTable();
                SqlDataAdapter ObjDataAdapter = new SqlDataAdapter(Comando);

                if (Conexion.State != ConnectionState.Open)
                {
                    Conexion.Open();
                }

                ObjDataAdapter.Fill(DtResultado);
                return DtResultado;
            }
            catch (Exception ex)
            {
                logErrores(Procedimiento, "Este procedimiento no tiene parámetros asociados", "GetDataTable", ex.Message.ToString());
                throw ex;
            }
            finally
            {
                if (Conexion.State != ConnectionState.Closed)
                {
                    Conexion.Close();
                }
                Comando.Parameters.Clear();
                Comando.Dispose();
            }
        }



        public static DataTable ObtenerDataTable(string sql)
        {
            SqlConnection Conexion = new SqlConnection();
            Conexion.ConnectionString = SqlConnectionString;
            SqlCommand Comando = new SqlCommand();

            try
            {
                Comando.Connection = Conexion;
                Comando.CommandType = CommandType.Text;
                Comando.CommandText = sql;
                Comando.CommandTimeout = 60000;
                DataTable DtResultado = new DataTable();
                SqlDataAdapter ObjDataAdapter = new SqlDataAdapter(Comando);

                if (Conexion.State != ConnectionState.Open)
                {
                    Conexion.Open();
                }

                ObjDataAdapter.Fill(DtResultado);
                return DtResultado;
            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                if (Conexion.State != ConnectionState.Closed)
                {
                    Conexion.Close();
                }
                Comando.Parameters.Clear();
                Comando.Dispose();
            }
        }
        #endregion

        #region Eventos ObtenerDataset

        public static DataSet GetDataSet(string Procedimiento)
        {
            SqlConnection Conexion = new SqlConnection();
            Conexion.ConnectionString = SqlConnectionString;
            SqlCommand Comando = new SqlCommand();

            try
            {
                Comando.Connection = Conexion;
                Comando.CommandType = CommandType.StoredProcedure;
                Comando.CommandText = Procedimiento;
                Comando.CommandTimeout = 60000;
                DataSet DtResultado = new DataSet();

                SqlDataAdapter ObjDataAdapter = new SqlDataAdapter(Comando);

                if (Conexion.State != ConnectionState.Open)
                {
                    Conexion.Open();
                }

                ObjDataAdapter.Fill(DtResultado);
                return DtResultado;
            }
            catch (Exception ex)
            {
                logErrores(Procedimiento, "Este procedimiento no tiene parámetros asociados", "GetDataSet", ex.Message.ToString());
                throw ex;
            }
            finally
            {
                if (Conexion.State != ConnectionState.Closed)
                {
                    Conexion.Close();
                }

                Comando.Parameters.Clear();
                Comando.Dispose();
            }
        }

        public static DataSet GetDataSet(string Procedimiento, List<Parametros> Parametros)
        {
            SqlConnection Conexion = new SqlConnection();
            Conexion.ConnectionString = SqlConnectionString;
            StringBuilder CadenaSQL = new StringBuilder();
            SqlCommand Comando = new SqlCommand();

            try
            {
                Comando.Connection = Conexion;
                Comando.CommandType = CommandType.StoredProcedure;
                Comando.CommandText = Procedimiento;
                Comando.CommandTimeout = 60000;

                SqlParameter ObjSqlParametro;

                foreach (Parametros ObjSpParametro in Parametros)
                {
                    ObjSqlParametro = new SqlParameter();
                    ObjSqlParametro.ParameterName = ObjSpParametro.Nombre;
                    ObjSqlParametro.SqlDbType = ObjSpParametro.Tipo;
                    ObjSqlParametro.Value = ObjSpParametro.Valor;

                    if (ObjSpParametro.Tipo == SqlDbType.VarChar)
                    {
                        ObjSqlParametro.Size = ObjSpParametro.Longitud;
                    }
                    CadenaSQL.Append(ObjSpParametro.Nombre);
                    CadenaSQL.Append(" = ");
                    CadenaSQL.Append(ObjSpParametro.Valor.ToString());
                    CadenaSQL.Append(" | ");

                    Comando.Parameters.Add(ObjSqlParametro);

                }

                DataSet DtResultado = new DataSet();
                SqlDataAdapter ObjDataAdapter = new SqlDataAdapter(Comando);

                if (Conexion.State != ConnectionState.Open)
                {
                    Conexion.Open();
                }

                ObjDataAdapter.Fill(DtResultado);

                return DtResultado;
            }
            catch (Exception ex)
            {
                logErrores(Procedimiento, CadenaSQL.ToString(), "GetDataSet", ex.Message.ToString());
                throw ex;
            }
            finally
            {
                if (Conexion.State != ConnectionState.Closed)
                {
                    Conexion.Close();
                }

                Comando.Parameters.Clear();
                Comando.Dispose();
            }
        }

        #endregion

        #region Eventos GetScalar

        public static Object SPGetEscalar(string Procedimiento, List<Parametros> Parametros)
        {
            SqlConnection Conexion = new SqlConnection();
            Conexion.ConnectionString = SqlConnectionString;
            StringBuilder CadenaSQL = new StringBuilder();
            SqlCommand Comando = new SqlCommand();

            try
            {
                Comando.Connection = Conexion;
                Comando.CommandType = CommandType.StoredProcedure;
                Comando.CommandText = Procedimiento;
                Comando.CommandTimeout = 60000;

                SqlParameter ObjSqlParametro;

                foreach (Parametros ObjSpParametro in Parametros)
                {
                    ObjSqlParametro = new SqlParameter();
                    ObjSqlParametro.ParameterName = ObjSpParametro.Nombre;
                    ObjSqlParametro.SqlDbType = ObjSpParametro.Tipo;
                    ObjSqlParametro.Value = ObjSpParametro.Valor;

                    ObjSqlParametro.Direction = ObjSpParametro.Direccion;

                    if (ObjSpParametro.Tipo == SqlDbType.VarChar)
                    {
                        ObjSqlParametro.Size = ObjSpParametro.Longitud;
                    }
                    CadenaSQL.Append(ObjSpParametro.Nombre);
                    CadenaSQL.Append(" = ");
                    CadenaSQL.Append(ObjSpParametro.Valor.ToString());
                    CadenaSQL.Append(" | ");

                    Comando.Parameters.Add(ObjSqlParametro);
                }

                if (Conexion.State != ConnectionState.Open)
                {
                    Conexion.Open();
                }

                Object DtResultado;
                DtResultado = Comando.ExecuteScalar();

                //if (Parametros.Count(p => p.Direccion == ParameterDirection.Output) > 0)
                //{
                //    Parametros.First(p => p.Direccion == ParameterDirection.Output).Valor= Comando.Parameters.OfType<SqlParameterCollection>().ToList().First(c => c.OfType<
                //        //Comando.Parameters

                //}

                return DtResultado;
            }
            catch (Exception ex)
            {
                logErrores(Procedimiento, CadenaSQL.ToString(), "GetEscalar", ex.Message.ToString());
                throw ex;
            }
            finally
            {
                if (Conexion.State != ConnectionState.Closed)
                {
                    Conexion.Close();
                }

                Comando.Parameters.Clear();
                Comando.Dispose();

            }
        }

        public static Object SPGetEscalar(string Procedimiento)
        {
            SqlConnection Conexion = new SqlConnection();
            Conexion.ConnectionString = SqlConnectionString;
            SqlCommand Comando = new SqlCommand();

            try
            {
                Comando.Connection = Conexion;
                Comando.CommandType = CommandType.StoredProcedure;
                Comando.CommandText = Procedimiento;
                Comando.CommandTimeout = 60000;

                if (Conexion.State != ConnectionState.Open)
                {
                    Conexion.Open();
                }

                Object DtResultado;
                DtResultado = Comando.ExecuteScalar();

                return DtResultado;
            }
            catch (Exception ex)
            {
                logErrores(Procedimiento, "Este procedimiento no tiene parámetros asociados", "GetEscalar", ex.Message.ToString());
                throw ex;
            }
            finally
            {
                if (Conexion.State != ConnectionState.Closed)
                {
                    Conexion.Close();
                }

                Comando.Parameters.Clear();
                Comando.Dispose();
            }
        }


        public static Object GetEscalar(string sql)
        {
            SqlConnection Conexion = new SqlConnection();
            Conexion.ConnectionString = SqlConnectionString;
            SqlCommand Comando = new SqlCommand();

            try
            {
                Comando.Connection = Conexion;
                Comando.CommandType = CommandType.Text;
                Comando.CommandText = sql;
                Comando.CommandTimeout = 60000;

                if (Conexion.State != ConnectionState.Open)
                {
                    Conexion.Open();
                }

                Object DtResultado;
                DtResultado = Comando.ExecuteScalar();

                return DtResultado;
            }
            catch (Exception ex)
            {
                //logErrores(Procedimiento, "Este procedimiento no tiene parámetros asociados", "GetEscalar", ex.Message.ToString());
                throw ex;
            }
            finally
            {
                if (Conexion.State != ConnectionState.Closed)
                {
                    Conexion.Close();
                }

                Comando.Parameters.Clear();
                Comando.Dispose();
            }
        }

        #endregion

        #region Eventos Execute

        //insert->
        public static void Execute(string Procedimiento, List<Parametros> Parametros)
        {
            SqlConnection Conexion = new SqlConnection();
            Conexion.ConnectionString = SqlConnectionString;
            StringBuilder CadenaSQL = new StringBuilder();
            SqlCommand Comando = new SqlCommand();

            try
            {
                Comando.Connection = Conexion;
                Comando.CommandType = CommandType.StoredProcedure;
                Comando.CommandText = Procedimiento;
                Comando.CommandTimeout = 60000;

                SqlParameter ObjSqlParametro;

                foreach (Parametros ObjSpParametro in Parametros)
                {
                    ObjSqlParametro = new SqlParameter();
                    ObjSqlParametro.ParameterName = ObjSpParametro.Nombre;
                    ObjSqlParametro.SqlDbType = ObjSpParametro.Tipo;
                    ObjSqlParametro.Value = ObjSpParametro.Valor;

                    ObjSqlParametro.Direction = ObjSpParametro.Direccion;

                    if (ObjSpParametro.Tipo == SqlDbType.VarChar)
                    {
                        ObjSqlParametro.Size = ObjSpParametro.Longitud;
                    }
                    CadenaSQL.Append(ObjSpParametro.Nombre);
                    CadenaSQL.Append(" = ");
                    CadenaSQL.Append(ObjSpParametro.Valor.ToString());
                    CadenaSQL.Append(" | ");

                    Comando.Parameters.Add(ObjSqlParametro);
                }

                if (Conexion.State != ConnectionState.Open)
                {
                    Conexion.Open();
                }

                Object DtResultado;
                DtResultado = Comando.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                logErrores(Procedimiento, CadenaSQL.ToString(), "Execute", ex.Message.ToString());
                throw ex;
            }
            finally
            {
                if (Conexion.State != ConnectionState.Closed)
                {
                    Conexion.Close();
                }

                Comando.Parameters.Clear();
                Comando.Dispose();
            }
        }

        public static void Execute(string Procedimiento)
        {
            SqlConnection Conexion = new SqlConnection();
            Conexion.ConnectionString = SqlConnectionString;
            SqlCommand Comando = new SqlCommand();

            try
            {
                Comando.Connection = Conexion;

                Comando.CommandType = CommandType.StoredProcedure;
                Comando.CommandText = Procedimiento;
                Comando.CommandTimeout = 60000;

                if (Conexion.State != ConnectionState.Open)
                {
                    Conexion.Open();
                }

                Object DtResultado;
                DtResultado = Comando.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                logErrores(Procedimiento, "Este procedimiento no tiene parámetros asociados", "Execute", ex.Message.ToString());
                throw ex;
            }
            finally
            {
                if (Conexion.State != ConnectionState.Closed)
                {
                    Conexion.Close();
                }

                Comando.Parameters.Clear();
                Comando.Dispose();
            }
        }

        #endregion

        #region Encriptacion

        public static string generarClaveSHA1(string nombre)
        {

            UTF8Encoding enc = new UTF8Encoding();
            byte[] data = enc.GetBytes(nombre);
            byte[] result;

            SHA1CryptoServiceProvider sha = new SHA1CryptoServiceProvider();

            result = sha.ComputeHash(data);

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < result.Length; i++)
            {
                if (result[i] < 16)
                {
                    sb.Append("0");
                }
                sb.Append(result[i].ToString("x"));
            }

            return sb.ToString().ToUpper();
        }

        #endregion




        #endregion

    }

}
