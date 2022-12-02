using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CONTROLDEINGRESOS.Models
{
    public class Empleado
    {
        public decimal cedula { get; set; }
        public string nombre { get; set; }
        public int sedeID { get; set; }
        public string sede { get; set; }
        public string imagen { get; set; }
        public string mensaje { get; set; }
        public string fechaUltimo { get; set; }
        public string tipo { get; set; }


    }
}