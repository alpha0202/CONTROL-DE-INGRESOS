using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CONTROLDEINGRESOS.Models
{
    public class Visitante
    {
        public decimal cedula { get; set; }
        public string nombre { get; set; }
        public string arl { get; set; }
        public string fechaCreacion { get; set; }
        public string usuarioCreacion { get; set; }
        public string empleadoAutoriza { get; set; }
        public string motivoVisita { get; set; }
        public string placa { get; set; }
        public string empresa { get; set; }
        public bool frecuente { get; set; }
        public string fechaIniFrecuente { get; set; }
        public string fechaFinFrecuente { get; set; }
        public string mensaje { get; set; }
        public string imagen { get; set; }
        public string fechaUltimo { get; set; }
        public string tipo { get; set; }

        public Visitante()
        {
            cedula = 0;
            nombre = "";
            arl = "";
            fechaCreacion = "";
            usuarioCreacion = "";
            empleadoAutoriza = "";
            motivoVisita = "";
            placa = "";
            empresa = "";
            frecuente = false;
            fechaIniFrecuente = "";
            fechaFinFrecuente = "";
            mensaje = "";
        }




    }
}