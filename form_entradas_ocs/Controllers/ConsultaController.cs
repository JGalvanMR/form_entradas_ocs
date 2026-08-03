using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using form_entradas_ocs.Models;
using System.Data;

namespace form_entradas_ocs.Controllers
{
    public class ConsultaController
    {
        private ConsultaModels _consultaModel;

        public ConsultaController()
        {
            _consultaModel = new ConsultaModels();
        }

        public DataTable VerificarOcReparacion(string folio)
        {
            return _consultaModel.VerificarOcReparacion(folio);
        }

        public void ActualizaCostoCatalogoEmpaquesOrdenReparacion(string emp, string alm, string cto, string fecha, string nom_compu, string nom_usu, string tipo_mov, string op_clave,
            string folio, string detalle, string sistema, string motivo, string obs)
        {
            _consultaModel.ActualizaCostoCatalogoEmpaquesOrdenReparacion(emp, alm, cto, fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, motivo, obs);
        }

        public void ActualizaCostoAguilares(string emp, string alm, string cto)
        {
            _consultaModel.ActualizaCostoAguilares(emp, alm, cto);
        }

        public void GuardarPedimentoFactura(string ent, string alm, string url, string mov)
        {
            _consultaModel.GuardarPedimentoFactura(ent, alm, url, mov);
        }

        public DataTable EliminarPedimentoFactura(string ent, string alm, string url, string mov, string i_d)
        {
            return _consultaModel.EliminarPedimentoFactura(ent, alm, url, mov, i_d);
        }

        public void RegistroMovimiento(string compu, string usu, string mov, string opc, string folio, string detalle, string sis, string mot)
        {
            _consultaModel.RegistroMovimiento(compu, usu, mov, opc, folio, detalle, sis, mot);
        }
    }
}
