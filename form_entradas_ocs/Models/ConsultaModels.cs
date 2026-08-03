using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using form_entradas_ocs.Data;
using System.Data;
using System.Data.SqlClient;

namespace form_entradas_ocs.Models
{
    public class ConsultaModels
    {
        public DataTable VerificarOcReparacion(string folio)
        {
            DataTable dtConsulta = new DataTable();
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("spSISEMPEntradasNorVerificarOcReparacion", conn);
                adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapter.SelectCommand.Parameters.AddWithValue("@noc", folio);
                adapter.Fill(dtConsulta);
            }

            return dtConsulta;
        }

        public void ActualizaCostoCatalogoEmpaquesOrdenReparacion(string emp, string alm, string cto, string fecha, string nom_compu, string nom_usu, string tipo_mov, string op_clave,
            string folio, string detalle, string sistema, string motivo, string obs)
        {
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand("spSISEMPEntradasNorModificarCostoCatEmpaques", conn))
                {
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@emp", emp);
                    cmd.Parameters.AddWithValue("@alm", alm);
                    cmd.Parameters.AddWithValue("@cto", cto);
                    cmd.Parameters.AddWithValue("@fecha", Convert.ToDateTime(fecha));
                    cmd.Parameters.AddWithValue("@nom_compu", nom_compu);
                    cmd.Parameters.AddWithValue("@nom_usu", nom_usu);
                    cmd.Parameters.AddWithValue("@tipo_mov", tipo_mov);
                    cmd.Parameters.AddWithValue("@op_clave", op_clave);
                    cmd.Parameters.AddWithValue("@folio", folio);
                    cmd.Parameters.AddWithValue("@detalle", detalle);
                    cmd.Parameters.AddWithValue("@sistema", sistema);
                    cmd.Parameters.AddWithValue("@motivo", motivo);
                    cmd.Parameters.AddWithValue("@obs", obs);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        public void ActualizaCostoAguilares(string emp, string alm, string cto)
        {
            DataTable dtConsulta = new DataTable();
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionAgui))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("spSISEMPActualizaCostoEmpaqueGAB", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@emp", emp);
                cmd.Parameters.AddWithValue("@alm", alm);
                cmd.Parameters.AddWithValue("@cto", Convert.ToDecimal(cto));
                cmd.ExecuteNonQuery();
            }
        }

        public void GuardarPedimentoFactura(string ent, string alm, string url, string mov)
        {
            string ENTR = ent;
            string ALMA = alm;
            string URLX = url;
            string MOVI = mov;

            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlCommand command = new SqlCommand("spSISEMPEntradasPedimentoFactura", conn);
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@mov", Convert.ToInt32(MOVI));
                command.Parameters.AddWithValue("@url", URLX);
                command.Parameters.AddWithValue("@fol", ENTR);
                command.Parameters.AddWithValue("@alm", ALMA);
                command.Parameters.AddWithValue("@i_d", 0);
                command.ExecuteNonQuery();
            }
        }

        public DataTable EliminarPedimentoFactura(string ent, string alm, string url, string mov, string idx)
        {
            string ENTR = ent;
            string ALMA = alm;
            string URLX = url;
            string MOVI = mov;
            string IDXX = idx;

            DataTable dtConsulta = new DataTable();
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("spSISEMPEntradasPedimentoFactura", conn);
                adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapter.SelectCommand.Parameters.AddWithValue("@mov", Convert.ToInt32(MOVI));
                adapter.SelectCommand.Parameters.AddWithValue("@url", URLX);
                adapter.SelectCommand.Parameters.AddWithValue("@fol", ENTR);
                adapter.SelectCommand.Parameters.AddWithValue("@alm", ALMA);
                adapter.SelectCommand.Parameters.AddWithValue("@i_d", IDXX);
                adapter.Fill(dtConsulta);
            }

            return dtConsulta;
        }

        public void RegistroMovimiento(string compu, string usu, string mov, string opc, string folio, string detalle, string sis, string mot)
        {
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlCommand command = new SqlCommand("spSISEMPRegistroMovimiento", conn);
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@fecha", DateTime.Now);
                command.Parameters.AddWithValue("@nom_compu", compu);
                command.Parameters.AddWithValue("@nom_usu", usu);
                command.Parameters.AddWithValue("@tipo_mov", mov);
                command.Parameters.AddWithValue("@op_clave", opc);
                command.Parameters.AddWithValue("@folio", folio);
                command.Parameters.AddWithValue("@detalle", detalle);
                command.Parameters.AddWithValue("@sistema", sis);
                command.Parameters.AddWithValue("@motivo", mot);
                command.Parameters.AddWithValue("@obs", "");
                command.ExecuteNonQuery();
            }
        }
    }
}
