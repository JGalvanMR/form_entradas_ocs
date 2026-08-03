using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.IO;
using form_entradas_ocs.Controllers;

namespace form_entradas_ocs
{
    public partial class fact_extras : Form
    {
        private ConsultaController _consultaController;

        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);//
        SqlDataReader reader1, reader2, reader3, reader4;
        SqlCommand cmnd4;
        SqlCommand cmnd3;
        SqlCommand cmnd2;
        SqlCommand cmnd1;
        SqlDataAdapter adapter1;
        DataSet set = new DataSet();
        string prv = "";
        string ent = "";
        string tip = "";
        string alm = "";
        public fact_extras(string entrada, string cve_alm, string nom_alm, string cve_prv, string f_ent, string tipo)
        {
            InitializeComponent();

            _consultaController = new ConsultaController();
            
            prv = cve_prv;
            ent = entrada;
            tip = tipo;
            alm = cve_alm;

            if (tipo == "0")
            {
                thisConnection.Open();
                adapter1 = new SqlDataAdapter("SELECT factura AS Factura, CONVERT(DATETIME, CONVERT(varchar, fecha_factura, 103)) AS Fecha_Factura, FORMAT(monto, 'C2', 'es-mx') AS Monto " +
                    "FROM tb_det_oc_facturas WHERE numero_oc = '" + entrada + "' and tipo = '" + nom_alm.ToUpper() + "' ORDER BY fecha_factura", thisConnection);
                adapter1.Fill(set, "facturas");
                dtgFacturas.DataSource = set.Tables["facturas"];
                thisConnection.Close();

                dtgFacturas.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dtgFacturas.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            else
            {
                cargar_datos_pedimento();
            }
            
        }

        private void dtgFacturas_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            

            
        }

        public void cargar_datos_pedimento()
        {
            DataTable dtPedi = _consultaController.EliminarPedimentoFactura(ent, alm, "0", "0", "0");

            this.Text = "Pedimento y Factura";
            dtgFacturas.DataSource = dtPedi;
            //thisConnection.Open();
            //adapter1 = new SqlDataAdapter("SELECT ruta AS Ruta, id AS ID" +
            //    "FROM tb_ent_pedimento WHERE folio = '" + entrada + "' and almacen = '" + cve_alm + "'", thisConnection);
            //adapter1.Fill(set, "facturas");
            //dtgFacturas.DataSource = set.Tables["facturas"];
            //thisConnection.Close();

            //dtgFacturas.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dtgFacturas.Columns[0].Width = 250;
            dtgFacturas.Columns[1].Width = 50;
        }

        private void dtgFacturas_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                string cad = prv + "_" + dtgFacturas.CurrentRow.Cells[0].Value.ToString() + "_" + ent.PadLeft(6, '0').ToString() + ".pdf";
                if (tip == "0")
                    Process.Start(@"\\gabira1\FacturasExternos\\" + cad);
                else
                    Process.Start(@dtgFacturas.CurrentRow.Cells[0].Value.ToString());//"\\gabira1\FacturasExternos\\" + cad
            }
            else
            {
                if (tip == "1")
                {
                    if (MessageBox.Show("¿Desea modificar el archivo?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        return;

                    openFileDialog1.InitialDirectory = "c:\\";
                    openFileDialog1.Filter = "pdf files (*.pdf)|*.pdf";
                    openFileDialog1.FilterIndex = 2;
                    openFileDialog1.RestoreDirectory = true;
                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        string textoABuscar = openFileDialog1.FileName;

                        string pattern = @"[\\\/:\*\?""<>|#\$]";

                        Regex regex = new Regex(pattern);
                        bool containsSpecialCharacters = regex.IsMatch(Path.GetFileName(openFileDialog1.FileName));
                        if (containsSpecialCharacters == true)
                        {
                            MessageBox.Show("El nombre del archivo contiene caracteres especiales ( \\ / : * ?  < > | # $ ) favor de removerlos e intentar nuevamente", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        bool encontrado = false;
                        foreach (DataGridViewRow item in dtgFacturas.Rows)
                        {
                            if (item.ToString().Equals(textoABuscar, StringComparison.OrdinalIgnoreCase))
                            {
                                encontrado = true;
                                break;
                            }
                        }
                        if (encontrado)
                        {
                            MessageBox.Show("Archivo seleccionado anteriormente", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }


                        string path_destino = "\\\\gabira1\\Pedimentos\\";

                        string dest = path_destino + Path.GetFileName(textoABuscar);
                        _consultaController.GuardarPedimentoFactura(ent, alm, dest, "1");

                        FileInfo archivo = new FileInfo(textoABuscar);
                        FileInfo doc_copy = new FileInfo(@dest);
                        archivo.CopyTo(@dest, true);

                        string id = dtgFacturas.CurrentRow.Cells[1].Value.ToString();
                        DataTable dtPedi = _consultaController.EliminarPedimentoFactura(ent, alm, dest, "2", id);

                        cargar_datos_pedimento();
                    }
                }
            }
            
        }
    }
}
