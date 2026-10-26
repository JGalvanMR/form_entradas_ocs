using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using System.Collections;
using System.Text.RegularExpressions;
using System.Xml;
using System.Net.NetworkInformation;
using System.Net.Mail;
using System.Net;
using form_entradas_ocs.Properties;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Drawing.Printing;
using System.Printing;
using form_entradas_ocs.Controllers;
//using Ionic.Zip;




namespace form_entradas_ocs
{
    public partial class Form1 : Form
    {
        private const int EM_SETCUEBANNER = 0x1501;
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern Int32 SendMessage(IntPtr hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)]string lParam);

        private ConsultaController _consultaController;

        //SqlConnection thisConnection = new SqlConnection("Data Source=GABIRA1\\SQL2005;Initial Catalog=GAB_Irapuato;Connect Timeout=130;User ID=sa; MultipleActiveResultSets=True");//Utilerias.Class1.ConnectionString
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);//
        SqlDataReader reader1, reader2, reader3, reader4;
        SqlCommand cmnd4;
        SqlCommand cmnd3;
        SqlCommand cmnd2;
        SqlCommand cmnd1;
        SqlDataAdapter adapter1;

        //SqlConnection conn = new SqlConnection("server=agricolasgrupou.com.mx.\\MSSQLSERVER2012; Initial Catalog = dbGuService; User Id = Cfdi; Password = Prv2017#i");
        //SqlConnection conn = new SqlConnection("server=grupou.net.,1433\\MSSQLSERVER2012; Initial Catalog = grupou_WMService; User Id = grupou_Cfdi; Password = Prv2017#i");
        SqlConnection conn = new SqlConnection("server=grupou-mrlucky.com.mx,15333; Initial Catalog = grupou_Service; User Id = mrlucky; Password = Cfdi79#!=");
        SqlCommand comm;
        SqlDataReader redr;

        DataTable dtordenes = new DataTable();

        public int glo_var_int_registro = 0;
        public string glo_var_chr_alm_clave;
        public string var_int_sis_folios = "";

        public List<string> provclave = new List<string>();
        public List<string> provnombre = new List<string>();

        DataTable filtroprov = new DataTable();

        string provnotfound = "";

        DataTable tmgrdpro = new DataTable();
        DataTable dtempoc = new DataTable();

        decimal iva1 = 0;
        decimal subtotal1 = 0;
        decimal totaladv = 0;
        decimal totaldta = 0;
        decimal totalieps = 0;
        decimal subtotal2 = 0;
        decimal iva2 = 0;
        decimal total = 0;
        
        List<string> facturas = new List<string>();
        List<string> pedimentos = new List<string>();

        public List<string> usuarionombre = new List<string>();
        

        decimal iepsporc = Convert.ToDecimal("0.08");

        public int ren = 0;
        public int renglonant = 0;

        public decimal nuevo_costo = 0;
        public decimal nueva_existencia = 0;
        public decimal nueva_costoprom = 0;

        public string operacion = "";

        public string subtipo = "";

        DataTable dtdetalles = new DataTable();
        DataTable dtoriginales = new DataTable();

        DataTable dtempaques = new DataTable();

        DataTable dtProveedor = new DataTable();

        DataTable dtPortal = new DataTable();

        bool valida_carga_xml = false;

        DataTable dtUsuarios1 = new DataTable();
        DataTable dtUsuarios2 = new DataTable();

        DataTable dtempaque = new DataTable();
        DataTable dtinv = new DataTable();
        DataTable dtconsal = new DataTable();
        DataTable dtconent = new DataTable();

        bool fnd_ieps = false;


        public Form1()
        {
            InitializeComponent();

            _consultaController = new ConsultaController();
            //conn.Open();
            //conn.Close();


            //glo_var_chr_alm_clave = "02";
            //decimal promedio = 0, C = 0;
            //promedio = Convert.ToDecimal(recalculo_costos_empaque("MC004", "20/01/2021"));
            //C = Convert.ToDecimal(recalculo_existencia_empaque("MC004"));


            SendMessage(txt_oc.Handle, EM_SETCUEBANNER, 0, "ORDEN DE COMPRA");

            deshabilitarnuevo();
            lblServidor.Text = Utilerias.Class1.ConnectionString;
            dtpFecha.Text = DateTime.Now.ToShortDateString();

            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT GETDATE() as fecha";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                dtpFecha.Text = Convert.ToDateTime(reader1["fecha"].ToString()).ToShortDateString();
                lblFechaDia.Text = Convert.ToDateTime(reader1["fecha"].ToString()).ToShortDateString();
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();


            dtgOrdenes.EnableHeadersVisualStyles = false;
            dtgOrdenes.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;

            dtgDetalle.EnableHeadersVisualStyles = false;
            dtgDetalle.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;

            dtgDetalle2.EnableHeadersVisualStyles = false;
            dtgDetalle2.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            //crear carpeta para los XML
            string pathEmp = @"c:\\EMPAQUES";
            string pathMat = @"c:\\MATERIALES";
            string pathCom = @"c:\\COMEDOR";
            string pathPap = @"c:\\PAPELERIA";
            string pathDen = @"c:\\DENTAL";
            string pathComb = @"c:\\COMBUSTIBLE";
            string pathSis = @"c:\\SISTEMAS";
            string pathGua = @"c:\\GUARDERIA";
            string pathGuaCom = @"c:\\GUARDERIACOM";
            string pathEven = @"c:\\EVENTOS";
            string pathMoto = @"c:\\MOTORES";


            if (!Directory.Exists(pathEmp))
            {
                Directory.CreateDirectory(pathEmp);
            }
            if (!Directory.Exists(pathMat))
            {
                Directory.CreateDirectory(pathMat);
            }
            if (!Directory.Exists(pathCom))
            {
                Directory.CreateDirectory(pathCom);
            }
            if (!Directory.Exists(pathPap))
            {
                Directory.CreateDirectory(pathPap);
            }
            if (!Directory.Exists(pathDen))
            {
                Directory.CreateDirectory(pathDen);
            }
            if (!Directory.Exists(pathComb))
            {
                Directory.CreateDirectory(pathComb);
            }
            if (!Directory.Exists(pathSis))
            {
                Directory.CreateDirectory(pathSis);
            }
            if (!Directory.Exists(pathGua))
            {
                Directory.CreateDirectory(pathGua);
            }
            if (!Directory.Exists(pathGuaCom))
            {
                Directory.CreateDirectory(pathGuaCom);
            }
            if (!Directory.Exists(pathEven))
            {
                Directory.CreateDirectory(pathEven);
            }
            if (!Directory.Exists(pathMoto))
            {
                Directory.CreateDirectory(pathMoto);
            }


            dtordenes.Columns.Add("numero_oc", typeof(string));
            dtordenes.Columns.Add("cveprov_oc", typeof(string));
            dtordenes.Columns.Add("nomprov_oc", typeof(string));
            dtordenes.Columns.Add("tipo_oc", typeof(string));

            filtroprov.Columns.Add("numero_oc", typeof(string));
            filtroprov.Columns.Add("cveprov_oc", typeof(string));
            filtroprov.Columns.Add("nomprov_oc", typeof(string));
            filtroprov.Columns.Add("tipo_oc", typeof(string));

            dtProveedor.Columns.Add("prov_clave", typeof(string));
            dtProveedor.Columns.Add("prov_xml", typeof(string));
            dtProveedor.Columns.Add("prov_descuento", typeof(string));
            dtProveedor.Columns.Add("prov_rfc", typeof(string));
            dtProveedor.Columns.Add("prov_factnal", typeof(string));
            dtProveedor.Columns.Add("prov_pedi", typeof(string));


            tmgrdpro.Columns.Add("clave", Type.GetType("System.String"));//0
            tmgrdpro.Columns.Add("nombre", Type.GetType("System.String"));//1
            tmgrdpro.Columns.Add("cantidad", Type.GetType("System.String"));//2
            tmgrdpro.Columns.Add("unitario", Type.GetType("System.String"));//3
            tmgrdpro.Columns.Add("costo", Type.GetType("System.String"));//4
            //tmgrdpro.Columns.Add("costo_mn", Type.GetType("System.String"));//5
            tmgrdpro.Columns.Add("adv", Type.GetType("System.String"));//5
            tmgrdpro.Columns.Add("dta", Type.GetType("System.String"));//6
            tmgrdpro.Columns.Add("lote", Type.GetType("System.String"));//7
            tmgrdpro.Columns.Add("ieps", Type.GetType("System.Boolean"));//8
            tmgrdpro.Columns.Add("iepsimporte", Type.GetType("System.String"));//9
            tmgrdpro.Columns.Add("gastoaduanal", Type.GetType("System.String"));//10
            tmgrdpro.Columns.Add("flete", Type.GetType("System.String"));//11
            tmgrdpro.Columns.Add("oc", typeof(string));//12
            tmgrdpro.Columns.Add("req", typeof(string));//13
            tmgrdpro.Columns.Add("surtido", typeof(string));//14
            tmgrdpro.Columns.Add("cantorig", typeof(string));//15
            tmgrdpro.Columns.Add("unitorig", typeof(string));//16
            tmgrdpro.Columns.Add("borrar", typeof(string));//17
            tmgrdpro.Columns.Add("oper", typeof(string));//18
            tmgrdpro.Columns.Add("iva", typeof(string));//18

            dtempoc.Columns.Add("numero_oc", typeof(string));
            dtempoc.Columns.Add("prod_clave", typeof(string));
            dtempoc.Columns.Add("cantidad_oc", typeof(string));
            dtempoc.Columns.Add("surtido_oc", typeof(string));

            dtdetalles.Columns.Add("emp_clave", typeof(string));
            dtdetalles.Columns.Add("orden_compra", typeof(string));
            dtdetalles.Columns.Add("cantidad", typeof(string));
            dtdetalles.Columns.Add("importe", typeof(string));
            dtdetalles.Columns.Add("requisicion", typeof(string));
            dtdetalles.Columns.Add("costo", typeof(string));
            dtdetalles.Columns.Add("borrar", typeof(string));
            dtdetalles.Columns.Add("operacion", typeof(string));
            dtdetalles.Columns.Add("renglon", typeof(string));
            dtdetalles.Columns.Add("ieps", typeof(string));
            dtdetalles.Columns.Add("lote", typeof(string));
            dtdetalles.Columns.Add("importe_org", typeof(string));
            dtdetalles.Columns.Add("costo_org", typeof(string));
            dtdetalles.Columns.Add("iva", typeof(string));
            dtdetalles.Columns.Add("tiva", typeof(string));
            dtdetalles.Columns.Add("importe_orig", typeof(string));
            dtdetalles.Columns.Add("urgente", typeof(string));
            dtdetalles.Columns.Add("cant_orig", typeof(string));
            dtdetalles.Columns.Add("resp", typeof(string));

            dtoriginales.Columns.Add("emp_clave", typeof(string));
            dtoriginales.Columns.Add("orden_compra", typeof(string));
            dtoriginales.Columns.Add("cantidad", typeof(string));
            dtoriginales.Columns.Add("importe", typeof(string));
            dtoriginales.Columns.Add("requisicion", typeof(string));
            dtoriginales.Columns.Add("costo", typeof(string));
            dtoriginales.Columns.Add("borrar", typeof(string));
            dtoriginales.Columns.Add("operacion", typeof(string));
            dtoriginales.Columns.Add("renglon", typeof(string));
            dtoriginales.Columns.Add("ieps", typeof(string));
            dtoriginales.Columns.Add("iva", typeof(string));


            dtempaques.Columns.Add("emp_clave", typeof(string));
            dtempaques.Columns.Add("emp_nombre", typeof(string));

            dtPortal.Columns.Add("cfd_keyfec", typeof(string));
            dtPortal.Columns.Add("cfd_timfis", typeof(string));
            dtPortal.Columns.Add("cfd_totalv", typeof(string));
            dtPortal.Columns.Add("cfd_sercfd", typeof(string));
            dtPortal.Columns.Add("cfd_folioc", typeof(string));
            dtPortal.Columns.Add("cfd_keyems", typeof(string));

            dtUsuarios1.Columns.Add("nombre", typeof(string));
            dtUsuarios1.Columns.Add("estatus", typeof(string));

            dtUsuarios2.Columns.Add("nombre", typeof(string));
            dtUsuarios2.Columns.Add("estatus", typeof(string));

            iepsporc = Convert.ToDecimal("0.08");

            thisConnection.Open();
            #region proveedores
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT prov_clave, prov_nombre, prov_xml, prov_descuento, prov_rfc, prov_factnal, prov_pedi FROM tb_cat_proveedor ORDER BY prov_nombre";
            reader1 = cmnd1.ExecuteReader();
            DataRow re;
            while (reader1.Read())
            {
                provclave.Add(reader1.GetValue(0).ToString().Trim());
                provnombre.Add(reader1.GetValue(1).ToString().Trim());

                re = dtProveedor.NewRow();
                re["prov_clave"] = reader1["prov_clave"].ToString();
                re["prov_xml"] = reader1["prov_xml"].ToString();
                re["prov_descuento"] = reader1["prov_descuento"].ToString().Trim();
                re["prov_rfc"] = reader1["prov_rfc"].ToString().Trim();
                re["prov_factnal"] = reader1["prov_factnal"].ToString().Trim();
                re["prov_pedi"] = reader1["prov_pedi"].ToString().Trim();
                dtProveedor.Rows.Add(re);
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            for (int i = 0; i < provnombre.Count; i++)
            {
                DDLProveedor.Items.Add(provnombre[i].ToString());
            }
            #endregion
            #region usuarios
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT usu_nombre FROM tb_cat_usuarios_emp WHERE grp_clave = '01' OR grp_clave = '02' ORDER BY usu_nombre";
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                usuarionombre.Add(reader1.GetValue(0).ToString().Trim());
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            //for (int i = 0; i < usuarionombre.Count; i++)
            //{
            //    DDLElaboro.Items.Add(usuarionombre[i]);
            //    DDLAutorizo.Items.Add(usuarionombre[i]);
            //}
            cmnd1 = thisConnection.CreateCommand();

            cmnd1.CommandText = "SELECT grp_clave FROM tb_cat_usuarios_emp WHERE usu_login = '" + Utilerias.Class1.Usu_login + "'";
            lblgrupo.Text = Convert.ToString(cmnd1.ExecuteScalar().ToString().Trim());
            cmnd1.Dispose();

            cmnd1 = thisConnection.CreateCommand();
            if (lblgrupo.Text == "01")
                cmnd1.CommandText = "SELECT usu_nombre, estatus FROM tb_cat_usuarios_emp WHERE grp_clave = '01' OR grp_clave = '02' ORDER BY usu_nombre";
            else
                cmnd1.CommandText = "SELECT usu_nombre, estatus FROM tb_cat_usuarios_emp WHERE grp_clave = '" + lblgrupo.Text + "' ORDER BY usu_nombre";
            reader1 = cmnd1.ExecuteReader();
            usuarionombre.Clear();
            while (reader1.Read())
            {
                usuarionombre.Add(reader1.GetValue(0).ToString().Trim());
            }
            //DataRow ru;
            //if(reader1.HasRows)
            //{
            //    while(reader1.Read())
            //    {
            //        ru = dtUsuarios1.NewRow();
            //        ru["nombre"] = reader1["usu_nombre"].ToString().Trim();
            //        ru["estatus"] = reader1["estatus"].ToString().Trim();
            //        dtUsuarios1.Rows.Add(ru);

            //        ru = dtUsuarios2.NewRow();
            //        ru["nombre"] = reader1["usu_nombre"].ToString().Trim();
            //        ru["estatus"] = reader1["estatus"].ToString().Trim();
            //        dtUsuarios2.Rows.Add(ru);
            //    }
            //}
            //reader1.Close();
            //reader1.Dispose();
            //cmnd1.Dispose();

            //DDLAutorizo.ValueMember = "nombre";
            //DDLAutorizo.DisplayMember = "nombre";

            //object[] elementos;
            //object elemento;
            //foreach (DataRow ru1 in dtUsuarios1.Rows)
            //{
            //    elementos = new [] {new ComboBoxItem(){Selectable = false, Text = ru1[0].ToString(), Value = ru1[0].ToString()}};
            //    //if(ru1[1].ToString() == "B")
            //    //    DDLAutorizo.Items.Add(new ComboBoxItem() { Selectable = false, Text = ru1[0].ToString(), Value=ru1[0].ToString()});
            //    //else
            //    //    DDLAutorizo.Items.Add(new ComboBoxItem() { Selectable = true, Text = ru1[0].ToString(), Value = ru1[0].ToString()});
            //}
            //DDLAutorizo.SelectedIndexChanged += (cbSender, cbe) =>
            //{
            //    var cb = cbSender as ComboBox;
            //    if (cb.SelectedItem != null && cb.SelectedItem is ComboBoxItem && ((ComboBoxItem)cb.SelectedItem).Selectable == false)
            //    {
            //        cb.SelectedIndex = -1;
            //        MessageBox.Show("You cannot select this Item");
            //    }
            //};
            
            //DDLAutorizo.Items.AddRange(elementos);

            for (int i = 0; i < usuarionombre.Count; i++)
            {

                DDLElaboro.Items.Add(usuarionombre[i]);
                DDLAutorizo.Items.Add(usuarionombre[i]);
            }
            #endregion

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT TOP 1 invemp_fecha FROM tb_mstr_inventario_emp ORDER BY invemp_fecha DESC, emp_clave DESC";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    txtCierre.Text = Convert.ToDateTime(reader1.GetValue(0).ToString().Trim()).ToShortDateString();
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            thisConnection.Close();

            btnBuscar.Enabled = true;
            btnCancelar.Enabled = false;
            btnComedor.Enabled = true;
            btnEditar.Enabled = false;
            btnGuardar.Enabled = true;
            btnImprimir.Enabled = false;
            btnLimpiar.Enabled = true;


            string filelog = "C:\\SisEmpWeb\\eventlog.txt";
            using (StreamWriter sw = File.AppendText(filelog))
            {
                sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Acceso 2.4 Entradas de Almacen");
                sw.Close();
            }


            dtempaque.Columns.Add("emp_clave", typeof(string));
            dtempaque.Columns.Add("emp_nombre", typeof(string));
            dtempaque.Columns.Add("invemp_inicial", typeof(string));
            dtempaque.Columns.Add("tent", typeof(string));
            dtempaque.Columns.Add("tsal", typeof(string));
            dtempaque.Columns.Add("inv_actual", typeof(string));
            dtempaque.Columns.Add("costo", typeof(string));

            dtinv.Columns.Add("emp_clave", typeof(string));
            dtinv.Columns.Add("inicial", typeof(string));
            dtinv.Columns.Add("costo", typeof(string));

            dtconent.Columns.Add("emp_clave", typeof(string));
            dtconent.Columns.Add("total", typeof(string));

            dtconsal.Columns.Add("emp_clave", typeof(string));
            dtconsal.Columns.Add("total", typeof(string));

            string str1_doc = "c:\\Reportes\\document2.pdf";
            if(!File.Exists(@str1_doc))
            {
                File.Copy(@"c:\\Reportes\\document.pdf", @"c:\\Reportes\\document2.pdf"); 
            }

            if (Utilerias.Class1.Usu_login.Trim() == "PEDRO" || Utilerias.Class1.Usu_login.Trim() == "N")
            {
                btnFactExtra.Visible = true;
            }


            //string rec = recalculo_existencia_empaque("C0143");
            //Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "MF", "2.5", "0000", "FACTURA REGISTRADA Entrada: 0 Factura: 0 Fecha Factura: 0 Fecha Registro: 0 Almacen: 0", "SISEMP");

            //string pathito = "";//System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            //pathito = @"c:\\Reportes\\document.pdf";
            //string salida = ReadPdfFile(pathito);
        }

        private class ComboBoxItem
        {
            public string Value { get; set; }
            public string Text { get; set; }
            public bool Selectable { get; set; }
        }

        private void DDLAlmacen1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                if (DDLAlmacen1.SelectedIndex == 0)
                {
                    glo_var_int_registro = 5;
                    lblCierreInventario.Text = "Entradas al Almacén de Materiales";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "01";
                    
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_mat";
                    reader1 = cmnd1.ExecuteReader();
                    string folmat = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        folmat = reader1.GetValue(0).ToString().Trim();
                    }
                    folmat = (Convert.ToInt32(folmat) + 1).ToString();
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    txtFolio.Text = folmat.PadLeft(6, '0');
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "SELECT TOP 1 ent_folio FROM tb_mstr_entradas WHERE alm_clave = '" + glo_var_chr_alm_clave + "' ORDER BY ent_folio DESC";// ent_fecha DESC, 
                    //reader1 = cmnd1.ExecuteReader();
                    //reader1.Read();
                    //var_int_sis_folios = Convert.ToInt32(reader1.GetValue(0).ToString()) + 1;
                    //reader1.Close();
                    //reader1.Dispose();
                    //txtFolio.Text = Convert.ToString(var_int_sis_folios).PadLeft(6, '0');
                    subtipo = "";
                    txt_ent_iva.Text = "0";
                    //var_int_sis_folios = Convert.ToInt32(sisfolio[4]) + 1;
                    //folio_sis = 4;
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }
                if (DDLAlmacen1.SelectedIndex == 1)
                {
                    glo_var_int_registro = 6;
                    lblCierreInventario.Text = "Entradas al Almacén de Empaques";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "02";
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_emp";
                    reader1 = cmnd1.ExecuteReader();
                    string folemp = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        folemp = reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folemp = (Convert.ToInt32(folemp) + 1).ToString();
                    txtFolio.Text = folemp.PadLeft(6, '0');
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "SELECT TOP 1 ent_folio FROM tb_mstr_entradas WHERE alm_clave = '" + glo_var_chr_alm_clave + "' ORDER BY  ent_folio DESC";//ent_fecha DESC,
                    //reader1 = cmnd1.ExecuteReader();
                    //reader1.Read();
                    //var_int_sis_folios = Convert.ToInt32(reader1.GetValue(0).ToString()) + 1;
                    //reader1.Close();
                    //reader1.Dispose();
                    //txtFolio.Text = Convert.ToString(var_int_sis_folios).PadLeft(6, '0');
                    subtipo = "";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    /////btnXml.Visible = false;
                }
                if (DDLAlmacen1.SelectedIndex == 2)
                {
                    glo_var_int_registro = 7;
                    lblCierreInventario.Text = "Entradas al Almacén del Comedor";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "03";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_com";
                    reader1 = cmnd1.ExecuteReader();
                    string folcom = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        folcom = reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folcom = (Convert.ToInt32(folcom) + 1).ToString();
                    txtFolio.Text = folcom.PadLeft(6, '0');
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "SELECT TOP 1 ent_folio FROM tb_mstr_entradas WHERE alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_subtipo = 'COM' ORDER BY ent_folio DESC";// ent_fecha DESC,
                    //reader1 = cmnd1.ExecuteReader();
                    //reader1.Read();
                    //var_int_sis_folios = Convert.ToInt32(reader1.GetValue(0).ToString()) + 1;
                    //reader1.Close();
                    //reader1.Dispose();
                    //txtFolio.Text = Convert.ToString(var_int_sis_folios).PadLeft(6, '0');
                    subtipo = "COM";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }
                if (DDLAlmacen1.SelectedIndex == 3)
                {
                    glo_var_int_registro = 9;
                    lblCierreInventario.Text = "Entradas al Almacén del Comedor Guardería";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "03";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_gua";
                    reader1 = cmnd1.ExecuteReader();
                    string folgua = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        folgua = reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folgua = (Convert.ToInt32(folgua) + 1).ToString();
                    txtFolio.Text = folgua.PadLeft(6, '0');

                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "SELECT TOP 1 ent_folio FROM tb_mstr_entradas WHERE alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_subtipo = 'GUA' ORDER BY ent_folio DESC";//ent_fecha DESC,
                    //reader1 = cmnd1.ExecuteReader();
                    //reader1.Read();
                    //var_int_sis_folios = Convert.ToInt32(reader1.GetValue(0).ToString()) + 1;
                    //reader1.Close();
                    //reader1.Dispose();
                    //txtFolio.Text = Convert.ToString(var_int_sis_folios).PadLeft(6, '0');
                    subtipo = "GUA";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }

                if (DDLAlmacen1.SelectedIndex == 4)
                {
                    glo_var_int_registro = 10;
                    lblCierreInventario.Text = "Entradas al Almacén de Enfermería";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "04";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_enf";
                    reader1 = cmnd1.ExecuteReader();
                    string folenf = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        folenf = reader1.GetValue(0).ToString().Trim();
                    }
                    folenf = (Convert.ToInt32(folenf) + 1).ToString();
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    txtFolio.Text = folenf.PadLeft(6, '0');
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }

                if (DDLAlmacen1.SelectedIndex == 5)
                {
                    glo_var_int_registro = 6;
                    lblCierreInventario.Text = "Entradas al Almacén de Papeleria";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "05";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_pap";
                    reader1 = cmnd1.ExecuteReader();
                    string folpap = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        folpap = reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folpap = (Convert.ToInt32(folpap) + 1).ToString();
                    txtFolio.Text = folpap.PadLeft(6, '0');
                    btnComedor.Enabled = false;
                    subtipo = "";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }

                if (DDLAlmacen1.SelectedIndex == 6)
                {
                    glo_var_int_registro = 6;
                    lblCierreInventario.Text = "Entradas al Almacén Dental";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "07";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_den";
                    reader1 = cmnd1.ExecuteReader();
                    string folpap = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        folpap = (reader1.GetValue(0).ToString().Trim() == "") ? "0" : reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folpap = (Convert.ToInt32(folpap) + 1).ToString();
                    txtFolio.Text = folpap.PadLeft(6, '0');
                    btnComedor.Enabled = false;
                    subtipo = "";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }

                if (DDLAlmacen1.SelectedIndex == 7)
                {
                    glo_var_int_registro = 6;
                    lblCierreInventario.Text = "Entradas al Almacén de Combustible";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "08";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_comb";
                    reader1 = cmnd1.ExecuteReader();
                    string folpap = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        //folpap = reader1.GetValue(0).ToString().Trim();
                        folpap = (reader1.GetValue(0).ToString().Trim() == "") ? "0" : reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folpap = (Convert.ToInt32(folpap) + 1).ToString();
                    txtFolio.Text = folpap.PadLeft(6, '0');
                    btnComedor.Enabled = false;
                    subtipo = "";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }

                if (DDLAlmacen1.SelectedIndex == 8)
                {
                    glo_var_int_registro = 6;
                    lblCierreInventario.Text = "Entradas al Almacén de Sistemas";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "09";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_sis";
                    reader1 = cmnd1.ExecuteReader();
                    string folpap = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        //folpap = reader1.GetValue(0).ToString().Trim();
                        folpap = (reader1.GetValue(0).ToString().Trim() == "") ? "0" : reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folpap = (Convert.ToInt32(folpap) + 1).ToString();
                    txtFolio.Text = folpap.PadLeft(6, '0');
                    btnComedor.Enabled = false;
                    subtipo = "";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }
                if (DDLAlmacen1.SelectedIndex == 9)
                {
                    glo_var_int_registro = 6;
                    lblCierreInventario.Text = "Entradas al Almacén de Guarderia";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "10";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_guar";
                    reader1 = cmnd1.ExecuteReader();
                    string folpap = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        //folpap = reader1.GetValue(0).ToString().Trim();
                        folpap = (reader1.GetValue(0).ToString().Trim() == "") ? "0" : reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folpap = (Convert.ToInt32(folpap) + 1).ToString();
                    txtFolio.Text = folpap.PadLeft(6, '0');
                    btnComedor.Enabled = false;
                    subtipo = "";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }

                if (DDLAlmacen1.SelectedIndex == 10)
                {
                    glo_var_int_registro = 6;
                    lblCierreInventario.Text = "Entradas al Almacén de Guarderia Comedor";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "11";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_guar_com";
                    reader1 = cmnd1.ExecuteReader();
                    string folpap = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        //folpap = reader1.GetValue(0).ToString().Trim();
                        folpap = (reader1.GetValue(0).ToString().Trim() == "") ? "0" : reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folpap = (Convert.ToInt32(folpap) + 1).ToString();
                    txtFolio.Text = folpap.PadLeft(6, '0');
                    btnComedor.Enabled = false;
                    subtipo = "";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }

                if (DDLAlmacen1.SelectedIndex == 11)
                {
                    glo_var_int_registro = 6;
                    lblCierreInventario.Text = "Entradas al Almacén de Eventos";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "12";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_even";
                    reader1 = cmnd1.ExecuteReader();
                    string folpap = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        //folpap = reader1.GetValue(0).ToString().Trim();
                        folpap = (reader1.GetValue(0).ToString().Trim() == "") ? "0" : reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folpap = (Convert.ToInt32(folpap) + 1).ToString();
                    txtFolio.Text = folpap.PadLeft(6, '0');
                    btnComedor.Enabled = false;
                    subtipo = "";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }

                if (DDLAlmacen1.SelectedIndex == 12)
                {
                    glo_var_int_registro = 6;
                    lblCierreInventario.Text = "Entradas al Almacén de Motores";
                    lblCierreInventario.Location = new Point((panel1.Width - lblCierreInventario.Width) / 2, (panel1.Height - lblCierreInventario.Height) / 2);
                    glo_var_chr_alm_clave = "13";

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT MAX(ent_folio) FROM tb_mstr_entradas_mot";
                    reader1 = cmnd1.ExecuteReader();
                    string folpap = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        //folpap = reader1.GetValue(0).ToString().Trim();
                        folpap = (reader1.GetValue(0).ToString().Trim() == "") ? "0" : reader1.GetValue(0).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    folpap = (Convert.ToInt32(folpap) + 1).ToString();
                    txtFolio.Text = folpap.PadLeft(6, '0');
                    btnComedor.Enabled = false;
                    subtipo = "";
                    txt_ent_iva.Text = "0";
                    empaques();
                    txtAcuse.Enabled = true;
                    btnValidar.Enabled = true;
                    //////btnXml.Visible = false;
                }

                DDL_ent_procedencia.SelectedIndex = 0;
                DDL_ent_procedencia.Focus();
                DDL_ent_procedencia_SelectionChangeCommitted(null, null);
                //DDL_ent_procedencia.DroppedDown = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.4", ex.ToString(), "SISEMP");
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + ex.ToString());
            }
        }

        private void DDL_ent_procedencia_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (DDLAlmacen1.SelectedIndex == -1)
            {
                MessageBox.Show("No ha seleccionado el almacen", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DDL_ent_procedencia.SelectedIndex = -1;
                DDL_ent_procedencia.Text = "";
                return;
            }
            //else
            //  {
            //    this.txt_gto_aduanal.Enabled = true;
            //    this.txt_sis_tipo_cambio.Enabled = true;
            //    this.txt_ent_pedimento.Enabled = true;
            //    this.txtFleteImp.Enabled = true;
            //    if (this.DDL_ent_procedencia.SelectedItem.ToString() == "Nacional")
            //    {
            //      this.txt_gto_aduanal.Enabled = false;
            //      this.txt_sis_tipo_cambio.Enabled = false;
            //      this.txt_sis_tipo_cambio.Text = "1";
            //      this.txt_ent_pedimento.Enabled = false;
            //      this.txtFleteImp.Enabled = false;
            //    }
            //    this.txt_prov_clave.Focus();
            //  }
            else
            {
                txt_gto_aduanal.Enabled = true;
                txt_sis_tipo_cambio.Enabled = true;
                txt_ent_pedimento.Enabled = true;
                txtFleteImp.Enabled = true;

                if (this.operacion == "" && DDL_ent_procedencia.SelectedIndex == 1)
                {
                    txt_prov_clave.Focus();
                    DateTime now = DateTime.Now;
                    thisConnection.Open();
                    cmnd1 = this.thisConnection.CreateCommand();
                    SqlCommand cmnd1_1 = cmnd1;
                    string[] strArray1 = new string[7];
                    strArray1[0] = "SELECT valor FROM tb_cat_tipocambio WHERE dia = '";
                    string[] strArray2 = strArray1;
                    int num = now.Day;
                    string str1 = num.ToString();
                    strArray2[1] = str1;
                    strArray1[2] = "' and mes = '";
                    string[] strArray3 = strArray1;
                    num = now.Month;
                    string str2 = num.ToString();
                    strArray3[3] = str2;
                    strArray1[4] = "' AND año = '";
                    string[] strArray4 = strArray1;
                    num = now.Year;
                    string str3 = num.ToString();
                    strArray4[5] = str3;
                    strArray1[6] = "' ORDER BY dia DESC";
                    string str4 = string.Concat(strArray1);
                    cmnd1_1.CommandText = str4;
                    reader1 = cmnd1.ExecuteReader();
                    bool flag = false;
                    if (this.reader1.HasRows)
                    {
                        this.reader1.Read();
                        this.txt_sis_tipo_cambio.Text = this.reader1["valor"].ToString();
                        this.txt_sis_tipo_cambio_KeyPress((object)this, new KeyPressEventArgs('\r'));
                        flag = true;
                    }
                    this.reader1.Close();
                    this.reader1.Dispose();
                    this.cmnd1.Dispose();
                    if (!flag)
                    {
                        DateTime dateTime = now.AddDays(-1.0);
                        cmnd1 = thisConnection.CreateCommand();
                        SqlCommand cmnd1_2 = cmnd1;
                        string[] strArray5 = new string[7];
                        strArray5[0] = "SELECT valor FROM tb_cat_tipocambio WHERE dia = '";
                        string[] strArray6 = strArray5;
                        num = dateTime.Day;
                        string str5 = num.ToString();
                        strArray6[1] = str5;
                        strArray5[2] = "' and mes = '";
                        string[] strArray7 = strArray5;
                        num = dateTime.Month;
                        string str6 = num.ToString();
                        strArray7[3] = str6;
                        strArray5[4] = "' AND año = '";
                        string[] strArray8 = strArray5;
                        num = dateTime.Year;
                        string str7 = num.ToString();
                        strArray8[5] = str7;
                        strArray5[6] = "' ORDER BY dia DESC";
                        string str8 = string.Concat(strArray5);
                        cmnd1_2.CommandText = str8;
                        reader1 = cmnd1.ExecuteReader();
                        if (reader1.HasRows)
                        {
                            reader1.Read();
                            txt_sis_tipo_cambio.Text = reader1["valor"].ToString();
                            txt_sis_tipo_cambio_KeyPress((object)this, new KeyPressEventArgs('\r'));
                        }
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();
                    }
                    thisConnection.Close();
                }
                if (DDL_ent_procedencia.SelectedItem.ToString() == "Nacional")
                {
                    txt_gto_aduanal.Enabled = false;
                    txt_sis_tipo_cambio.Enabled = false;
                    txt_sis_tipo_cambio.Text = "1";
                    txt_ent_pedimento.Enabled = false;
                    txtFleteImp.Enabled = false;
                }

                txt_prov_clave.Focus();
                
            }

            

            ////BUSCAR TIPO DE CAMBIO EN BASE A LA FECHA ACTUAL
            //if (DDL_ent_procedencia.SelectedIndex == 1)
            //{
            //    thisConnection.Open();
            //    cmnd1 = thisConnection.CreateCommand();
            //    cmnd1.CommandText = "SELECT valor, dia FROM tb_cat_tipocambio WHERE dia in ('" + DateTime.Now.AddDays(-1).Day.ToString() + "', '" + DateTime.Now.Day.ToString() + "')" +
            //        " and mes = '" + DateTime.Now.Month.ToString() + "' AND año = '" + DateTime.Now.Year.ToString() + "' ORDER BY dia DESC";
            //    reader1 = cmnd1.ExecuteReader();
            //    if (reader1.HasRows)
            //    {
            //        while (reader1.Read())
            //        {
            //            if (DateTime.Now.Day.ToString() == reader1["dia"].ToString().Trim())
            //            {
            //                txt_sis_tipo_cambio.Text = reader1["valor"].ToString();
            //                //llamar evento keypress
            //                txt_sis_tipo_cambio_KeyPress(this, new KeyPressEventArgs((char)(Keys.Enter)));
            //                break;
            //            }
            //            else
            //            {
            //                if (MessageBox.Show("El tipo de cambio no corresponde al del día actual. ¿Desea cargar el del día anterior? o ingreselo manualmente", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            //                {
            //                    txt_sis_tipo_cambio.Text = reader1["valor"].ToString();
            //                    //llamar evento keypress
            //                    txt_sis_tipo_cambio_KeyPress(this, new KeyPressEventArgs((char)(Keys.Enter)));
            //                    break;
            //                }
            //                else
            //                {

            //                }
            //            }
            //        }
            //    }
            //    reader1.Close();
            //    reader1.Dispose();
            //    thisConnection.Close();
            //}            
        }

        private void txt_prov_clave_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                if (DDLAlmacen1.SelectedIndex == -1)
                {
                    MessageBox.Show("No ha seleccionado el almacen", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txt_prov_clave.Text = "";
                    return;
                }
                if (DDL_ent_procedencia.SelectedIndex == -1)
                {
                    MessageBox.Show("No ha seleccionado si es nacional o importación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txt_prov_clave.Text = "";
                    return;
                }
                for (int i = 0; i < provclave.Count; i++)
                {
                    if (txt_prov_clave.Text == provclave[i].ToString())
                    {
                        DDLProveedor.SelectedIndex = i;
                        //15/06/2018
                        foreach (DataRow re in dtProveedor.Select("prov_clave = '" + txt_prov_clave.Text + "'"))
                        {
                            lblXmlProv.Text = re["prov_xml"].ToString();
                            lblPorcentajeDesc.Text = Convert.ToDecimal(re["prov_descuento"].ToString()).ToString("#.0");
                            lblRFC.Text = re["prov_rfc"].ToString();
                            lblFactNal.Text = re["prov_factnal"].ToString();
                            lblPedimento.Text = re["prov_pedi"].ToString();
                        }
                        //FIN 15/06/2018
                        thisConnection.Open();
                        buscarproveedor(txt_prov_clave.Text);
                        thisConnection.Close();
                        if (provnotfound == "0")
                        {
                            MessageBox.Show("No hay ordenes de compra para ese proveedor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        else
                        {
                            cargardatatables();
                            dtgOrdenes.Rows.Clear();
                            filtroprov.Clear();
                            foreach (DataRow rw in dtordenes.Select("cveprov_oc = '" + txt_prov_clave.Text + "'"))
                            {
                                filtroprov.Rows.Add(rw.ItemArray);
                            }
                            foreach (DataRow rw in filtroprov.Rows)
                            {
                                dtgOrdenes.Rows.Add(rw.ItemArray);
                            }
                            if (txt_prov_clave.Text == "81")
                            {
                                txt_ent_factura.ReadOnly = false;
                                break;
                            }
                        }
                        break;
                    }
                }

                DDL_ent_procedencia.Enabled = false;
                txt_ent_factura.Focus();
            }
        }

        private void DDLProveedor_SelectionChangeCommitted(object sender, EventArgs e)
        {
            txt_prov_clave.Text = provclave[DDLProveedor.SelectedIndex].ToString();
            thisConnection.Open();
            buscarproveedor(txt_prov_clave.Text);
            thisConnection.Close();

            //20/06/2018
            foreach (DataRow re in dtProveedor.Select("prov_clave = '" + txt_prov_clave.Text + "'"))
            {
                lblXmlProv.Text = re["prov_xml"].ToString();
                lblPorcentajeDesc.Text = Convert.ToDecimal(re["prov_descuento"].ToString()).ToString("#.0");
                lblFactNal.Text = re["prov_factnal"].ToString();
                lblPedimento.Text = re["prov_pedi"].ToString();
            }
            //FIN 20/06/2018

            //cargar ordenes de compra por proveedor
            if (provnotfound == "0")
            {
                MessageBox.Show("No hay ordenes de compra para ese proveedor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return; 
            }
            else
            {
                cargardatatables();
                dtgOrdenes.Rows.Clear();
                filtroprov.Clear();
                foreach (DataRow rw in dtordenes.Select("cveprov_oc = '" + txt_prov_clave.Text + "'"))
                {
                    filtroprov.Rows.Add(rw.ItemArray);
                }
                foreach (DataRow rw in filtroprov.Rows)
                {
                    dtgOrdenes.Rows.Add(rw.ItemArray);
                }
                if (txt_prov_clave.Text == "81")
                {
                    txt_ent_factura.ReadOnly = false;
                }
            }
            DDL_ent_procedencia.Enabled = false;
            txt_ent_factura.Focus();
        }

        public void buscarproveedor(string prov)
        {
            cmnd1.CommandText = "SELECT COUNT(cveprov_oc) FROM tb_mstr_ordencompra WHERE cveprov_oc = '" + prov + "' AND status_oc <> 'C'  and cerrada IS NULL AND factura_oc = 'ALMACEN'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                if (Convert.ToInt32(reader1.GetValue(0).ToString()) > 0)
                    provnotfound = "1";
                else
                    provnotfound = "0";
                
            }
            else
                provnotfound = "0";
        }

        public void cargardatatables()
        {
            dtordenes.Clear();
            filtroprov.Clear();
            if (DDL_ent_procedencia.SelectedIndex == 0)
            {
                //Cargar ordenes de almacen
                thisConnection.Open();
                if (this.lblFactNal.Text == "1")
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT numero_oc, cveprov_oc, fecha_oc, tipo_oc FROM tb_mstr_ordencompra WHERE status_oc <> 'C' /*and cerrada IS NULL*/ AND factura_oc = 'ALMACEN' AND aut_oc = '1' ORDER BY numero_oc DESC";
                    reader1 = cmnd1.ExecuteReader();
                    DataRow rw;
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            rw = dtordenes.NewRow();
                            rw["numero_oc"] = reader1.GetValue(0).ToString().Trim();
                            rw["cveprov_oc"] = reader1.GetValue(1).ToString().Trim();
                            rw["nomprov_oc"] = Convert.ToDateTime(reader1.GetValue(2).ToString().Trim()).ToShortDateString();
                            rw["tipo_oc"] = reader1["tipo_oc"].ToString().Trim();
                            dtordenes.Rows.Add(rw);
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                }
                else
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT numero_oc, cveprov_oc, fecha_oc, tipo_oc FROM tb_mstr_ordencompra WHERE tipo_oc = 'P' AND status_oc <> 'C'  /*and cerrada IS NULL*/ AND factura_oc = 'ALMACEN' AND aut_oc = '1' ORDER BY numero_oc DESC";
                    reader1 = cmnd1.ExecuteReader();
                    DataRow rw;
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            rw = dtordenes.NewRow();
                            rw["numero_oc"] = reader1.GetValue(0).ToString().Trim();
                            rw["cveprov_oc"] = reader1.GetValue(1).ToString().Trim();
                            rw["nomprov_oc"] = Convert.ToDateTime(reader1.GetValue(2).ToString().Trim()).ToShortDateString();
                            rw["tipo_oc"] = reader1["tipo_oc"].ToString().Trim();
                            dtordenes.Rows.Add(rw);
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                }
                
                thisConnection.Close();
            }
            if (DDL_ent_procedencia.SelectedIndex == 1)
            {
                //Cargar ordenes de almacen
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT numero_oc, cveprov_oc, fecha_oc, tipo_oc FROM tb_mstr_ordencompra WHERE tipo_oc IN ('D', 'E') AND status_oc <> 'C' /*and cerrada IS NULL*/ AND factura_oc = 'ALMACEN' AND aut_oc = '1' ORDER BY numero_oc DESC";
                reader1 = cmnd1.ExecuteReader();
                DataRow rw;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rw = dtordenes.NewRow();
                        rw["numero_oc"] = reader1.GetValue(0).ToString().Trim();
                        rw["cveprov_oc"] = reader1.GetValue(1).ToString().Trim();
                        rw["nomprov_oc"] = Convert.ToDateTime(reader1.GetValue(2).ToString().Trim()).ToShortDateString();
                        rw["tipo_oc"] = reader1["tipo_oc"].ToString().Trim();
                        dtordenes.Rows.Add(rw);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();
            }
        }

        private void dtgOrdenes_CellDoubleClick(object sender, DataGridViewCellEventArgs e) 
        {
            string oc = dtgOrdenes.CurrentRow.Cells[0].Value.ToString();
            string fecha_oc = dtgOrdenes.CurrentRow.Cells[2].Value.ToString();
            lblToc.Text = dtgOrdenes.CurrentRow.Cells[3].Value.ToString();

            //validacion de orden de compra cerrada
            bool found_cerrada = busca_orden_cerrada(oc);
            if (found_cerrada == true)
            {
                MessageBox.Show("La orden no puede ser surtida porque ya esta cerrada, cualquier duda favor de comunicarse al área de compras, gracias", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (operacion != "consulta" && txt_prov_clave.Text != "81")
            {
                if (lblFactNal.Text == "1")
                {
                    if ((lblToc.Text == "D" || lblToc.Text == "E" || DDL_ent_procedencia.SelectedIndex == 0) && txtFolioFiscal.Text == "")
                    {
                        MessageBox.Show("Debe ingresar primeramente el XML", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else if (DDL_ent_procedencia.SelectedIndex == 1 && txt_prov_clave.Text != "81" && (txtFolioFiscal.Text == "" && lblXmlProv.Text == "0"))
                {
                    MessageBox.Show("Selecione el XML", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btnXml_Click((object)null, (EventArgs)null);
                    if (valida_carga_xml)
                        return;
                    valida_carga_xml = false;
                }
            }
            //if (buscarordenengrid(oc) == true)
            //{
                //if (MessageBox.Show("¿Desea ingresar nuevamente la Orden de Compra?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                //{
                //    foreach (DataRow rr in tmgrdpro.Select("oc = '" + oc + "'"))
                //    {
                //        rr.Delete();
                //    }
                //}
                //else
                //{
                //    MessageBox.Show("La orden de compra ya fué cargada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    return;
                //}
            //}

            bool entro = false;
            if (DDL_ent_procedencia.SelectedIndex == 1)
            {
                if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) == 1)
                {
                    MessageBox.Show("Las ordenes son de importación, debe ingresar el tipo de cambio");
                    return;
                }
            }

            if (Convert.ToDateTime(dtpFecha.Text) < Convert.ToDateTime(fecha_oc))
            {
                MessageBox.Show("La fecha de la entrada no debe ser menor a la fecha de la orden de compra", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txt_oc.Text = "";
                fecha_oc = "";
                return;
            }
            else
            {
                pbxCalendarUlEn.Enabled = false;
            }

            //VERIFICAR QUE LA ORDEN DE COMPRA SEA DE REPARACION 10/02/2025
            //NO VALIDO POR MOTIVO DE QUE PUEDE HABER ORDENES QUE SEAN DE UN PRODUCTO NUEVO Y NO UNA REPARACION
            //DataTable dt = _consultaController.VerificarOcReparacion(oc);
            //if (glo_var_chr_alm_clave == "13")
            //{
            //    if (dt.Rows[0]["urgente_oc"].ToString() != "0")
            //        lblReparacion.Text = dt.Rows[0]["urgente_oc"].ToString();
            //    if (lblReparacion.Text == "0")
            //    {
            //        MessageBox.Show("", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //        lblReparacion.Text = "-";
            //        return;
            //    }
            //}

            if (MessageBox.Show("¿Desea agregar esta orden de compra a la entrada?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                dtgDetalle2.Rows.Clear();
                //cargar datos de la oc
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = "SELECT A.subtotal_oc, A.status_oc, A.tipo_oc, A.req_oc, B.cantidad_oc, B.unidad_oc, B.prod_clave, B.nomprod_oc, B.orden_oc, B.cargo_oc," +
                //    " B.precio_oc, B.importe_oc, B.ieps_oc, B.surtido_oc, B.iva_oc FROM tb_mstr_ordencompra A, tb_det_ordenescompra B WHERE A.numero_oc = B.numero_oc AND A.numero_oc = '" + oc + "' AND B.prod_clave <> '' ORDER BY conse";
                cmnd1.CommandText = "SELECT A.subtotal_oc, A.status_oc, A.tipo_oc, A.req_oc, B.cantidad_oc, B.unidad_oc, B.prod_clave, B.nomprod_oc, B.orden_oc, B.cargo_oc, B.precio_oc, B.importe_oc, B.ieps_oc, " +
                    "(SELECT isnull(SUM(C.ent_cantidad), 0) FROM tb_det_entrada_oc C WHERE C.numero_oc =  A.numero_oc AND C.emp_clave = B.prod_clave AND C.estatus <> 'C') AS surtido_oc, B.iva_oc, C.urgente_oc " +
                    "FROM tb_mstr_ordencompra A INNER JOIN tb_det_ordenescompra B ON A.numero_oc = B.numero_oc " +
                    "INNER JOIN tb_oc_cfdi C ON A.numero_oc = C.numero_oc " +
                    "WHERE A.numero_oc = '" + oc + "' AND B.prod_clave <> '' ORDER BY conse";
                reader1 = cmnd1.ExecuteReader();
                DataRow rw;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        //if (DDL_ent_procedencia.SelectedIndex == 0)
                        //{
                        //    if (reader1.GetValue(2).ToString().Trim() != "P")
                        //    {
                        //        MessageBox.Show("La orden de compra es en dólares o euros, no se puede cargar por que se esta realizando una entrada de NACIONAL");
                        //        return;
                        //    }
                        //}
                        //if (DDL_ent_procedencia.SelectedIndex == 1)
                        //{
                        //    if (reader1.GetValue(2).ToString().Trim() == "P")
                        //    {
                        //        MessageBox.Show("La orden de compra es pesos, no se puede cargar por que se esta realizando una entrada de IMPORTACIÓN");
                        //        return;
                        //    }
                        //}

                        decimal x = Convert.ToDecimal(reader1.GetValue(4).ToString().Trim());
                        decimal y = Convert.ToDecimal(reader1.GetValue(13).ToString().Trim());
                        
                        
                        if (Convert.ToDecimal(reader1.GetValue(4).ToString().Trim()) <= Convert.ToDecimal(reader1.GetValue(13).ToString().Trim()))
                        {
                            continue;
                            //MessageBox.Show("La orden de compra no se puede cargar porque ya fué surtida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //thisConnection.Close();
                            //return;
                        }

                        if (Convert.ToDecimal(reader1.GetValue(13).ToString().Trim()) < 0)
                        {
                            continue;
                            //MessageBox.Show("La orden de compra no se puede cargar porque ya fué surtida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //thisConnection.Close();
                            //return;
                        }

                        //VALIDAR EMPAQUES YA CARGADOS
                        bool enc = false;
                        if (dtdetalles.Rows.Count > 0)
                        {
                            foreach (DataRow rx in dtdetalles.Select("emp_clave = '" + reader1.GetValue(6).ToString().Trim() + "' AND orden_compra = '" + oc + "'"))
                            {
                                enc = true;       
                            }

                            if (enc == true)
                            {
                                enc = false;
                                continue;
                            }
                        }

                        rw = dtdetalles.NewRow();
                        
                        //Se evalua que el empaque sea del almacen seleccionado
                        if (product_almacen(reader1.GetValue(6).ToString().Trim()) == true)
                        {
                            rw["emp_clave"] = reader1.GetValue(6).ToString().Trim();
                            //rw["nombre"] = reader1.GetValue(7).ToString().Trim();
                            decimal surt = 0;
                            surt = Convert.ToDecimal(reader1.GetValue(4).ToString().Trim()) - Convert.ToDecimal(reader1.GetValue(13).ToString().Trim());
                            if (surt > 0)
                            {
                                rw["cantidad"] = surt;
                                //if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) > 1)
                                //{
                                //    rw["unitario"] = (Convert.ToDecimal(reader1.GetValue(10).ToString().Trim()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                                //}
                                //else
                                //    rw["unitario"] = Convert.ToDecimal(reader1.GetValue(10).ToString().Trim());
                                rw["importe"] = (surt * Convert.ToDecimal(reader1.GetValue(10).ToString().Trim()));
                                rw["importe_org"] = rw["importe"];
                                rw["costo_org"] = rw["costo"];
                            }
                                
                            else
                            {
                                rw["cantidad"] = reader1.GetValue(4).ToString().Trim();
                                //if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) > 1)
                                //{
                                //    rw["unitario"] = (Convert.ToDecimal(reader1.GetValue(10).ToString().Trim()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                                //}
                                //else
                                //    rw["unitario"] = reader1.GetValue(10).ToString().Trim();
                                rw["importe"] = reader1.GetValue(11).ToString().Trim();
                                rw["importe_org"] = rw["importe"];
                                rw["costo_org"] = rw["costo"];
                            }
                            //rw["adv"] = "0";
                            //rw["dta"] = "0";
                            //rw["lote"] = "";
                            //rw["iepsimporte"] = reader1.GetValue(12).ToString().Trim();
                            //if (Convert.ToDecimal(reader1.GetValue(12).ToString().Trim()) > 0)
                            //{
                            //    rw["ieps"] = true;
                            //}
                            //else
                            //{
                            rw["ieps"] = reader1.GetValue(12).ToString().Trim();
                            //}
                            //rw["gastoaduanal"] = "0";
                            //rw["flete"] = "0";
                            rw["orden_compra"] = oc;
                            rw["requisicion"] = reader1.GetValue(3).ToString().Trim();
                            rw["costo"] = reader1.GetValue(10).ToString().Trim();
                            if (operacion == "consulta")
                            {
                                //rw["oper"] = "nueva";
                            }
                            rw["importe_org"] = rw["importe"];
                            rw["costo_org"] = rw["costo"];

                            if (Convert.ToDecimal(reader1["iva_oc"]) > 0)
                            {
                                decimal iva_t = ((surt * Convert.ToDecimal(reader1.GetValue(10).ToString().Trim())) + Convert.ToDecimal(reader1.GetValue(12).ToString().Trim())) * (Convert.ToDecimal(16) / 100); 
                                rw["iva"] = iva_t.ToString("###,###,##0.000");
                                rw["tiva"] = (Convert.ToDecimal(reader1["iva_oc"]) > 0) ? "1" : "0";
                            }
                            else
                            {
                                rw["iva"] = reader1["iva_oc"];
                                rw["tiva"] = (Convert.ToDecimal(reader1["iva_oc"]) > 0) ? "1" : "0";
                            }

                            rw["importe_orig"] = rw["importe"];
                            rw["urgente"] = reader1["urgente_oc"];
                            
                            if (surt > 0)
                                rw["cant_orig"] = surt;
                            else
                                rw["cant_orig"] = reader1.GetValue(4).ToString().Trim();
                            
                            //COMENTADO 02/07/2020
                            //////if (surt > 0)
                            //////    rw["iva"] = (surt * Convert.ToDecimal(reader1.GetValue(10).ToString().Trim()) * Convert.ToDecimal("0.16")).ToString();
                            //////else
                            //////    rw["iva"] = reader1["iva_oc"].ToString().Trim();
                            entro = true;
                            dtdetalles.Rows.Add(rw);
                        }
                        else
                        {
                            MessageBox.Show("Los productos de esa orden de compra no corresponden a el almacen seleccionado", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    
                }

                if (dtdetalles.Rows.Count == 0)
                {
                    MessageBox.Show("La Orden de Compra ya fué surtida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (thisConnection.State == ConnectionState.Open)
                        thisConnection.Close();
                    return;
                }

                
                thisConnection.Close();

                


                if (lblTCFact.Text != "1")
                {
                    foreach (DataRow row in dtdetalles.Rows)
                    {
                        if (row["borrar"].ToString() != "A")
                        {
                            DataRow dataRow1 = row;
                            Decimal num2 = Convert.ToDecimal(row["importe"].ToString()) * Convert.ToDecimal(lblTCFact.Text);
                            string str3 = num2.ToString();
                            dataRow1["importe"] = str3;
                            DataRow dataRow2 = row;
                            num2 = Convert.ToDecimal(row["costo"].ToString()) * Convert.ToDecimal(lblTCFact.Text);
                            string str4 = num2.ToString();
                            dataRow2["costo"] = str4;
                            row["borrar"] = "A";
                            DataRow dataRow3 = row;
                            num2 = Convert.ToDecimal(row["importe_org"].ToString()) * Convert.ToDecimal(lblTCFact.Text);
                            string str5 = num2.ToString();
                            dataRow3["importe_org"] = str5;
                            DataRow dataRow4 = row;
                            num2 = Convert.ToDecimal(row["costo_org"].ToString()) * Convert.ToDecimal(lblTCFact.Text);
                            string str6 = num2.ToString();
                            dataRow4["costo_org"] = str6;
                        }
                    }
                }

                //Validamos que los productos de la orden de compra no hayan sido ya surtidos
                //validarempaquessurtidos();
                button2_Click(null, null);
                //Cargar datos al datagrid
                foreach (DataRow rr in dtdetalles.Rows)
                {
                    //dtgDetalle.Rows.Add(rr.ItemArray);
                    //dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"), 
                    //    Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"), 
                    //    Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                    //    Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                    //    Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                    //    rr["borrar"].ToString(), rr["oper"].ToString());
                    dtgDetalle2.Rows.Add(rr["emp_clave"].ToString(), rr["orden_compra"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["importe"].ToString()).ToString("###,###,##0.000"),
                        rr["requisicion"].ToString(), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.00000"), rr["borrar"].ToString(), rr["operacion"].ToString(), "",
                        Convert.ToDecimal(rr["ieps"].ToString()).ToString("###,###,##0.000"), "", rr["lote"].ToString(), rr["iva"].ToString(), rr["tiva"].ToString(), Convert.ToDecimal(rr["importe"].ToString()).ToString("###,###,##0.000"), rr["urgente"].ToString(),
                        Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"), "0");
                }

                txt_oc.Text = "";
                ////Cargar datos del mstr_oc

                ////Cargar datos del detalle_oc
                //DataRow dr;
                //thisConnection.Open();
                //cmnd2 = thisConnection.CreateCommand();
                //cmnd2.CommandText = "SELECT cantidad_oc, unidad_oc, prod_clave, nomprod_oc, orden_oc, cargo_oc, precio_oc, importe_oc, ieps_oc FROM tb_det_ordenescompra WHERE numero_oc = '" + oc + "'";
                //reader2 = cmnd2.ExecuteReader();
                //if (reader2.HasRows)
                //{
                //    while (reader2.Read())
                //    {
                //        if (product_almacen(reader2.GetValue(2).ToString().Trim()) == true)
                //        {
                //            dr = tmgrdpro.NewRow();
                //            dr["clave"] = reader2.GetValue(2).ToString().Trim();
                //            dr["nombre"] = reader2.GetValue(3).ToString().Trim();
                //            dr["cantidad"] = reader2.GetValue(0).ToString().Trim();
                //            //if (reader1.GetValue(113).ToString().Trim() == "D")
                //            //    dr["costo_usd"] = reader2.GetValue(7).ToString().Trim();
                //            //dr["unitario"] = reader2.GetValue(6).ToString().Trim();
                //            //if (reader1.GetValue(113).ToString().Trim() == "P")
                //            //    dr["costo_mn"] = reader2.GetValue(7).ToString().Trim(); ;
                //            dr["adv"] = "0.00";
                //            dr["dta"] = "0.00";
                //            dr["lote"] = "";
                //            //dr["ieps"] = (reader2.GetValue(8).ToString().Trim() != "") ? true : false;
                //            dr["ieps"] = (Convert.ToDecimal(reader2.GetValue(8).ToString()) > 0) ? true : false;
                //            dr["iepsimporte"] = (reader2.GetValue(8).ToString().Trim() != "") ? reader2.GetDecimal(8) : Convert.ToDecimal("0.00");
                //            dr["gastoaduanal"] = "0.00";
                //            dr["flete"] = "0.00";
                //            tmgrdpro.Rows.Add(dr);
                //            //tieps = tieps + reader2.GetDecimal(8);
                //        }
                //    }
                //}
                //reader2.Close();
                //reader2.Dispose();
                //cmnd2.Dispose();
                ////}
                //thisConnection.Close();
                //dtgDetalle.Rows.Clear();
                //foreach (DataRow drw in tmgrdpro.Rows)
                //{
                //    dtgDetalle.Rows.Add(drw.ItemArray);
                //}
                if (entro == true)
                {
                    bool fnd = false;
                    foreach (var item in lstOC.Items)
                    {
                        if (item.ToString() == oc)
                        {
                            fnd = true;
                            break;
                        }
                    }
                    if (fnd == false)
                    {
                        lstOC.Items.Add(oc);
                    }
                }
                entro = false;

                calculatotales();
                recalculo_isr_descuento();

                
            }
        }

        //validar que el empaque pertenezca al almacen seleccionado
        private bool product_almacen(string num)
        {
            cmnd3 = thisConnection.CreateCommand();
            cmnd3.CommandText = "SELECT alm_clave FROM tb_cat_empaques WHERE emp_clave = '" + num + "'";
            reader3 = cmnd3.ExecuteReader();
            bool fnd = true;
            if (reader3.HasRows)
            {
                reader3.Read();
                if (glo_var_chr_alm_clave != reader3.GetValue(0).ToString().Trim())
                    fnd = false;
            }
            reader3.Close();
            reader3.Dispose();
            cmnd3.Dispose();
            return fnd;

        }

        //validar surtido de empaque
        private void validarempaquessurtidos()
        { 
            
            //traer empaques de la orden de compra
            //thisConnection.Open();
            //DataRow r1;
            //cmnd1 = thisConnection.CreateCommand();
            //cmnd1.CommandText = "SELECT numero_oc, prod_clave, cantidad_oc, surtido_oc FROM tb_det_ordenescompra WHERE numero_oc = '" + numoc + "'";
            //reader1 = cmnd1.ExecuteReader();
            //if (reader1.HasRows)
            //{
            //    while (reader1.Read())
            //    {
            //        r1 = dtempoc.NewRow();
            //        r1["numero_oc"] = reader1.GetValue(0).ToString().Trim();
            //        r1["prod_clave"] = reader1.GetValue(1).ToString().Trim();
            //        r1["cantidad_oc"] = reader1.GetValue(2).ToString().Trim();
            //        r1["surtido_oc"] = reader1.GetValue(3).ToString().Trim();
            //        dtempoc.Rows.Add(r1);
            //    }
            //}
            //reader1.Close();
            //reader1.Dispose();
            //cmnd1.Dispose();

            foreach (DataRow r in dtempoc.Rows)
            {
                if (Convert.ToDecimal(r["cantidad"].ToString()) == Convert.ToDecimal(r["surtido"].ToString()))
                {
                    r.Delete();
                }
            }
            //thisConnection.Close();
        }

        private void validarordenrequisicion()
        {
            thisConnection.Open();
            //Validacion de que el empaque con su requisicion y orden de compra no hayan sido cargadoss en otra entrada
            foreach (DataRow rw in tmgrdpro.Rows)
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT req_folio, emp_clave, numero_oc, req_xpedir, req_surtido FROM tb_det_requisiciones WHERE req_folio = '" + rw["req"] + "' AND emp_clave = '" + rw["clave"].ToString() + "' AND numero_oc = '" + rw["oc"] + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        if (Convert.ToDecimal(reader1.GetValue(3).ToString().Trim()) == Convert.ToDecimal(reader1.GetValue(4).ToString().Trim()))
                        {
                            //Si xpedir y surtido son iguales borrar el renglon de la tabla
                            rw.Delete();
                        }
                        else
                        {
                            //Si xpedir es mayor a surtido sacar la diferencia y modificar la cantidad a pedir
                            rw["cantidad"] = (Convert.ToDecimal(reader1.GetValue(3).ToString().Trim()) - Convert.ToDecimal(reader1.GetValue(4).ToString().Trim())).ToString();
                        }
                    }
                }
            }
            thisConnection.Close();
        }

        private bool buscarordenengrid(string oc)
        {
            bool bandera = false;
            foreach (DataRow rw in tmgrdpro.Select("oc = '" + oc + "'"))
            {
                bandera = true;
            }
            return bandera;
        }

        private void quitarempaque(string oc, string emp)
        {
            //foreach (DataRow rw in tmgrdpro.Select("oc = '" + oc + "' AND clave = '" + emp + "'"))
            foreach (DataRow rw in tmgrdpro.Select("clave = '" + emp + "'"))
            {
                //rw["borrar"] = "1";
                rw.Delete();
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            ////if (operacion != "consulta")
            if (lblPathXml.Text == "-")//txtAcuse.Text != "" && 
            {
                string importacion = "0";
                thisConnection.Open();
                SqlCommand cmndimportacion = thisConnection.CreateCommand();
                cmndimportacion.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
                try
                {
                    importacion = Convert.ToString(cmndimportacion.ExecuteScalar().ToString().Trim());
                }
                catch {
                    importacion = "";
                }

                cmndimportacion.Dispose();
                thisConnection.Close();

                if (importacion == "1") { 
                
                }

                /*if (txt_prov_clave.Text == "81" || txt_prov_clave.Text == "404" || txt_prov_clave.Text == "11" || txt_prov_clave.Text == "8851" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694" || txt_prov_clave.Text.Trim() == "9957" || txt_prov_clave.Text.Trim() == "833" || txt_prov_clave.Text.Trim() == "1550" || txt_prov_clave.Text.Trim() == "1761" || txt_prov_clave.Text.Trim() == "935" || txt_prov_clave.Text.Trim() == "6854" || txt_prov_clave.Text == "354" || txt_prov_clave.Text == "1250" || txt_prov_clave.Text == "1536" || txt_prov_clave.Text == "9758" || txt_prov_clave.Text == "1799" || txt_prov_clave.Text == "1244" || txt_prov_clave.Text == "191" || txt_prov_clave.Text == "812" || txt_prov_clave.Text == "8014" || txt_prov_clave.Text == "305" || txt_prov_clave.Text == "106")
                {
                }*/
                else
                {
                    MessageBox.Show("Debe seleccionar el XML de la factura, recuerde tener el PDF en la misma carpeta del XML, por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
            }

            string oc_val = dtgDetalle2.Rows[0].Cells["ordencompra"].Value.ToString().Trim();
            string pr_val = "";
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT RTRIM(cveprov_oc) AS cveprov_oc FROM tb_mstr_ordencompra WHERE numero_oc = '" + oc_val + "'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                pr_val = reader1["cveprov_oc"].ToString();
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();

            if (pr_val != txt_prov_clave.Text.Trim())
            {
                MessageBox.Show("El proveedor fue cambiado manualmente y es diferente al de las ordenes de compra, favor de verificarlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            if (lblPedimento.Text == "1")
            {
                if (lsbPediFact.Items.Count < 2)
                {
                    MessageBox.Show("No se han cargado los dos archivos de pedimento y factura obligatorios para este proveedor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            

            btnGuardar.Enabled = false;

            

            guardar();

            //

            

            btnGuardar.Enabled = true;


        }

        private void dtgDetalle_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.ColumnIndex == 1)
            //{
            //    if (dtgDetalle.Rows.Count == 0)
            //    {
            //        return;
            //    }
            //    if (MessageBox.Show("¿Desea quitar el empaque seleccionado?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //    {
            //        //VALIDA PARA NO BORRAR EMPAQUES DE LA ORDEN DE COMPRA
            //        string c = dtgDetalle.CurrentRow.Cells[0].Value.ToString();
            //        foreach (DataRow a in dtdetalles.Select("emp_clave = '" + c + "'"))
            //        {
            //            MessageBox.Show("El empaque no pude ser borrado porque se introdujo por medio de una orden de compra", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //            return;
            //        }
            //        //FIN VALIDA PARA NO BORRAR EMPAQUES DE LA ORDEN DE COMPRA

            //        quitarempaque(dtgDetalle.CurrentRow.Cells[12].Value.ToString(), dtgDetalle.CurrentRow.Cells[0].Value.ToString());

            //        //Cargar datos al datagrid
            //        dtgDetalle.Rows.Clear();
            //        foreach (DataRow rr in tmgrdpro.Rows)
            //        {
            //            if (rr["borrar"].ToString() != "1")
            //            {
            //                dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
            //                    Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
            //                    Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
            //                    Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
            //                    Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
            //                    rr["borrar"].ToString(), rr["oper"].ToString());
            //            }
            //                //dtgDetalle.Rows.Add(rr.ItemArray);
            //        }
            //        if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
            //        {
            //            txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
            //        }
            //        if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
            //        {
            //            txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
            //        }
            //        calculatotales();
            //    }
            //}
        }

        private void calculatotales()
        {
            decimal subtotal1 = 0;
            decimal totaladv = 0;
            decimal totaldta = 0;
            decimal totalieps = 0;
            decimal subtotal2 = 0;
            decimal iva2 = 0;
            decimal total = 0;
            decimal tip_cambi = 0;
            decimal gtoadu = 0;

            decimal total_iva = 0;

            decimal gtos_importacion = 0;

            tip_cambi = Convert.ToDecimal(txt_sis_tipo_cambio.Text);
            gtoadu = Convert.ToDecimal(txt_gto_aduanal.Text);
            foreach (DataRow rw in tmgrdpro.Rows)
            {
                if (rw["borrar"].ToString() != "1")
                {
                    subtotal1 = subtotal1 + Convert.ToDecimal(rw[4].ToString());
                    totaladv = totaladv + Convert.ToDecimal(rw[5].ToString());
                    totaldta = totaldta + Convert.ToDecimal(rw[6].ToString());
                    totalieps = totalieps + Convert.ToDecimal(rw[9].ToString());
                    total_iva = total_iva + Convert.ToDecimal(rw[19].ToString());
                }
            }
            if (tip_cambi > 0)
            {
                subtotal1 = subtotal1 * tip_cambi;

            }
            //if (gtoadu > 0)
            //{
            //    gtoadu = gtoadu * tip_cambi;
            //}

            gtos_importacion = gtoadu + (Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text));

            subtotal1 = subtotal1 + gtos_importacion;

            subtotal2 = subtotal1 + totaladv + totaldta + totalieps;

            //////subtotal2 = subtotal1 + totaladv + totaldta + totalieps + gtoadu + Convert.ToDecimal(txtFlete.Text);
            //////if (tip_cambi > 0)
            //////{
            //////    subtotal1 = subtotal1 * tip_cambi;
            //////    subtotal2 = subtotal2 * tip_cambi;
            //////    foreach (DataRow rw in tmgrdpro.Rows)
            //////    {
            //////        rw["unitario"] = (Convert.ToDecimal(rw["unitario"].ToString()) * tip_cambi).ToString();
            //////    }
            //////}

            //if (Convert.ToDecimal(txt_ent_iva.Text) > 0)
            //{
            //    iva2 = subtotal2 * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
            //    total_iva = iva2;
            //}
            ////if (Convert.ToDecimal(txt_ent_iva.Text) > 0)
            ////{
            ////    iva2 = subtotal2 * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
            ////}
            ////total = subtotal2 + iva2;

            total = subtotal2 + total_iva - Convert.ToDecimal(txtDesc.Text);

            txt_subtotal1.Text = subtotal1.ToString("###,###,##0.00");  //subtotal1.ToString("###,###,##0.00");
            txt_adv.Text = totaladv.ToString("###,###,##0.00");
            txt_dta.Text = totaldta.ToString("###,###,##0.00");
            txt_ent_ieps.Text = totalieps.ToString("###,###,##0.00");
            txt_subtotal2.Text = subtotal2.ToString("###,###,##0.00");
            //txt_iva.Text = iva2.ToString("###,###,##0.00");
            txt_iva.Text = total_iva.ToString("###,###,##0.00");
            txt_total.Text = total.ToString("###,###,##0.00");
        }

        private void dtgDetalle_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //Calculo cuando se modifica la cantidad de producto
            if (e.ColumnIndex == 2)
            {
                if (validarnumero(dtgDetalle.CurrentRow.Cells[2].Value.ToString()) == false)
                {
                    MessageBox.Show("El valor ingresado no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                //VALIDA PARA NO MODIFICAR EMPAQUES DE LA ORDEN DE COMPRA
                string c = dtgDetalle.CurrentRow.Cells[0].Value.ToString();
                foreach (DataRow r in dtdetalles.Select("emp_clave = '" + c + "'"))
                {
                    MessageBox.Show("El empaque no pude ser modificado porque se introdujo por medio de una orden de compra", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    string x = r["cantidad"].ToString();
                    for (int y = 0; y < dtgDetalle.Rows.Count - 1; y++)
                    {
                        if (dtgDetalle.Rows[y].Cells["clave"].Value.ToString() == r["emp_clave"].ToString())
                            dtgDetalle.Rows[y].Cells["cantidad"].Value = Convert.ToDecimal(x).ToString("###,###,##0.000");
                    }
                    return;
                }
                //FIN VALIDA PARA NO MODIFICAR EMPAQUES DE LA ORDEN DE COMPRA

                //////dividir importe entre cantidad
                ////decimal a = 0;
                ////decimal b = 0;
                ////a = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value) / Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[2].Value);//costo unitario
                ////b = a;
                //////dtgDetalle.CurrentRow.Cells[4].Value = (Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[2].Value.ToString()) * Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[3].Value.ToString())).ToString();
                ////if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) > 0)
                ////    b = a * (Convert.ToDecimal(txt_sis_tipo_cambio.Text));

                ////tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][2] = dtgDetalle.CurrentRow.Cells[2].Value.ToString();
                ////tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][4] = dtgDetalle.CurrentRow.Cells[4].Value.ToString();
                ////tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][3] = b.ToString();

                ////dtgDetalle.CurrentRow.Cells[2].Value = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");
                ////dtgDetalle.CurrentRow.Cells[4].Value = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()).ToString("###,###,##0.000");
                ////dtgDetalle.CurrentRow.Cells[3].Value = b.ToString("###,###,##0.000");

                //dividir importe entre cantidad
                decimal a = 0;
                decimal b = 0;
                a = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[2].Value) * Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[3].Value);//costo unitario
                b = a;
                //dtgDetalle.CurrentRow.Cells[4].Value = (Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[2].Value.ToString()) * Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[3].Value.ToString())).ToString();
                if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) > 1)
                    b = a / (Convert.ToDecimal(txt_sis_tipo_cambio.Text));

                //foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "' and oc = '" + dtgDetalle.CurrentRow.Cells[12].Value + "'"))
                foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "'"))
                {
                    rr["cantidad"] = dtgDetalle.CurrentRow.Cells[2].Value.ToString();
                    rr["costo"] = b.ToString();
                }
                
                //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][3] = b.ToString();

                dtgDetalle.CurrentRow.Cells[2].Value = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");//cantidad
                dtgDetalle.CurrentRow.Cells[4].Value = b.ToString("###,###,##0.000");//costo
                //dtgDetalle.CurrentRow.Cells[3].Value = b.ToString("###,###,##0.000");
                
            }

            //Calculo cuando se modifica el costo del producto
            if (e.ColumnIndex == 3)
            {
                //if (validarnumero(dtgDetalle.CurrentRow.Cells[3].Value.ToString()) == false)
                //{
                //    MessageBox.Show("El valor ingresado no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    return;
                //}
                //dtgDetalle.CurrentRow.Cells[4].Value = (Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[2].Value.ToString()) * Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[3].Value.ToString())).ToString();
                //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][3] = dtgDetalle.CurrentRow.Cells[3].Value.ToString();
                //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][4] = dtgDetalle.CurrentRow.Cells[4].Value.ToString();
            }

            //Calculo cuando se modifica el importe del producto
            if (e.ColumnIndex == 4)
            {
                if (validarnumero(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) == false)
                {
                    MessageBox.Show("El valor ingresado no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                //VALIDA PARA NO MODIFICAR EMPAQUES DE LA ORDEN DE COMPRA
                //string c = dtgDetalle.CurrentRow.Cells[0].Value.ToString();
                //foreach (DataRow r in dtdetalles.Select("emp_clave = '" + c + "'"))
                //{
                //    MessageBox.Show("El empaque no pude ser modificado porque se introdujo por medio de una orden de compra", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    string x = r["importe"].ToString();
                //    for (int y = 0; y < dtgDetalle.Rows.Count - 1; y++)
                //    {
                //        if (dtgDetalle.Rows[y].Cells["clave"].Value.ToString() == r["emp_clave"].ToString())
                //            dtgDetalle.Rows[y].Cells["costo"].Value = Convert.ToDecimal(x).ToString("###,###,##0.000");
                //    }
                //    return;
                //}
                //FIN VALIDA PARA NO MODIFICAR EMPAQUES DE LA ORDEN DE COMPRA

                decimal a = 0; decimal b = 0;
                a = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value) / Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[2].Value);//costo unitario
                b = a;
                if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) > 1)
                    b = a * (Convert.ToDecimal(txt_sis_tipo_cambio.Text));

                //VERIFICAR SI EXISTE CARGO POR IEPS
                if ((bool)dtgDetalle.CurrentRow.Cells[8].EditedFormattedValue == true)
                {
                    if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) > 1)
                    {
                        //foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "' and oc = '" + dtgDetalle.CurrentRow.Cells[12].Value + "'"))
                        foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "'"))
                        {
                            rr["iepsimporte"] = ((Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) * iepsporc) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                        }
                        //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][9] = ((Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) * iepsporc) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                        dtgDetalle.CurrentRow.Cells[9].Value = ((Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) * iepsporc) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString("###,###,##0.000");//iepsimporte
                    }
                    else
                    {
                        //foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "' and oc = '" + dtgDetalle.CurrentRow.Cells[12].Value + "'"))
                        foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "'"))
                        {
                            rr["iepsimporte"] = (Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) * iepsporc).ToString();
                        }
                        //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][9] = (Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) * iepsporc).ToString();
                        dtgDetalle.CurrentRow.Cells[9].Value = (Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) * iepsporc).ToString("###,###,##0.000");//iepsimporte
                    }
                }
                else
                {
                    if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) > 1)
                    {
                        //foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "' and oc = '" + dtgDetalle.CurrentRow.Cells[12].Value + "'"))
                        foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "'"))
                        {
                            rr["iepsimporte"] = "0.000";
                        }
                        //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][9] = ((Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) * iepsporc) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                        dtgDetalle.CurrentRow.Cells[9].Value = ((Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) * iepsporc) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString("###,###,##0.000");//iepsimporte
                    }
                    else
                    {
                        //foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "' and oc = '" + dtgDetalle.CurrentRow.Cells[12].Value + "'"))
                        foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "'"))
                        {
                            rr["iepsimporte"] = "0.000";
                        }
                        //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][9] = (Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) * iepsporc).ToString();
                        dtgDetalle.CurrentRow.Cells[9].Value = (Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) * iepsporc).ToString("###,###,##0.000");//iepsimporte
                    }
                }

                //foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "' and oc = '" + dtgDetalle.CurrentRow.Cells[12].Value + "'"))
                foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "'"))
                {
                    rr["unitario"] = b.ToString();
                    rr["costo"] = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()).ToString("0.000");
                }
                //dtgDetalle.CurrentRow.Cells[3].Value = (Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()) / Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[2].Value.ToString())).ToString();
                ////tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][3] = b.ToString();//
                ////tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][4] = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value.ToString()).ToString("0.000");
                dtgDetalle.CurrentRow.Cells[3].Value = b.ToString("###,###,##0.000");
                dtgDetalle.CurrentRow.Cells[4].Value = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[4].Value).ToString("###,###,##0.000");

                //calculatotales();
            }

            //Calculo cuando se modifica el ADV
            if (e.ColumnIndex == 5)
            {
                if (validarnumero(dtgDetalle.CurrentRow.Cells[5].Value.ToString()) == false)
                {
                    MessageBox.Show("El valor ingresado no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "' and oc = '" + dtgDetalle.CurrentRow.Cells[12].Value + "'"))
                foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "'"))
                {
                    rr["adv"] = dtgDetalle.CurrentRow.Cells[5].Value.ToString();
                }
                dtgDetalle.CurrentRow.Cells[5].Value = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[5].Value.ToString()).ToString("###,###,##0.000");
                //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][5] = dtgDetalle.CurrentRow.Cells[5].Value.ToString();
            }

            //Calculo cuando se modifica el DTA
            if (e.ColumnIndex == 6)
            {
                if (validarnumero(dtgDetalle.CurrentRow.Cells[6].Value.ToString()) == false)
                {
                    MessageBox.Show("El valor ingresado no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "' and oc = '" + dtgDetalle.CurrentRow.Cells[12].Value + "'"))
                foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "'"))
                {
                    rr["dta"] = dtgDetalle.CurrentRow.Cells[6].Value.ToString();
                }
                dtgDetalle.CurrentRow.Cells[6].Value = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[6].Value.ToString()).ToString("###,###,##0.000");
                //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][6] = dtgDetalle.CurrentRow.Cells[6].Value.ToString();
            }

            //Calculo cuando se modifica el lote
            if (e.ColumnIndex == 7)
            {
                //foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "' and oc = '" + dtgDetalle.CurrentRow.Cells[12].Value + "'"))
                foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "'"))
                {
                    rr["lote"] = dtgDetalle.CurrentRow.Cells[7].Value.ToString();
                }
                //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][7] = dtgDetalle.CurrentRow.Cells[7].Value.ToString();
            }

            if (e.ColumnIndex == 8)//Realizar calculo de IEPS
            {
                //if (validarnumero(dtgDetalle.CurrentRow.Cells[6].Value.ToString()) == false)
                //{
                //    MessageBox.Show("El valor ingresado no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    return;
                //}
                //foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "'"))
                //{
                //    rr["dta"] = dtgDetalle.CurrentRow.Cells[6].Value.ToString();
                //}
                //dtgDetalle.CurrentRow.Cells[6].Value = Convert.ToDecimal(dtgDetalle.CurrentRow.Cells[6].Value.ToString()).ToString("###,###,##0.000");
            }

            ////Calculo cuando se modifica la factua
            //if (e.ColumnIndex == 15)
            //{
            //    foreach (DataRow rr in tmgrdpro.Select("clave = '" + dtgDetalle.CurrentRow.Cells[0].Value + "' and oc = '" + dtgDetalle.CurrentRow.Cells[12].Value + "'"))
            //    {
            //        rr["lote"] = dtgDetalle.CurrentRow.Cells[7].Value.ToString();
            //    }
            //    //tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][15] = dtgDetalle.CurrentRow.Cells[15].Value.ToString();
            //}

            ////Calculo cuando se modifica el lote
            //if (e.ColumnIndex == 16)
            //{
            //    tmgrdpro.Rows[dtgDetalle.CurrentRow.Index][16] = dtgDetalle.CurrentRow.Cells[16].Value.ToString();
            //}
            if (e.ColumnIndex == 0)//Carga de datos desde el DataGridView
            {
                if (dtgDetalle.CurrentRow.Cells[0].Value == null)
                {
                    return;
                }
                if (txtFlete.Text == "")
                {
                    MessageBox.Show("El campo de flete esta vacio, debe ingresar un valor(ej. 0 o el costo del flete en moneda nacional)", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                cargaempaque(dtgDetalle.CurrentRow.Cells[0].Value.ToString());
            }
            calculatotales();
        }

        public bool validarnumero(string valor)
        {
            try
            {
                decimal val = 0;
                val = Convert.ToDecimal(valor);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private void dtgDetalle_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void txt_gto_aduanal_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (DDL_ent_procedencia.SelectedIndex == 0)
            {
                txt_gto_aduanal.Text = "0";
            }
            else
            { 
                //realizar calculo para asignar el gasto aduanal a cada producto(prorrateo)
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (txt_sis_tipo_cambio.Text == "" || Convert.ToDecimal(txt_sis_tipo_cambio.Text) == 1)
                    {
                        MessageBox.Show("Ingrese un de tipo de cambio", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    else
                    {
                        if (validarnumero(txt_sis_tipo_cambio.Text) == false)
                        {
                            MessageBox.Show("El valor de tipo cambio no es numérico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                    if (txt_gto_aduanal.Text == "")// || Convert.ToDecimal(txt_gto_aduanal.Text) == 0
                    {
                        MessageBox.Show("Ingrese un valor al gasto aduanal", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (validarnumero(txt_gto_aduanal.Text) == false)
                    {
                        MessageBox.Show("El valor de gasto aduanal no es numérico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    decimal totgtoadu = 0;
                    if (Convert.ToDecimal(txt_gto_aduanal.Text) == 0)
                    {
                        foreach (DataRow rw in tmgrdpro.Rows)
                        {
                            rw["gastoaduanal"] = "0.00";
                        }
                        foreach (DataRow rw in tmgrdpro.Rows)
                        {
                            if (rw["borrar"].ToString() != "1")
                                totgtoadu = totgtoadu + Convert.ToDecimal(rw["costo"].ToString());
                        }
                        if (totgtoadu > 0)
                        {
                            foreach (DataRow rw in tmgrdpro.Rows)
                            {
                                //if (rw["borrar"].ToString() != "1")
                                //{
                                //    rw["gastoaduanal"] = ((Convert.ToDecimal(rw["costo"]) / totgtoadu) * Convert.ToDecimal(txt_gto_aduanal.Text) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                                //    if(DDL_ent_procedencia.SelectedIndex == 1)
                                //        rw["unitario"] = ((Convert.ToDecimal(rw["costo"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                //    else
                                //        rw["unitario"] = (Convert.ToDecimal(rw["costo"]) + Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                //}
                                //    //rw["gastoaduanal"] = ((Convert.ToDecimal(rw["costo"]) / totgtoadu) * Convert.ToDecimal(txt_gto_aduanal.Text) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                                if (rw["borrar"].ToString() != "1")
                                {
                                    //rw["gastoaduanal"] = ((Convert.ToDecimal(rw["costo"]) / totgtoadu) * Convert.ToDecimal(txt_gto_aduanal.Text) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                                    if (chkFactura.Checked == true)
                                    {
                                        rw["gastoaduanal"] = ((Convert.ToDecimal(rw["costo"]) / totgtoadu) * Convert.ToDecimal(txt_gto_aduanal.Text)).ToString();
                                        rw["unitario"] = ((Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]));
                                    }
                                    else
                                    {
                                        rw["gastoaduanal"] = ((Convert.ToDecimal(rw["costo"]) / totgtoadu) * Convert.ToDecimal(txt_gto_aduanal.Text)).ToString();
                                        decimal a1 = (Convert.ToDecimal(rw["costo"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                                        decimal a2 = Convert.ToDecimal(rw["gastoaduanal"]);
                                        decimal a3 = Convert.ToDecimal(rw["flete"]);
                                        decimal a4 = Convert.ToDecimal(rw["adv"]);
                                        decimal a5 = Convert.ToDecimal(rw["dta"]);
                                        decimal a6 = Convert.ToDecimal(rw["cantidad"]);
                                        if (a6 == 0)
                                            a6 = 1;
                                        decimal un = ((a1 + a2 + a3 + a4 + a5) / a6);
                                        rw["unitario"] = un;
                                    }

                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (DataRow rw in tmgrdpro.Rows)
                        {
                            if(rw["borrar"].ToString() != "1")
                                totgtoadu = totgtoadu + Convert.ToDecimal(rw["costo"].ToString());
                        }
                        if (totgtoadu > 0)
                        {
                            foreach (DataRow rw in tmgrdpro.Rows)
                            {
                                //if (rw["borrar"].ToString() != "1")
                                //{
                                //    rw["gastoaduanal"] = ((Convert.ToDecimal(rw["costo"]) / totgtoadu) * Convert.ToDecimal(txt_gto_aduanal.Text) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                                //    if(DDL_ent_procedencia.SelectedIndex == 1)
                                //        rw["unitario"] = ((Convert.ToDecimal(rw["costo"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                //    else
                                //        rw["unitario"] = (Convert.ToDecimal(rw["costo"]) + Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                //}
                                //    //rw["gastoaduanal"] = ((Convert.ToDecimal(rw["costo"]) / totgtoadu) * Convert.ToDecimal(txt_gto_aduanal.Text) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                                if (rw["borrar"].ToString() != "1")
                                {
                                    //rw["gastoaduanal"] = ((Convert.ToDecimal(rw["costo"]) / totgtoadu) * Convert.ToDecimal(txt_gto_aduanal.Text) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                                    if (chkFactura.Checked == true)
                                    {
                                        rw["gastoaduanal"] = ((Convert.ToDecimal(rw["costo"]) / totgtoadu) * Convert.ToDecimal(txt_gto_aduanal.Text)).ToString();
                                        rw["unitario"] = ((Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]));
                                    }
                                    else
                                    {
                                        rw["gastoaduanal"] = ((Convert.ToDecimal(rw["costo"]) / totgtoadu) * Convert.ToDecimal(txt_gto_aduanal.Text)).ToString();
                                        decimal a1 = 0;
                                        if (Convert.ToDecimal(lblTCFact.Text) > 1)
                                        {
                                            a1 = Convert.ToDecimal(rw["costo"]) * Convert.ToDecimal(this.lblTCFact.Text);
                                        }
                                        else
                                        { 
                                            a1 = Convert.ToDecimal(rw["costo"]) * Convert.ToDecimal(this.txt_sis_tipo_cambio.Text);
                                        }
                                        //!(Convert.ToDecimal(this.lblTCFact.Text) > new Decimal(0)) ?  : ;
                                        //a1 = (Convert.ToDecimal(rw["costo"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) * Convert.ToDecimal(this.lblTCFact.Text);
                                        decimal a2 = Convert.ToDecimal(rw["gastoaduanal"]);
                                        decimal a3 = Convert.ToDecimal(rw["flete"]);
                                        decimal a4 = Convert.ToDecimal(rw["adv"]);
                                        decimal a5 = Convert.ToDecimal(rw["dta"]);
                                        decimal a6 = Convert.ToDecimal(rw["cantidad"]);
                                        if (a6 == 0)
                                            a6 = 1;
                                        decimal un = ((a1 + a2 + a3 + a4 + a5) / a6);
                                        rw["unitario"] = un;
                                    }
                                        
                                }
                            }
                        }
                    }
                    
                    
                    
                    //Cargar datos al datagrid
                    dtgDetalle.Rows.Clear();
                    foreach (DataRow rr in tmgrdpro.Rows)
                    {
                        if (rr["borrar"].ToString() != "1")
                        {
                            dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
                                Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
                                Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                                Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                                Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                                rr["borrar"].ToString(), rr["oper"].ToString(), "");//, Convert.ToDecimal(rr["iva"].ToString()).ToString("###,###,##0.000")
                        }
                            //dtgDetalle.Rows.Add(rr.ItemArray);
                    }
                    calculatotales();
                }
            }
        }

        private void txtFlete_KeyPress(object sender, KeyPressEventArgs e)
        {
            //realizar calculo para asignar el flete a cada producto(prorrateo)
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (txtFlete.Text == "")
                {
                    MessageBox.Show("Ingrese un valor al flete", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (validarnumero(txtFlete.Text) == false)
                {
                    MessageBox.Show("El valor de flete no es numérico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                decimal totflete = 0;
                foreach (DataRow rw in tmgrdpro.Rows)
                {
                    if(rw["borrar"].ToString() != "1")
                        totflete = totflete + Convert.ToDecimal(rw["costo"].ToString());
                }
                if (totflete > 0)
                {
                    foreach (DataRow rw in tmgrdpro.Rows)
                    {
                        if (rw["borrar"].ToString() != "1")
                        {
                            //if (DDL_ent_procedencia.SelectedIndex == 0)
                            //{
                            //    rw["flete"] = ((Convert.ToDecimal(rw["costo"]) / totflete) * Convert.ToDecimal(txtFlete.Text)).ToString();
                            //    if (Convert.ToDecimal(rw["cantidad"]) > 0)
                            //    {
                            //        rw["unitario"] = (Convert.ToDecimal(rw["costo"]) + Convert.ToDecimal(rw["flete"])) / Convert.ToDecimal(rw["cantidad"]);
                            //    }
                            //}
                            //else
                            //{
                            //    rw["flete"] = ((Convert.ToDecimal(rw["costo"]) / totflete) * Convert.ToDecimal(txtFlete.Text)).ToString();
                            //    if (Convert.ToDecimal(rw["cantidad"]) > 0)
                            //    {
                            //        rw["unitario"] = ((Convert.ToDecimal(rw["costo"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                            //        //rw["unitario"] = (Convert.ToDecimal(rw["costo"]) + Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"])) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"]) / Convert.ToDecimal(rw["cantidad"]);
                            //    }
                            //}
                            if (DDL_ent_procedencia.SelectedIndex == 0)
                            {
                                rw["flete"] = ((Convert.ToDecimal(rw["costo"]) / totflete) * (Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text))).ToString();
                                if (Convert.ToDecimal(rw["cantidad"]) > 0)
                                {
                                    if(chkFactura.Checked == true)
                                        rw["unitario"] = (Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                    else
                                        rw["unitario"] = (Convert.ToDecimal(rw["costo"]) + Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                    //rw["unitario"] = (Convert.ToDecimal(rw["unitario"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString("###,###,##0.000");
                                }
                            }
                            else
                            {
                                rw["flete"] = ((Convert.ToDecimal(rw["costo"]) / totflete) * (Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text))).ToString();//flete en pesos
                                if (Convert.ToDecimal(rw["cantidad"]) > 0)
                                {
                                    if (chkFactura.Checked == true)
                                        rw["unitario"] = (Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                    else
                                        rw["unitario"] = ((Convert.ToDecimal(rw["costo"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                    //rw["unitario"] = (Convert.ToDecimal(rw["unitario"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString("###,###,##0.000");
                                }
                            }   
                        }
                    }
                }

                //Cargar datos al datagrid
                dtgDetalle.Rows.Clear();
                foreach (DataRow rr in tmgrdpro.Rows)
                {
                    if (rr["borrar"].ToString() != "1")
                    {
                        dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
                            Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
                            Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                            Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                            Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                            rr["borrar"].ToString(), rr["oper"].ToString(), "");//, Convert.ToDecimal(rr["iva"].ToString()).ToString("###,###,##0.000")
                    }
                        //dtgDetalle.Rows.Add(rr.ItemArray);
                }
                calculatotales();
            }
        }

        private void txt_ent_iva_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                calculatotales();
            }
        }

        private void envia_correo(string data)
        {
            form_entradas_ocs.Data.GraphMailService.Enviar(
                new string[] { "aescamilla@mrlucky.com.mx" }, null, new string[] { "aescamilla@mrlucky.com.mx" },
                "Envio de Orden de Compra de entrada.", data, null);
        }

        public void guardar()
        {
            //VALIDACION POR CIERRE DE MES
            if (Convert.ToDateTime(dtpFecha.Text) < Convert.ToDateTime(txtCierre.Text))
            {
                MessageBox.Show("No puede modificar la entrada por que es de fecha anterior al cierre de mes", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //return;
            }
            if (validacontroles() == false)
            {
                return;
            }

            //VALIDAR CAPTURA DE LOTE A EMPAQUES OBLIGATORIOS
            if (DDLAlmacen1.SelectedIndex == 1)
            {
                if (validar_empaque_lote() == true)
                {
                    return;
                }
            }

            if (validagrid_N_datatable() == true)
            {
                MessageBox.Show("Los empaques no pueden tener costos en cero, favor de verificarlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (lblXmlProv.Text == "1")//XML TIENE QUE SER OMITIDO
            { 
            }
            else//SI NO DEBE DE SER OBLIGATORIO
            {
                if (txt_prov_clave.Text != "81")//SI EL PROVEEDOR NO ES INTERNATIONAL PAPER EL XML DEBE SER OBLIGATORIO
                {
                    //validacion de seleccion de XML
                    if (txtFolioFiscal.Text == "")//VALIDACION SI YA SELECCIONO EL XML
                    {
                        MessageBox.Show("No ha seleccionado el XML, favor de verificarlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }    
            }
            
            

            imprimirtxt();
            if (MessageBox.Show("¿Desea realizar algún cambio en la orden de entrada?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                return;
            }


            bool encontrado = false;

            //string fac = "";
            //fac = txt_ent_factura.Text;
            //if (fac == "")
            //{
            //    if (MessageBox.Show("No ha ingresado el número de factura, ¿Desea ingresarlo?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //        return;
            //    else
            //    {
            //        txt_ent_factura.Text = "S/N";
            //    }
            //}
            //verificarfolio();
            //Insercion en tb_mstr_entradas
            string query = "";
            List<string> queries = new List<string>();
            try
            {
                

                string tabla = "";
                switch (DDLAlmacen1.SelectedIndex)
                {
                    case 0:
                        tabla = "tb_mstr_entradas_mat";
                        break;
                    case 1:
                        tabla = "tb_mstr_entradas_emp";
                        break;
                    case 2:
                        tabla = "tb_mstr_entradas_com";
                        break;
                    case 3:
                        tabla = "tb_mstr_entradas_gua";
                        break;
                    case 4:
                        tabla = "tb_mstr_entradas_enf";
                        break;
                    case 5:
                        tabla = "tb_mstr_entradas_pap";
                        break;
                    case 6:
                        tabla = "tb_mstr_entradas_den";
                        break;
                    case 7:
                        tabla = "tb_mstr_entradas_comb";
                        break;
                    case 8:
                        tabla = "tb_mstr_entradas_sis";
                        break;
                    case 9:
                        tabla = "tb_mstr_entradas_guar";
                        break;
                    case 10:
                        tabla = "tb_mstr_entradas_guar_com";
                        break;
                    case 11:
                        tabla = "tb_mstr_entradas_even";
                        break;
                    case 12:
                        tabla = "tb_mstr_entradas_mot";
                        break;
                }

                string str3 = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                string tip = "NOR";
                string reg = "DEF";
                string pro = ((DDL_ent_procedencia.SelectedItem.ToString() == "Nacional") ? "NAL" : "IMP");

                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();

                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();

                string q_u_e_r_y = "IF NOT EXISTS(SELECT ent_folio, compu, usuario FROM " + tabla + " " +
                    "WHERE ent_folio = '" + txtFolio.Text.TrimStart('0') + "' AND compu = '" + Environment.MachineName + "' AND usuario = '" + Utilerias.Class1.Usu_login + "') " +
                    "INSERT INTO " + tabla + " (ent_fecha, compu, usuario) VALUES('" + str3 + "'," +
                    " '" + Environment.MachineName + "', '" + Utilerias.Class1.Usu_login + "') SELECT SCOPE_IDENTITY()";

                cmnd1.CommandText = q_u_e_r_y;
                var_int_sis_folios = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                cmnd1.Dispose();

                if (this.var_int_sis_folios == null || this.var_int_sis_folios == "")
                {
                    MessageBox.Show("No se generó el folio de entrada, vuelva a guardar la entrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    thisConnection.Close();
                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + q_u_e_r_y);
                    return;
                }

                var_int_sis_folios = var_int_sis_folios.PadLeft(6, '0').ToString();

                

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT ent_folio, compu, usuario FROM " + tabla + " WHERE ent_folio = '" + var_int_sis_folios + "' AND compu = '" + Environment.MachineName + "' " +
                    "AND usuario = '" +Utilerias.Class1.Usu_login + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    
                }
                else
                {
                    cmnd2.CommandText = "INSERT INTO " + tabla + " (ent_fecha, compu, usuario) VALUES('" + str3 + "'," +
                    " '" + Environment.MachineName + "', '" + Utilerias.Class1.Usu_login + "') SELECT SCOPE_IDENTITY()";
                    var_int_sis_folios = Convert.ToString(cmnd2.ExecuteScalar()).Trim();
                    cmnd2.Dispose();
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                if (var_int_sis_folios.PadLeft(6, '0') != txtFolio.Text)
                {
                    MessageBox.Show("Por movimientos en la red la Entrada se generó con el siguiente número de recibo " + var_int_sis_folios, "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtFolio.Text = var_int_sis_folios.PadLeft(6, '0');
                }
                else
                {
                    txtFolio.Text = var_int_sis_folios.PadLeft(6, '0');
                }

                //15/01/2025 CREAR PDF DE LA ENTRADA Y GUARDARLO EN EL SERVIDOR Y SUBIR DOCUMENTOS DE LA FACTURA
                guarda_genera_pdf_servidor();

                #region archivos_factura
                string importacion = "0";
                thisConnection.Open();
                SqlCommand cmndimportacion = thisConnection.CreateCommand();
                cmndimportacion.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
                try
                {
                    importacion = Convert.ToString(cmndimportacion.ExecuteScalar().ToString().Trim());
                }
                catch
                {
                    importacion = "0";
                }

                cmndimportacion.Dispose();
                thisConnection.Close();

                if (importacion == "0") //El proveedor require archivos
                {
                    if (txtFolioFiscal.Text != "")
                    {
                        FileInfo archivo = new FileInfo(@lblPathXml.Text);
                        string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                        FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        if (liq_copy.Exists == true)
                        {
                            liq_copy.Delete();
                        }
                        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        if (liq_copy.Exists != true)
                        {
                            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de XML " + file_name);
                        }


                        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de XML");

                        file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        if (liq_copy.Exists == true)
                        {
                            liq_copy.Delete();
                        }
                        archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        if (liq_copy.Exists != true)
                        {
                            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de PDF " + file_name);
                        }

                        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de PDF");
                        //if (txtFolioFiscal.Text != "")
                        //{

                        //}
                    }
                }
                else
                {
                    if (txtFolioFiscal.Text != "")
                    {
                        FileInfo archivo = new FileInfo(@lblPathXml.Text);
                        string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                        FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        if (liq_copy.Exists == true)
                        {
                            liq_copy.Delete();
                        }
                        if (liq_copy.Exists == false)
                        {
                            Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + file_name + " No existe", "SISEMP");
                        }
                        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

                        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        if (liq_copy.Exists != true)
                        {
                            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de XML " + file_name);
                            Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + lblPathXml.Text, "SISEMP");
                        }

                        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de PDF Ruta: " + lblPathXml.Text);

                        file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        if (liq_copy.Exists == true)
                        {
                            liq_copy.Delete();
                        }
                        if (liq_copy.Exists == false)
                        {
                            Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + file_name + " No existe", "SISEMP");
                        }
                        archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

                        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        if (liq_copy.Exists != true)
                        {
                            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de PDF " + file_name);
                            Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + lblPathXml.Text, "SISEMP");
                        }

                        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de PDF Ruta: " + lblPathXml.Text);
                    }
                }
                #endregion
                //FIN 15/01/2025 CREAR PDF DE LA ENTRADA Y GUARDARLO EN EL SERVIDOR

                //////BUSCAR FOLIO DE ENTRADA EN EL MAESTRO (CONSECUTIVOS)
                ////cmnd1 = thisConnection.CreateCommand();
                ////cmnd1.CommandText = "SELECT ent_folio, compu, usuario FROM " + tabla + " WHERE ent_folio = '" + var_int_sis_folios + "'" +
                ////    " AND compu = '" + Environment.MachineName + "' AND usuario = '" + Utilerias.Class1.Usu_login + "'";
                ////reader1 = cmnd1.ExecuteReader();
                ////bool bandera = false;
                ////if (reader1.HasRows)
                ////{
                ////    bandera = true;
                ////}
                ////else
                ////{
                ////    bandera = false;
                ////}
                ////reader1.Close();
                ////reader1.Dispose();
                ////cmnd1.Dispose();

                ////if (bandera == false)
                ////{
                ////    cmnd1 = thisConnection.CreateCommand();
                ////    cmnd1.CommandText = "INSERT INTO " + tabla + " (ent_fecha, compu, usuario) VALUES('" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "'," +
                ////        " '" + Environment.MachineName + "', '" + Utilerias.Class1.Usu_login + "') SELECT SCOPE_IDENTITY()";
                ////    var_int_sis_folios = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                ////    cmnd1.Dispose();
                ////}

                //09/04/2021 mandar detalle por fallas al guardar en tb_det_entradas_oc
                string asztal = "<html><body><table border='1'><thead><tr><th>ENTRADA</th><th>ORDEN_COMPRA</th><th>EMPAQUE</th><th>CANTIDAD</th><th>IMPORTE</th><th>REQUISICION</th><th>LOTE</th></tr></thead><tbody>";
                foreach (DataRow rw in dtdetalles.Rows)
                {
                    asztal += "<tr><td>" + var_int_sis_folios + "</td><td>" + rw["orden_compra"].ToString() + "</td><td>" + rw["emp_clave"].ToString() + "</td>" +
                        "<td>" + rw["cantidad"].ToString() + "</td><td>" + rw["importe"].ToString() + "</td><td>" + rw["requisicion"].ToString() + "</td>" +
                        "<td>" + rw["lote"].ToString() + "</td></tr>";
                }
                asztal += "</tbody></table></body></html>";
                envia_correo(asztal);
                //mandar detalle por fallas al guardar en tb_det_entradas_oc

                

                lblXmlFile.Text = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";

                
                //string sub = ((glo_var_int_registro == 7) ? ", 'COM'" : (glo_var_int_registro == 9) ? ", 'GUA'" : ", ''");
                //cmnd1 = thisConnection.CreateCommand();
                //if (lblXmlProv.Text == "1")//EL INSERT DEL PROVEEDOR DEBE SER FECHA ACTUAL Y FOLIO FISCAL DEBE DECIR 'SIN FOLIO FISCAL'
                //{
                //    query = "INSERT INTO tb_mstr_entradas(ent_folio, ent_tipo, ent_fecha, ent_destino, ent_elaboro, ent_autorizo, ent_recibio," +
                //            " ent_observaciones, alm_clave, ent_estatus, prov_clave, ent_iva, ent_monto, ent_ord_compra," +
                //            " ent_factura, sis_tipo_cambio, ent_regimen, ent_procedencia, ent_pedimento, ent_subtipo, ent_gastoaduanal, ent_iepstotal, ent_flete, ent_flete_imp, ent_sin_fact, ent_fechafact, ent_foliofiscal, ent_alta)" +
                //            " VALUES('" + var_int_sis_folios + "', '" + tip + "', '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '', '" + DDLElaboro.SelectedItem.ToString() + "', '" + DDLAutorizo.SelectedItem.ToString() + "'," +
                //            " '', '" + txtObservaciones.Text + "', '" + glo_var_chr_alm_clave + "', '', '" + txt_prov_clave.Text + "', '" + Convert.ToDecimal(txt_ent_iva.Text).ToString("0.0000") + "'," +
                //            " '" + Convert.ToDecimal(txt_total.Text).ToString("0.0000") + "', '', '" + txt_ent_factura.Text + "', '" + txt_sis_tipo_cambio.Text + "', '" + reg + "', '" + pro + "', '" + txt_ent_pedimento.Text + "', '" + subtipo + "'," +
                //            " '" + ((txt_gto_aduanal.Text == "0") ? "0.0000" : Convert.ToDecimal(txt_gto_aduanal.Text).ToString("0.0000")) + "', '" + ((txt_ent_ieps.Text == "0") ? "0.0000" : Convert.ToDecimal(txt_ent_ieps.Text).ToString()) + "'," +
                //            " '" + ((txtFlete.Text == "0") ? "0.0000" : Convert.ToDecimal(txtFlete.Text).ToString("0.0000")) + "', '" + ((txtFleteImp.Text == "0") ? "0.0000" : Convert.ToDecimal(txtFleteImp.Text).ToString("0.0000")) + "', '" + ((chkFactura.Checked == true) ? "1" : "0") + "', " +
                //            " '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '" + txtFolioFiscal.Text + "', '" + DateTime.Now.ToString() + "')";
                //    cmnd1.CommandText = query;
                //    cmnd1.ExecuteNonQuery();
                //    cmnd1.Dispose();
                //}
                //else//EL PROVEEDOR REQUIERE EL XML
                //{


                

                if (thisConnection.State == ConnectionState.Closed)
                    thisConnection.Open();

                cmnd1 = thisConnection.CreateCommand();
                if (txt_prov_clave.Text == "81")//SI EL PROVEEDOR ES INTERNATIONAL PAPER DEJAR PASAR AL GUARDAR
                {
                    if (txtFolioFiscal.Text == "")//VALIDACION SI YA SELECCIONO EL XML
                    {
                        query = "INSERT INTO tb_mstr_entradas(ent_folio, ent_tipo, ent_fecha, ent_destino, ent_elaboro, ent_autorizo, ent_recibio," +
                            " ent_observaciones, alm_clave, ent_estatus, prov_clave, ent_iva, ent_monto, ent_ord_compra," +
                            " ent_factura, sis_tipo_cambio, ent_regimen, ent_procedencia, ent_pedimento, ent_subtipo, ent_gastoaduanal, ent_iepstotal, ent_flete, ent_flete_imp, ent_sin_fact, ent_fechafact, ent_foliofiscal, ent_alta," +
                            " descuento, ret_iva, ret_isr, cedular, ent_iva_imp, ent_desc_pp, ent_desc_im)" +
                            " VALUES('" + var_int_sis_folios + "', '" + tip + "', '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '', '" + DDLElaboro.SelectedItem.ToString() + "', '" + DDLAutorizo.SelectedItem.ToString() + "'," +
                            " '', '" + txtObservaciones.Text + "', '" + glo_var_chr_alm_clave + "', '', '" + txt_prov_clave.Text + "', '" + Convert.ToDecimal(txt_ent_iva.Text).ToString("0.0000") + "'," +
                            " '" + Convert.ToDecimal(txt_total.Text).ToString("0.0000") + "', '', '" + txt_ent_factura.Text + "', '" + txt_sis_tipo_cambio.Text + "', '" + reg + "', '" + pro + "', '" + txt_ent_pedimento.Text + "', '" + subtipo + "'," +
                            " '" + ((txt_gto_aduanal.Text == "0") ? "0.0000" : Convert.ToDecimal(txt_gto_aduanal.Text).ToString("0.0000")) + "', '" + ((txt_ent_ieps.Text == "0") ? "0.0000" : Convert.ToDecimal(txt_ent_ieps.Text).ToString()) + "'," +
                            " '" + ((txtFlete.Text == "0") ? "0.0000" : Convert.ToDecimal(txtFlete.Text).ToString("0.0000")) + "', '" + ((txtFleteImp.Text == "0") ? "0.0000" : Convert.ToDecimal(txtFleteImp.Text).ToString("0.0000")) + "', '" + ((chkFactura.Checked == true) ? "1" : "0") + "', " +
                            " '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '" + txtFolioFiscal.Text + "', '" + DateTime.Now.ToString() + "'," +
                            " '" + Convert.ToDecimal(txtDesc.Text).ToString() + "', '" + Convert.ToDecimal(txtRetIva.Text).ToString() + "', '" + Convert.ToDecimal(txtIsr.Text).ToString() + "', '" + Convert.ToDecimal(txtCedular.Text).ToString() + "', '" + Convert.ToDecimal(txt_iva.Text).ToString() + "'," +
                            " '" + lblPorcentajeDesc.Text + "', '" + lblImporteDescuento.Text + "')";
                        //queries.Add(query);
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();
                    }
                    else//SI AL MOMENTO DE GUARDAR YA REGISTRO XML GUARDAR NORMALMENTE
                    {
                        query = "INSERT INTO tb_mstr_entradas(ent_folio, ent_tipo, ent_fecha, ent_destino, ent_elaboro, ent_autorizo, ent_recibio," +
                            " ent_observaciones, alm_clave, ent_estatus, prov_clave, ent_iva, ent_monto, ent_ord_compra," +
                            " ent_factura, sis_tipo_cambio, ent_regimen, ent_procedencia, ent_pedimento, ent_subtipo, ent_gastoaduanal, ent_iepstotal, ent_flete, ent_flete_imp, ent_sin_fact, ent_fechafact, ent_foliofiscal, ent_alta," +
                            " descuento, ret_iva, ret_isr, cedular, ent_iva_imp, ent_desc_pp, ent_desc_im)" +
                            " VALUES('" + var_int_sis_folios + "', '" + tip + "', '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '', '" + DDLElaboro.SelectedItem.ToString() + "', '" + DDLAutorizo.SelectedItem.ToString() + "'," +
                            " '', '" + txtObservaciones.Text + "', '" + glo_var_chr_alm_clave + "', '', '" + txt_prov_clave.Text + "', '" + Convert.ToDecimal(txt_ent_iva.Text).ToString("0.0000") + "'," +
                            " '" + Convert.ToDecimal(txt_total.Text).ToString("0.0000") + "', '', '" + txt_ent_factura.Text + "', '" + txt_sis_tipo_cambio.Text + "', '" + reg + "', '" + pro + "', '" + txt_ent_pedimento.Text + "', '" + subtipo + "'," +
                            " '" + ((txt_gto_aduanal.Text == "0") ? "0.0000" : Convert.ToDecimal(txt_gto_aduanal.Text).ToString("0.0000")) + "', '" + ((txt_ent_ieps.Text == "0") ? "0.0000" : Convert.ToDecimal(txt_ent_ieps.Text).ToString()) + "'," +
                            " '" + ((txtFlete.Text == "0") ? "0.0000" : Convert.ToDecimal(txtFlete.Text).ToString("0.0000")) + "', '" + ((txtFleteImp.Text == "0") ? "0.0000" : Convert.ToDecimal(txtFleteImp.Text).ToString("0.0000")) + "', '" + ((chkFactura.Checked == true) ? "1" : "0") + "'," +
                            " '" + txtFechaFact.Text + "', '" + txtFolioFiscal.Text + "', '" + DateTime.Now.ToString() + "'," +
                            " '" + Convert.ToDecimal(txtDesc.Text).ToString() + "', '" + Convert.ToDecimal(txtRetIva.Text).ToString() + "', '" + Convert.ToDecimal(txtIsr.Text).ToString() + "', '" + Convert.ToDecimal(txtCedular.Text).ToString() + "', '" + Convert.ToDecimal(txt_iva.Text).ToString() + "'," +
                            " '" + lblPorcentajeDesc.Text + "', '" + lblImporteDescuento.Text + "')";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();
                        //queries.Add(query);

                        //GUARDAR DATOS EN TB_DET_OC_FACTURAS
                        if (lblPortal.Text == "-")
                        {
                            query = "INSERT INTO tb_det_oc_facturas(numero_oc, factura, fecha_factura, monto, fecha_registro, folio_fiscal, tipo, estatus, portal, estatus_portal) " +
                            "VALUES('" + var_int_sis_folios.TrimStart('0').ToString() + "', '" + txt_ent_factura.Text + "', '" + txtFechaFact.Text + "', '" + lblTotal.Text + "', '" + dtpFecha.Text + "', " +
                            "'" + txtFolioFiscal.Text + "', '" + DDLAlmacen1.Text.ToUpper() + "', 'A', '', '0')";
                            queries.Add(query);
                        }
                        else
                        {
                            query = "INSERT INTO tb_det_oc_facturas(numero_oc, factura, fecha_factura, monto, fecha_registro, folio_fiscal, tipo, estatus, portal, estatus_portal) " +
                            "VALUES('" + var_int_sis_folios.TrimStart('0').ToString() + "', '" + txt_ent_factura.Text + "', '" + txtFechaFact.Text + "', '" + lblTotal.Text + "', '" + dtpFecha.Text + "', " +
                            "'" + txtFolioFiscal.Text + "', '" + DDLAlmacen1.Text.ToUpper() + "', 'A', '" + txtAcuse.Text + "', '1')";
                            queries.Add(query);
                        }

                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                    }//84353
                }
                else//SI ES CUALQUIER OTRO PROVEEDOR GUARDAR LOS DATOS DEL XML
                {
                        

                    query = "INSERT INTO tb_mstr_entradas(ent_folio, ent_tipo, ent_fecha, ent_destino, ent_elaboro, ent_autorizo, ent_recibio," +
                        " ent_observaciones, alm_clave, ent_estatus, prov_clave, ent_iva, ent_monto, ent_ord_compra," +
                        " ent_factura, sis_tipo_cambio, ent_regimen, ent_procedencia, ent_pedimento, ent_subtipo, ent_gastoaduanal, ent_iepstotal, ent_flete, ent_flete_imp, ent_sin_fact, ent_fechafact, ent_foliofiscal, ent_alta," +
                        " descuento, ret_iva, ret_isr, cedular, ent_iva_imp, ent_desc_pp, ent_desc_im)" +
                        " VALUES('" + var_int_sis_folios + "', '" + tip + "', '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '', '" + DDLElaboro.SelectedItem.ToString() + "', '" + DDLAutorizo.SelectedItem.ToString() + "'," +
                        " '', '" + txtObservaciones.Text + "', '" + glo_var_chr_alm_clave + "', '', '" + txt_prov_clave.Text + "', '" + Convert.ToDecimal(txt_ent_iva.Text).ToString("0.0000") + "'," +
                        " '" + Convert.ToDecimal(txt_total.Text).ToString("0.0000") + "', '', '" + txt_ent_factura.Text + "', '" + txt_sis_tipo_cambio.Text + "', '" + reg + "', '" + pro + "', '" + txt_ent_pedimento.Text + "', '" + subtipo + "'," +
                        " '" + ((txt_gto_aduanal.Text == "0") ? "0.0000" : Convert.ToDecimal(txt_gto_aduanal.Text).ToString("0.0000")) + "', '" + ((txt_ent_ieps.Text == "0") ? "0.0000" : Convert.ToDecimal(txt_ent_ieps.Text).ToString()) + "'," +
                        " '" + ((txtFlete.Text == "0") ? "0.0000" : Convert.ToDecimal(txtFlete.Text).ToString("0.0000")) + "', '" + ((txtFleteImp.Text == "0") ? "0.0000" : Convert.ToDecimal(txtFleteImp.Text).ToString("0.0000")) + "', '" + ((chkFactura.Checked == true) ? "1" : "0") + "'," +
                        " '" + txtFechaFact.Text + "', '" + txtFolioFiscal.Text + "', '" + DateTime.Now.ToShortDateString() + "'," +
                        " '" + Convert.ToDecimal(txtDesc.Text).ToString() + "', '" + Convert.ToDecimal(txtRetIva.Text).ToString() + "', '" + Convert.ToDecimal(txtIsr.Text).ToString() + "', '" + Convert.ToDecimal(txtCedular.Text).ToString() + "', '" + Convert.ToDecimal(txt_iva.Text).ToString() + "'," +
                        " '" + lblPorcentajeDesc.Text + "', '" + lblImporteDescuento.Text + "')";
                    cmnd1.CommandText = query;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();
                    //queries.Add(query);

                    //if(lblXmlProv.Text != "1")
                    //{
                    if (lblPortal.Text == "-")
                    {
                        query = "INSERT INTO tb_det_oc_facturas(numero_oc, factura, fecha_factura, monto, fecha_registro, folio_fiscal, tipo, estatus, portal, estatus_portal) " +
                            "VALUES('" + var_int_sis_folios.TrimStart('0').ToString() + "', '" + txt_ent_factura.Text + "', '" + txtFechaFact.Text + "', '" + lblTotal.Text + "', '" + dtpFecha.Text + "', " +
                            "'" + txtFolioFiscal.Text + "', '" + DDLAlmacen1.Text.ToUpper() + "', 'A', '" + txtAcuse.Text + "', '0')";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();
                        //queries.Add(query);
                    }
                    else
                    {
                        query = "INSERT INTO tb_det_oc_facturas(numero_oc, factura, fecha_factura, monto, fecha_registro, folio_fiscal, tipo, estatus, portal, estatus_portal) " +
                            "VALUES('" + var_int_sis_folios.TrimStart('0').ToString() + "', '" + txt_ent_factura.Text + "', '" + txtFechaFact.Text + "', '" + lblTotal.Text + "', '" + dtpFecha.Text + "', " +
                            "'" + txtFolioFiscal.Text + "', '" + DDLAlmacen1.Text.ToUpper() + "', 'A', '" + txtAcuse.Text + "', '1')";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();
                        //queries.Add(query);
                    }
                            
                    //}
                    ////GUARDAR DATOS EN TB_DET_OC_FACTURAS
                    //if (lblPortal.Text == "")
                    //{
                    //    query = "INSERT INTO tb_det_oc_facturas(numero_oc, factura, fecha_factura, monto, fecha_registro, folio_fiscal, tipo, estatus, portal, estatus_portal) " +
                    //    "VALUES('" + var_int_sis_folios.TrimStart('0').ToString() + "', '" + txt_ent_factura.Text + "', '" + txtFechaFact.Text + "', '" + lblTotal.Text + "', '" + dtpFecha.Text + "', " +
                    //    "'" + txtFolioFiscal.Text + "', '" + DDLAlmacen1.Text.ToUpper() + "', 'A', '', '0')";
                    //}
                    //else
                    //{
                    //    query = "INSERT INTO tb_det_oc_facturas(numero_oc, factura, fecha_factura, monto, fecha_registro, folio_fiscal, tipo, estatus, portal, estatus_portal) " +
                    //    "VALUES('" + var_int_sis_folios.TrimStart('0').ToString() + "', '" + txt_ent_factura.Text + "', '" + txtFechaFact.Text + "', '" + lblTotal.Text + "', '" + dtpFecha.Text + "', " +
                    //    "'" + txtFolioFiscal.Text + "', '" + DDLAlmacen1.Text.ToUpper() + "', 'A', '" + lblPortal.Text + "', '1')";
                    //}
                        
                        
                }
                //}
                thisConnection.Close();

                
            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de guardar la entrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.5", txtFolio.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }

            //SUBIR DOCUMENTOS DE FACTURA Y XML
            #region docs upload

            

            //if (lblXmlProv.Text == "0")//SI EL PROVEEDOR REQUIERE XML SUBIR ARCHIVO
            //{
            //    string importacion = "0";
            //    thisConnection.Open();
            //    SqlCommand cmndimportacion = thisConnection.CreateCommand();
            //    cmndimportacion.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
            //    try
            //    {
            //        importacion = Convert.ToString(cmndimportacion.ExecuteScalar().ToString().Trim());
            //    }
            //    catch
            //    {
            //        importacion = "";
            //    }

            //    cmndimportacion.Dispose();
            //    thisConnection.Close();

            //    if (importacion == "1")
            //    {
            //        /*if (txt_prov_clave.Text == "81" || txt_prov_clave.Text == "404" || txt_prov_clave.Text == "11" || txt_prov_clave.Text == "8851" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694" || txt_prov_clave.Text.Trim() == "9957" || txt_prov_clave.Text.Trim() == "833" || txt_prov_clave.Text.Trim() == "1550" || txt_prov_clave.Text.Trim() == "1761" || txt_prov_clave.Text.Trim() == "935" || txt_prov_clave.Text.Trim() == "6854" || txt_prov_clave.Text == "354" || txt_prov_clave.Text == "1250" || txt_prov_clave.Text == "1536" || txt_prov_clave.Text == "9758" || txt_prov_clave.Text == "1799" || txt_prov_clave.Text == "1244" || txt_prov_clave.Text == "191" || txt_prov_clave.Text == "812" || txt_prov_clave.Text == "8014" || txt_prov_clave.Text == "305" || txt_prov_clave.Text == "106")
            //        {*/
            //        if (txtFolioFiscal.Text != "")//SI SELECCIONO EL FOLIO FISCAL SUBIR EL XML
            //        {
            //            FileInfo archivo = new FileInfo(@lblPathXml.Text);
            //            string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
            //            FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists == true)
            //            {
            //                liq_copy.Delete();
            //            }
            //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists != true)
            //            {
            //                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de XML " + file_name);
            //            }


            //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de XML");

            //            file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists == true)
            //            {
            //                liq_copy.Delete();
            //            }
            //            archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
            //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists != true)
            //            {
            //                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de PDF " + file_name);
            //            }

            //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de PDF");


            //            ////SUBIR ARCHIVO xml
            //            //string origen1 = @lblPathXml.Text;//nombre de archivo
            //            //string destino1 = @lblDestino.Text + lblXmlFile.Text;
            //            ////if (!File.Exists(origen1))
            //            ////{
            //            ////    //using (FileStream fs = File.Create(origen1)) { }

            //            ////}
            //            //File.Copy(origen1, destino1, true);
            //            ////if (File.Exists(lblDestino.Text))
            //            ////File.Delete(lblDestino.Text);


            //            ////SUBIR PDF
            //            //origen1 = @lblPathXml.Text.Replace(".xml", ".pdf");//nombre de archivo
            //            //destino1 = @lblDestino.Text + lblXmlFile.Text.Replace(".xml", ".pdf");
            //            ////if (!File.Exists(origen1))
            //            ////{
            //            ////    //using (FileStream fs = File.Create(origen1)) { }

            //            ////}
            //            //File.Copy(origen1, destino1, true);
            //            ////if (File.Exists(lblDestino.Text))
            //            ////File.Delete(lblDestino.Text);

            //        }

            //    }
            //    else
            //    {
            //        if (txtFolioFiscal.Text != "")
            //        {
            //            FileInfo archivo = new FileInfo(@lblPathXml.Text);
            //            string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
            //            FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists == true)
            //            {
            //                liq_copy.Delete();
            //            }
            //            if (liq_copy.Exists == false)
            //            {
            //                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + file_name + " No existe", "SISEMP");
            //            }
            //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists != true)
            //            {
            //                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
            //                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de XML " + file_name);
            //                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + lblPathXml.Text, "SISEMP");
            //            }

            //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de PDF Ruta: " + lblPathXml.Text);

            //            file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists == true)
            //            {
            //                liq_copy.Delete();
            //            }
            //            if (liq_copy.Exists == false)
            //            {
            //                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + file_name + " No existe", "SISEMP");
            //            }
            //            archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
            //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists != true)
            //            {
            //                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
            //                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de PDF " + file_name);
            //                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + lblPathXml.Text, "SISEMP");
            //            }

            //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de PDF Ruta: " + lblPathXml.Text);
            //        }


            //        //string origen1 = @lblPathXml.Text;
            //        //string destino1 = @lblDestino.Text + lblXmlFile.Text;
            //        ////if (!File.Exists(origen1))
            //        ////{
            //        ////    //using (FileStream fs = File.Create(origen1)) { }

            //        ////}
            //        //File.Copy(origen1, destino1, true);
            //        ////if (File.Exists(lblDestino.Text))
            //        ////File.Delete(lblDestino.Text);



            //        ////SUBIR PDF
            //        //origen1 = @lblPathXml.Text.Replace(".xml", ".pdf");//nombre de archivo
            //        //destino1 = @lblDestino.Text + lblXmlFile.Text.Replace(".xml", ".pdf");
            //        ////if (!File.Exists(origen1))
            //        ////{
            //        ////    //using (FileStream fs = File.Create(origen1)) { }

            //        ////}
            //        //File.Copy(origen1, destino1, true);
            //        ////if (File.Exists(lblDestino.Text))
            //        ////File.Delete(lblDestino.Text);

            //    }
            //}
            //else
            //{
            //    string importacion = "0";
            //    thisConnection.Open();
            //    SqlCommand cmndimportacion = thisConnection.CreateCommand();
            //    cmndimportacion.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
            //    try
            //    {
            //        importacion = Convert.ToString(cmndimportacion.ExecuteScalar().ToString().Trim());
            //    }
            //    catch
            //    {
            //        importacion = "";
            //    }

            //    cmndimportacion.Dispose();
            //    thisConnection.Close();

            //    if (importacion == "1")
            //    {
            //        /*if (txt_prov_clave.Text == "81" || txt_prov_clave.Text == "404" || txt_prov_clave.Text == "11" || txt_prov_clave.Text == "8851" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694" || txt_prov_clave.Text.Trim() == "9957" || txt_prov_clave.Text.Trim() == "833" || txt_prov_clave.Text.Trim() == "1550" || txt_prov_clave.Text.Trim() == "1761" || txt_prov_clave.Text.Trim() == "935" || txt_prov_clave.Text.Trim() == "6854" || txt_prov_clave.Text == "354" || txt_prov_clave.Text == "1250" || txt_prov_clave.Text == "1536" || txt_prov_clave.Text == "9758" || txt_prov_clave.Text == "1799" || txt_prov_clave.Text == "1244" || txt_prov_clave.Text == "191" || txt_prov_clave.Text == "812" || txt_prov_clave.Text == "8014" || txt_prov_clave.Text == "305" || txt_prov_clave.Text == "106")
            //        {*/
            //        if (txtFolioFiscal.Text != "")//SI SELECCIONO EL FOLIO FISCAL SUBIR EL XML
            //        {
            //            FileInfo archivo = new FileInfo(@lblPathXml.Text);
            //            string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
            //            FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists == true)
            //            {
            //                liq_copy.Delete();
            //            }

            //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists != true)
            //            {
            //                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
            //                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de XML " + file_name);

            //            }
            //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de XML Ruta: " + lblPathXml.Text);

            //            file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists == true)
            //            {
            //                liq_copy.Delete();
            //            }

            //            archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
            //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists != true)
            //            {
            //                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de PDF " + file_name);
            //            }

            //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de PDF");

            //        }

            //    }
            //    else
            //    {
            //        if (txtFolioFiscal.Text != "")
            //        {
            //            FileInfo archivo = new FileInfo(@lblPathXml.Text);
            //            string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
            //            FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists == true)
            //            {
            //                liq_copy.Delete();
            //            }
            //            if (liq_copy.Exists == false)
            //            {
            //                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + file_name + " No existe", "SISEMP");
            //            }
            //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists != true)
            //            {
            //                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
            //                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de XML " + file_name);
            //                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + lblPathXml.Text, "SISEMP");
            //            }

            //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de XML Ruta: " + lblPathXml.Text);

            //            file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists == true)
            //            {
            //                liq_copy.Delete();
            //            }
            //            if (liq_copy.Exists == false)
            //            {
            //                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + file_name + " No existe", "SISEMP");
            //            }
            //            archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
            //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

            //            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
            //            if (liq_copy.Exists != true)
            //            {
            //                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
            //                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de PDF " + file_name);
            //                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + lblPathXml.Text, "SISEMP");
            //            }

            //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de PDF");
            //        }
            //    }
            //}
            #endregion

            try
            {
                string var_chr_emp_clave = "";
                decimal var_dec_entd_cantidad = 0;
                string iepscargo = "";

                //INICIO DE EMPAQUES
                queries.Add("EMPAQUES");
                foreach (DataRow rw in tmgrdpro.Rows)
                {
                    if (rw["borrar"].ToString() == "1")
                        continue;
                    var_chr_emp_clave = rw["clave"].ToString();
                    var_dec_entd_cantidad = Convert.ToDecimal(rw["cantidad"].ToString());
                    iepscargo = (Convert.ToBoolean(rw["ieps"].ToString())) ? "T" : "F";

                    //traer datos de cat_empaques
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    query = "SELECT emp_costo, emp_existencia, alm_clave FROM tb_cat_empaques WHERE emp_clave = '" + rw["clave"].ToString() + "' ORDER BY emp_clave";
                    cmnd1.CommandText = query;
                    reader1 = cmnd1.ExecuteReader();
                    decimal costoactual = 0;
                    decimal existencia = 0;
                    string cvealm = "";
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        costoactual = Convert.ToDecimal(reader1.GetValue(0).ToString().Trim());
                        existencia = Convert.ToDecimal(reader1.GetValue(1).ToString().Trim());
                        cvealm = reader1.GetValue(2).ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();

                    decimal existencia_act = existencia + var_dec_entd_cantidad;

                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "INSERT INTO tb_existencias_emp (ent_sal, existencia_act, existencia_ant, almacen, movimiento, empaque) " +
                        "VALUES('" + var_int_sis_folios + "','" + existencia_act.ToString() + "','" + existencia.ToString() + "','" + glo_var_chr_alm_clave + "','EA', '" + var_chr_emp_clave + "')";
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();

                    //insercion en det_ent_temporal
                    //string sub2 = (glo_var_int_registro == 7) ? ", 'COM'" : (glo_var_int_registro == 9) ? ", 'GUA'" : ", ''";
                    cmnd1 = thisConnection.CreateCommand();
                    query = "INSERT INTO tb_det_ent_temporal(ent_folio, emp_clave, alm_clave, entd_cantidad, entd_costo, entd_adv, entd_dta, lote_ras, ent_tipo, ent_ieps, ent_ieps_importe, ent_gastoaduanal, " +
                        "ent_flete, ord_compra, ent_iva, existencia_ant, existencia_act)" +
                                " VALUES ('" + var_int_sis_folios + "', '" + var_chr_emp_clave + "', '" + glo_var_chr_alm_clave + "', '" + Convert.ToDecimal(var_dec_entd_cantidad).ToString() + "', '" + Convert.ToDecimal(rw["costo"].ToString()) + "'" +
                                ", '" + Convert.ToDecimal(rw["adv"].ToString()).ToString() + "', '" + Convert.ToDecimal(rw["dta"].ToString()).ToString() + "', '" + rw["lote"].ToString() + "', '" + subtipo + "', '" + iepscargo + "'," +
                                " '" + Convert.ToDecimal(rw["iepsimporte"].ToString()).ToString() + "', '" + Convert.ToDecimal(rw["gastoaduanal"].ToString()).ToString() + "', '" + Convert.ToDecimal(rw["flete"].ToString()).ToString() + "', '" + rw["oc"].ToString() + "', " +
                                "'" + Convert.ToDecimal(rw["iva"].ToString()).ToString() + "', '" + existencia.ToString() + "', '" + existencia_act.ToString() + "')";
                    cmnd1.CommandText = query;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();
                    //queries.Add(query);

                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + query);

                    if (glo_var_chr_alm_clave == "02")
                    {
                        if (rw["lote"].ToString() != "")
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            query = "INSERT INTO tb_det_ent_lote(ent_folio, emp_clave, entd_cantidad, estatus, lote_ras, alm_clave, tipo) " +
                                "VALUES('" + var_int_sis_folios + "', '" + var_chr_emp_clave + "', '" + Convert.ToDecimal(var_dec_entd_cantidad).ToString() + "', " +
                                "'A', '" + rw["lote"].ToString() + "', '" + glo_var_chr_alm_clave + "', 'NOR')";
                            cmnd1.CommandText = query;
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                            //queries.Add(query);

                            //BUSCAR SI YA SE REGISTRO EL LOTE
                            cmnd1 = thisConnection.CreateCommand();
                            query = "SELECT lote_ras FROM tb_det_lote_ras WHERE lote_ras = '" + rw["lote"].ToString() + "' AND empaque = '" + var_chr_emp_clave + "'";
                            cmnd1.CommandText = query;
                            reader1 = cmnd1.ExecuteReader();
                            bool found = false;
                            if (reader1.HasRows)
                            {
                                found = true;
                            }
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();

                            cmnd1 = thisConnection.CreateCommand();
                            if (found == false)
                            {
                                query = "INSERT INTO tb_det_lote_ras(lote_ras, cantidad, usado, empaque) " +
                                "VALUES('" + rw["lote"].ToString() + "', '" + Convert.ToDecimal(var_dec_entd_cantidad).ToString() + "', '0', '" + var_chr_emp_clave + "')";
                                //queries.Add(query);
                            }
                            else
                            {
                                query = "UPDATE tb_det_lote_ras set cantidad = cantidad + '" + Convert.ToDecimal(var_dec_entd_cantidad).ToString() + "' WHERE lote_ras = '" + rw["lote"].ToString() + "'";
                                //queries.Add(query);

                            }
                            cmnd1.CommandText = query;
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();

                            if (rw["lote"].ToString() != "")
                                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", txtFolio.Text, "INSERCION DE ENTRADA(OC): " + txtFolio.Text + " " + var_chr_emp_clave + " " + rw["lote"].ToString() + ", ALMACEN: " + glo_var_chr_alm_clave, "SISEMP");
                        }
                    }
                    //thisConnection.Close();
                    //COMENTADO 20/01/2021
                    //SE HARA EL CALCULO DE EMPAQUE TANTO PARA COSTO Y EXISTENCIA SEGUN EL ULTIMO CIERRE DETECTADO PARA ESE PRODUCTO
                    ////mandar al historico el registro
                    //decimal A = 0, B = 0, C = 0, D = 0, promedio = 0;
                    //A = Math.Abs(existencia * costoactual);
                    //B = var_dec_entd_cantidad * Convert.ToDecimal(rw["unitario"].ToString());
                    //C = Math.Abs(existencia) + var_dec_entd_cantidad;
                    //D = (C == 0) ? Convert.ToDecimal(rw["unitario"].ToString()) : (A + B) / C;
                    //C = existencia + var_dec_entd_cantidad;
                    //promedio = Math.Abs(D);

                    ////graba ultimo costo
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "SELECT emp_entrada_ultima FROM tb_cat_empaques WHERE emp_clave = '" + var_chr_emp_clave + "' AND emp_entrada_ultima < '" + dtpFecha.Text + "'";
                    //reader1 = cmnd1.ExecuteReader();
                    //if (reader1.HasRows)
                    //{
                    //    cmnd4 = thisConnection.CreateCommand();
                    //    ////cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "', emp_entrada_ultima = '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', emp_existencia = " + C.ToString() + " WHERE emp_clave = '" + var_chr_emp_clave + "' AND emp_entrada_ultima < '" + dtpFecha.Text + "'";
                    //    query = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "', emp_entrada_ultima = '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "' WHERE emp_clave = '" + var_chr_emp_clave + "' AND emp_entrada_ultima < '" + dtpFecha.Text + "'";
                    //    cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "', emp_entrada_ultima = '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "' WHERE emp_clave = '" + var_chr_emp_clave + "' AND emp_entrada_ultima < '" + dtpFecha.Text + "'";
                    //    cmnd4.ExecuteNonQuery();
                    //    cmnd4.Dispose();
                    //    //queries.Add(query);
                    //}
                    //else
                    //{
                    //    cmnd4 = thisConnection.CreateCommand();
                    //    //query = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "' WHERE emp_clave = '" + var_chr_emp_clave + "'";
                    //    cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "' WHERE emp_clave = '" + var_chr_emp_clave + "'";
                    //    cmnd4.ExecuteNonQuery();
                    //    cmnd4.Dispose();
                    //    //queries.Add(query);
                    //}
                    //reader1.Close();
                    //reader1.Dispose();
                    //cmnd1.Dispose();
                    //COMENTADO 20/01/2021




                    //thisConnection.Open();
                    string tip2 = "NOR";
                    string reg2 = "DEF";
                    string pro2 = (DDL_ent_procedencia.SelectedItem.ToString() == "Nacional") ? "NAL" : "IMP";
                    cmnd1 = thisConnection.CreateCommand();
                    query = "INSERT INTO tb_historico_recepcion(hrp_recibo, hrp_fecha, emp_clave, hrp_cantidad, hrp_estatus, hrp_tipo_recepcion, alm_clave, hrp_situacion," +
                        " hrp_costo, hrp_regimen, hrp_procedencia, hrp_sobrante, lote_ras, hrp_costo_prom, hrp_existencia, hrp_tipo, hrp_fechamov, hrp_importe, hrp_tipo_cambio, hrp_iva, hrp_costo_resp) VALUES ('" + var_int_sis_folios + "', " +
                        " '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + Convert.ToDecimal(var_dec_entd_cantidad.ToString()).ToString() + "', 'T', 'ENT', '" + glo_var_chr_alm_clave + "'," +
                        " '" + tip2 + "', '" + Convert.ToDecimal(rw["unitario"].ToString()).ToString() + "', " +
                        " '" + reg2 + "', '" + pro2 + "', '" + Convert.ToDecimal(var_dec_entd_cantidad.ToString()).ToString() + "', '" + rw["lote"].ToString() + "', '0', '0'," +
                        " '" + subtipo + "', '" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "', '" + Convert.ToDecimal(rw["costo"].ToString()) + "', '" + txt_sis_tipo_cambio.Text + "', '" + Convert.ToDecimal(rw["iva"].ToString()).ToString() + "'," +
                        " '" + Convert.ToDecimal(rw["unitario"].ToString()).ToString() + "')";
                    //query = "INSERT INTO tb_historico_recepcion(hrp_recibo, hrp_fecha, emp_clave, hrp_cantidad, hrp_estatus, hrp_tipo_recepcion, alm_clave, hrp_situacion," +
                    //    " hrp_costo, hrp_regimen, hrp_procedencia, hrp_sobrante, lote_ras, hrp_costo_prom, hrp_existencia, hrp_tipo, hrp_fechamov, hrp_importe, hrp_tipo_cambio, hrp_iva) VALUES ('" + var_int_sis_folios + "', " +
                    //    " '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + Convert.ToDecimal(var_dec_entd_cantidad.ToString()).ToString() + "', 'T', 'ENT', '" + glo_var_chr_alm_clave + "'," +
                    //    " '" + tip2 + "', '" + Convert.ToDecimal(rw["unitario"].ToString()).ToString() + "', " +
                    //    " '" + reg2 + "', '" + pro2 + "', '" + Convert.ToDecimal(var_dec_entd_cantidad.ToString()).ToString() + "', '" + rw["lote"].ToString() + "', '" + Convert.ToDecimal(promedio.ToString()) + "', '" + Convert.ToDecimal(C.ToString()).ToString() + "'," +
                    //    " '" + subtipo + "', '" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "', '" + Convert.ToDecimal(rw["costo"].ToString()) + "', '" + txt_sis_tipo_cambio.Text + "', '" + Convert.ToDecimal(rw["iva"].ToString()).ToString() + "')";
                    //((glo_var_int_registro == 7) ? "COM" : (glo_var_int_registro == 9) ? "GUA" : "")
                    cmnd1.CommandText = query;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();
                    //queries.Add(query);

                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + query);

                    //insercion en mstr_transaccion_emp
                    cmnd1 = thisConnection.CreateCommand();
                    query = "INSERT INTO tb_mstr_transaccion_emp(tra_folio, tra_fecha, emp_clave, tra_cantidad, tra_tipo_mov, tra_codigo_mov, tra_codigo_tra) VALUES " +
                        "('" + var_int_sis_folios + "', '" + DateTime.Now.ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + Convert.ToDecimal(var_dec_entd_cantidad.ToString()).ToString() + "'," +
                        " 'E', 'EN', 'A')";
                    cmnd1.CommandText = query;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();
                    //queries.Add(query);

                    thisConnection.Close();

                    //COMENTADO 24/02/2022
                    //decimal promedio = 0, C = 0;
                    //promedio = Convert.ToDecimal(recalculo_costos_empaque(var_chr_emp_clave, dtpFecha.Text));
                    //C = Convert.ToDecimal(recalculo_existencia_empaque(var_chr_emp_clave));

                    //thisConnection.Open();
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_costo_prom = '" + Convert.ToDecimal(promedio.ToString()) + "', hrp_existencia = '" + Convert.ToDecimal(C.ToString()).ToString() + "' " +
                    //    "WHERE hrp_recibo = '" + var_int_sis_folios + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND emp_clave = '" + var_chr_emp_clave + "' AND hrp_tipo_recepcion = 'ENT'";
                    //cmnd1.ExecuteNonQuery();
                    //cmnd1.Dispose();
                    //thisConnection.Close();
                    //FIN COMENTADO 24/02/2022

                }
                //queries.Add("FIN_EMPAQUES");
                //FIN DE EMPAQUES

                

                foreach (DataRow rw in dtdetalles.Rows)
                {
                    if (rw["orden_compra"].ToString() == "")
                        continue;

                    query = "INSERT INTO tb_det_entrada_oc(ent_folio, numero_oc, emp_clave, ent_cantidad, ent_importe, estatus, alm_clave," +
                        " ent_fecha, req_folio, ent_subtipo, ent_lote, ent_iva, ent_importe_orig, urgente_oc, cant_orig, resp, usu, machine) VALUES(" +
                        " '" + var_int_sis_folios + "', '" + rw["orden_compra"].ToString() + "', '" + rw["emp_clave"].ToString() + "'," +
                        " '" + Convert.ToDecimal(rw["cantidad"].ToString()).ToString() + "', '" + Convert.ToDecimal(rw["importe"].ToString()).ToString() + "'," +
                        " 'A', '" + glo_var_chr_alm_clave + "', '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '" + rw["requisicion"].ToString() + "'," +
                        " '" + subtipo + "', '" + rw["lote"].ToString() + "', '" + Convert.ToDecimal(rw["iva"].ToString()).ToString() + "'," +
                        " '" + Convert.ToDecimal(rw["importe_orig"].ToString()).ToString() + "', '" + rw["urgente"].ToString() + "'," +
                        " '" + Convert.ToDecimal(rw["cant_orig"].ToString()).ToString() + "', '" + rw["resp"].ToString() + "'," +
                        " '" + Utilerias.Class1.Usu_login.ToString().Trim() + "', '" + Environment.MachineName.ToString() + "')";

                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + query);

                    thisConnection.Open();
                    //insercion en det_entrada_oc
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = query;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();
                    //queries.Add(query);




                    //14/11/2018 cambio en la forma de saldar orden de compra
                    string sumatoria = "";
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT SUM(ent_cantidad) AS sumatoria FROM tb_det_entrada_oc WHERE numero_oc = '" + rw["orden_compra"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "' AND estatus = 'A'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        sumatoria = reader1["sumatoria"].ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();

                    if (sumatoria == "")
                        sumatoria = rw["cantidad"].ToString();

                    //actualizar det_ordenescompra
                    cmnd1 = thisConnection.CreateCommand();
                    //query = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc + '" + Convert.ToDecimal(rw["cantidad"].ToString()).ToString() + "' WHERE numero_oc = '" + rw["orden_compra"].ToString() + "' AND prod_clave = '" + rw["emp_clave"].ToString() + "'";
                    query = "UPDATE tb_det_ordenescompra SET surtido_oc = '" + sumatoria + "' WHERE numero_oc = '" + rw["orden_compra"].ToString() + "' AND prod_clave = '" + rw["emp_clave"].ToString() + "'";
                    cmnd1.CommandText = query;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();
                    //queries.Add(query);
                    //14/11/2018 fin cambio en la forma de saldar orden de compra

                    thisConnection.Close();

                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + query);

                    thisConnection.Open();



                    //Actualizar requisicion si hubiera empaques
                    if (rw["requisicion"].ToString() != "")
                    {
                        //buscar que existe el empaque dentro de det_requisiciones
                        if (DDLAlmacen1.SelectedIndex == 4)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            query = "SELECT req_xpedir, req_surtido FROM tb_det_requisiciones_enf_pap WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "' AND req_tipo = 'ENF'";
                            cmnd1.CommandText = query;
                            reader1 = cmnd1.ExecuteReader();
                            if (reader1.HasRows)
                            {
                                //actualizacion de surtido de requisicion
                                cmnd2 = thisConnection.CreateCommand();
                                query = "UPDATE tb_det_requisiciones_enf_pap SET req_surtido = req_surtido + '" + Convert.ToDecimal(rw["cantidad"].ToString()).ToString() + "', req_fecharec = '" + DateTime.Now.ToShortDateString() + "'," +
                                    " WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "' AND req_tipo = 'ENF'";
                                cmnd2.CommandText = query;
                                cmnd2.ExecuteNonQuery();
                                cmnd2.Dispose();
                                //queries.Add(query);
                            }
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();
                        }
                        else if (DDLAlmacen1.SelectedIndex == 5)
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            query = "SELECT req_xpedir, req_surtido FROM tb_det_requisiciones_enf_pap WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "'  AND req_tipo = 'PAP'";
                            cmnd1.CommandText = query;
                            reader1 = cmnd1.ExecuteReader();
                            if (reader1.HasRows)
                            {
                                //actualizacion de surtido de requisicion
                                cmnd2 = thisConnection.CreateCommand();
                                query = "UPDATE tb_det_requisiciones_enf_pap SET req_surtido = req_surtido + '" + Convert.ToDecimal(rw["cantidad"].ToString()).ToString() + "', req_fecharec = '" + DateTime.Now.ToShortDateString() + "'," +
                                    " WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "' AND req_tipo = 'PAP'";
                                cmnd2.CommandText = query;
                                cmnd2.ExecuteNonQuery();
                                cmnd2.Dispose();
                                //queries.Add(query);
                            }
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();
                        }
                        else
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            query = "SELECT req_xpedir, req_surtido FROM tb_det_requisiciones WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "'";
                            cmnd1.CommandText = query;
                            reader1 = cmnd1.ExecuteReader();
                            if (reader1.HasRows)
                            {
                                //actualizacion de surtido de requisicion
                                cmnd2 = thisConnection.CreateCommand();
                                query = "UPDATE tb_det_requisiciones SET req_surtido = req_surtido + '" + Convert.ToDecimal(rw["cantidad"].ToString()).ToString() + "', req_fecharec = '" + DateTime.Now.ToShortDateString() + "'," +
                                    " alm_clave = '" + glo_var_chr_alm_clave + "' WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "'";
                                cmnd2.CommandText = query;
                                cmnd2.ExecuteNonQuery();
                                cmnd2.Dispose();
                                //queries.Add(query);
                            }
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();
                        }

                    }

                    thisConnection.Close();
                }

                

                //XmlDocument doc = new XmlDocument();
                //XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
                //XmlElement root = doc.DocumentElement;
                //doc.InsertBefore(xmlDeclaration, root);
                //XmlElement element1 = doc.CreateElement(string.Empty, "consultas", string.Empty);
                //doc.AppendChild(element1);

                //XmlElement element2 = doc.CreateElement(string.Empty, "folio", string.Empty);
                //XmlText text2 = doc.CreateTextNode(txtFolio.Text);
                //element2.AppendChild(text2);
                //element1.AppendChild(element2);

                //element2 = doc.CreateElement(string.Empty, "almacen", string.Empty);
                //text2 = doc.CreateTextNode(glo_var_chr_alm_clave);
                //element2.AppendChild(text2);
                //element1.AppendChild(element2);

                //int i = 0;
                //foreach (string a in queries)
                //{
                //    i++;
                //    XmlElement elemento = doc.CreateElement(string.Empty, "query_" + i, string.Empty);
                //    XmlText texto = doc.CreateTextNode(a);
                //    elemento.AppendChild(texto);
                //    element1.AppendChild(elemento);


                //}



                #region docs upload 2
                if (lblXmlProv.Text == "0")//SI EL PROVEEDOR REQUIERE XML SUBIR ARCHIVO
                {
                    string importacion = "0";
                    thisConnection.Open();
                    SqlCommand cmndimportacion = thisConnection.CreateCommand();
                    cmndimportacion.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
                    try
                    {
                        importacion = Convert.ToString(cmndimportacion.ExecuteScalar().ToString().Trim());
                    }
                    catch
                    {
                        importacion = "";
                    }

                    cmndimportacion.Dispose();
                    thisConnection.Close();

                    if (importacion == "1")
                    {

                        /*if (txt_prov_clave.Text == "81" || txt_prov_clave.Text == "404" || txt_prov_clave.Text == "11" || txt_prov_clave.Text == "8851" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694" || txt_prov_clave.Text.Trim() == "9957" || txt_prov_clave.Text.Trim() == "833" || txt_prov_clave.Text.Trim() == "1550" || txt_prov_clave.Text.Trim() == "1761" || txt_prov_clave.Text.Trim() == "935" || txt_prov_clave.Text.Trim() == "6854" || txt_prov_clave.Text == "354" || txt_prov_clave.Text == "1250" || txt_prov_clave.Text == "1536" || txt_prov_clave.Text == "9758" || txt_prov_clave.Text == "1799" || txt_prov_clave.Text == "1244" || txt_prov_clave.Text == "191" || txt_prov_clave.Text == "812" || txt_prov_clave.Text == "8014" || txt_prov_clave.Text == "305" || txt_prov_clave.Text == "106")
                        {*/
                        if (txtFolioFiscal.Text != "")//SI SELECCIONO EL FOLIO FISCAL SUBIR EL XML
                        {
                            FileInfo archivo = new FileInfo(@lblPathXml.Text);
                            string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                            FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists == true)
                            {
                                encontrado = true;
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Encontrado Guardado(INS) de XML");
                                //liq_copy.Delete();
                            }
                            if (encontrado == false)
                            {
                                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                                liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                                if (liq_copy.Exists != true)
                                {
                                    archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". No encontrado Guardado 2 (INS) de XML " + file_name);

                                }
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Vuelto a subir Guardado(INS) de XML");
                            }

                            encontrado = false;

                            file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists == true)
                            {
                                encontrado = true;
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Encontrado Guardado 2 (INS) de PDF");
                                //liq_copy.Delete();
                            }

                            if (encontrado == false)
                            {
                                archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                                liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                                if (liq_copy.Exists != true)
                                {
                                    archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". No encontrado Guardado 2 (INS) de PDF " + file_name);
                                }

                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Vuelto a subir Guardado(INS) de PDF");
                            }
                            


                            ////SUBIR ARCHIVO xml
                            //string origen1 = @lblPathXml.Text;//nombre de archivo
                            //string destino1 = @lblDestino.Text + lblXmlFile.Text;
                            ////if (!File.Exists(origen1))
                            ////{
                            ////    //using (FileStream fs = File.Create(origen1)) { }

                            ////}
                            //File.Copy(origen1, destino1, true);
                            ////if (File.Exists(lblDestino.Text))
                            ////File.Delete(lblDestino.Text);


                            ////SUBIR PDF
                            //origen1 = @lblPathXml.Text.Replace(".xml", ".pdf");//nombre de archivo
                            //destino1 = @lblDestino.Text + lblXmlFile.Text.Replace(".xml", ".pdf");
                            ////if (!File.Exists(origen1))
                            ////{
                            ////    //using (FileStream fs = File.Create(origen1)) { }

                            ////}
                            //File.Copy(origen1, destino1, true);
                            ////if (File.Exists(lblDestino.Text))
                            ////File.Delete(lblDestino.Text);

                        }

                    }
                    else
                    {
                        if (txtFolioFiscal.Text != "")
                        {
                            FileInfo archivo = new FileInfo(@lblPathXml.Text);
                            string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                            FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists == true)
                            {
                                encontrado = true;
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Encontrado Guardado(INS) de XML");
                                //liq_copy.Delete();
                            }
                            if (encontrado == false)
                            {
                                if (liq_copy.Exists == false)
                                {
                                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + file_name + " No existe 2", "SISEMP");
                                }
                                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

                                liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                                if (liq_copy.Exists != true)
                                {
                                    archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
                                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". No encontrado Guardado 2 (INS) de XML " + file_name);
                                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Vuelto a subir 2 Ruta: " + lblPathXml.Text, "SISEMP");
                                }

                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ".Vuelto a subir Guardado(INS) de PDF Ruta: " + lblPathXml.Text);

                            }

                            encontrado = false;

                            file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists == true)
                            {
                                //liq_copy.Delete();
                                encontrado = true;
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Encontrado Guardado(INS) de PDF");
                            }
                            if (encontrado == false)
                            {
                                if (liq_copy.Exists == false)
                                {
                                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + file_name + " No existe 2", "SISEMP");
                                }
                                archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

                                liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                                if (liq_copy.Exists != true)
                                {
                                    archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
                                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". No encontrado Guardado 2 (INS) de PDF " + file_name);
                                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Vuelto a subir 2 Ruta: " + lblPathXml.Text, "SISEMP");
                                }

                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Vuelto a subirGuardado(INS) de PDF Ruta: " + lblPathXml.Text);
                            }
                            
                        }


                        //string origen1 = @lblPathXml.Text;
                        //string destino1 = @lblDestino.Text + lblXmlFile.Text;
                        ////if (!File.Exists(origen1))
                        ////{
                        ////    //using (FileStream fs = File.Create(origen1)) { }

                        ////}
                        //File.Copy(origen1, destino1, true);
                        ////if (File.Exists(lblDestino.Text))
                        ////File.Delete(lblDestino.Text);



                        ////SUBIR PDF
                        //origen1 = @lblPathXml.Text.Replace(".xml", ".pdf");//nombre de archivo
                        //destino1 = @lblDestino.Text + lblXmlFile.Text.Replace(".xml", ".pdf");
                        ////if (!File.Exists(origen1))
                        ////{
                        ////    //using (FileStream fs = File.Create(origen1)) { }

                        ////}
                        //File.Copy(origen1, destino1, true);
                        ////if (File.Exists(lblDestino.Text))
                        ////File.Delete(lblDestino.Text);

                    }
                }
                else
                {
                    string importacion = "0";
                    thisConnection.Open();
                    SqlCommand cmndimportacion = thisConnection.CreateCommand();
                    cmndimportacion.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
                    try
                    {
                        importacion = Convert.ToString(cmndimportacion.ExecuteScalar().ToString().Trim());
                    }
                    catch
                    {
                        importacion = "";
                    }

                    cmndimportacion.Dispose();
                    thisConnection.Close();

                    if (importacion == "1")
                    {
                        /*if (txt_prov_clave.Text == "81" || txt_prov_clave.Text == "404" || txt_prov_clave.Text == "11" || txt_prov_clave.Text == "8851" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694" || txt_prov_clave.Text.Trim() == "9957" || txt_prov_clave.Text.Trim() == "833" || txt_prov_clave.Text.Trim() == "1550" || txt_prov_clave.Text.Trim() == "1761" || txt_prov_clave.Text.Trim() == "935" || txt_prov_clave.Text.Trim() == "6854" || txt_prov_clave.Text == "354" || txt_prov_clave.Text == "1250" || txt_prov_clave.Text == "1536" || txt_prov_clave.Text == "9758" || txt_prov_clave.Text == "1799" || txt_prov_clave.Text == "1244" || txt_prov_clave.Text == "191" || txt_prov_clave.Text == "812" || txt_prov_clave.Text == "8014" || txt_prov_clave.Text == "305" || txt_prov_clave.Text == "106")
                        {*/
                        if (txtFolioFiscal.Text != "")//SI SELECCIONO EL FOLIO FISCAL SUBIR EL XML
                        {
                            FileInfo archivo = new FileInfo(@lblPathXml.Text);
                            string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                            FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists == true)
                            {
                                liq_copy.Delete();
                            }

                            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

                            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists != true)
                            {
                                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de XML " + file_name);

                            }
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de XML Ruta: " + lblPathXml.Text);

                            file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists == true)
                            {
                                liq_copy.Delete();
                            }

                            archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

                            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists != true)
                            {
                                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de PDF " + file_name);
                            }

                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de PDF");

                        }

                    }
                    else
                    {
                        if (txtFolioFiscal.Text != "")
                        {
                            FileInfo archivo = new FileInfo(@lblPathXml.Text);
                            string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                            FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists == true)
                            {
                                liq_copy.Delete();
                            }
                            if (liq_copy.Exists == false)
                            {
                                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + file_name + " No existe", "SISEMP");
                            }
                            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

                            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists != true)
                            {
                                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de XML " + file_name);
                                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + lblPathXml.Text, "SISEMP");
                            }

                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de XML Ruta: " + lblPathXml.Text);

                            file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists == true)
                            {
                                liq_copy.Delete();
                            }
                            if (liq_copy.Exists == false)
                            {
                                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + file_name + " No existe", "SISEMP");
                            }
                            archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);

                            liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            if (liq_copy.Exists != true)
                            {
                                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name, true);
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (INS) de PDF " + file_name);
                                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "Ruta: " + lblPathXml.Text, "SISEMP");
                            }

                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado(INS) de PDF");
                        }
                    }
                }
                #endregion


                string filelog2 = "C:\\SisEmpWeb\\eventlog.txt";
                using (StreamWriter sw = File.AppendText(filelog2))
                {
                    sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Inserción de entrada(oc): " + var_int_sis_folios + ", Almacen: " + glo_var_chr_alm_clave);
                    sw.Close();
                }

                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", var_int_sis_folios, "INSERCION DE ENTRADA(OC): " + var_int_sis_folios + ", ALMACEN: " + glo_var_chr_alm_clave, "SISEMP");

                //RECORRER RECALCULO DE EMPAQUES GUARDADOS
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + "RECALCULO DE EMPAQUES GUARDADOS COSTO Y EXISTENCIA " + DateTime.Now.ToString());

                foreach (DataRow rw in tmgrdpro.Rows)
                {
                    if (glo_var_chr_alm_clave == "13")
                    {
                        //PROCESO PARA ACTUALIZAR EL COSTO DEL CATALOGO DE EMPAQUES AL COSTO DE LA ENTRADA O REPARACION
                        string emp = rw["clave"].ToString();
                        string alm = glo_var_chr_alm_clave;
                        string cto = Convert.ToDecimal(rw["unitario"].ToString()).ToString();

                        string FECHA = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                        string NOM_COMPU = Environment.MachineName;
                        string NOM_USU = Utilerias.Class1.Usu_login.Trim();
                        string TIPO_MOV = "A";
                        string OP_CLAVE = "2.5";
                        string FOLIO = var_int_sis_folios;
                        string DETALLE = "Modificación de costo en catalogo de empaques por orden de reparacion";
                        string SISTEMA = "SISEMP";
                        string MOTIVO = "ENTRADA POR ORDEN DE REPARACION";
                        string OBS = "Empaque: " + rw["clave"].ToString() + " Costo: " + cto;

                        _consultaController.ActualizaCostoCatalogoEmpaquesOrdenReparacion(emp, alm, cto, FECHA, NOM_COMPU, NOM_USU, TIPO_MOV, OP_CLAVE, FOLIO, DETALLE, SISTEMA, MOTIVO, OBS);
                        decimal inv = Convert.ToDecimal(recalculo_existencia_empaque(emp));

                    }
                    else //PROCESO NORMAL DE LA ENTRADA
                    {
                        string e_m_p = rw["clave"].ToString();
                        decimal promedio = 0, C = 0;
                        promedio = Convert.ToDecimal(recalculo_costos_empaque(e_m_p, dtpFecha.Text));
                        C = Convert.ToDecimal(recalculo_existencia_empaque(e_m_p));

                        thisConnection.Open();
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_costo_prom = '" + Convert.ToDecimal(promedio.ToString()) + "', hrp_existencia = '" + Convert.ToDecimal(C.ToString()).ToString() + "' " +
                            "WHERE hrp_recibo = '" + var_int_sis_folios + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND emp_clave = '" + e_m_p + "' AND hrp_tipo_recepcion = 'ENT'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();
                        thisConnection.Close();

                        //if (glo_var_chr_alm_clave == "02")
                        //{
                        //    _consultaController.ActualizaCostoAguilares(e_m_p, glo_var_chr_alm_clave, promedio.ToString());
                        //}
                    }
                    
                }


                //GUARDADO DE PEDIMENTO Y FACTURA
                if (lblPedimento.Text == "1")
                {
                    CargarPedimentoFactura();
                    //string path_destino = "\\\\gabira1:81\\Pedimentos\\";
                    //foreach (string item in lsbPediFact.Items)
                    //{
                    //    string destino = path_destino + Path.GetFileName(item);
                    //}
                }

                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + "FIN RECALCULO DE EMPAQUES GUARDADOS COSTO Y EXISTENCIA " + DateTime.Now.ToString());

                //28/04/2026 AJUSTE DE ORDEN DE COMPRA SEGUN LA FACTURA
                if (glo_var_chr_alm_clave == "03" || glo_var_chr_alm_clave == "11" || glo_var_chr_alm_clave == "08")
                {
                    ModificaOrdenCompraEntrada(dtdetalles);
                }
                //28/04/2026 FIN AJUSTE DE ORDEN DE COMPRA SEGUN LA FACTURA

                MessageBox.Show("Datos guardados", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //imprmir_entrada_servidor();               

                if (MessageBox.Show("Imprimir Entrada en Impresora", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    btnImprimir_Click(null, null);

                //guarda_genera_pdf_servidor();

                if (glo_var_chr_alm_clave != "13")
                {
                    foreach (DataRow rw in tmgrdpro.Rows)
                    {
                        string e_m_p = rw["clave"].ToString();
                        recalculo_costos(e_m_p);
                        recalculo_existencia(e_m_p);
                        modificar_ultima_entrada(e_m_p);
                    }
                }
                

                if(glo_var_chr_alm_clave == "01")
                    verificar_oc_datos_sin_clave();

                btnLimpiar_Click(null, null);

                DDLAlmacen1_SelectionChangeCommitted(null, null);
            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de guardar la entrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.5", var_int_sis_folios + " " + sqlex.ToString() + " " + query, "SISEMP");
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }
        }

        

        public string verifica_salidas_trazabilidad()
        {
            string empaques_salida = "";
            string valor_encontrado = "";
            thisConnection.Open();
            foreach (DataRow ry in dtdetalles.Rows)
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT COUNT(lote_ras) AS lote_ras FROM tb_historico_recepcion WHERE lote_ras = '" + ry["lote"].ToString() + "' AND hrp_tipo_recepcion = 'SAL' AND hrp_estatus = 'T'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    valor_encontrado = reader1["lote_ras"].ToString().Trim();
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                if (valor_encontrado != "0")
                {
                    empaques_salida = ry["emp_clave"].ToString() + "\n";
                }
            }
            thisConnection.Close();

            return empaques_salida.TrimEnd('\n').ToString();
        }

        public void modificar2()
        { 
            //VALIDACION POR CIERRE DE MES
            if (Convert.ToDateTime(dtpFecha.Text) < Convert.ToDateTime(txtCierre.Text))
            {
                MessageBox.Show("No puede modificar la entrada por que es de fecha anterior al cierre de mes", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //return;
            }

            //validacion de seleccion de XML
            if (txtFolioFiscal.Text == "" && lblXmlProv.Text == "0")
            {
                MessageBox.Show("No ha seleccionado el XML, favor de verificarlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //MODIFICACION DE LA ENTRADA
            if (validacontroles() == false)
            {
                return;
            }

            //VALIDAR CAPTURA DE LOTE A EMPAQUES OBLIGATORIOS
            if (DDLAlmacen1.SelectedIndex == 1)
            {
                if (validar_empaque_lote() == true)
                {
                    return;
                }
            }

            if (validagrid_N_datatable() == true)
            {
                MessageBox.Show("Los empaques no pueden tener costos en cero, favor de verificarlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //VALIDAR QUE LOS DATOS A MODIFICAR NO TENGAN SALIDAS YA REGISTRADAS CON ESOS NUMEROS DE TRAZABILIDAD
            

            imprimirtxt();

            if (MessageBox.Show("¿Desea realizar algún cambio en la orden de entrada?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                return;
            }

            string importacion = "0";
            thisConnection.Open();
            SqlCommand cmndimportacion = thisConnection.CreateCommand();
            cmndimportacion.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
            try
            {
                importacion = Convert.ToString(cmndimportacion.ExecuteScalar().ToString().Trim());
            }
            catch
            {
                importacion = "";
            }

            cmndimportacion.Dispose();
            thisConnection.Close();

            if (importacion == "1")
            {

            /*if (txt_prov_clave.Text == "81" || txt_prov_clave.Text == "404" || txt_prov_clave.Text == "11" || txt_prov_clave.Text == "8851" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694" || txt_prov_clave.Text.Trim() == "9957" || txt_prov_clave.Text.Trim() == "833" || txt_prov_clave.Text.Trim() == "1550" || txt_prov_clave.Text.Trim() == "1761" || txt_prov_clave.Text.Trim() == "935" || txt_prov_clave.Text.Trim() == "6854" || txt_prov_clave.Text == "354" || txt_prov_clave.Text == "1250" || txt_prov_clave.Text == "1536" || txt_prov_clave.Text == "9758" || txt_prov_clave.Text == "1799" || txt_prov_clave.Text == "1244" || txt_prov_clave.Text == "191" || txt_prov_clave.Text == "812" || txt_prov_clave.Text == "8014" || txt_prov_clave.Text == "305" || txt_prov_clave.Text == "106")
            {*/
                if (lblValidaXML.Text == "1")
                {
                    //if (txtFolioFiscal.Text == "")
                    //{
                    //    //MessageBox.Show("No sa ha seleccionado el XML, favor de verificarlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //    MessageBox.Show("No se ha ingresado el Acuse de la factura, favor de verificarlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    //    return;
                    //}
                    ////if(lblXmlFile.Text == "-")
                    ////    lblXmlFile.Text = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                }
            }
            
            //string fac = "";
            //fac = txt_ent_factura.Text;
            //if (fac == "")
            //{
            //    if (MessageBox.Show("No ha ingresado el número de factura, ¿Desea ingresarlo?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //        return;
            //    else
            //    {
            //        txt_ent_factura.Text = "S/N";
            //    }
            //}

            //15/01/2025 CREACION DE ARCHIVO DE ENTRADA
            

            string importacionx = "0";
            thisConnection.Open();

            guarda_genera_pdf_servidor();

            SqlCommand cmndimportacionx = thisConnection.CreateCommand();
            cmndimportacionx.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
            try
            {
                importacionx = Convert.ToString(cmndimportacionx.ExecuteScalar().ToString().Trim());
            }
            catch
            {
                importacionx = "";
            }

            cmndimportacionx.Dispose();
            thisConnection.Close();
            if (importacionx == "1")
            {

                /*if (txt_prov_clave.Text == "81" || txt_prov_clave.Text == "404" || txt_prov_clave.Text == "11" || txt_prov_clave.Text == "8851" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694" || txt_prov_clave.Text.Trim() == "9957" || txt_prov_clave.Text.Trim() == "833" || txt_prov_clave.Text.Trim() == "1550" || txt_prov_clave.Text.Trim() == "1761" || txt_prov_clave.Text.Trim() == "935" || txt_prov_clave.Text.Trim() == "6854" || txt_prov_clave.Text == "354" || txt_prov_clave.Text == "1250" || txt_prov_clave.Text == "1536" || txt_prov_clave.Text == "9758" || txt_prov_clave.Text == "1799" || txt_prov_clave.Text == "1244" || txt_prov_clave.Text == "191" || txt_prov_clave.Text == "812" || txt_prov_clave.Text == "8014" || txt_prov_clave.Text == "305" || txt_prov_clave.Text == "106")
                {*/
                //subir archivo XML
                //if (lblValidaXML.Text == "1")
                //{
                if (lblPathXml.Text != "-")
                {
                    FileInfo archivo = new FileInfo(@lblPathXml.Text);
                    string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                    FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    if (liq_copy.Exists == true)
                    {
                        liq_copy.Delete();
                    }
                    archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                    liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    if (liq_copy.Exists != true)
                    {
                        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (MOD) de XML " + file_name);
                    }

                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ".  Guardado(MOD) de archivo XML " + file_name);

                    file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                    liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    if (liq_copy.Exists == true)
                    {
                        liq_copy.Delete();
                    }
                    archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                    archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                    liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    if (liq_copy.Exists != true)
                    {
                        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (MOD) de PDF " + file_name);
                    }

                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ".  Guardado(MOD) de archivo PDF " + file_name);
                }

            }
            else
            {
                if (lblPathXml.Text != "-")
                {
                    FileInfo archivo = new FileInfo(@lblPathXml.Text);
                    string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                    FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    if (liq_copy.Exists == true)
                    {
                        liq_copy.Delete();
                    }
                    archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                    liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    if (liq_copy.Exists != true)
                    {
                        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (MOD) de XML " + file_name);
                    }

                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ".  Guardado(MOD) de archivo XML " + file_name);

                    file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                    liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    if (liq_copy.Exists == true)
                    {
                        liq_copy.Delete();
                    }
                    archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                    archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                    liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    if (liq_copy.Exists != true)
                    {
                        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (MOD) de PDF " + file_name);
                    }

                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ".  Guardado(MOD) de archivo PDF " + file_name);
                }
            }
            //FIN 15/01/2025
            

            string query = "";
            string tipo = "";
            try
            {
                thisConnection.Open();
                tipo = "NOR";
                string reg = "";
                reg = "DEF";
                string pro = (DDL_ent_procedencia.SelectedIndex == 0) ? "NAL" : "IMP";

                cmnd1 = thisConnection.CreateCommand();
                //ACTUALIZACION DE ENTRADA
                query = "UPDATE tb_mstr_entradas SET ent_tipo = '" + tipo + "', ent_fecha = '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', ent_elaboro = '" + DDLElaboro.SelectedItem.ToString() + "'," +
                    " ent_autorizo = '" + DDLAutorizo.SelectedItem.ToString() + "', ent_observaciones = '" + txtObservaciones.Text + "', alm_clave = '" + glo_var_chr_alm_clave + "', prov_clave = '" + txt_prov_clave.Text + "'," +
                    " ent_iva = '" + Convert.ToDecimal(txt_ent_iva.Text) + "', ent_monto = '" + Convert.ToDecimal(txt_total.Text) + "', ent_factura = '" + txt_ent_factura.Text + "', sis_tipo_cambio = '" + Convert.ToDecimal(txt_sis_tipo_cambio.Text) + "'," +
                    " ent_regimen = '" + reg + "', ent_procedencia = '" + pro + "', ent_pedimento = '" + txt_ent_pedimento.Text + "', ent_subtipo = '" + subtipo + "', ent_gastoaduanal = '" + Convert.ToDecimal(txt_gto_aduanal.Text) + "'," +
                    " ent_iepstotal = '" + Convert.ToDecimal(txt_ent_ieps.Text) + "', ent_flete = '" + Convert.ToDecimal(txtFlete.Text) + "', ent_flete_imp = '" + Convert.ToDecimal(txtFleteImp.Text) + "', ent_sin_fact = '" + ((chkFactura.Checked == true) ? "1" : "0") + "'," +
                    " ent_fechafact = '" + txtFechaFact.Text + "', ent_foliofiscal = '" + txtFolioFiscal.Text + "', ent_mod = '" + DateTime.Now.ToString() + "'," +
                    " descuento = '" + Convert.ToDecimal(txtDesc.Text).ToString() + "', ret_iva = '" + Convert.ToDecimal(txtRetIva.Text).ToString() + "', ret_isr = '" + Convert.ToDecimal(txtIsr.Text).ToString() + "', cedular = '" + Convert.ToDecimal(txtCedular.Text).ToString() + "'," +
                    " ent_desc_pp = '" + lblPorcentajeDesc.Text + "', ent_desc_im = '" + lblImporteDescuento.Text + "' WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_subtipo = '" + subtipo + "'";
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();

                if (txt_prov_clave.Text == "81")
                {
                    //if (lblValidaXML.Text == "1")
                    //{
                        //GUARDAR DATOS EN TB_DETOC_FACTURAS
                        //if (lblPortal.Text == "-")
                        //{
                        //    query = "INSERT INTO tb_det_oc_facturas(numero_oc, factura, fecha_factura, monto, fecha_registro, folio_fiscal, tipo, estatus, portal, estatus_portal) " +
                        //    "VALUES('" + txtFolio.Text.TrimStart('0').ToString() + "', '" + txt_ent_factura.Text + "', '" + txtFechaFact.Text + "', '" + lblTotal.Text + "', '" + dtpFecha.Text + "', " +
                        //    "'" + txtFolioFiscal.Text + "', '" + DDLAlmacen1.Text.ToUpper() + "', 'A', '', '0')";
                        //}
                        //else
                        //{
                    bool fnd_reg = false;
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT numero_oc FROM tb_det_oc_facturas WHERE numero_oc = '" + txtFolio.Text.TrimStart('0').ToString() + "' AND folio_fiscal = '" + txtFolioFiscal.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        fnd_reg = true;
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();

                    if (fnd_reg == false)
                    {
                        cmnd1 = thisConnection.CreateCommand();
                        query = "INSERT INTO tb_det_oc_facturas(numero_oc, factura, fecha_factura, monto, fecha_registro, folio_fiscal, tipo, estatus, portal, estatus_portal) " +
                        "VALUES('" + txtFolio.Text.TrimStart('0').ToString() + "', '" + txt_ent_factura.Text + "', '" + txtFechaFact.Text + "', '" + lblTotal.Text + "', '" + dtpFecha.Text + "', " +
                        "'" + txtFolioFiscal.Text + "', '" + DDLAlmacen1.Text.ToUpper() + "', 'A', '" + txtAcuse.Text + "', '1')";
                        //}
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();
                    }
                        
                        
                    //}
                }

                //INSTRUCCION AGREGADA PARA ACTUALIZAR LA INFORMACION DE TB_DET_OC_FACTURAS AGREGADA EN COVID 08/12/2020
                cmnd4 = thisConnection.CreateCommand();
                cmnd4.CommandText = "UPDATE tb_det_oc_facturas SET factura = '" + txt_ent_factura.Text + "', fecha_factura = '" + txtFechaFact.Text + "', " +
                        "folio_fiscal = '" + txtFolioFiscal.Text + "', monto = '" + lblTotal.Text + "', portal = '" + txtAcuse.Text + "', estatus_portal = '1' " +
                        "WHERE numero_oc = '" + txtFolio.Text.TrimStart('0').ToString() + "' AND tipo = '" + DDLAlmacen1.Text.ToUpper() + "'";
                cmnd4.ExecuteNonQuery();
                cmnd4.Dispose();

                Utilerias.Class1.registrar_movimiento3(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "AF", "2.5", txtFolio.Text.Trim(), "ACTUALIZACION DE FACTURA", "SISEMP", "", "fACTURA REGISTRADA DESDE FORM1 MODULO EDITAR");

                
                           

                //***********TERMINO DE INSTRUCCION AGREGADA EN ERA COVID************************************************

                //ELIMINAR REGISTROS DEL DETALLE DE ENTRADAS
                cmnd1 = thisConnection.CreateCommand();
                query = "DELETE FROM tb_det_ent_temporal WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_tipo = '" + subtipo + "'";
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();

                //ELIMINAR REGISTRO DE HISTORICO
                cmnd1 = thisConnection.CreateCommand();
                query = "DELETE FROM tb_historico_recepcion WHERE hrp_recibo = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND hrp_tipo_recepcion = 'ENT' AND hrp_tipo = '" + subtipo + "'";
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();

                //ELIMINAR REGISTRO DEL DETALLE DE ENTRADAS POR OC
                cmnd1 = thisConnection.CreateCommand();
                query = "DELETE FROM tb_det_entrada_oc WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_subtipo = '" + subtipo + "'";
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();

                //ELIMINAR REGISTROS DE TB_DET_ENT_LOTE

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT lote_ras FROM tb_det_ent_lote WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND tipo = 'NOR'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        cmnd2 = thisConnection.CreateCommand();
                        query = "DELETE FROM tb_det_lote_ras WHERE lote_ras = '" + reader1["lote_ras"].ToString().Trim() + "'";
                        cmnd2.CommandText = query;
                        cmnd2.ExecuteNonQuery();
                        cmnd2.Dispose();
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                cmnd1 = thisConnection.CreateCommand();
                query = "DELETE FROM tb_det_ent_lote WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND tipo = 'NOR'";
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();



                //RESTAR CANTIDADES A LA EXISTENCIA DE EMPAQUES
                
                foreach (DataRow re in dtoriginales.Rows)
                {
                    //Se va a cambiar por el recalculo
                    //cmnd3.CommandText = "UPDATE tb_cat_empaques SET emp_existencia = emp_existencia - '" + re["cantidad"].ToString() + "' WHERE emp_clave = '" + re["emp_clave"].ToString() + "'";
                    //cmnd3.ExecuteNonQuery();

                    //Recalculo de surtido
                    cmnd3 = thisConnection.CreateCommand();
                    cmnd3.CommandText = "SELECT ISNULL(SUM(ent_cantidad), 0) AS cantidad FROM tb_det_entrada_oc WHERE numero_oc = '" + re["orden_compra"].ToString() + "' AND emp_clave = '" + re["emp_clave"].ToString() + "' " +
                        "AND estatus = 'A'";
                    reader3 = cmnd3.ExecuteReader();
                    string canti = "";
                    if (reader3.HasRows)
                    {
                        reader3.Read();
                        canti = reader3["cantidad"].ToString().Trim();
                    }
                    reader3.Close();
                    reader3.Dispose();
                    cmnd3.Dispose();

                    cmnd3 = thisConnection.CreateCommand();
                    cmnd3.CommandText = "UPDATE tb_det_ordenescompra SET surtido_oc = '" + canti + "' WHERE prod_clave = '" + re["emp_clave"].ToString() + "' AND numero_oc = '" + re["orden_compra"].ToString() + "'";
                    cmnd3.ExecuteNonQuery();

                    cmnd3 = thisConnection.CreateCommand();
                    if (DDLAlmacen1.SelectedIndex == 4)
                        cmnd3.CommandText = "UPDATE tb_det_requisiciones_enf_pap SET req_surtido = req_surtido - '" + re["cantidad"].ToString() + "' WHERE emp_clave = '" + re["emp_clave"].ToString() + "' AND req_folio = '" + re["requisicion"].ToString() + "' AND req_tipo = 'ENF'";
                    else if (DDLAlmacen1.SelectedIndex == 5)
                        cmnd3.CommandText = "UPDATE tb_det_requisiciones_enf_pap SET req_surtido = req_surtido - '" + re["cantidad"].ToString() + "' WHERE emp_clave = '" + re["emp_clave"].ToString() + "' AND req_folio = '" + re["requisicion"].ToString() + "' AND req_tipo = 'ENF'";
                    else
                    {
                        try
                        {
                            decimal valor = Convert.ToDecimal(re["requisicion"].ToString());
                            if (valor > 0)
                            {
                                cmnd3.CommandText = "UPDATE tb_det_requisiciones SET req_surtido = req_surtido - '" + re["cantidad"].ToString() + "' WHERE emp_clave = '" + re["emp_clave"].ToString() + "' AND req_folio = '" + re["requisicion"].ToString() + "'";
                                //cmnd3.ExecuteNonQuery();
                            }
                                
                        }
                        catch (Exception ex)
                        { 

                        }   
                    }
                    if(cmnd3.CommandText != "")
                        cmnd3.ExecuteNonQuery();
                }
                cmnd3.Dispose();
                
                thisConnection.Close();

                
                

                try
                {
                    string var_chr_emp_clave = "";
                    decimal var_dec_entd_cantidad = 0;
                    string iepscargo = "";
                    foreach (DataRow rw in tmgrdpro.Rows)
                    {
                        if (this.glo_var_chr_alm_clave == "01" || this.glo_var_chr_alm_clave == "02")
                            //form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " Empaque: " + rw["clave"].ToString() + ", Cantidad: " + rw["cantidad"].ToString() + ", Unitario: " + rw["unitario"].ToString() + ", Costo: " + rw["costo"].ToString());
                        if (rw["borrar"].ToString() == "1")
                            continue;
                        var_chr_emp_clave = rw["clave"].ToString();
                        var_dec_entd_cantidad = Convert.ToDecimal(rw["cantidad"].ToString());
                        iepscargo = (Convert.ToBoolean(rw["ieps"].ToString())) ? "T" : "F";

                        //traer datos de cat_empaques
                        thisConnection.Open();
                        cmnd1 = thisConnection.CreateCommand();
                        query = "SELECT emp_costo, emp_existencia, alm_clave FROM tb_cat_empaques WHERE emp_clave = '" + rw["clave"].ToString() + "' ORDER BY emp_clave";
                        cmnd1.CommandText = query;
                        reader1 = cmnd1.ExecuteReader();
                        decimal costoactual = 0;
                        decimal existencia = 0;
                        string cvealm = "";
                        if (reader1.HasRows)
                        {
                            reader1.Read();
                            costoactual = Convert.ToDecimal(reader1.GetValue(0).ToString().Trim());
                            existencia = Convert.ToDecimal(reader1.GetValue(1).ToString().Trim());
                            cvealm = reader1.GetValue(2).ToString().Trim();
                        }
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();

                        //decimal existencia_act = 0;
                        //existencia_act = existencia + var_dec_entd_cantidad;

                        //cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = "INSERT INTO tb_existencias_emp (ent_sal, existencia_act, existencia_ant, almacen, movimiento) " +
                        //    "VALUES('" + var_int_sis_folios + "','" + existencia_act.ToString() + "','" + existencia.ToString() + "','" + glo_var_chr_alm_clave + "','A')";
                        //cmnd1.ExecuteNonQuery();
                        //cmnd1.Dispose();

                        //02/07/2020

                        ////CONSULTAR DATOS DE TRAZABILIDAD
                        //query = "SELECT id_lote, lote_ras, cantidad, usado FROM tb_det_ent_lote WHERE ent_folio = '" + txtFolio.Text + "' AND tipo = 'NOR'";
                        //adapter1 = new SqlDataAdapter(query, thisConnection);
                        //adapter1.Fill(ds, "tb_det_lote_ras");

                        ////VERIFICAR QUE EL EMPAQUE EXISTA SINO INSERTARLO
                        //bool encontrado = false;
                        //foreach (DataRow rx in ds.Tables["tb_det_ent_lote"].Select("emp_clave = '" + var_chr_emp_clave + "'"))
                        //{
                        //    encontrado = true;
                        //}
                        //if (encontrado == false)//INGRESAR REGISTRO SI EMPAQUE NO EXISTE EN LA BASE DE DATOS
                        //{
                        //    query = "INSERT INTO tb_det_ent_lote(ent_folio, emp_clave, entd_cantidad, estatus, lote_ras, alm_clave, tipo) " +
                        //        "VALUES('" + txtFolio.Text + "', '" + var_chr_emp_clave + "', '" + var_dec_entd_cantidad.ToString() + "', 'A', '" + rw["lote"].ToString() + "', " +
                        //        "'" + glo_var_chr_alm_clave + "', 'NOR')";
                        //    cmnd1 = thisConnection.CreateCommand();
                        //    cmnd1.CommandText = query;
                        //    cmnd1.ExecuteNonQuery();
                        //}
                        ////VERIFICAR SI CAMBIO EL LOTE DE TRAZABILIDAD O LA CANTIDAD
                        //cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = "SELECT entd_cantidad, lote_ras FROM tb_det_ent_lote WHERE ent_folio = '" + txtFolio.Text + "' AND emp_clave = '" + var_chr_emp_clave + "'";
                        //reader1 = cmnd1.ExecuteReader();
                        //if (reader1.HasRows)
                        //{
                        //    while (reader1.Read())
                        //    {
                        //        if (Convert.ToDecimal(reader1["entd_cantidad"].ToString().Trim()) != var_dec_entd_cantidad && reader1["lote_ras"].ToString().Trim() == rw["lote"].ToString())
                        //        {
                        //            cmnd2 = thisConnection.CreateCommand();
                        //            cmnd2.CommandText = "UPDATE tb_det_ent_lote SET entd_cantidad = '" + var_dec_entd_cantidad.ToString() + "' WHERE ent_folio = '" + txtFolio.Text + "' AND emp_clave = '" + var_chr_emp_clave + "'";
                        //            cmnd2.ExecuteNonQuery();
                        //            cmnd2.Dispose();

                        //            cmnd2 = thisConnection.CreateCommand();
                        //            cmnd2.CommandText = "UPDATE tb_det_lote_ras SET cantidad = '" + var_dec_entd_cantidad.ToString() + "' WHERE lote_ras = '" + rw["lote"].ToString() + "' AND empaque = '" + var_chr_emp_clave + "'";
                        //            cmnd2.ExecuteNonQuery();
                        //            cmnd2.Dispose();
                        //        }
                        //        if (Convert.ToDecimal(reader1["entd_cantidad"].ToString().Trim()) == var_dec_entd_cantidad && reader1["lote_ras"].ToString().Trim() != rw["lote"].ToString())
                        //        {
                        //            cmnd2 = thisConnection.CreateCommand();
                        //            cmnd2.CommandText = "UPDATE tb_det_ent_lote SET lote_ras = '" + rw["lote"].ToString() + "' WHERE ent_folio = '" + txtFolio.Text + "'  AND emp_clave = '" + var_chr_emp_clave + "'";
                        //            cmnd2.ExecuteNonQuery();
                        //            cmnd2.Dispose();

                        //            cmnd2 = thisConnection.CreateCommand();
                        //            cmnd2.CommandText = "UPDATE tb_det_lote_ras SET lote_ras = '" + rw["lote"].ToString() + "' WHERE ent_folio = '" + txtFolio.Text + "'  AND empaque = '" + var_chr_emp_clave + "'";
                        //            cmnd2.ExecuteNonQuery();
                        //            cmnd2.Dispose();
                        //        }
                        //        if (Convert.ToDecimal(reader1["entd_cantidad"].ToString().Trim()) != var_dec_entd_cantidad && reader1["lote_ras"].ToString().Trim() != rw["lote"].ToString())
                        //        {
                        //            cmnd2 = thisConnection.CreateCommand();
                        //            cmnd2.CommandText = "UPDATE tb_det_ent_lote SET lote_ras = '" + var_dec_entd_cantidad.ToString() + "', entd_cantidad = '" + var_dec_entd_cantidad.ToString() + "' WHERE ent_folio = '" + txtFolio.Text + "'  AND emp_clave = '" + var_chr_emp_clave + "'";
                        //            cmnd2.ExecuteNonQuery();
                        //            cmnd2.Dispose();

                        //            cmnd2 = thisConnection.CreateCommand();
                        //            cmnd2.CommandText = "UPDATE tb_det_lote_ras SET cantidad = '" + var_dec_entd_cantidad.ToString() + "' WHERE lote_ras = '" + rw["lote"].ToString() + "' AND empaque = '" + var_chr_emp_clave + "'";
                        //            cmnd2.ExecuteNonQuery();
                        //            cmnd2.Dispose();
                        //        }
                        //    }
                        //}
                        //else //CAMBIO EL LOTE DE TRAZABILIDAD
                        //{
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = "SELECT entd_cantidad FROM tb_det_ent_lote WHERE ent_folio = '" + txtFolio.Text + "' AND lote_ras = '" + rw["lote"].ToString() + "'";
                        //}

                        //insercion en det_ent_temporal
                        //string sub2 = (glo_var_int_registro == 7) ? ", 'COM'" : (glo_var_int_registro == 9) ? ", 'GUA'" : ", ''";
                        cmnd1 = thisConnection.CreateCommand();
                        query = "INSERT INTO tb_det_ent_temporal(ent_folio, emp_clave, alm_clave, entd_cantidad, entd_costo, entd_adv, entd_dta, lote_ras, ent_tipo, ent_ieps, ent_ieps_importe, ent_gastoaduanal, ent_flete, ord_compra)" +
                                    " VALUES ('" + txtFolio.Text + "', '" + var_chr_emp_clave + "', '" + glo_var_chr_alm_clave + "', '" + var_dec_entd_cantidad.ToString() + "', '" + Convert.ToDecimal(rw["costo"].ToString()).ToString() + "'" +
                                    ", '" + rw["adv"].ToString() + "', '" + rw["dta"].ToString() + "', '" + rw["lote"].ToString() + "', '" + subtipo + "', '" + iepscargo + "'," +
                                    " '" + rw["iepsimporte"].ToString() + "', '" + rw["gastoaduanal"].ToString() + "', '" + rw["flete"].ToString() + "', '" + rw["oc"].ToString() + "')";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        if (glo_var_chr_alm_clave == "02")
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            query = "INSERT INTO tb_det_ent_lote(ent_folio, emp_clave, entd_cantidad, estatus, lote_ras, alm_clave, tipo) " +
                                "VALUES('" + txtFolio.Text + "', '" + var_chr_emp_clave + "', '" + Convert.ToDecimal(var_dec_entd_cantidad).ToString() + "', " +
                                "'A', '" + rw["lote"].ToString() + "', '" + glo_var_chr_alm_clave + "', 'NOR')";
                            cmnd1.CommandText = query;
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();

                            //BUSCAR SI YA SE REGISTRO EL LOTE
                            cmnd1 = thisConnection.CreateCommand();
                            query = "SELECT lote_ras FROM tb_det_lote_ras WHERE lote_ras = '" + rw["lote"].ToString() + "' AND empaque = '" + var_chr_emp_clave + "'";
                            cmnd1.CommandText = query;
                            reader1 = cmnd1.ExecuteReader();
                            bool found = false;
                            if (reader1.HasRows)
                            {
                                found = true;
                            }
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();

                            cmnd1 = thisConnection.CreateCommand();
                            if (found == false)
                            {
                                query = "INSERT INTO tb_det_lote_ras(lote_ras, cantidad, usado, empaque) " +
                                "VALUES('" + rw["lote"].ToString() + "', '" + Convert.ToDecimal(var_dec_entd_cantidad).ToString() + "', '0', '" + var_chr_emp_clave + "')";
                            }
                            else
                            {
                                query = "UPDATE tb_det_lote_ras set cantidad = cantidad + '" + Convert.ToDecimal(var_dec_entd_cantidad).ToString() + "' WHERE lote_ras = '" + rw["lote"].ToString() + "'";

                            }
                            cmnd1.CommandText = query;
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();

                            if (rw["lote"].ToString() != "")
                                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "2.5", txtFolio.Text, "MODIFICACION DE ENTRADA(OC): " + txtFolio.Text + " " + var_chr_emp_clave + " " + rw["lote"].ToString() + ", ALMACEN: " + glo_var_chr_alm_clave, "SISEMP");

                        }
                        //thisConnection.Close();
                        //COMENTADO 20/01/2021
                        //SE HARA EL CALCULO DE EMPAQUE TANTO PARA COSTO Y EXISTENCIA SEGUN EL ULTIMO CIERRE DETECTADO PARA ESE PRODUCTO
                        ////mandar al historico el registro
                        //decimal A = 0, B = 0, C = 0, D = 0, promedio = 0;
                        //A = Math.Abs(existencia * costoactual);
                        //B = var_dec_entd_cantidad * Convert.ToDecimal(rw["unitario"].ToString());
                        //C = Math.Abs(existencia) + var_dec_entd_cantidad;
                        //D = (C == 0) ? Convert.ToDecimal(rw["unitario"].ToString()) : (A + B) / C;
                        //C = existencia + var_dec_entd_cantidad;
                        //promedio = Math.Abs(D);

                        ////graba ultimo costo
                        //cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = "SELECT emp_entrada_ultima FROM tb_cat_empaques WHERE emp_clave = '" + var_chr_emp_clave + "' AND emp_entrada_ultima < '" + dtpFecha.Text + "'";
                        //reader1 = cmnd1.ExecuteReader();
                        //if (reader1.HasRows)
                        //{
                        //    cmnd4 = thisConnection.CreateCommand();
                        //    cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "', emp_entrada_ultima = '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', emp_existencia = " + C.ToString() + " WHERE emp_clave = '" + var_chr_emp_clave + "' AND emp_entrada_ultima < '" + dtpFecha.Text + "'";
                        //    cmnd4.ExecuteNonQuery();
                        //    cmnd4.Dispose();
                        //}
                        //else
                        //{
                        //    cmnd4 = thisConnection.CreateCommand();
                        //    cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "', emp_existencia = " + C.ToString() + " WHERE emp_clave = '" + var_chr_emp_clave + "'";
                        //    cmnd4.ExecuteNonQuery();
                        //    cmnd4.Dispose();
                        //}
                        //reader1.Close();
                        //reader1.Dispose();
                        //cmnd1.Dispose();
                        //COMENTADO 20/01/2021

                        //decimal promedio = 0, C = 0;
                        //promedio = Convert.ToDecimal(recalculo_costos_empaque(var_chr_emp_clave, dtpFecha.Text));
                        //C = Convert.ToDecimal(recalculo_existencia_empaque(var_chr_emp_clave));

                        //thisConnection.Open();
                        string tip2 = "NOR";
                        string reg2 = "DEF";
                        string pro2 = (DDL_ent_procedencia.SelectedItem.ToString() == "Nacional") ? "NAL" : "IMP";
                        cmnd1 = thisConnection.CreateCommand();
                        //query = "INSERT INTO tb_historico_recepcion(hrp_recibo, hrp_fecha, emp_clave, hrp_cantidad, hrp_estatus, hrp_tipo_recepcion, alm_clave, hrp_situacion," +
                        //    " hrp_costo, hrp_regimen, hrp_procedencia, hrp_sobrante, lote_ras, hrp_costo_prom, hrp_existencia, hrp_tipo, hrp_fechamov, hrp_importe, hrp_tipo_cambio) VALUES ('" + txtFolio.Text + "', " +
                        //    " '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + var_dec_entd_cantidad.ToString() + "', 'T', 'ENT', '" + glo_var_chr_alm_clave + "'," +
                        //    " '" + tip2 + "', '" + rw["unitario"].ToString() + "', " +
                        //    " '" + reg2 + "', '" + pro2 + "', '" + var_dec_entd_cantidad.ToString() + "', '" + rw["lote"].ToString() + "', '" + promedio.ToString() + "', '" + C.ToString() + "'," +
                        //    " '" + subtipo + "', '" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "', '" + Convert.ToDecimal(rw["costo"].ToString()).ToString() + "', '" + txt_sis_tipo_cambio.Text + "')";
                        query = "INSERT INTO tb_historico_recepcion(hrp_recibo, hrp_fecha, emp_clave, hrp_cantidad, hrp_estatus, hrp_tipo_recepcion, alm_clave, hrp_situacion," +
                            " hrp_costo, hrp_regimen, hrp_procedencia, hrp_sobrante, lote_ras, hrp_costo_prom, hrp_existencia, hrp_tipo, hrp_fechamov, hrp_importe, hrp_tipo_cambio, hrp_costo_resp) VALUES ('" + txtFolio.Text + "', " +
                            " '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + var_dec_entd_cantidad.ToString() + "', 'T', 'ENT', '" + glo_var_chr_alm_clave + "'," +
                            " '" + tip2 + "', '" + rw["unitario"].ToString() + "', " +
                            " '" + reg2 + "', '" + pro2 + "', '" + var_dec_entd_cantidad.ToString() + "', '" + rw["lote"].ToString() + "', '0', '0'," +
                            " '" + subtipo + "', '" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "', '" + Convert.ToDecimal(rw["costo"].ToString()).ToString() + "', '" + txt_sis_tipo_cambio.Text + "'," +
                            " '" + rw["unitario"].ToString() + "')";
                        //((glo_var_int_registro == 7) ? "COM" : (glo_var_int_registro == 9) ? "GUA" : "")
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //insercion en mstr_transaccion_emp
                        cmnd1 = thisConnection.CreateCommand();
                        query = "INSERT INTO tb_mstr_transaccion_emp(tra_folio, tra_fecha, emp_clave, tra_cantidad, tra_tipo_mov, tra_codigo_mov, tra_codigo_tra) VALUES " +
                            "('" + txtFolio.Text + "', '" + DateTime.Now.ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + var_dec_entd_cantidad.ToString() + "'," +
                            " 'E', 'EN', 'A')";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        thisConnection.Close();

                        //decimal promedio = 0, C = 0;
                        //promedio = Convert.ToDecimal(recalculo_costos_empaque(var_chr_emp_clave, dtpFecha.Text));
                        //C = Convert.ToDecimal(recalculo_existencia_empaque(var_chr_emp_clave));

                        //thisConnection.Open();
                        //cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_costo_prom = '" + Convert.ToDecimal(promedio.ToString()) + "', hrp_existencia = '" + Convert.ToDecimal(C.ToString()).ToString() + "' " +
                        //    "WHERE hrp_recibo = '" + var_int_sis_folios + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND emp_clave = '" + var_chr_emp_clave + "' AND hrp_tipo_recepcion = 'ENT'";
                        //cmnd1.ExecuteNonQuery();
                        //cmnd1.Dispose();
                        //thisConnection.Close();
                    }

                    foreach (DataRow rw in dtdetalles.Rows)
                    {
                        if (rw["orden_compra"].ToString() == "")
                            continue;
                        thisConnection.Open();
                        //insercion en det_entrada_oc
                        cmnd1 = thisConnection.CreateCommand();
                        query = "INSERT INTO tb_det_entrada_oc(ent_folio, numero_oc, emp_clave, ent_cantidad, ent_importe, estatus," +
                            " alm_clave, ent_fecha, req_folio, ent_subtipo, ent_lote, ent_iva, ent_importe_orig, cant_orig, resp, usu, machine) VALUES(" +
                            " '" + txtFolio.Text + "', '" + rw["orden_compra"].ToString() + "', '" + rw["emp_clave"].ToString() + "'," +
                            " '" + Convert.ToDecimal(rw["cantidad"].ToString()).ToString() + "', '" +Convert.ToDecimal(rw["importe"].ToString()).ToString() + "', 'A', '" + glo_var_chr_alm_clave + "'," +
                            " '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '" + rw["requisicion"].ToString() + "'," +
                            " '" + subtipo + "', '" + rw["lote"].ToString() + "', '" + Convert.ToDecimal(rw["iva"].ToString()).ToString() + "', '" + Convert.ToDecimal(rw["importe_orig"].ToString()).ToString() + "'," +
                            " '" + Convert.ToDecimal(rw["cant_orig"].ToString()).ToString() + "', '" + rw["resp"].ToString() + "'," +
                            " '" + Utilerias.Class1.Usu_login.ToString().Trim() + "', '" + Environment.MachineName.ToString() + "')";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        if (this.glo_var_chr_alm_clave == "01" || this.glo_var_chr_alm_clave == "02")
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " Consulta(MODIFICACION): " + query);

                        //14/11/2018 cambio en la forma de saldar orden de compra
                        string sumatoria = "";
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "SELECT SUM(ent_cantidad) AS sumatoria FROM tb_det_entrada_oc WHERE numero_oc = '" + rw["orden_compra"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "' AND estatus = 'A'";
                        reader1 = cmnd1.ExecuteReader();
                        if (reader1.HasRows)
                        {
                            reader1.Read();
                            sumatoria = reader1["sumatoria"].ToString().Trim();
                        }
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();

                        //actualizar det_ordenescompra
                        cmnd1 = thisConnection.CreateCommand();
                        //query = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc + '" + Convert.ToDecimal(rw["cantidad"].ToString()).ToString() + "' WHERE numero_oc = '" + rw["orden_compra"].ToString() + "' AND prod_clave = '" + rw["emp_clave"].ToString() + "'";
                        query = "UPDATE tb_det_ordenescompra SET surtido_oc = '" + sumatoria + "' WHERE numero_oc = '" + rw["orden_compra"].ToString() + "' AND prod_clave = '" + rw["emp_clave"].ToString() + "'";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();
                        //14/11/2018 fin cambio en la forma de saldar orden de compra
                        ////actualizar det_ordenescompra
                        //cmnd1 = thisConnection.CreateCommand();
                        //query = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc + '" + rw["cantidad"].ToString() + "' WHERE numero_oc = '" + rw["orden_compra"].ToString() + "' AND prod_clave = '" + rw["emp_clave"].ToString() + "'";
                        //cmnd1.CommandText = query;
                        //cmnd1.ExecuteNonQuery();
                        //cmnd1.Dispose();

                        //Actualizar requisicion si hubiera empaques
                        if (rw["requisicion"].ToString() != "")
                        {
                            if (DDLAlmacen1.SelectedIndex == 4)
                            {
                                //buscar que existe el empaque dentro de det_requisiciones
                                cmnd1 = thisConnection.CreateCommand();
                                query = "SELECT req_xpedir, req_surtido FROM tb_det_requisiciones_enf_pap WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "' AND req_tipo = 'ENF'";
                                cmnd1.CommandText = query;
                                reader1 = cmnd1.ExecuteReader();
                                if (reader1.HasRows)
                                {
                                    //actualizacion de surtido de requisicion
                                    cmnd2 = thisConnection.CreateCommand();
                                    query = "UPDATE tb_det_requisiciones_enf_pap SET req_surtido = req_surtido + '" + rw["cantidad"].ToString() + "', req_fecharec = '" + DateTime.Now.ToShortDateString() + "'," +
                                        " alm_clave = '" + glo_var_chr_alm_clave + "' WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "' and req_tipo = 'ENF'";
                                    cmnd2.CommandText = query;
                                    cmnd2.ExecuteNonQuery();
                                    cmnd2.Dispose();
                                }
                                reader1.Close();
                                reader1.Dispose();
                                cmnd1.Dispose();
                            }
                            else if (DDLAlmacen1.SelectedIndex == 5)
                            {
                                //buscar que existe el empaque dentro de det_requisiciones
                                cmnd1 = thisConnection.CreateCommand();
                                query = "SELECT req_xpedir, req_surtido FROM tb_det_requisiciones_enf_pap WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "' AND req_tipo = 'PAP'";
                                cmnd1.CommandText = query;
                                reader1 = cmnd1.ExecuteReader();
                                if (reader1.HasRows)
                                {
                                    //actualizacion de surtido de requisicion
                                    cmnd2 = thisConnection.CreateCommand();
                                    query = "UPDATE tb_det_requisiciones_enf_pap SET req_surtido = req_surtido + '" + rw["cantidad"].ToString() + "', req_fecharec = '" + DateTime.Now.ToShortDateString() + "'," +
                                        " alm_clave = '" + glo_var_chr_alm_clave + "' WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "' and req_tipo = 'PAP'";
                                    cmnd2.CommandText = query;
                                    cmnd2.ExecuteNonQuery();
                                    cmnd2.Dispose();
                                }
                                reader1.Close();
                                reader1.Dispose();
                                cmnd1.Dispose();
                            }
                            else
                            {
                                //buscar que existe el empaque dentro de det_requisiciones
                                cmnd1 = thisConnection.CreateCommand();
                                query = "SELECT req_xpedir, req_surtido FROM tb_det_requisiciones WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "'";
                                cmnd1.CommandText = query;
                                reader1 = cmnd1.ExecuteReader();
                                if (reader1.HasRows)
                                {
                                    //actualizacion de surtido de requisicion
                                    cmnd2 = thisConnection.CreateCommand();
                                    query = "UPDATE tb_det_requisiciones SET req_surtido = req_surtido + '" + rw["cantidad"].ToString() + "', req_fecharec = '" + DateTime.Now.ToShortDateString() + "'," +
                                        " alm_clave = '" + glo_var_chr_alm_clave + "' WHERE req_folio = '" + rw["requisicion"].ToString() + "' AND emp_clave = '" + rw["emp_clave"].ToString() + "'";
                                    cmnd2.CommandText = query;
                                    cmnd2.ExecuteNonQuery();
                                    cmnd2.Dispose();
                                }
                                reader1.Close();
                                reader1.Dispose();
                                cmnd1.Dispose();
                            }
                        }

                        thisConnection.Close();

                        
                    }

                    //guarda_genera_pdf_servidor();


                    //string importacionx = "0";
                    //thisConnection.Open();
                    //SqlCommand cmndimportacionx = thisConnection.CreateCommand();
                    //cmndimportacionx.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
                    //try
                    //{
                    //    importacionx = Convert.ToString(cmndimportacionx.ExecuteScalar().ToString().Trim());
                    //}
                    //catch
                    //{
                    //    importacionx = "";
                    //}

                    //cmndimportacionx.Dispose();
                    //thisConnection.Close();
                    //if (importacionx == "1")
                    //{

                    //    /*if (txt_prov_clave.Text == "81" || txt_prov_clave.Text == "404" || txt_prov_clave.Text == "11" || txt_prov_clave.Text == "8851" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694" || txt_prov_clave.Text.Trim() == "9957" || txt_prov_clave.Text.Trim() == "833" || txt_prov_clave.Text.Trim() == "1550" || txt_prov_clave.Text.Trim() == "1761" || txt_prov_clave.Text.Trim() == "935" || txt_prov_clave.Text.Trim() == "6854" || txt_prov_clave.Text == "354" || txt_prov_clave.Text == "1250" || txt_prov_clave.Text == "1536" || txt_prov_clave.Text == "9758" || txt_prov_clave.Text == "1799" || txt_prov_clave.Text == "1244" || txt_prov_clave.Text == "191" || txt_prov_clave.Text == "812" || txt_prov_clave.Text == "8014" || txt_prov_clave.Text == "305" || txt_prov_clave.Text == "106")
                    //    {*/
                    //    //subir archivo XML
                    //    //if (lblValidaXML.Text == "1")
                    //    //{
                    //    if (lblPathXml.Text != "-")
                    //    {
                    //        FileInfo archivo = new FileInfo(@lblPathXml.Text);
                    //        string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                    //        FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //        if (liq_copy.Exists == true)
                    //        {
                    //            liq_copy.Delete();
                    //        }
                    //        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                    //        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //        if (liq_copy.Exists != true)
                    //        {
                    //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (MOD) de XML " + file_name);
                    //        }

                    //        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ".  Guardado(MOD) de archivo XML " + file_name);

                    //        file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                    //        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //        if (liq_copy.Exists == true)
                    //        {
                    //            liq_copy.Delete();
                    //        }
                    //        archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                    //        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                    //        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //        if (liq_copy.Exists != true)
                    //        {
                    //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (MOD) de PDF " + file_name);
                    //        }

                    //        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ".  Guardado(MOD) de archivo PDF " + file_name);
                    //    }


                    //    ////SUBIR ARCHIVO xml
                    //    //string origen1 = @lblPathXml.Text;
                    //    //string destino1 = @lblDestino.Text + lblXmlFile.Text;
                    //    //if (!File.Exists(origen1))
                    //    //{
                    //    //    //using (FileStream fs = File.Create(origen1)) { }
                    //    //}
                    //    ////if (File.Exists(lblDestino.Text))
                    //    ////    File.Delete(lblDestino.Text);
                    //    //File.Copy(origen1, destino1, true);

                    //    ////SUBIR PDF
                    //    //origen1 = @lblPathXml.Text.Replace(".xml", ".pdf");//nombre de archivo
                    //    //destino1 = @lblDestino.Text + lblXmlFile.Text.Replace(".xml", ".pdf");
                    //    //if (!File.Exists(origen1))
                    //    //{
                    //    //    //using (FileStream fs = File.Create(origen1)) { }
                    //    //}
                    //    ////if (File.Exists(lblDestino.Text))
                    //    ////File.Delete(lblDestino.Text);
                    //    //File.Copy(origen1, destino1, true);
                    //    //}
                    //}
                    //else
                    //{
                    //    if (lblPathXml.Text != "-")
                    //    {
                    //        FileInfo archivo = new FileInfo(@lblPathXml.Text);
                    //        string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                    //        FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //        if (liq_copy.Exists == true)
                    //        {
                    //            liq_copy.Delete();
                    //        }
                    //        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                    //        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //        if (liq_copy.Exists != true)
                    //        {
                    //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (MOD) de XML " + file_name);
                    //        }

                    //        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ".  Guardado(MOD) de archivo XML " + file_name);

                    //        file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                    //        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //        if (liq_copy.Exists == true)
                    //        {
                    //            liq_copy.Delete();
                    //        }
                    //        archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                    //        archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                    //        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //        if (liq_copy.Exists != true)
                    //        {
                    //            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                    //            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Guardado 2 (MOD) de PDF " + file_name);
                    //        }

                    //        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ".  Guardado(MOD) de archivo PDF " + file_name);
                    //    }
                    //}


                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "M", "2.5", txtFolio.Text, "MODIFICACION DE ENTRADA(OC): " + txtFolio.Text + ", ALMACEN: " + glo_var_chr_alm_clave, "SISEMP");
                    string filelog2 = "C:\\SisEmpWeb\\eventlog.txt";
                    using (StreamWriter sw = File.AppendText(filelog2))
                    {
                        sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Actualización de entrada(oc): " + txtFolio.Text + ", Almacen: " + glo_var_chr_alm_clave);
                        sw.Close();
                    }

                    if (lblPathXml.Text != "-")
                    {
                        
                        FileInfo archivo = new FileInfo(@lblPathXml.Text);

                        string file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".xml";
                        FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                        if (liq_copy.Exists == true)
                        {
                            //liq_copy.Delete();
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Archivo existe XML" + file_name);
                        }
                        else
                        {
                            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Archivo no existe se volvio a subir XML " + file_name);
                        }
                        

                        file_name = txt_prov_clave.Text + "_" + txt_ent_factura.Text + "_" + txtFolio.Text + ".pdf";
                        liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                        if (liq_copy.Exists == true)
                        {
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Archivo existe PDF " + file_name);
                        }
                        else
                        {
                            //archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                            archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ". Archivo no existe se volvio a subir PDF " + file_name);
                        }
                        

                        //form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + ".  Guardado(MOD) de archivo PDF " + file_name);
                    }
                    

                    

                    //imprmir_entrada_servidor();
                    //RECORRER RECALCULO DE EMPAQUES GUARDADOS
                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + "RECALCULO DE EMPAQUES MODIFICADOS COSTO Y EXISTENCIA " + DateTime.Now.ToString());
                    foreach (DataRow rw in tmgrdpro.Rows)
                    {
                        if (glo_var_chr_alm_clave == "13")
                        {
                            //PROCESO PARA ACTUALIZAR EL COSTO DEL CATALOGO DE EMPAQUES AL COSTO DE LA ENTRADA O REPARACION
                            string emp = rw["clave"].ToString();
                            string alm = glo_var_chr_alm_clave;
                            string cto = Convert.ToDecimal(rw["unitario"].ToString()).ToString();

                            string FECHA = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                            string NOM_COMPU = Environment.MachineName;
                            string NOM_USU = Utilerias.Class1.Usu_login.Trim();
                            string TIPO_MOV = "A";
                            string OP_CLAVE = "2.5";
                            string FOLIO = var_int_sis_folios;
                            string DETALLE = "Modificación de costo en catalogo de empaques por orden de reparacion(Consulta)";
                            string SISTEMA = "SISEMP";
                            string MOTIVO = "ENTRADA POR ORDEN DE REPARACION";
                            string OBS = "Empaque: " + rw["clave"].ToString() + " Costo: " + cto;

                            _consultaController.ActualizaCostoCatalogoEmpaquesOrdenReparacion(emp, alm, cto, FECHA, NOM_COMPU, NOM_USU, TIPO_MOV, OP_CLAVE, FOLIO, DETALLE, SISTEMA, MOTIVO, OBS);
                            decimal inv = Convert.ToDecimal(recalculo_existencia_empaque(emp));
                        }
                        else
                        {
                            decimal promedio = 0, C = 0;
                            promedio = Convert.ToDecimal(recalculo_costos_empaque(var_chr_emp_clave, dtpFecha.Text));
                            C = Convert.ToDecimal(recalculo_existencia_empaque(var_chr_emp_clave));

                            thisConnection.Open();
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_costo_prom = '" + Convert.ToDecimal(promedio.ToString()) + "', hrp_existencia = '" + Convert.ToDecimal(C.ToString()).ToString() + "' " +
                                "WHERE hrp_recibo = '" + var_int_sis_folios + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND emp_clave = '" + var_chr_emp_clave + "' AND hrp_tipo_recepcion = 'ENT'";
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                            thisConnection.Close();
                        }
                    }
                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + "FIN RECALCULO DE EMPAQUES MODIFICADOS COSTO Y EXISTENCIA " + DateTime.Now.ToString());
                    MessageBox.Show("Datos modificados", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    if (MessageBox.Show("Imprimir Entrada en Impresora", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        btnImprimir_Click(null, null);

                    

                    foreach (DataRow re in dtoriginales.Rows)
                    {
                        bool fnd = false;
                        string e_m_p = re["emp_clave"].ToString();
                        foreach (DataRow rw in tmgrdpro.Select("clave = '" + e_m_p + "'"))
                        {
                            fnd = true;
                            break;
                        }
                        if (fnd == false)
                        {
                            recalculo_costos(re["emp_clave"].ToString());
                            recalculo_existencia(re["emp_clave"].ToString());
                            modificar_ultima_entrada(re["emp_clave"].ToString());
                        }
                    }

                    //foreach (DataRow rc in tmgrdpro.Rows)
                    //{
                    //    recalculo_costos(rc["clave"].ToString());
                    //    recalculo_existencia(rc["clave"].ToString());
                    //}

                    if (glo_var_chr_alm_clave != "13")
                    {
                        foreach (DataRow rw in tmgrdpro.Rows)
                        {
                            string e_m_p = rw["clave"].ToString();
                            recalculo_costos(e_m_p);
                            recalculo_existencia(e_m_p);
                            modificar_ultima_entrada(e_m_p);
                        }
                    }
                    

                    if (Utilerias.Class1.Usu_login.Trim() == "PEDRO")
                    {
                        string tbl = "<p><h2>Entrada: " + txtFolio.Text + "</h2></p><table border='1'><tr><td><b>Empaque</b></td><td><b>Cantidad</b></td><td><b>Costo</b></td><td><b>Importe</b></td><td><b>OC</b></td></tr>";
                        foreach (DataGridViewRow gr in dtgDetalle2.Rows)
                        {
                            tbl += "<tr><td>" + gr.Cells["clave1"].Value.ToString() + "</td><td align='right'>" + gr.Cells["cant"].Value.ToString() + "</td><td align='right'>" + gr.Cells["cost"].Value.ToString() + "</td><td align='right'>" +
                                gr.Cells["import"].Value.ToString() + "</td><td>" + gr.Cells["ordencompra"].Value.ToString() + "</td></tr>";
                        }
                        tbl += "</table>";
                        correo_error(tbl, "Modificacion de Entrada Pedro");
                    }

                    btnLimpiar_Click(null, null);

                    DDLAlmacen1_SelectionChangeCommitted(null, null);
                }
                catch (SqlException sqlex)
                {
                    MessageBox.Show("Error de sistema, no se termino de guardar la entrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (thisConnection.State == ConnectionState.Open)
                        thisConnection.Close();
                    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.5", txtFolio.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + sqlex.ToString());
                    //this.Close();                                
                    return;
                }

            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de modificar la entrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.5", txtFolio.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }
        }

        public void modificar()
        { 
            //Modificacion de la entrada
            if (validacontroles() == false)
            {
                return;
            }
            imprimirtxt();
            if (MessageBox.Show("¿Desea realizar algún cambio en la orden de entrada?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                return;
            }
            string query = "";
            try
            {
                thisConnection.Open();
                string tipo = "NOR";
                string reg = "DEF";
                string pro = ((DDL_ent_procedencia.SelectedItem.ToString() == "Nacional") ? "NAL" : "IMP");
                //string sub = ((glo_var_int_registro == 7) ? "COM" : (glo_var_int_registro == 9) ? "GUA" : "");
                cmnd1 = thisConnection.CreateCommand();
                //ACTUALIZACION DE ENTRADA
                query = "UPDATE tb_mstr_entradas SET ent_tipo = '" + tipo + "', ent_fecha = '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', ent_elaboro = '" + DDLElaboro.SelectedItem.ToString() + "'," +
                    " ent_autorizo = '" + DDLAutorizo.SelectedItem.ToString() + "', ent_observaciones = '" + txtObservaciones.Text + "', alm_clave = '" + glo_var_chr_alm_clave + "', prov_clave = '" + txt_prov_clave.Text + "'," +
                    " ent_iva = '" + Convert.ToDecimal(txt_ent_iva.Text) + "', ent_monto = '" + Convert.ToDecimal(txt_total.Text) + "', ent_factura = '" + txt_ent_factura.Text + "', sis_tipo_cambio = '" + Convert.ToDecimal(txt_sis_tipo_cambio.Text) + "'," +
                    " ent_regimen = '" + reg + "', ent_procedencia = '" + pro + "', ent_pedimento = '" + txt_ent_pedimento.Text + "', ent_subtipo = '" + subtipo + "', ent_gastoaduanal = '" + Convert.ToDecimal(txt_gto_aduanal.Text) + "'," +
                    " ent_iepstotal = '" + Convert.ToDecimal(txt_ent_ieps.Text) + "', ent_flete = '" + Convert.ToDecimal(txtFlete.Text) + "' WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();



                int oper = 0;
                foreach (DataRow rw in tmgrdpro.Rows)
                {
                    if (rw["oper"].ToString() == "nueva")//MODIFICACION CUANDO SE AGREGA UN NUEVO RENGLON
                    {
                        #region nueva
                        string var_chr_emp_clave = "";
                        decimal var_dec_entd_cantidad = 0;
                        string iepscargo = "";

                        var_chr_emp_clave = rw["clave"].ToString();
                        var_dec_entd_cantidad = Convert.ToDecimal(rw["cantidad"].ToString());
                        iepscargo = (Convert.ToBoolean(rw["ieps"].ToString())) ? "T" : "F";

                        //traer datos de cat_empaques
                        //thisConnection.Open();
                        cmnd1 = thisConnection.CreateCommand();
                        query = "SELECT emp_costo, emp_existencia, alm_clave FROM tb_cat_empaques WHERE emp_clave = '" + rw["clave"].ToString() + "' ORDER BY emp_clave";
                        cmnd1.CommandText = query;
                        reader1 = cmnd1.ExecuteReader();
                        decimal costoactual = 0;
                        decimal existencia = 0;
                        string cvealm = "";
                        if (reader1.HasRows)
                        {
                            reader1.Read();
                            costoactual = Convert.ToDecimal(reader1.GetValue(0).ToString().Trim());
                            existencia = Convert.ToDecimal(reader1.GetValue(1).ToString().Trim());
                            cvealm = reader1.GetValue(2).ToString().Trim();
                        }
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();

                        //insercion en det_ent_temporal
                        //string sub2 = (glo_var_int_registro == 7) ? ", 'COM'" : (glo_var_int_registro == 9) ? ", 'GUA'" : ", ''";
                        cmnd1 = thisConnection.CreateCommand();
                        query = "INSERT INTO tb_det_ent_temporal(ent_folio, emp_clave, alm_clave, entd_cantidad, entd_costo, entd_adv, entd_dta, lote_ras, ent_tipo, ent_ieps, ent_ieps_importe, ent_gastoaduanal, ent_flete, ord_compra)" +
                                    " VALUES ('" + txtFolio.Text + "', '" + var_chr_emp_clave + "', '" + glo_var_chr_alm_clave + "', '" + var_dec_entd_cantidad.ToString() + "', '" + rw["costo"].ToString() + "'" +
                                    ", '" + rw["adv"].ToString() + "', '" + rw["dta"].ToString() + "', '" + rw["lote"].ToString() + "', '" + subtipo + "', '" + iepscargo + "'," +
                                    " '" + rw["iepsimporte"].ToString() + "', '" + rw["gastoaduanal"].ToString() + "', '" + rw["flete"].ToString() + "', '" + rw["oc"].ToString() + "')";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //mandar al historico el registro
                        decimal A = 0, B = 0, C = 0, D = 0, promedio = 0;
                        A = Math.Abs(existencia * costoactual);
                        B = var_dec_entd_cantidad * Convert.ToDecimal(rw["unitario"].ToString());
                        C = Math.Abs(existencia) + var_dec_entd_cantidad;
                        D = (C == 0) ? Convert.ToDecimal(rw["unitario"].ToString()) : (A + B) / C;
                        C = existencia + var_dec_entd_cantidad;
                        promedio = Math.Abs(D);

                        //graba ultimo costo
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "SELECT emp_entrada_ultima FROM tb_cat_empaques WHERE emp_clave = '" + var_chr_emp_clave + "' AND emp_entrada_ultima < '" + dtpFecha.Text + "'";
                        reader1 = cmnd1.ExecuteReader();
                        if (reader1.HasRows)
                        {
                            cmnd4 = thisConnection.CreateCommand();
                            cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "', emp_entrada_ultima = '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', emp_existencia = " + C.ToString() + " WHERE emp_clave = '" + var_chr_emp_clave + "' AND emp_entrada_ultima < '" + dtpFecha.Text + "'";
                            cmnd4.ExecuteNonQuery();
                            cmnd4.Dispose();
                        }
                        else
                        {
                            cmnd4 = thisConnection.CreateCommand();
                            cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "', emp_existencia = " + C.ToString() + " WHERE emp_clave = '" + var_chr_emp_clave + "'";
                            cmnd4.ExecuteNonQuery();
                            cmnd4.Dispose();
                        }
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();

                        string tip2 = "NOR";
                        string reg2 = "DEF";
                        string pro2 = (DDL_ent_procedencia.SelectedItem.ToString() == "Nacional") ? "NAL" : "IMP";
                        cmnd1 = thisConnection.CreateCommand();
                        query = "INSERT INTO tb_historico_recepcion(hrp_recibo, hrp_fecha, emp_clave, hrp_cantidad, hrp_estatus, hrp_tipo_recepcion, alm_clave, hrp_situacion," +
                            " hrp_costo, hrp_regimen, hrp_procedencia, hrp_sobrante, lote_ras, hrp_costo_prom, hrp_existencia, hrp_tipo, hrp_fechamov) VALUES ('" + txtFolio.Text + "', " +
                            " '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + var_dec_entd_cantidad.ToString() + "', 'T', 'ENT', '" + glo_var_chr_alm_clave + "'," +
                            " '" + tip2 + "', '" + rw["unitario"].ToString() + "', " +
                            " '" + reg2 + "', '" + pro2 + "', '" + var_dec_entd_cantidad.ToString() + "', '" + rw["lote"].ToString() + "', '" + promedio.ToString() + "', '" + C.ToString() + "'," +
                            " '" + subtipo + "', '" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + "')";
                        //((glo_var_int_registro == 7) ? "COM" : (glo_var_int_registro == 9) ? "GUA" : "")
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //insercion en mstr_transaccion_emp
                        cmnd1 = thisConnection.CreateCommand();
                        query = "INSERT INTO tb_mstr_transaccion_emp(tra_folio, tra_fecha, emp_clave, tra_cantidad, tra_tipo_mov, tra_codigo_mov, tra_codigo_tra) VALUES " +
                            "('" + txtFolio.Text + "', '" + DateTime.Now.ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + var_dec_entd_cantidad.ToString() + "'," +
                            " 'E', 'EN', 'A')";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //insercion en det_entrada_oc
                        cmnd1 = thisConnection.CreateCommand();
                        query = "INSERT INTO tb_det_entrada_oc(ent_folio, numero_oc, emp_clave, ent_cantidad, estatus, alm_clave, ent_fecha, req_folio) VALUES(" +
                            " '" + txtFolio.Text + "', '" + rw["oc"].ToString() + "', '" + rw["clave"] + "'," +
                            " '" + rw["cantidad"].ToString() + "', 'A', '" + glo_var_chr_alm_clave + "', '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "'," +
                            " '" + rw["req"].ToString() + "')";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //actualizar det_ordenescompra
                        cmnd1 = thisConnection.CreateCommand();
                        query = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc + '" + var_dec_entd_cantidad.ToString() + "' WHERE numero_oc = '" + rw["oc"].ToString() + "' AND prod_clave = '" + var_chr_emp_clave + "'";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //Actualizar requisicion si hubiera empaques
                        if (rw["req"].ToString() != "")
                        {
                            //buscar que existe el empaque dentro de det_requisiciones
                            cmnd1 = thisConnection.CreateCommand();
                            query = "SELECT req_xpedir, req_surtido FROM tb_det_requisiciones WHERE req_folio = '" + rw["req"].ToString() + "' AND emp_clave = '" + var_chr_emp_clave + "'";
                            cmnd1.CommandText = query;
                            reader1 = cmnd1.ExecuteReader();
                            if (reader1.HasRows)
                            {
                                //actualizacion de surtido de requisicion
                                cmnd2 = thisConnection.CreateCommand();
                                query = "UPDATE tb_det_requisiciones SET req_surtido = req_surtido + '" + var_dec_entd_cantidad.ToString() + "', req_fecharec = '" + DateTime.Now.ToShortDateString() + "'," +
                                    " alm_clave = '" + glo_var_chr_alm_clave + "' WHERE req_folio = '" + rw["req"].ToString() + "' AND emp_clave = '" + var_chr_emp_clave + "'";
                                cmnd2.CommandText = query;
                                cmnd2.ExecuteNonQuery();
                                cmnd2.Dispose();
                            }
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();
                        }
                        #endregion
                    }
                    else if(rw["borrar"].ToString() == "1")//CUANDO SE ELIMINO UN REGISTRO DEL GRID
                    {
                        #region eliminar_empaque
                        //ELIMINACION DE REGISTROS EN DETALLE DE ENTRADA
                        cmnd1 = thisConnection.CreateCommand();
                        query = "DELETE FROM tb_det_ent_temporal WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND emp_clave = '" + rw["clave"].ToString() + "' " +
                            "AND ord_compra = '" + rw["oc"].ToString() + "'";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ELIMINAR REGISTRO DE HISTORICO
                        cmnd1 = thisConnection.CreateCommand();
                        query = "DELETE FROM tb_historico_recepcion WHERE hrp_recibo = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND hrp_tipo_recepcion = 'ENT' " +
                            " AND hrp_situacion = 'NOR' AND emp_clave = '" + rw["clave"].ToString() + "'";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ELIMINAR DE ENTRADA_OC
                        cmnd1 = thisConnection.CreateCommand();
                        query = "DELETE FROM tb_det_entrada_oc WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' " +
                            "AND emp_clave = '" + rw["clave"].ToString() + "' AND numero_oc = '" + rw["oc"].ToString() +"'";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ACTUALIZA DET_ORDENESCOMPRA QUITA LA CANTIDAD SURTIDA
                        cmnd1 = thisConnection.CreateCommand();
                        query = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc - '" + rw["cantidad"].ToString() + "' WHERE numero_oc = '" + rw["oc"].ToString() + "' AND prod_clave = '" + rw["clave"].ToString() + "'";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //VERIFICAMOS SI VIENE DE UNA REQUISICION PARA QUITAR LA CANTIDAD SURTIDA
                        if (rw["req"].ToString() != "")
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            query = "UPDATE tb_det_requisiciones SET req_surtido = req_surtido - '" + rw["cantidad"].ToString() + "' WHERE req_folio = '" + rw["req"].ToString() + "' AND emp_clave = '" + rw["clave"].ToString() + "'";
                            cmnd1.CommandText = query;
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                        }

                        //MODIFICAMOS CAT_EMPAQUES CON EL NUEVO COSTO PROMEDIO
                        traer_costopromedio(txtFolio.Text, rw["clave"].ToString(), dtpFecha.Text);
                        #endregion
                    }
                    //else//SI NO ES NUEVA 
                    //{
                        #region modificacion
                    //    if (Convert.ToDecimal(rw["cantidad"].ToString()) == Convert.ToDecimal(rw["cantorig"].ToString()))
                    //    {
                    //        if (Convert.ToDecimal(rw["unitario"].ToString()) == Convert.ToDecimal(rw["unitorig"].ToString()))
                    //        {
                    //            oper = 1;
                    //        }
                    //    }
                    //    if (Convert.ToDecimal(rw["unitario"].ToString()) != Convert.ToDecimal(rw["unitorig"].ToString()))
                    //    {
                    //        oper = 2;
                    //    }
                    //    if (Convert.ToDecimal(rw["cantidad"].ToString()) != Convert.ToDecimal(rw["cantorig"].ToString()))
                    //    {
                    //        oper = 2;
                    //    }
                    //    switch (oper)
                    //    {
                    //        case 1:
                    //            //SOLO BUSCAMOS EN DETALLE DE ENTRADAS PARA MODIFICAR LOS DATOS
                    //            cmnd2 = thisConnection.CreateCommand();
                    //            cmnd2.CommandText = "UPDATE tb_det_ent_temporal SET entd_adv = '" + rw["adv"].ToString() + "', entd_dta = '" + rw["dta"].ToString() + "', lote_ras = '" + rw["lote"].ToString() + "'" +
                    //                ", ent_ieps = '" + ((rw["ieps"].ToString() == "True") ? "T" : "F") + "'" +
                    //                ", ent_ieps_importe = '" + rw["iepsimporte"] + "', ent_gastoaduanal = '" + rw["gastoaduanal"].ToString() + "', ent_flete = '" + rw["flete"] + "' " +
                    //                "WHERE ent_folio = '" + txtFolio.Text + "' AND emp_clave = '" + rw["clave"] + "' AND ord_compra = '" + rw["oc"] + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
                    //            cmnd2.ExecuteNonQuery();
                    //            cmnd2.Dispose();
                    //            break;
                    //        case 2:
                    //            //HACEMOS LAS ACTUALIZACIONES CUANDO HAY CAMBION EN CANTIDAD O IMPORTE
                    //            //if (Convert.ToDecimal(rw["cantidad"].ToString()) != Convert.ToDecimal(rw["cantorig"].ToString()))//SI HAY DIFERENCIAS HACEMOS EL CAMBIO
                    //            //{
                    //            //TB_MSTR_TRANSACCION_EMP
                    //            cmnd4 = thisConnection.CreateCommand();
                    //            cmnd4.CommandText = "UPDATE tb_mstr_transaccion_emp SET tra_cantidad = '" + rw["cantidad"].ToString() + "' WHERE emp_clave = '" + rw["clave"].ToString() + "'" +
                    //                " AND tra_folio = '" + txtFolio.Text + "' AND tra_tipo_mov = 'E' AND tra_codigo_mov = 'EN' AND tra_codigo_tra = 'A'";
                    //            cmnd4.ExecuteNonQuery();
                    //            cmnd4.Dispose();

                    //            //TB_DET_ENTRADA_OC
                    //            cmnd4 = thisConnection.CreateCommand();
                    //            cmnd4.CommandText = "UPDATE tb_det_entrada_oc SET ent_cantidad = '" + rw["cantidad"].ToString() + "' WHERE emp_clave = '" + rw["clave"].ToString() + "'" +
                    //                " AND ent_folio = '" + txtFolio.Text + "' AND numero_oc = '" + rw["oc"].ToString() + "'";
                    //            cmnd4.ExecuteNonQuery();
                    //            cmnd4.Dispose();


                    //            //TB_DET_ORDENES_COMPRA
                    //            decimal num_x = 0;
                    //            if (Convert.ToDecimal(rw["cantidad"].ToString()) > Convert.ToDecimal(rw["cantorig"].ToString()))
                    //            {
                    //                num_x = Convert.ToDecimal(rw["cantidad"].ToString()) - Convert.ToDecimal(rw["cantorig"].ToString());
                    //                cmnd4 = thisConnection.CreateCommand();
                    //                cmnd4.CommandText = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc + '" + num_x.ToString() + "' WHERE prod_clave = '" + rw["clave"].ToString() + "'" +
                    //                    " AND numero_oc = '" + rw["oc"].ToString() + "'";
                    //                cmnd4.ExecuteNonQuery();
                    //                cmnd4.Dispose();

                    //            }
                    //            else
                    //            {
                    //                num_x = Convert.ToDecimal(rw["cantorig"].ToString()) - Convert.ToDecimal(rw["cantidad"].ToString());
                    //                cmnd4 = thisConnection.CreateCommand();
                    //                cmnd4.CommandText = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc - '" + num_x.ToString() + "' WHERE prod_clave = '" + rw["clave"].ToString() + "'" +
                    //                    " AND numero_oc = '" + rw["oc"].ToString() + "'";
                    //                cmnd4.ExecuteNonQuery();
                    //                cmnd4.Dispose();
                    //            }

                    //            if (rw["req"].ToString() != "")
                    //            {
                    //                decimal addsubs = 0;
                    //                //Convert.ToDecimal(rw["cantidad"].ToString()) != Convert.ToDecimal(rw["cantorig"].ToString())
                    //                if (Convert.ToDecimal(rw["cantidad"].ToString()) > Convert.ToDecimal(rw["cantorig"].ToString()))
                    //                {
                    //                    addsubs = Convert.ToDecimal(rw["cantidad"].ToString()) - Convert.ToDecimal(rw["cantorig"].ToString());
                    //                    cmnd4 = thisConnection.CreateCommand();
                    //                    cmnd4.CommandText = "UPDATE tb_det_requisiciones SET req_surtido = req_surtido + '" + addsubs + "' WHERE req_folio = '" + rw["req"].ToString() + "'" +
                    //                        " AND emp_clave = '" + rw["clave"].ToString() + "'";
                    //                    cmnd4.ExecuteNonQuery();
                    //                    cmnd4.Dispose();
                    //                }
                    //                else
                    //                {
                    //                    addsubs = Convert.ToDecimal(rw["cantorig"].ToString()) - Convert.ToDecimal(rw["cantidad"].ToString());
                    //                    cmnd4 = thisConnection.CreateCommand();
                    //                    cmnd4.CommandText = "UPDATE tb_det_requisiciones SET req_surtido = req_surtido - '" + addsubs + "' WHERE req_folio = '" + rw["req"].ToString() + "'" +
                    //                        " AND emp_clave = '" + rw["clave"].ToString() + "'";
                    //                    cmnd4.ExecuteNonQuery();
                    //                    cmnd4.Dispose();
                    //                }
                    //            }


                    //            //}
                    //            //if (Convert.ToDecimal(rw["unitario"].ToString()) != Convert.ToDecimal(rw["unitorig"].ToString()))
                    //            //{
                    //            //CUANDO COSTO ES DIFERENTE
                    //            decimal num = 0;
                    //            if (Convert.ToDecimal(rw["cantidad"].ToString()) > Convert.ToDecimal(rw["cantorig"].ToString()))
                    //            {
                    //                num = Convert.ToDecimal(rw["cantidad"].ToString()) - Convert.ToDecimal(rw["cantorig"].ToString());
                    //                //TB_HISTORICO_RECEPCION
                    //                cmnd4 = thisConnection.CreateCommand();
                    //                cmnd4.CommandText = "UPDATE tb_historico_recepcion SET hrp_cantidad = '" + rw["cantidad"].ToString() + "', hrp_costo = '" + rw["unitario"].ToString() + "'," +
                    //                    " hrp_existencia = hrp_existencia + '" + num + "'" +
                    //                    " WHERE emp_clave = '" + rw["clave"].ToString() + "' AND hrp_recibo = '" + txtFolio.Text + "'  AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR' AND alm_clave = '" + glo_var_chr_alm_clave + "' ";
                    //                cmnd4.ExecuteNonQuery();
                    //                cmnd4.Dispose();
                    //            }
                    //            else
                    //            {
                    //                num = Convert.ToDecimal(rw["cantorig"].ToString()) - Convert.ToDecimal(rw["cantidad"].ToString());
                    //                //TB_HISTORICO_RECEPCION
                    //                cmnd4 = thisConnection.CreateCommand();
                    //                cmnd4.CommandText = "UPDATE tb_historico_recepcion SET hrp_cantidad = '" + rw["cantidad"].ToString() + "', hrp_costo = '" + rw["unitario"].ToString() + "'," +
                    //                    " hrp_existencia = hrp_existencia - '" + num + "'" +
                    //                    " WHERE emp_clave = '" + rw["clave"].ToString() + "' AND hrp_recibo = '" + txtFolio.Text + "'  AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR' AND alm_clave = '" + glo_var_chr_alm_clave + "' ";
                    //                cmnd4.ExecuteNonQuery();
                    //                cmnd4.Dispose();
                    //            }



                    //            cmnd2 = thisConnection.CreateCommand();
                    //            cmnd2.CommandText = "UPDATE tb_det_ent_temporal SET entd_cantidad = '" + rw["cantidad"].ToString() + "', entd_costo = '" + rw["costo"].ToString() + "'" +
                    //                ", entd_adv = '" + rw["adv"].ToString() + "', entd_dta = '" + rw["dta"].ToString() + "'" +
                    //                ", lote_ras = '" + rw["lote"].ToString() + "'" +
                    //                ", ent_ieps = '" + ((rw["ieps"].ToString() == "True") ? "T" : "F") + "'" +
                    //                ", ent_ieps_importe = '" + rw["iepsimporte"] + "', ent_gastoaduanal = '" + rw["gastoaduanal"].ToString() + "', ent_flete = '" + rw["flete"] + "' " +
                    //                "WHERE ent_folio = '" + txtFolio.Text + "' AND emp_clave = '" + rw["clave"] + "' AND ord_compra = '" + rw["oc"] + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
                    //            cmnd2.ExecuteNonQuery();
                    //            cmnd2.Dispose();

                    //            traer_costopromedio(txtFolio.Text, rw["clave"].ToString(), dtpFecha.Text);
                    //            //}
                    //            break;
                    //    }
                        #endregion
                    //}
                    
                }

                #region comentada
                //consulta ultimo cierre de mes
                //cmnd1 = thisConnection.CreateCommand();
                //string var_date_fecha = "";
                //query = "SELECT TOP 1 invemp_fecha tb_mstr_inventario_emp WHERE invemp_fecha < '" + dtpFecha.Text + "' ORDER BY invemp_fecha DESC";
                //if (reader1.HasRows)
                //{
                //    reader1.Read();
                //    var_date_fecha = reader1.GetValue(0).ToString().Trim();

                //}

                //cmnd1.CommandText = "";
                //foreach (DataRow rw in tmgrdpro.Rows)
                //{
                //    cmnd1 = thisConnection.CreateCommand();
                //    query = "SELECT emp_clave, hrp_cantidad, hrp_costo FROM tb_historico_recepcion WHERE hrp_fecha >= '" + var_date_fecha + "' AND hrp_fecha <= '" + dtpFecha.Text + "'" +
                //        " AND ent_folio < '" + rw["clave"] + "' AND hrp_tipo_rece";
                //}
                

                ////CONSULTA TODAS LAS SALIDAS DESDE EL ULTIMO CIERRE

                ////ELIMINACION DE REGISTROS EN DETALLE DE ENTRADA
                //cmnd1 = thisConnection.CreateCommand();
                //query = "DELETE FROM tb_det_ent_temporal WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
                //cmnd1.CommandText = query;
                //cmnd1.ExecuteNonQuery();
                //cmnd1.Dispose();

                ////ELIMINAR REGISTRO DE HISTORICO
                //cmnd1 = thisConnection.CreateCommand();
                //query = "DELETE FROM tb_historico_recepcion WHERE hrp_recibo = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND hrp_tipo_recepcion = 'ENT'";
                //cmnd1.CommandText = query;
                //cmnd1.ExecuteNonQuery();
                //cmnd1.Dispose();

                ////ELIMINAR DE TRANSACCION_EMP
                //cmnd1 = thisConnection.CreateCommand();
                //query = "DELETE FROM tb_mstr_transaccion_emp WHERE tra_folio = '" + txtFolio.Text + "' AND tra_tipo_mov = 'E' AND tra_codigo_mov = 'EN' AND tra_codigo_tra = 'A'";
                //cmnd1.CommandText = query;
                //cmnd1.ExecuteNonQuery();
                //cmnd1.Dispose();

                ////ELIMINAR DE ENTRADA_OC
                //cmnd1 = thisConnection.CreateCommand();
                //query = "DELETE FROM tb_det_entrada_oc WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
                //cmnd1.CommandText = query;
                //cmnd1.ExecuteNonQuery();
                //cmnd1.Dispose();

                //string var_chr_emp_clave = "";
                //decimal var_dec_entd_cantidad = 0;
                //string iepscargo = "";
                //string sub2 = (glo_var_int_registro == 7) ? ", 'COM'" : (glo_var_int_registro == 9) ? ", 'GUA'" : ", ''";
                //foreach (DataRow rw in tmgrdpro.Rows)
                //{
                //    //INSERCION EN DETALLE DE ENTRADA
                //    var_chr_emp_clave = rw["clave"].ToString();
                //    var_dec_entd_cantidad = Convert.ToDecimal(rw["cantidad"].ToString());
                //    iepscargo = (Convert.ToBoolean(rw["ieps"].ToString())) ? "T" : "F";
                //    cmnd1 = thisConnection.CreateCommand();
                //    query = "INSERT INTO tb_det_ent_temporal(ent_folio, emp_clave, alm_clave, entd_cantidad, entd_costo, entd_adv, entd_dta, lote_ras, ent_tipo, ent_ieps, ent_ieps_importe, ent_gastoaduanal, ent_flete, ord_compra)" +
                //        "VALUES('" + txtFolio.Text + "', '" + var_chr_emp_clave + "', '" + glo_var_chr_alm_clave + "', '" + var_dec_entd_cantidad.ToString() + "', '" + rw["costo"].ToString() + "'," +
                //        " '" + rw["adv"].ToString() + "', '" + rw["dta"].ToString() + "', '" + rw["lote"].ToString() + "'" + sub2 + ", '" + iepscargo + "', '" + rw["iepsimporte"].ToString() + "', '" + rw["gastoaduanal"].ToString() + "'," +
                //        " '" + rw["flete"].ToString() + "', '" + rw["oc"].ToString() + "')";
                //    cmnd1.CommandText = query;
                //    cmnd1.ExecuteNonQuery();
                //    cmnd1.Dispose();

                //    //ACTUALIZAR CAT_EMPAQUES SI HAY CAMBIO EN CANTIDAD
                //    if (Convert.ToDecimal(rw["cantidad"].ToString()) != Convert.ToDecimal(rw["cantorig"]))
                //    {
                //        cmnd1 = thisConnection.CreateCommand();
                //        query = "UPDATE tb_cat_empaques SET emp_existencia = emp_existencia - " + rw["cantorig"].ToString() + " WHERE emp_clave = '" + rw["clave"].ToString() + "'";
                //        cmnd1.CommandText = query;
                //        cmnd1.ExecuteNonQuery();
                //        cmnd1.Dispose();
                //    }

                //    //REECALCULO DE COSTO Y EXISTENCIA
                //    costo_promedio2(Convert.ToDateTime(dtpFecha.Text), rw["clave"].ToString());

                //    cmnd1 = thisConnection.CreateCommand();
                //    query = "SELECT emp_costo, emp_existencia, alm_clave FROM tb_cat_empaques WHERE emp_clave = '" + rw["clave"].ToString() + "' ORDER BY emp_clave";
                //    cmnd1.CommandText = query;
                //    reader1 = cmnd1.ExecuteReader();
                //    decimal costoactual = 0;
                //    decimal existencia = 0;
                //    string cvealm = "";
                //    if (reader1.HasRows)
                //    {
                //        reader1.Read();
                //        costoactual = Convert.ToDecimal(reader1.GetValue(0).ToString().Trim());
                //        existencia = Convert.ToDecimal(reader1.GetValue(1).ToString().Trim());
                //        cvealm = reader1.GetValue(2).ToString().Trim();
                //    }
                //    reader1.Close();
                //    reader1.Dispose();
                //    cmnd1.Dispose();

                //    decimal A = 0, B = 0, C = 0, D = 0, promedio = 0;
                //    A = Math.Abs(existencia * costoactual);
                //    B = var_dec_entd_cantidad * Convert.ToDecimal(rw["unitario"].ToString());
                //    C = Math.Abs(existencia) + var_dec_entd_cantidad;
                //    D = (C == 0) ? Convert.ToDecimal(rw["unitario"].ToString()) : (A + B) / C;
                //    C = existencia + var_dec_entd_cantidad;
                //    promedio = Math.Abs(D);
                //    string tip2 = "NOR";
                //    string reg2 = "DEF";
                //    string pro2 = (DDL_ent_procedencia.SelectedItem.ToString() == "Nacional") ? "NAL" : "IMP";

                //    //INSERCION EN HISTORICO RECEPCION
                //    cmnd1 = thisConnection.CreateCommand();
                //    query = "INSERT INTO tb_historico_recepcion(hrp_recibo, hrp_fecha, emp_clave, hrp_cantidad, hrp_estatus, hrp_tipo_recepcion, alm_clave, hrp_situacion," +
                //        " hrp_costo, hrp_regimen, hrp_procedencia, hrp_sobrante, lote_ras, hrp_costo_prom, hrp_existencia, hrp_tipo, hrp_fechamov) VALUES ('" + txtFolio.Text + "', " +
                //        " '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + var_dec_entd_cantidad.ToString() + "', 'T', 'ENT', '" + glo_var_chr_alm_clave + "'," +
                //        " '" + tip2 + "', '" + rw["unitario"].ToString() + "', " +
                //        " '" + reg2 + "', '" + pro2 + "', '" + var_dec_entd_cantidad.ToString() + "', '" + rw["lote"].ToString() + "', '" + promedio.ToString() + "', '" + C.ToString() + "'," +
                //        " '" + ((glo_var_int_registro == 7) ? "COM" : (glo_var_int_registro == 9) ? "GUA" : "") + "', '" + DateTime.Now.ToShortDateString() + "')";
                //    cmnd1.CommandText = query;
                //    cmnd1.ExecuteNonQuery();
                //    cmnd1.Dispose();

                //    //GRABAR ULTIMO COSTO
                //    cmnd1 = thisConnection.CreateCommand();
                //    cmnd1.CommandText = "SELECT emp_entrada_ultima FROM tb_cat_empaques WHERE emp_clave = '" + var_chr_emp_clave + "' AND emp_entrada_ultima < '" + dtpFecha.Text + "'";
                //    reader1 = cmnd1.ExecuteReader();
                //    if (reader1.HasRows)
                //    {
                //        cmnd4 = thisConnection.CreateCommand();
                //        cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "', emp_entrada_ultima = '" + Convert.ToDateTime(dtpFecha.Text).ToShortDateString() + "' , emp_existencia = " + C.ToString() + " WHERE emp_clave = '" + var_chr_emp_clave + "' AND emp_entrada_ultima < '" + dtpFecha.Text + "'";
                //        cmnd4.ExecuteNonQuery();
                //        cmnd4.Dispose();
                //    }
                //    else
                //    {
                //        cmnd4 = thisConnection.CreateCommand();
                //        cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + promedio + "', emp_existencia = " + C.ToString() + " WHERE emp_clave = '" + var_chr_emp_clave + "'";
                //        cmnd4.ExecuteNonQuery();
                //        cmnd4.Dispose();
                //    }
                //    reader1.Close();
                //    reader1.Dispose();
                //    cmnd1.Dispose();

                //    //INSERCION EN TRANSACCION_EMP
                //    cmnd1 = thisConnection.CreateCommand();
                //    query = "INSERT INTO tb_mstr_transaccion_emp(tra_folio, tra_fecha, emp_clave, tra_cantidad, tra_tipo_mov, tra_codigo_mov, tra_codigo_tra) VALUES " +
                //        "('" + txtFolio.Text + "', '" + DateTime.Now.ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + var_dec_entd_cantidad.ToString() + "'," +
                //        " 'E', 'EN', 'A')";
                //    cmnd1.CommandText = query;
                //    cmnd1.ExecuteNonQuery();
                //    cmnd1.Dispose();

                //    //INSERCION ENTRADAS_OC
                //    cmnd1 = thisConnection.CreateCommand();
                //    query = "INSERT INTO tb_det_entrada_oc(ent_folio, numero_oc, emp_clave, ent_cantidad, estatus, alm_clave) VALUES(" +
                //        " '" + txtFolio.Text + "', '" + rw["oc"].ToString() + "', '" + rw["clave"] + "'," +
                //        " '" + rw["cantidad"].ToString() + "', 'A', '" + glo_var_chr_alm_clave + "')";
                //    cmnd1.CommandText = query;
                //    cmnd1.ExecuteNonQuery();
                //    cmnd1.Dispose();

                //    //ACTUALIZAR ORDENESCOMPRA (DISMINUIR CANTIDAD EN CAMPO SURTIDO_OC) POR CAMBIO EN LA CANTIDAD
                //    cmnd1 = thisConnection.CreateCommand();
                //    query = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc - '" + rw["cantorig"].ToString() + "' WHERE numero_oc = '" + rw["oc"].ToString() + "' AND prod_clave = '" + var_chr_emp_clave + "'";
                //    cmnd1.CommandText = query;
                //    cmnd1.ExecuteNonQuery();
                //    cmnd1.Dispose();

                //    //ACTUALIZAR ORDENESCOMPRA (AUMENTAR CANTIDAD EN CAMPO SURTIDO_OC) NUEVA CANTIDAD
                //    cmnd1 = thisConnection.CreateCommand();
                //    query = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc + '" + var_dec_entd_cantidad.ToString() + "' WHERE numero_oc = '" + rw["oc"].ToString() + "' AND prod_clave = '" + var_chr_emp_clave + "'";
                //    cmnd1.CommandText = query;
                //    cmnd1.ExecuteNonQuery();
                //    cmnd1.Dispose();

                //    //VERIFICAR SI EMPAQUE ES DE UNA REQUISICION
                //    if (rw["req"].ToString() != "")
                //    {
                //        //DISMINUIR CANTIDAD ORIGINAL
                //        cmnd1 = thisConnection.CreateCommand();
                //        query = "UPDATE tb_det_requisiciones SET req_surtido = req_surtido - '" + rw["cantorig"] + "' WHERE req_folio = '" + rw["req"].ToString() + "' AND emp_clave = '" + rw["clave"].ToString() + "'";
                //        cmnd1.CommandText = query;
                //        cmnd1.ExecuteNonQuery();
                //        cmnd1.Dispose();

                //        //AUMENTAR CANTIDAD NUEVA
                //        cmnd1 = thisConnection.CreateCommand();
                //        query = "UPDATE tb_det_requisiciones SET req_surtido = req_surtido + '" + var_dec_entd_cantidad + "' WHERE req_folio = '" + rw["req"].ToString() + "' AND emp_clave = '" + rw["clave"].ToString() + "'";
                //        cmnd1.CommandText = query;
                //        cmnd1.ExecuteNonQuery();
                //        cmnd1.Dispose();
                //    }

                //}
                #endregion
                thisConnection.Close();

                //costo_promedio2(Convert.ToDateTime(dtpFecha.Text));
                MessageBox.Show("Entrada modificada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (MessageBox.Show("Imprimir Entrada", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    btnImprimir_Click(null, null);
            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de modificar la entrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.5", txtFolio.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }
        }

        public void cancelar()
        {
            if (Convert.ToDateTime(dtpFecha.Text) < Convert.ToDateTime(txtCierre.Text))
            {
                MessageBox.Show("No de puede cancelar movimientos anteriores al cierre del mes", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (MessageBox.Show("¿Desea realizar la cancelacion de la entrada?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                //subtipo = "TOM";
                string query = "";
                try
                {
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    query = "UPDATE tb_mstr_entradas SET ent_estatus = 'C' WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_subtipo = '" + subtipo + "'";
                    cmnd1.CommandText = query;
                    cmnd1.ExecuteNonQuery();
                    //thisConnection.Close()

                    //ACTUALIZAR REGISTRO DE HISTORICO
                    cmnd1 = thisConnection.CreateCommand();
                    query = "UPDATE tb_historico_recepcion SET hrp_estatus = 'C' WHERE hrp_recibo = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND hrp_tipo_recepcion = 'ENT' AND hrp_tipo = '" + subtipo + "'";
                    //" AND hrp_situacion = 'NOR'";
                    cmnd1.CommandText = query;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();

                    //CANCELAR REGISTRO EN TB_DET_ENT_LOTE
                    cmnd1 = thisConnection.CreateCommand();
                    query = "UPDATE tb_det_ent_lote SET estatus = 'C' WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND tipo = 'NOR'";
                    //" AND hrp_situacion = 'NOR'";
                    cmnd1.CommandText = query;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();

                    //CANCELAR REGISTRO EN TB_DET_ENT_LOTE
                    string a = "";
                    if (glo_var_chr_alm_clave == "01")
                        a = "MATERIALES";
                    if (glo_var_chr_alm_clave == "02")
                        a = "EMPAQUES";
                    if (glo_var_chr_alm_clave == "03" && subtipo == "COM")
                        a = "COMEDOR";
                    if (glo_var_chr_alm_clave == "03" && subtipo == "GUA")
                        a = "GUARDERIA";
                    if (glo_var_chr_alm_clave == "04")
                        a = "ENFERMERIA";
                    if (glo_var_chr_alm_clave == "05")
                        a = "PAPELERIA";
                    if (glo_var_chr_alm_clave == "07")
                        a = "DENTAL";
                    if (glo_var_chr_alm_clave == "08")
                        a = "COMBUSTIBLE";
                    if (glo_var_chr_alm_clave == "09")
                        a = "SISTEMAS";
                    if (glo_var_chr_alm_clave == "10")
                        a = "GUARDERIAALM";
                    if (glo_var_chr_alm_clave == "11")
                        a = "GUARDERIACOM";
                    if (glo_var_chr_alm_clave == "12")
                        a = "EVENTOS";
                    if (glo_var_chr_alm_clave == "13")
                        a = "MOTORES";
                    cmnd1 = thisConnection.CreateCommand();
                    query = "UPDATE tb_det_oc_facturas SET estatus = 'C' WHERE numero_oc = '" + txtFolio.Text.TrimStart('0') + "' AND tipo = '" + a + "' AND estatus = 'A'";
                    //" AND hrp_situacion = 'NOR'";
                    cmnd1.CommandText = query;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();

                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT lote_ras, emp_clave FROM tb_det_ent_lote WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND tipo = 'NOR'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            cmnd2 = thisConnection.CreateCommand();
                            cmnd2.CommandText = "UPDATE tb_det_lote_ras SET estatus = 'C' WHERE lote_ras = '" + reader1["lote_ras"].ToString().Trim() + "' AND empaque = '" + reader1["emp_clave"].ToString().Trim() + "'";
                            cmnd2.ExecuteNonQuery();
                            cmnd2.Dispose();
                        }
                    }

                    //RESTAR ENTRADAS A LA EXISTENCIA DEL EMPAQUE
                    cmnd3 = thisConnection.CreateCommand();
                    foreach (DataRow re in dtoriginales.Rows)
                    {
                        cmnd3.CommandText = "UPDATE tb_cat_empaques SET emp_existencia = emp_existencia - '" + re["cantidad"].ToString() + "' WHERE emp_clave = '" + re["emp_clave"].ToString() + "'";
                        cmnd3.ExecuteNonQuery();

                        query = "INSERT INTO tb_mstr_transaccion_emp(tra_folio, tra_fecha, emp_clave, tra_cantidad, tra_tipo_mov, tra_codigo_mov, tra_codigo_tra) VALUES " +
                            "('" + txtFolio.Text + "', '" + DateTime.Now.ToShortDateString() + "', '" + re["emp_clave"].ToString() + "', '" + re["cantidad"].ToString() + "'," +
                            " 'S', 'EN', 'B')";
                        cmnd3.CommandText = query;
                        cmnd3.ExecuteNonQuery();

                        //ACTUALIZAR ENTRADA_OC
                        query = "UPDATE tb_det_entrada_oc SET estatus = 'C' WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' " +
                            "AND emp_clave = '" + re["emp_clave"].ToString() + "' AND numero_oc = '" + re["orden_compra"].ToString() + "' AND ent_subtipo = '" + subtipo + "'";
                        cmnd3.CommandText = query;
                        cmnd3.ExecuteNonQuery();

                        //VERIFICAMOS SI VIENE DE UNA REQUISICION PARA QUITAR LA CANTIDAD SURTIDA
                        if (re["requisicion"].ToString() != "")
                        {
                            if (DDLAlmacen1.SelectedIndex == 4)
                                query = "UPDATE tb_det_requisiciones_enf_pap SET req_surtido = req_surtido - '" + re["cantidad"].ToString() + "' WHERE req_folio = '" + re["requisicion"].ToString() + "' AND emp_clave = '" + re["emp_clave"].ToString() + "' AND req_tipo = 'ENF'";
                            else if(DDLAlmacen1.SelectedIndex == 5)
                                query = "UPDATE tb_det_requisiciones_enf_pap SET req_surtido = req_surtido - '" + re["cantidad"].ToString() + "' WHERE req_folio = '" + re["requisicion"].ToString() + "' AND emp_clave = '" + re["emp_clave"].ToString() + "' AND req_tipo = 'PAP'";
                            else
                                query = "UPDATE tb_det_requisiciones SET req_surtido = req_surtido - '" + re["cantidad"].ToString() + "' WHERE req_folio = '" + re["requisicion"].ToString() + "' AND emp_clave = '" + re["emp_clave"].ToString() + "'";
                            cmnd3.CommandText = query;
                            cmnd3.ExecuteNonQuery();
                        }

                        //ACTUALIZA DET_ORDENESCOMPRA QUITA LA CANTIDAD SURTIDA
                        query = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc - '" + re["cantidad"].ToString() + "' WHERE numero_oc = '" + re["orden_compra"].ToString() + "' AND prod_clave = '" + re["emp_clave"].ToString() + "'";
                        cmnd3.CommandText = query;
                        cmnd3.ExecuteNonQuery();
                    }
                    cmnd3.Dispose();
                    thisConnection.Close();
                    

                    

                    btnBuscar.Enabled = false;
                    btnCancelar.Enabled = false;
                    btnComedor.Enabled = false;
                    btnEditar.Enabled = false;
                    btnGuardar.Enabled = false;
                    btnImprimir.Enabled = false;

                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "B", "2.4", txtFolio.Text, "CANCELACION DE ENTRADA(OC): " + txtFolio.Text + ", ALMACEN: " + glo_var_chr_alm_clave, "SISEMP");

                    string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                    using (StreamWriter sw = File.AppendText(filelog))
                    {
                        sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Cancelación de entrada(oc): " + txtFolio.Text + ", Almacen: " + glo_var_chr_alm_clave);
                        sw.Close();
                    }

                    //recalcular existencias y costos
                    foreach (DataRow re in dtoriginales.Rows)
                    {
                        string e_m_p_a_q_u_e = re["emp_clave"].ToString().Trim();
                        string f_e_c_h_a = dtpFecha.Text;

                        //recalculo_costos_empaque(e_m_p_a_q_u_e, f_e_c_h_a);
                        //recalculo_existencia_empaque(e_m_p_a_q_u_e);
                        recalculo_costos(e_m_p_a_q_u_e);
                        recalculo_existencia(e_m_p_a_q_u_e);
                        modificar_ultima_entrada(e_m_p_a_q_u_e);
                    }

                    MessageBox.Show("Entrada Cancelada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    btnLimpiar_Click(null, null);
                }
                catch (SqlException sqlex)
                {
                    MessageBox.Show("Error de sistema, no se termino de modificar la entrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (thisConnection.State == ConnectionState.Open)
                        thisConnection.Close();
                    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.5", txtFolio.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + sqlex.ToString());
                    //this.Close();                                
                    return;
                }
                #region comentada
                ////INICIO DE CANCELACION DE ENTRADA
                //string query = "";
                //string estatus = "";
                //try
                //{
                //    thisConnection.Open();
                //    //VERIFICAR ESTATUS DE ENTRADA
                //    cmnd1 = thisConnection.CreateCommand();
                //    query = "SELECT ent_estatus FROM tb_mstr_entradas WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_tipo = 'NOR'";
                //    cmnd1.CommandText = query;
                //    reader1 = cmnd1.ExecuteReader();
                //    if (reader1.HasRows)
                //    {
                //        while (reader1.Read())
                //        {
                //            if (reader1.GetValue(0).ToString().Trim() == "C")
                //            {
                //                MessageBox.Show("La entrada fué cancelada con anterioridad", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //                thisConnection.Close();
                //                return;
                //            }
                //        }
                //    }
                //    reader1.Close();
                //    reader1.Dispose();
                //    cmnd1.Dispose();
                //    //CANCELAR EN MSTR_ENTRADAS
                //    cmnd1 = thisConnection.CreateCommand();
                //    query = "UPDATE tb_mstr_entradas SET ents_estatus = 'C' WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
                //    cmnd1.CommandText = query;
                //    cmnd1.ExecuteNonQuery();
                //    cmnd1.Dispose();

                //    string var_chr_empaque = "";
                //    bool var_bool_bandera = false;
                //    bool found = false;
                //    decimal var_dec_hrp_costo = 0;
                //    //SE REGRESAN LOS COSTOS DE LOS PRODUCTOS IGUALES A LA ULTIMA ENTRADA
                //    foreach (DataRow rw in tmgrdpro.Rows)
                //    {
                //        var_chr_empaque = rw["clave"].ToString();
                //        var_bool_bandera = false;
                //        cmnd1 = thisConnection.CreateCommand();
                //        query = "SELECT TOP 1 hrp_costo FROM tb_historico_recepcion WHERE hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR' AND hrp_estatus <> 'C' AND emp_clave = '" + var_chr_empaque + "' ORDER BY hrp_fecha DESC, hrp_tipo_recepcion, hrp_situacion, emp_clave";
                //        cmnd1.CommandText = query;
                //        reader1 = cmnd1.ExecuteReader();
                //        if (reader1.HasRows)
                //        {
                //            while (reader1.Read())
                //            {
                //                var_dec_hrp_costo = Convert.ToDecimal(reader1.GetValue(0).ToString().Trim());
                //                var_bool_bandera = true;
                //                found = true;
                //            }
                //        }
                //        reader1.Close();
                //        reader1.Dispose();
                //        cmnd1.Dispose();
                //        if (found == false)
                //        {
                //            cmnd1 = thisConnection.CreateCommand();
                //            query = "SELECT TOP 1 hrp_costo FROM tb_historico_recepcion WHERE hrp_tipo_recepcion = 'ENT' AND hrp_estatus <> 'C' AND emp_clave = '" + var_chr_empaque + "' ORDER BY hrp_fecha DESC, hrp_tipo_recepcion, hrp_situacion, emp_clave";
                //            cmnd1.CommandText = query;
                //            reader1 = cmnd1.ExecuteReader();
                //            if (reader1.HasRows)
                //            {
                //                while (reader1.Read())
                //                {
                //                    var_dec_hrp_costo = Convert.ToDecimal(reader1.GetValue(0).ToString().Trim());
                //                    var_bool_bandera = true;
                //                }
                //            }
                //            reader1.Close();
                //            reader1.Dispose();
                //            cmnd1.Dispose();
                //        }
                //        if (var_bool_bandera == true)
                //        {
                //            cmnd1 = thisConnection.CreateCommand();
                //            query = "UPDATE tb_cat_empaques SET emp_costo = '" + var_dec_hrp_costo + "' WHERE emp_clave = '" + var_chr_empaque + "'";
                //            cmnd1.CommandText = query;
                //            cmnd1.ExecuteNonQuery();
                //            cmnd1.Dispose();
                //        }
                //    }

                //    //SI YA SE AFECTO EL INVENTARIO, SE RESTAN LAS ENTRADAS CORRESPONDIENTES
                //    string mtipo = "";
                //    mtipo = (glo_var_int_registro == 7) ? " AND hrp_tipo = 'COM'" : (glo_var_int_registro == 9) ? " AND hrp_tipo = 'GUA'" : "";
                //    cmnd1 = thisConnection.CreateCommand();
                //    if (mtipo != "")
                //    {
                //        query = "SELECT emp_clave, hrp_cantidad, hrp_estatus FROM tb_historico_recepcion WHERE hrp_tipo_recepcion = 'ENT' AND alm_clave = '" + glo_var_chr_alm_clave + "'" +
                //            " AND hrp_recibo = '" + txtFolio.Text + "'" + mtipo + " ORDER BY hrp_tipo_recepcion, alm_clave, hrp_tipo, hrp_recibo, emp_clave";
                //    }

                //    else
                //    {
                //        query = "SELECT emp_clave, hrp_cantidad, hrp_estatus FROM tb_historico_recepcion WHERE hrp_tipo_recepcion = 'ENT' AND alm_clave = '" + glo_var_chr_alm_clave + "'" +
                //            " AND hrp_recibo = '" + txtFolio.Text + "'" + mtipo + " ORDER BY hrp_tipo_recepcion, alm_clave, hrp_recibo, emp_clave";
                //    }
                //    cmnd1.CommandText = query;
                //    reader1 = cmnd1.ExecuteReader();
                //    string var_chr_emp_clave = "";
                //    decimal var_dec_hrp_cantidad = 0;
                //    if (reader1.HasRows)
                //    {
                //        while (reader1.Read())
                //        {
                //            var_chr_emp_clave = reader1.GetValue(0).ToString().Trim();
                //            var_dec_hrp_cantidad = Convert.ToDecimal(reader1.GetValue(1).ToString().Trim());

                //            cmnd2 = thisConnection.CreateCommand();
                //            cmnd2.CommandText = "UPDATE tb_historico_recepcion SET hrp_estatus = 'C' WHERE hrp_tipo_recepcion = 'ENT' AND alm_clave = '" + glo_var_chr_alm_clave + "'" +
                //                " AND hrp_recibo = '" + txtFolio.Text + "' AND emp_clave = '" + var_chr_emp_clave + "'" + mtipo;
                //            cmnd2.ExecuteNonQuery();
                //            cmnd2.Dispose();

                //            //INSERCION EN TRANSACCION_EMP CANCELACIONES
                //            cmnd2 = thisConnection.CreateCommand();
                //            cmnd2.CommandText = "INSERT INTO tb_mstr_transaccion_emp(tra_folio, tra_fecha, emp_clave, tra_cantidad, tra_tipo_mov, tra_codigo_mov, tra_codigo_tra) VALUES " +
                //                "('" + txtFolio.Text + "', '" + DateTime.Now.ToShortDateString() + "', '" + var_chr_emp_clave + "', '" + var_dec_hrp_cantidad.ToString() + "'," +
                //                " 'S', 'EN', 'B')";
                //            cmnd2.ExecuteNonQuery();
                //            cmnd2.Dispose();
                //        }
                //    }
                //    reader1.Close();
                //    reader1.Dispose();
                //    cmnd1.Dispose();

                //    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "B", "2.5", txtFolio.Text, "CANCELACION DE ENTRADA: " + txtFolio.Text + ", ALMACEN: " + glo_var_chr_alm_clave, "SISEMP");

                //    string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                //    using (StreamWriter sw = File.AppendText(filelog))
                //    {
                //        sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Cancelación de entrada: " + txtFolio.Text + ", Almacen: " + glo_var_chr_alm_clave);
                //        sw.Close();
                //    }

                //    thisConnection.Close();

                //    MessageBox.Show("Entrada Cancelada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
                //}
                //catch (SqlException sqlex)
                //{
                //    MessageBox.Show("Error de sistema, no se termino de modificar la entrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    if (thisConnection.State == ConnectionState.Open)
                //        thisConnection.Close();
                //    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.5", txtFolio.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                //    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + sqlex.ToString());
                //    //this.Close();                                
                //    return;
                //}
                #endregion
            }
        }

        public void imprimir()
        {
            imprimirtxt();
        }

        public bool validacontroles()
        {
            bool bandera = true;
            if (DDLElaboro.SelectedIndex == -1 || DDLAutorizo.SelectedIndex == -1)
            {
                MessageBox.Show("No ha seleccionado campo Elaboró y/o Autorizó", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bandera = false;
            }

            if (DDLProveedor.SelectedIndex == -1)
            {
                MessageBox.Show("No ha seleccionado proveedor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bandera = false;
            }

            if (dtgDetalle.Rows.Count == 0)
            {
                MessageBox.Show("No se han introducido productos", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bandera = false;
            }

            if (Convert.ToDateTime(dtpFecha.Text) < Convert.ToDateTime("01/01/2000"))
            {
                MessageBox.Show("La fecha no es válida, favor de verificar", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bandera = false;
            }

            decimal n1 = 0;
            try
            {
                n1 = Convert.ToDecimal(txt_sis_tipo_cambio.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("El valor introducido en tipo cambio no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bandera = false;
            }

            decimal n2 = 0;
            try
            {
                n2 = Convert.ToDecimal(txt_gto_aduanal.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("El valor introducido en gasto aduanal no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bandera = false;
            }
            if (dtgDetalle.Rows.Count == 0)
            {
                MessageBox.Show("No hay registros para guardar, favor de seleccionar la orden de compra requerida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                bandera = false;
            }
            return bandera;
        }
        public void imprimirtxt()
        {
            TextWriter tw = new StreamWriter("c:\\empaque\\Gab\\pruebafile3.txt");

            tw.Write("\t\t\t\t\t\t\t\t\t" + txtFolio.Text + " " + txtTipo.Text);
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\t\t\t\t\t\t\t\t" + dtpFecha.Text);
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\n"); tw.WriteLine();
            tw.WriteLine();
            tw.Write("\t" + DDLProveedor.SelectedItem.ToString());
            tw.Write("\t" + txt_ent_factura.Text);
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\t\t\t\t\t\t\tLote Rastreabilidad");
            tw.WriteLine();

            for (int x = 0; x < dtgDetalle.Rows.Count - 1; x++)
            {
                if (DDL_ent_procedencia.SelectedItem.ToString() == "Nacional")
                {
                    tw.Write("\t");
                    tw.Write(Convert.ToDecimal(dtgDetalle.Rows[x].Cells[2].Value).ToString("###,###,##0.00"));
                    tw.Write("\t");
                    tw.Write((dtgDetalle.Rows[x].Cells[0].Value + " " + dtgDetalle.Rows[x].Cells[1].Value).PadRight(45, ' '));
                    tw.Write("\t");
                    tw.Write(dtgDetalle.Rows[x].Cells[7].Value);
                    tw.Write("\t");
                    tw.Write(Convert.ToDecimal(dtgDetalle.Rows[x].Cells[4].Value).ToString("###,###,##0.00"));
                    tw.Write("\t");
                    tw.Write(Convert.ToDecimal(dtgDetalle.Rows[x].Cells[5].Value).ToString("###,###,##0.000"));
                    tw.WriteLine();
                }
                else
                {
                    tw.Write("\t");
                    tw.Write(Convert.ToDecimal(dtgDetalle.Rows[x].Cells[2].Value).ToString("###,###,##0.00"));
                    tw.Write("\t");
                    tw.Write((dtgDetalle.Rows[x].Cells[0].Value + " " + dtgDetalle.Rows[x].Cells[1].Value).PadRight(45, ' '));
                    tw.Write("\t");
                    tw.Write(dtgDetalle.Rows[x].Cells[7].Value);
                    tw.Write("\t");
                    tw.Write(Convert.ToDecimal(dtgDetalle.Rows[x].Cells[4].Value).ToString("###,###,##0.00"));
                    tw.Write("\t");
                    tw.Write((Convert.ToDecimal(dtgDetalle.Rows[x].Cells[3].Value)).ToString("###,###,##0.000"));
                    tw.WriteLine();
                }

            }

            tw.Write("\n"); tw.WriteLine();
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\n"); tw.WriteLine();

            tw.Write("\t\t\t\t\t\t\t\tSubtotal:\t" + Convert.ToDecimal(txt_subtotal2.Text).ToString("###,###,##0.00"));
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\t\t\t\t\t\t\t\tI.V.A.   \t" + Convert.ToDecimal(txt_iva.Text).ToString("###,###,##0.00"));
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\t\t\t\t\t\t\t\tTotal:   \t" + Convert.ToDecimal(txt_total.Text).ToString("###,###,##0.00"));
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\t" + txtObservaciones.Text);

            tw.Write("\n"); tw.WriteLine();
            tw.Write("\n"); tw.WriteLine();
            tw.Write("\t" + DDLElaboro.SelectedItem.ToString());
            tw.Write("\t" + DDLAutorizo.SelectedItem.ToString());
            tw.Write("\t" + ((DDLElaboro.SelectedIndex == -1) ? "" : DDLAutorizo.SelectedItem.ToString()));
            tw.WriteLine();
            tw.Close();

            Process.Start("notepad.exe", "c:\\empaque\\Gab\\pruebafile3.txt");
        }
        public void verificarfolio()
        {
            //thisConnection.Open();
            //cmnd1 = thisConnection.CreateCommand();
            //cmnd1.CommandText = "SELECT ent_folio FROM tb_mstr_entradas WHERE ent_folio = '" + txtFolio.Text.Trim() + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'  ORDER BY ent_folio DESC";
            //if (DDLAlmacen1.SelectedIndex == 2)
            //{
            //    cmnd1.CommandText = "SELECT ent_folio FROM tb_mstr_entradas WHERE ent_folio = '" + txtFolio.Text.Trim() + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_subtipo = 'COM'  ORDER BY ent_folio DESC";
            //}
            //if (DDLAlmacen1.SelectedIndex == 3)
            //{
            //    cmnd1.CommandText = "SELECT ent_folio FROM tb_mstr_entradas WHERE ent_folio = '" + txtFolio.Text.Trim() + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_subtipo = 'GUA'  ORDER BY ent_folio DESC";
            //}
            //reader1 = cmnd1.ExecuteReader();
            //if (reader1.HasRows)
            //{
            //    if (DDLAlmacen1.SelectedIndex == 0)//Materiales
            //    {
            //        //volver a cargar el folio
            //        cmnd2 = thisConnection.CreateCommand();
            //        cmnd2.CommandText = "SELECT TOP 1 ent_folio FROM tb_mstr_entradas WHERE alm_clave = '" + glo_var_chr_alm_clave + "' ORDER BY ent_fecha DESC, ent_folio DESC";
            //        reader2 = cmnd2.ExecuteReader();
            //        reader2.Read();
            //        var_int_sis_folios = Convert.ToInt32(reader2.GetValue(0).ToString().Trim()) + 1;
            //        reader2.Close();
            //        reader2.Dispose();
            //        cmnd2.Dispose();
            //        txtFolio.Text = Convert.ToString(var_int_sis_folios).PadLeft(6, '0');
            //        MessageBox.Show("Debido al trabajo en red: el folio de orden de salida ha cambiado a " + txtFolio.Text.Trim(), "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //        subtipo = "";
            //    }
            //    if (DDLAlmacen1.SelectedIndex == 1)//Empaques
            //    {
            //        //volver a cargar el folio
            //        cmnd2 = thisConnection.CreateCommand();
            //        cmnd2.CommandText = "SELECT TOP 1 ent_folio FROM tb_mstr_entradas WHERE alm_clave = '" + glo_var_chr_alm_clave + "' ORDER BY ent_fecha DESC, ent_folio DESC";
            //        reader2 = cmnd2.ExecuteReader();
            //        reader2.Read();
            //        var_int_sis_folios = Convert.ToInt32(reader2.GetValue(0).ToString().Trim()) + 1;
            //        reader2.Close();
            //        reader2.Dispose();
            //        cmnd2.Dispose();
            //        txtFolio.Text = Convert.ToString(var_int_sis_folios).PadLeft(6, '0');
            //        MessageBox.Show("Debido al trabajo en red: el folio de orden de salida ha cambiado a " + txtFolio.Text.Trim(), "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //        subtipo = "";
            //    }
            //    if (DDLAlmacen1.SelectedIndex == 2)//Comedor
            //    {
            //        //volver a cargar el folio
            //        cmnd2 = thisConnection.CreateCommand();
            //        cmnd2.CommandText = "SELECT TOP 1 ent_folio FROM tb_mstr_entradas WHERE alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_subtipo = 'COM' ORDER BY ent_fecha DESC, ent_folio DESC";
            //        reader2 = cmnd1.ExecuteReader();
            //        reader2.Read();
            //        var_int_sis_folios = Convert.ToInt32(reader2.GetValue(0).ToString().Trim()) + 1;
            //        reader2.Close();
            //        reader2.Dispose();
            //        cmnd2.Dispose();
            //        txtFolio.Text = Convert.ToString(var_int_sis_folios).PadLeft(6, '0');
            //        subtipo = "COM";
            //        MessageBox.Show("Debido al trabajo en red: el folio de orden de salida ha cambiado a " + txtFolio.Text.Trim(), "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    }
            //    if (DDLAlmacen1.SelectedIndex == 3)//Guarderia
            //    {
            //        //volver a cargar el folio
            //        cmnd2 = thisConnection.CreateCommand();
            //        cmnd2.CommandText = "SELECT TOP 1 ent_folio FROM tb_mstr_entradas WHERE alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_subtipo = 'GUA' ORDER BY ent_fecha DESC, ent_folio DESC";
            //        reader2 = cmnd2.ExecuteReader();
            //        reader2.Read();
            //        var_int_sis_folios = Convert.ToInt32(reader2.GetValue(0).ToString().Trim()) + 1;
            //        reader2.Close();
            //        reader2.Dispose();
            //        cmnd2.Dispose();
            //        txtFolio.Text = Convert.ToString(var_int_sis_folios).PadLeft(6, '0');
            //        subtipo = "GUA";
            //        MessageBox.Show("Debido al trabajo en red: el folio de orden de salida ha cambiado a " + txtFolio.Text.Trim(), "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    }

            //}
            //reader1.Close();
            //reader1.Dispose();
            //cmnd1.Dispose();
            //thisConnection.Close();
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            if (DDLAlmacen1.SelectedIndex == -1)
            {
                MessageBox.Show("No ha seleccionado el almacen", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            panel2.Height = 70;
            txtBuscar.Focus();
        }

        private void txtBuscar_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                //VERIFICAR TIPO DE ENTRADA
                string tipo_ent = "";
                string sin_fact = "";
                string fecha_ent = "";
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT ent_tipo, ent_sin_fact, ent_fecha FROM tb_mstr_entradas WHERE alm_clave = '" + glo_var_chr_alm_clave + "' AND ent_folio = '" + txtBuscar.Text.PadLeft(6, '0') + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        tipo_ent = reader1.GetValue(0).ToString().Trim();
                        sin_fact = reader1["ent_sin_fact"].ToString().Trim();
                        fecha_ent = Convert.ToDateTime(reader1["ent_fecha"].ToString().Trim()).ToShortDateString();
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();

                if (tipo_ent == "DEV")
                {
                    MessageBox.Show("Entrada por Devolución, no  puede consultarla, use pantalla alterna", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                bool no_valida = false;
                if ((Convert.ToDateTime(lblFechaDia.Text) - Convert.ToDateTime(fecha_ent)).Days > 7)
                {
                    MessageBox.Show("La Entrada a consultar tiene más de 7 días, no es valida para modificación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    no_valida = true;
                    //return;
                    //btnEditar.Enabled = false;
                }
                

                operacion = "consulta";
                btnLimpiar_Click(null, null);
                string fol = txtBuscar.Text.PadLeft(6, '0');
                string mtipo = "";
                string qryex = "";
                if (glo_var_int_registro == 7 || glo_var_int_registro == 9)
                {
                    mtipo = (glo_var_int_registro == 7) ? "COM" : "GUA";
                    qryex = " AND ent_subtipo = '" + mtipo + "'";
                }
                thisConnection.Open();
                //VALIDACION DE ENTRADA POR ORDEN DE COMPRA
                string qry = "";
                cmnd1 = thisConnection.CreateCommand();
                qry = "SELECT COUNT(ent_folio) AS total FROM tb_det_entrada_oc WHERE ent_folio = '" + fol + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
                cmnd1.CommandText = qry;
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows == false)
                {
                    MessageBox.Show("La entrada no puede ser consultada porque no fué realizada por medio de Orden de Compra", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtBuscar.Height = 0;
                    txtBuscar.Text = "";
                    return;
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                bool entro = false;
                
                //CARGAR DATOS EN DETALLE
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT emp_clave, numero_oc, ent_cantidad, ent_importe, req_folio, ent_lote, ent_iva, ent_importe_orig, urgente_oc, cant_orig, CAST(resp AS INT) AS resp FROM tb_det_entrada_oc WHERE ent_folio = '" + fol + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' ";
                reader1 = cmnd1.ExecuteReader();
                dtdetalles.Clear();
                dtoriginales.Clear();
                DataRow rw;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rw = dtdetalles.NewRow();
                        rw["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                        rw["orden_compra"] = reader1.GetValue(1).ToString().Trim();
                        rw["cantidad"] = reader1.GetValue(2).ToString().Trim();
                        rw["importe"] = reader1.GetValue(3).ToString().Trim();
                        rw["requisicion"] = reader1.GetValue(4).ToString().Trim();
                        decimal costo = 0;
                        costo = Convert.ToDecimal(reader1.GetValue(3).ToString().Trim()) / Convert.ToDecimal(reader1.GetValue(2).ToString().Trim());
                        rw["costo"] = costo.ToString("0.00000");
                        rw["ieps"] = "0.0000";
                        rw["lote"] = reader1["ent_lote"].ToString().Trim();
                        rw["iva"] = reader1["ent_iva"].ToString().Trim();
                        rw["tiva"] = (Convert.ToDecimal(reader1["ent_iva"].ToString().Trim()) > 0) ? "1" : "0";
                        rw["importe_orig"] = reader1["ent_importe_orig"].ToString().Trim();
                        rw["urgente"] = reader1["urgente_oc"].ToString().Trim();
                        rw["cant_orig"] = reader1["cant_orig"].ToString().Trim();
                        rw["resp"] = reader1["resp"].ToString().Trim();
                        dtdetalles.Rows.Add(rw);

                        rw = dtoriginales.NewRow();
                        rw["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                        rw["orden_compra"] = reader1.GetValue(1).ToString().Trim();
                        rw["cantidad"] = reader1.GetValue(2).ToString().Trim();
                        rw["importe"] = reader1.GetValue(3).ToString().Trim();
                        rw["requisicion"] = reader1.GetValue(4).ToString().Trim();
                        rw["costo"] = costo.ToString("0.00000");
                        rw["ieps"] = "0.0000";
                        rw["iva"] = reader1["ent_iva"].ToString().Trim();
                        dtoriginales.Rows.Add(rw);
                    }
                }
                cmnd1.Dispose();

                
                string user = "";
                
                cmnd1 = thisConnection.CreateCommand();
                string query = "SELECT ent_folio, ent_fecha, prov_clave, ent_observaciones, ent_elaboro, ent_autorizo, ent_recibio, ent_tipo, ent_factura, ent_regimen, ent_procedencia, ent_pedimento, sis_tipo_cambio, " +//12
                    "ent_iva, ent_estatus, ent_gastoaduanal, ent_iepstotal, ent_flete, ent_flete_imp, ent_sin_fact, ent_fechafact, ent_foliofiscal, ent_alta, ent_mod, descuento, ret_iva, ret_isr, cedular, ent_monto, ent_desc_pp, ent_desc_im FROM tb_mstr_entradas WHERE ent_folio = '" + fol + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'" + qryex;
                cmnd1.CommandText = query;
                reader1 = cmnd1.ExecuteReader();
                int proce = -1;
                string tipo = "";
                string s_fact = "";
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        if (reader1.GetValue(14).ToString().Trim() == "C")
                        {
                            MessageBox.Show("Entrada cancelada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            lblEstatus.Text = "CANCELADA";
                            lblEstatus.ForeColor = Color.Red;

                            btnBuscar.Enabled = false;
                            btnCancelar.Enabled = false;
                            btnComedor.Enabled = false;
                            btnEditar.Enabled = false;
                            btnGuardar.Enabled = false;
                            btnImprimir.Enabled = false;
                            btnCambioFactura.Enabled = false;
                        }

                        txtFolio.Text = fol;
                        dtpFecha.Text = reader1.GetValue(1).ToString().Trim();
                        txt_prov_clave.Text = reader1.GetValue(2).ToString().Trim();
                        txtObservaciones.Text = reader1.GetValue(3).ToString().Trim();

                        user = reader1.GetValue(4).ToString().Trim();
                        for (int i = 0; i < usuarionombre.Count; i++)
                        {
                            if (usuarionombre[i].ToString() == reader1.GetValue(4).ToString().Trim())
                            {
                                DDLElaboro.SelectedIndex = i;
                                break;
                            }
                        }
                        for (int i = 0; i < usuarionombre.Count; i++)
                        {
                            if (usuarionombre[i].ToString() == reader1.GetValue(5).ToString().Trim())
                            {
                                DDLAutorizo.SelectedIndex = i;
                                break;
                            }
                        }

                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT usu_puesto FROM tb_cat_usuarios_emp WHERE usu_nombre = '" + user + "' AND usu_puesto = 'AGUILARES'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            reader2.Read();
                            lblAguilares.Text = "AGUILARES";
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        tipo = reader1.GetValue(7).ToString().Trim();
                        //txtTipo.Text = (reader1.GetValue(7).ToString().Trim() == "NOR") ? "Normal" : "";
                        txt_ent_factura.Text = reader1.GetValue(8).ToString().Trim();
                        txtRegimen.Text = (reader1.GetValue(9).ToString().Trim() == "DEF") ? "Definitivo" : "";
                        proce = (reader1.GetValue(10).ToString().Trim() == "NAL") ? 0 : 1;
                        //string pro = reader1.GetValue(10).ToString().Trim();
                        //if (pro == "NAL")
                        //    DDL_ent_procedencia.Text = "Nacional";
                        //else
                        //    DDL_ent_procedencia.Text = "Importación";
                        //DDL_ent_procedencia.SelectedIndex = (reader1.GetValue(10).ToString().Trim() == "NAL") ? 0 : 1;
                        txt_ent_pedimento.Text = reader1.GetValue(11).ToString().Trim();
                        txt_sis_tipo_cambio.Text = reader1.GetValue(12).ToString().Trim();
                        //16 octubre 2023 txt_ent_iva.Text = reader1.GetValue(13).ToString().Trim();
                        txt_gto_aduanal.Text = reader1.GetValue(15).ToString();
                        txt_ent_ieps.Text = reader1.GetValue(16).ToString();
                        txtFlete.Text = reader1.GetValue(17).ToString();
                        txtFleteImp.Text = reader1.GetValue(18).ToString();
                        //if (reader1["ent_sin_fact"].ToString().Trim() == "1")
                        //    chkFactura.Checked = true;
                        //else
                        //    chkFactura.Checked = false;
                        s_fact = reader1["ent_sin_fact"].ToString().Trim();
                        //dtpFechaFact.Text = Convert.ToDateTime(reader1["ent_fechafact"].ToString().Trim()).ToShortDateString();
                        txtFolioFiscal.Text = reader1["ent_foliofiscal"].ToString().Trim();
                        txtFechaFact.Text = (reader1["ent_fechafact"].ToString().Trim() != "") ? Convert.ToDateTime(reader1["ent_fechafact"].ToString().Trim()).ToShortDateString() : "";

                        if (txtFolioFiscal.Text != "")
                            btnCambioFactura.Enabled = true;


                        string importacion = "0";
                        
                        SqlCommand cmndimportacion = thisConnection.CreateCommand();
                        cmndimportacion.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
                        try
                        {
                            importacion = Convert.ToString(cmndimportacion.ExecuteScalar().ToString().Trim());
                        }
                        catch
                        {
                            importacion = "";
                        }

                        cmndimportacion.Dispose();
                        

                        if (importacion == "1")
                        {
                        /*if (txt_prov_clave.Text == "81" || txt_prov_clave.Text == "404" || txt_prov_clave.Text == "11" || txt_prov_clave.Text == "8851" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694")//txt_prov_clave.Text == "81"
                        {*/
                            if (txtFolioFiscal.Text == "")
                            {
                                lblValidaXML.Text = "1";
                            }
                        }

                        lblFechaEnt.Text = reader1["ent_alta"].ToString().Trim();
                        if (reader1["ent_alta"].ToString().Trim() == "")
                            lblFechaEnt.Text = DateTime.Now.ToString();
                        else if (reader1["ent_mod"].ToString().Trim() != "")
                            lblFechaEnt.Text = reader1["ent_mod"].ToString().Trim();

                        txtRetIva.Text = Convert.ToDecimal(reader1["ret_iva"].ToString().Trim()).ToString("###,###,##0.0000");
                        txtDesc.Text = Convert.ToDecimal(reader1["descuento"].ToString().Trim()).ToString("###,###,##0.0000");
                        txtIsr.Text = Convert.ToDecimal(reader1["ret_isr"].ToString().Trim()).ToString("###,###,##0.0000");
                        txtCedular.Text = Convert.ToDecimal(reader1["cedular"].ToString().Trim()).ToString("###,###,##0.0000");

                        lblImporteDescuento.Text = reader1["ent_desc_im"].ToString().Trim();
                        if (lblImporteDescuento.Text == "")
                            lblImporteDescuento.Text = "0";

                        if (Convert.ToDecimal(lblImporteDescuento.Text) > 0)
                        {
                            chkDescuento.Checked = true;
                            //chkDescuento.Enabled = false;
                            MessageBox.Show("La entrada fue generada con descuento por proveedor\n" +
                                "para modificarla debe regresar a los montos originales\n" +
                                "favor de quitar la seleccion de 'Descuento pronto pago'\n" +
                                "realizar los cambios necesarios y si aplica el descuento\n" +
                                "volver a seleccionar el 'Descuento pronto pago' para recalcular", "SISEMP",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();
                tmgrdpro.Clear();
                KeyPressEventArgs llave = new KeyPressEventArgs(Convert.ToChar(13));
                DDL_ent_procedencia.Text = (proce == 0) ? "Nacional" : "Importación";
                //if(Convert.ToDecimal(txt_sis_tipo_cambio.Text) < 1)
                    //DDL_ent_procedencia_SelectionChangeCommitted(null, null);
                txt_prov_clave_KeyPress(sender, llave);
                thisConnection.Open();
                
                //TRAER DETALLE DE ENTRADA
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT ent_folio, emp_clave, alm_clave, entd_cantidad, entd_costo, entd_adv, entd_dta, lote_ras, ent_tipo, ent_ieps, ent_ieps_importe, ent_gastoaduanal, ent_flete, ord_compra, ent_iva" +
                    " FROM tb_det_ent_temporal WHERE ent_folio = '" + fol + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
                reader1 = cmnd1.ExecuteReader();
                DataRow datarow;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        datarow = tmgrdpro.NewRow();
                        datarow["clave"] = reader1.GetValue(1).ToString().Trim();
                        datarow["nombre"] = buscaempaque(reader1.GetValue(1).ToString().Trim());
                        datarow["cantidad"] = reader1.GetValue(3).ToString().Trim();
                        //Math.Round(((((var_dec_entd_costo * (1 + (var_dec_entd_adv / 100))) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + var_dec_entd_dta)) / var_dec_entd_cantidad, 2);

                        decimal u_org = 0;
                        decimal entd_can = 0;
                        decimal entd_cos = 0;
                        decimal entd_adv = 0;
                        decimal entd_dta = 0;
                        decimal ent_ieps = 0;
                        decimal ent_gast = 0;
                        decimal ent_flet = 0;

                        entd_can = Convert.ToDecimal(reader1.GetValue(3).ToString().Trim());
                        //if(sin_fact != "")
                        //    entd_cos = 0;
                        //else
                        entd_cos = Convert.ToDecimal(reader1.GetValue(4).ToString().Trim());
                        entd_adv = Convert.ToDecimal(reader1.GetValue(5).ToString().Trim());
                        entd_dta = Convert.ToDecimal(reader1.GetValue(6).ToString().Trim());
                        ent_ieps = Convert.ToDecimal(reader1.GetValue(10).ToString().Trim());
                        ent_gast = Convert.ToDecimal(reader1.GetValue(11).ToString().Trim());
                        ent_flet = Convert.ToDecimal(reader1.GetValue(12).ToString().Trim());

                        decimal unit = 0;
                        unit = ((entd_cos * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + entd_adv + entd_dta + ent_ieps + ent_gast + ent_flet) / entd_can;

                        datarow["unitario"] = Math.Round(unit, 4);
                        u_org = Math.Round(unit, 4);

                        //if (DDL_ent_procedencia.SelectedIndex == 0)
                        //{
                        //    datarow["unitario"] = Math.Round(((((Convert.ToDecimal(reader1.GetValue(4).ToString().Trim()) * (1 + (Convert.ToDecimal(reader1.GetValue(5).ToString().Trim()) / 100)))) + Convert.ToDecimal(reader1.GetValue(5).ToString().Trim()))) / Convert.ToDecimal(reader1.GetValue(3).ToString().Trim()), 4);
                        //    u_org = Math.Round(((((Convert.ToDecimal(reader1.GetValue(4).ToString().Trim()) * (1 + (Convert.ToDecimal(reader1.GetValue(5).ToString().Trim()) / 100)))) + Convert.ToDecimal(reader1.GetValue(5).ToString().Trim()))) / Convert.ToDecimal(reader1.GetValue(3).ToString().Trim()), 4);
                        //}
                        //else
                        //{
                        //    datarow["unitario"] = Math.Round(((((Convert.ToDecimal(reader1.GetValue(4).ToString().Trim()) * (1 + (Convert.ToDecimal(reader1.GetValue(5).ToString().Trim()) / 100))) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + Convert.ToDecimal(reader1.GetValue(5).ToString().Trim()))) / Convert.ToDecimal(reader1.GetValue(3).ToString().Trim()), 4);
                        //    u_org = Math.Round(((((Convert.ToDecimal(reader1.GetValue(4).ToString().Trim()) * (1 + (Convert.ToDecimal(reader1.GetValue(5).ToString().Trim()) / 100))) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + Convert.ToDecimal(reader1.GetValue(5).ToString().Trim()))) / Convert.ToDecimal(reader1.GetValue(3).ToString().Trim()), 4);
                        //}
                        datarow["costo"] = reader1.GetValue(4).ToString().Trim();
                        datarow["adv"] = reader1.GetValue(5).ToString().Trim();
                        datarow["dta"] = reader1.GetValue(6).ToString().Trim();
                        datarow["lote"] = reader1.GetValue(7).ToString().Trim();
                        datarow["ieps"] = (reader1.GetValue(9).ToString().Trim() == "T") ? true : false;
                        datarow["iepsimporte"] = reader1.GetValue(10).ToString().Trim();
                        datarow["gastoaduanal"] = reader1.GetValue(11).ToString().Trim();
                        datarow["flete"] = reader1.GetValue(12).ToString().Trim();
                        datarow["iva"] = (reader1["ent_iva"].ToString().Trim() == "" || reader1["ent_iva"] == null) ? "0.0000" : reader1["ent_iva"].ToString().Trim();
                        //datarow["oc"] = reader1.GetValue(13).ToString().Trim();
                        //cmnd2 = thisConnection.CreateCommand();
                        //cmnd2.CommandText = "SELECT req_folio, ent_cantidad FROM tb_det_entrada_oc WHERE ent_folio = '" + fol + "' AND numero_oc = '" + reader1.GetValue(13).ToString().Trim() + "' AND emp_clave = '" + reader1.GetValue(1).ToString().Trim() + "'";
                        //reader2 = cmnd2.ExecuteReader();
                        //if (reader2.HasRows)
                        //{
                        //    while (reader2.Read())
                        //    {
                        //        datarow["req"] = reader2.GetValue(0).ToString().Trim();
                        //        datarow["surtido"] = reader1.GetValue(3).ToString().Trim();
                        //    }
                        //}
                        //reader2.Close();
                        //reader2.Dispose();
                        //cmnd2.Dispose();

                        //datarow["cantorig"] = reader1.GetValue(3).ToString().Trim();
                        //datarow["unitorig"] = u_org;
                        tmgrdpro.Rows.Add(datarow);
                    }
                }

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT portal, monto FROM tb_det_oc_facturas WHERE numero_oc = '" + txtFolio.Text.TrimStart('0').ToString() + "' AND tipo = '" + this.DDLAlmacen1.Text.ToUpper() + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        txtAcuse.Text = reader1["portal"].ToString().Trim();
                        lblTotal.Text = reader1["monto"].ToString().Trim();
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                if (txtAcuse.Text != "")
                {
                    txtAcuse.ReadOnly = true;
                    btnValidar.Enabled = false;
                }
                if (txt_prov_clave.Text == "81")
                {
                    txtAcuse.ReadOnly = false;
                    btnValidar.Enabled = true;
                }


                foreach (DataRow ra in tmgrdpro.Rows)
                { 
                    bool encontrado = false;
                    string clv = ra["clave"].ToString();
                    foreach (DataRow rb in dtdetalles.Select("emp_clave = '" + clv + "'"))
                    {
                        encontrado = true;
                    }
                    if (encontrado == false)
                    {
                        DataRow row = dtdetalles.NewRow();
                        row["emp_clave"] = ra["clave"].ToString();
                        row["cantidad"] = ra["cantidad"].ToString();
                        row["importe"] = ra["costo"].ToString();
                        row["costo"] = ra["unitario"].ToString();
                        row["ieps"] = "0.0000";
                        row["lote"] = ra["lote"].ToString();
                        row["iva"] = "0.0000";
                        row["tiva"] = "0";
                        row["cant_orig"] = ra["cantidad"].ToString();
                        row["resp"] = "0";
                        dtdetalles.Rows.Add(row);

                    }
                }

                thisConnection.Close();

                foreach (DataRow rr in dtdetalles.Rows)
                {
                    dtgDetalle2.Rows.Add(rr["emp_clave"].ToString(), rr["orden_compra"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"), 
                        Convert.ToDecimal(rr["importe"].ToString()).ToString("###,###,##0.000"), rr["requisicion"].ToString(),
                        Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.00000"), "", "", "", "", "", rr["lote"].ToString(), Convert.ToDecimal(rr["iva"].ToString()), rr["tiva"].ToString(),
                        ((rr["importe_orig"].ToString() == "") ? "0" : Convert.ToDecimal(rr["importe_orig"].ToString()).ToString("###,###,##0.000")),
                        rr["urgente"].ToString(), Convert.ToDecimal(rr["cant_orig"].ToString()).ToString("###,###,##0.000"), rr["resp"].ToString());
                    string oc = rr["orden_compra"].ToString();
                    entro = true;
                    if (entro == true)
                    {
                        bool fnd = false;
                        foreach (var item in lstOC.Items)
                        {
                            if (item.ToString() == oc)
                            {
                                fnd = true;
                                break;
                            }
                        }
                        if (fnd == false)
                        {
                            lstOC.Items.Add(oc);
                        }
                    }
                    entro = false;
                }

                decimal adv = 0;
                decimal dta1 = 0;
                decimal iva = 0;

                if (Convert.ToDecimal(lblPorcentajeDesc.Text) > 0)
                {
                    for (int i = 0; i < tmgrdpro.Rows.Count; i++)
                    {
                        if (Convert.ToDecimal(tmgrdpro.Rows[i]["cantidad"].ToString()) > 0)
                        {
                            if (DDL_ent_procedencia.SelectedIndex == 0)
                            {
                                tmgrdpro.Rows[i]["unitario"] = Math.Round(((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString())) / Convert.ToDecimal(tmgrdpro.Rows[i]["cantidad"].ToString()), 4);
                                subtotal1 = subtotal1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                                adv = adv + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)));
                                dta1 = dta1 + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString());
                                subtotal2 = subtotal2 + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()));
                                //iva = iva + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
                                total = total + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (1 + (Convert.ToDecimal(txt_ent_iva.Text) / 100));

                            }
                            else
                            {
                                tmgrdpro.Rows[i]["unitario"] = Math.Round(((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString())) / Convert.ToDecimal(tmgrdpro.Rows[i]["cantidad"].ToString()), 4);
                                subtotal1 = subtotal1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                                adv = adv + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                                dta1 = dta1 + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString());
                                subtotal2 = subtotal2 + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()));
                                //iva = iva + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
                                total = total + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (1 + (Convert.ToDecimal(txt_ent_iva.Text) / 100));
                            }
                            ////if (DDL_ent_procedencia.SelectedIndex == 0)
                            ////{
                            ////    //tmgrdpro.Rows[i]["unitario"] = Math.Round(((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) / Convert.ToDecimal(tmgrdpro.Rows[i]["cantidad"].ToString()), 4);
                            ////    //subtotal1 = subtotal1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                            ////    adv = adv + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)));
                            ////    dta1 = dta1 + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString());
                            ////    subtotal2 = subtotal2 + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()));
                            ////    if (txt_ent_iva.Text != "0")
                            ////    {
                            ////        iva = iva + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
                            ////    }
                            ////    else
                            ////    {
                            ////        iva = iva + Convert.ToDecimal(tmgrdpro.Rows[i]["iva"].ToString());
                            ////    }
                            ////    //

                            ////    total = total + subtotal2 + iva;
                            ////    //total = total + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (1 + (Convert.ToDecimal(txt_ent_iva.Text) / 100));
                            ////}
                            ////else
                            ////{
                            ////    decimal t1 = 0;

                            ////    //tmgrdpro.Rows[i]["unitario"] = Math.Round(((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["flete"].ToString())) / Convert.ToDecimal(tmgrdpro.Rows[i]["cantidad"].ToString()), 4);
                            ////    subtotal1 = subtotal1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                            ////    adv = adv + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                            ////    dta1 = dta1 + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString());
                            ////    subtotal2 = subtotal2 + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["flete"].ToString()));
                            ////    if (txt_ent_iva.Text != "0")
                            ////    {
                            ////        iva = iva + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["flete"].ToString())) * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
                            ////    }
                            ////    else
                            ////    {
                            ////        iva = iva + Convert.ToDecimal(tmgrdpro.Rows[i]["iva"].ToString());
                            ////    }


                            ////    //total = total + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["flete"].ToString())) * (1 + (Convert.ToDecimal(txt_ent_iva.Text) / 100));
                            ////    total = total + subtotal2 + iva;
                            ////}
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < tmgrdpro.Rows.Count; i++)
                    {
                        if (Convert.ToDecimal(tmgrdpro.Rows[i]["cantidad"].ToString()) > 0)
                        {
                            if (DDL_ent_procedencia.SelectedIndex == 0)
                            {
                                tmgrdpro.Rows[i]["unitario"] = Math.Round(((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString())) / Convert.ToDecimal(tmgrdpro.Rows[i]["cantidad"].ToString()), 4);
                                subtotal1 = subtotal1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                                adv = adv + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)));
                                dta1 = dta1 + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString());
                                subtotal2 = subtotal2 + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()));
                                //iva = iva + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
                                total = total + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (1 + (Convert.ToDecimal(txt_ent_iva.Text) / 100));
                            }
                            else
                            {
                                tmgrdpro.Rows[i]["unitario"] = Math.Round(((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString())) / Convert.ToDecimal(tmgrdpro.Rows[i]["cantidad"].ToString()), 4);
                                subtotal1 = subtotal1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                                adv = adv + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                                dta1 = dta1 + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString());
                                subtotal2 = subtotal2 + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()));
                                //iva = iva + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
                                total = total + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (1 + (Convert.ToDecimal(txt_ent_iva.Text) / 100));
                            }
                            ////if (DDL_ent_procedencia.SelectedIndex == 0)
                            ////{
                            ////    //tmgrdpro.Rows[i]["unitario"] = Math.Round(((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) / Convert.ToDecimal(tmgrdpro.Rows[i]["cantidad"].ToString()), 4);
                            ////    //subtotal1 = subtotal1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                            ////    adv = adv + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)));
                            ////    dta1 = dta1 + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString());
                            ////    subtotal2 = subtotal2 + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()));
                            ////    if (txt_ent_iva.Text != "0")
                            ////    {
                            ////        iva = iva + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
                            ////    }
                            ////    else
                            ////    {
                            ////        iva = iva + Convert.ToDecimal(tmgrdpro.Rows[i]["iva"].ToString());
                            ////    }
                            ////    //

                            ////    total = total + subtotal2 + iva;
                            ////    //total = total + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString())) * (1 + (Convert.ToDecimal(txt_ent_iva.Text) / 100));
                            ////}
                            ////else
                            ////{
                            ////    decimal t1 = 0;

                            ////    //tmgrdpro.Rows[i]["unitario"] = Math.Round(((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["flete"].ToString())) / Convert.ToDecimal(tmgrdpro.Rows[i]["cantidad"].ToString()), 4);
                            ////    subtotal1 = subtotal1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                            ////    adv = adv + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString()) * (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                            ////    dta1 = dta1 + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString());
                            ////    subtotal2 = subtotal2 + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["flete"].ToString()));
                            ////    if (txt_ent_iva.Text != "0")
                            ////    {
                            ////        iva = iva + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["flete"].ToString())) * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
                            ////    }
                            ////    else
                            ////    {
                            ////        iva = iva + Convert.ToDecimal(tmgrdpro.Rows[i]["iva"].ToString());
                            ////    }


                            ////    //total = total + ((Convert.ToDecimal(tmgrdpro.Rows[i]["costo"].ToString())) * (1 + (Convert.ToDecimal(tmgrdpro.Rows[i]["adv"].ToString()) / 100)) * Convert.ToDecimal(txt_sis_tipo_cambio.Text) + Convert.ToDecimal(tmgrdpro.Rows[i]["dta"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["iepsimporte"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[i]["flete"].ToString())) * (1 + (Convert.ToDecimal(txt_ent_iva.Text) / 100));
                            ////    total = total + subtotal2 + iva;
                            ////}
                        }
                    }
                }

                //sumatoria de ivas del det_ent_oc
                decimal total_iva_det = dtdetalles.AsEnumerable().Sum(row => Convert.ToDecimal(row.Field<String>("iva"))) * Convert.ToDecimal(txt_sis_tipo_cambio.Text);

                txt_subtotal1.Text = subtotal1.ToString("###,###,##0.0000");
                txt_adv.Text = adv.ToString("###,###,##0.0000");
                txt_dta.Text = dta1.ToString("###,###,##0.0000");
                txt_subtotal2.Text = subtotal2.ToString("###,###,##0.0000");
                txt_iva.Text = total_iva_det .ToString("###,###,##0.0000");

                total = total - (Convert.ToDecimal(txtRetIva.Text) + Convert.ToDecimal(txtIsr.Text) + Convert.ToDecimal(txtCedular.Text)) + total_iva_det - Convert.ToDecimal(txtDesc.Text);

                txt_total.Text = total.ToString("###,###,##0.0000");

                //if (Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
                //    txt_gto_aduanal_KeyPress(sender, llave);

                if (Convert.ToDecimal(txtFlete.Text) > 0)
                    txtFlete_KeyPress(sender, llave);

                if (Convert.ToDecimal(txtFleteImp.Text) > 0)
                    txtFleteImp_KeyPress(sender, llave);

                if (Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
                    txt_gto_aduanal_KeyPress(sender, llave);

                //recalculo_isr_descuento();
                //else
                //{
                dtgDetalle.Rows.Clear();
                foreach (DataRow rr in tmgrdpro.Rows)
                {
                    //dtgDetalle.Rows.Add(rr.ItemArray);
                    dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                        Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                        rr["borrar"].ToString(), rr["oper"].ToString(), "", Convert.ToDecimal(rr["iva"].ToString()).ToString("###,###,##0.000"));
                }
                //}
                
               

                panel2.Height = 0;
                operacion = "consulta";
                if (s_fact == "1")
                    chkFactura.Checked = true;
                else
                    chkFactura.Checked = false;
                if (lblEstatus.Text == "CANCELADA")
                {
                    btnBuscar.Enabled = true;
                    btnCancelar.Enabled = false;
                    btnComedor.Enabled = false;
                    btnEditar.Enabled = false;
                    btnGuardar.Enabled = false;
                    btnImprimir.Enabled = false;
                    btnLimpiar.Enabled = true;
                }
                else
                {
                    btnBuscar.Enabled = true;
                    btnCancelar.Enabled = true;
                    btnComedor.Enabled = false;
                    btnEditar.Enabled = true;
                    btnGuardar.Enabled = false;
                    btnImprimir.Enabled = true;
                    btnLimpiar.Enabled = true;
                }

                

                //VALIDACION DE SALIDAS POR LOTE DE TRAZABILIDAD
                string verifica_salidas_traza = verifica_salidas_trazabilidad();
                if (verifica_salidas_traza != "")
                {
                    MessageBox.Show("Se detectaron los siguientes lotes con registro de salida,\ndebe verificar esos movimientos para poder dar cambios o cancelar la entrada",
                        "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    MessageBox.Show(verifica_salidas_traza, "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btnEditar.Enabled = false;
                    btnCancelar.Enabled = false;
                }
                else
                {
                    if (btnEditar.Enabled == false)
                    {
                        btnEditar.Enabled = true;
                        btnCancelar.Enabled = true;
                    }
                }

                if (no_valida == true)
                {
                    btnEditar.Enabled = false;
                }

                if (lblAguilares.Text == "AGUILARES")
                {
                    MessageBox.Show("La entrada fue registrada desde AGUILARES no puede ser modificada desde GAB, solicitar movimiento a AGUILARES, por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                txt_prov_clave.ReadOnly = true;
                DDLProveedor.Enabled = false;
                DDL_ent_procedencia.Enabled = false;

                if(Utilerias.Class1.Usu_login.Trim() == "PEDRO")
                {
                    string tbl = "<p><h2>Entrada: " + txtFolio.Text + "</h2></p><table border='1'><tr><td><b>Empaque</b></td><td><b>Cantidad</b></td><td><b>Costo</b></td><td><b>Importe</b></td><td><b>OC</b></td></tr>";
                    foreach(DataGridViewRow gr in dtgDetalle2.Rows)
                    {
                        tbl += "<tr><td>" + gr.Cells["clave1"].Value.ToString() + "</td><td align='right'>" + gr.Cells["cant"].Value.ToString() + "</td><td align='right'>" + gr.Cells["cost"].Value.ToString() + "</td><td align='right'>" + 
                            gr.Cells["import"].Value.ToString() + "</td><td>" + gr.Cells["ordencompra"].Value.ToString() + "</td></tr>";
                    }
                    tbl += "</table>";
                    correo_error(tbl, "Consulta de Entrada Pedro");
                }

                if (operacion == "consulta")
                {
                    if (lblPedimento.Text == "1")
                    {
                        thisConnection.Open();
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "SELECT COUNT(*) FROM tb_ent_pedimento WHERE folio = '" + txtFolio.Text + "' AND almacen = '" + glo_var_chr_alm_clave + "' AND estatus = 1";
                        reader1 = cmnd1.ExecuteReader();
                        if (reader1.HasRows)
                        {
                            btnSubirArchivos.Visible = true;
                            btnSubirArchivos.Update();
                        }
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();
                        thisConnection.Close();
                    }

                    if (glo_var_chr_alm_clave == "11" || glo_var_chr_alm_clave == "03")
                    {
                        thisConnection.Open();
                        SqlCommand comm;
                        comm = thisConnection.CreateCommand();
                        comm.CommandText = "SELECT COUNT(numero_oc) AS Registros FROM tb_oc_autoriza_mod WHERE numero_oc = '" + txtFolio.Text + "' AND autorizacion = 2 AND estatus = 1";
                        SqlDataReader lector;
                        lector = comm.ExecuteReader();
                        int regi = 0;
                        if (lector.HasRows)
                        {
                            lector.Read();
                            regi = Convert.ToInt32(lector["Registros"].ToString());
                        }
                        lector.Close();
                        lector.Dispose();
                        comm.Dispose();
                        thisConnection.Close();

                        if (regi > 0)
                        {
                            MessageBox.Show("La Entrada fue autorizada para ajustar la orden de compra que se uso para hacer la entrega a almacen", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            btnReajustar.Visible = true;
                        }
                    }
                }
            }
            if (e.KeyChar == (char)Keys.Escape)
            {
                panel2.Height = 0;
            }
        }

        private string buscaempaque(string clave)
        {
            string nombre = "";
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "SELECT emp_nombre FROM tb_cat_empaques WHERE emp_clave = '" + clave + "'";
            reader2 = cmnd2.ExecuteReader();
            if (reader2.HasRows)
            {
                reader2.Read();
                nombre = reader2.GetValue(0).ToString().Trim();
            }
            reader2.Close();
            reader2.Dispose();
            cmnd2.Dispose();
            return nombre;
        }

        public void deshabilitarconsulta()
        {
            //btnGuardar.Enabled = false;
            //btnComedor.Enabled = false;
        }
        public void deshabilitarnuevo()
        {
            //btnImprimir.Enabled = false;
            //btnEditar.Enabled = false;
            //btnComedor.Enabled = false;
            //btnCancelar.Enabled = false;
        }
        public void habilitarconsulta()
        {
            //btnImprimir.Enabled = true;
            //btnEditar.Enabled = true;
            //btnComedor.Enabled = true;
            //btnCancelar.Enabled = true;
        }
        public void habilitarnuevolimpiar()
        {
            //deshabilitarnuevo();
            //FALTA LIMPIAR EL FORMULARIO
        }

        private void btnEditar_Click(object sender, EventArgs e)
        {
            if (glo_var_chr_alm_clave == "02")
            {
                //MessageBox.Show("Se verificaran los cambios realizados en los lotes para verificar salidas previas a la modificacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ////BUSCAR LOTES DE ENTRADA 
                //bool encontrada = false;
                //string salidas_fnd = "";
                //thisConnection.Open();
                //cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = "SELECT lote_ras, emp_clave FROM tb_det_ent_lote WHERE ent_folio = '" + txtFolio.Text + "' AND alm_clave = '02' AND tipo = 'NOR' AND estatus = 'A'";
                //reader1 = cmnd1.ExecuteReader();
                //if (reader1.HasRows)
                //{
                //    while (reader1.Read())
                //    {
                //        encontrada = dtgDetalle2.Rows.Cast<DataGridViewRow>().Any(x => x.Cells["lot"].Value.ToString() == reader1["lote_ras"].ToString().Trim());
                //        if (encontrada == false)
                //        {
                //            //SI EL LOTE NO SE ENCUENTRA EN EL GRID VERIFICAR SALIDAS POR LOTE A ELIMINAR
                //            cmnd2 = thisConnection.CreateCommand();
                //            cmnd2.CommandText = "SELECT lote_ras, sal_folio, emp_clave FROM tb_det_salida_lote WHERE lote_ras = '" + reader1["lote_ras"].ToString().Trim() + "' " +
                //                "AND estatus = 'A' AND emp_clave = '" + reader1["emp_clave"].ToString().Trim() + "'";
                //            reader2 = cmnd2.ExecuteReader();
                //            if (reader2.HasRows)
                //            {
                //                salidas_fnd = salidas_fnd + 
                //                    "Salida: " + reader2["sal_folio"].ToString().Trim() + ", Empaque: " + reader2["emp_clave"].ToString().Trim() + ", Lote: " + reader2["lote_ras"].ToString().Trim() + "\n";
                //            }
                //            reader2.Close();
                //            reader2.Dispose();
                //            cmnd2.Dispose();
                //        }
                //    }
                //    salidas_fnd = salidas_fnd.TrimEnd('\n');
                //    if (salidas_fnd != "")
                //    {
                //        MessageBox.Show("Se detectaron empaques a eliminar con registro de salidas:\n" + salidas_fnd + "\nPara modificar la entrada cancele la(s) salida(s)", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //        return;
                //    }
                //}
                //thisConnection.Close();
            }
            if (lblAguilares.Text == "AGUILARES")
            {
                MessageBox.Show("La entrada fue registrada desde AGUILARES no puede ser modificada desde GAB, solicitar movimiento a AGUILARES, por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            modificar2();
        }

        private void costo_promedio2(DateTime fecha, string clave)
        {
            //if (MessageBox.Show("¿En verdad desea realizar el costo promedio del catálogo de empaques?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            //{
            //    return;
            //}
            string mcod = "";
            decimal minvi = 0;
            decimal mcosp = 0;

            string mrec = "";
            string mtip = "";
            string malm = "";
            decimal mcant = 0;

            decimal A = 0, B = 0, C = 0, D = 0;

            DataTable tbhist = new DataTable();
            tbhist.Columns.Add("hrp_recibo", typeof(string));
            tbhist.Columns.Add("hrp_fecha", typeof(string));
            tbhist.Columns.Add("emp_clave", typeof(string));
            tbhist.Columns.Add("hrp_cantidad", typeof(string));
            tbhist.Columns.Add("hrp_tipo_recepcion", typeof(string));
            tbhist.Columns.Add("alm_clave", typeof(string));
            tbhist.Columns.Add("hrp_situacion", typeof(string));
            tbhist.Columns.Add("hrp_costo", typeof(string));
            tbhist.Columns.Add("hrp_regimen", typeof(string));
            tbhist.Columns.Add("hrp_procedencia", typeof(string));
            tbhist.Columns.Add("hrp_costo_prom", typeof(string));
            tbhist.Columns.Add("hrp_tipo", typeof(string));

            DataTable tbinv = new DataTable();
            tbinv.Columns.Add("emp_clave", typeof(string));
            tbinv.Columns.Add("invemp_inicial", typeof(string));
            tbinv.Columns.Add("invemp_salidas", typeof(string));
            tbinv.Columns.Add("invemp_entradas", typeof(string));
            tbinv.Columns.Add("alm_clave", typeof(string));
            tbinv.Columns.Add("invemp_costopro", typeof(string));

            

            try
            {
                //thisConnection.Open();

                //foreach (DataGridViewRow rz in dtgDetalle.Rows)
                //{
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT hrp_recibo, hrp_fecha, emp_clave, hrp_cantidad, hrp_tipo_recepcion, alm_clave, hrp_situacion, hrp_costo, hrp_regimen, hrp_procedencia, hrp_costo_prom, hrp_tipo" +
                    " FROM tb_historico_recepcion WHERE hrp_estatus <> 'C' AND hrp_fecha >= '" + fecha.ToShortDateString() + "' AND emp_clave = '" + clave + "' ORDER BY hrp_fecha, hrp_recibo";
                reader1 = cmnd1.ExecuteReader();
                DataRow rw;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rw = tbhist.NewRow();
                        rw["hrp_recibo"] = reader1.GetValue(0).ToString().Trim();
                        rw["hrp_fecha"] = reader1.GetValue(1).ToString().Trim();
                        rw["emp_clave"] = reader1.GetValue(2).ToString().Trim();
                        rw["hrp_cantidad"] = reader1.GetValue(3).ToString().Trim();
                        rw["hrp_tipo_recepcion"] = reader1.GetValue(4).ToString().Trim();
                        rw["alm_clave"] = reader1.GetValue(5).ToString().Trim();
                        rw["hrp_situacion"] = reader1.GetValue(6).ToString().Trim();
                        rw["hrp_costo"] = reader1.GetValue(7).ToString().Trim();
                        rw["hrp_regimen"] = reader1.GetValue(8).ToString().Trim();
                        rw["hrp_procedencia"] = reader1.GetValue(9).ToString().Trim();
                        rw["hrp_costo_prom"] = reader1.GetValue(10).ToString().Trim();
                        rw["hrp_tipo"] = reader1.GetValue(11).ToString().Trim();
                        tbhist.Rows.Add(rw);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                cmnd1.CommandText = "SELECT emp_clave, invemp_inicial, invemp_salidas, invemp_entradas, alm_clave, invemp_costopro FROM tb_mstr_inventario_emp WHERE invemp_fecha = '" + fecha.AddDays((fecha.Day - 1) * -1).ToShortDateString() + "' AND emp_clave = '" + clave + "' ORDER BY emp_clave";
                DataRow r1;
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        r1 = tbinv.NewRow();
                        r1["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                        r1["invemp_inicial"] = reader1.GetValue(1).ToString().Trim();
                        r1["invemp_salidas"] = reader1.GetValue(2).ToString().Trim();
                        r1["invemp_entradas"] = reader1.GetValue(3).ToString().Trim();
                        r1["alm_clave"] = reader1.GetValue(4).ToString().Trim();
                        r1["invemp_costopro"] = reader1.GetValue(5).ToString().Trim();
                        tbinv.Rows.Add(r1);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                //thisConnection.Close();

                //thisConnection.Open();
                bool hay = false;
                foreach (DataRow rinv in tbinv.Rows)
                {
                    mcod = rinv[0].ToString();
                    minvi = Convert.ToDecimal(rinv[1].ToString());
                    mcosp = Convert.ToDecimal(rinv[5].ToString());
                    //string mesf = fecha.Month.ToString();
                    //string mesh = "";

                    foreach (DataRow rhis in tbhist.Select("emp_clave = '" + mcod + "'"))
                    {
                        hay = true;
                        //mesh = Convert.ToDateTime(rhis[1].ToString()).Month.ToString();
                        mrec = rhis["hrp_recibo"].ToString();
                        mtip = rhis["hrp_tipo_recepcion"].ToString();
                        malm = rhis["alm_clave"].ToString();
                        if (rhis["hrp_tipo_recepcion"].ToString() == "ENT" && rhis["hrp_situacion"].ToString() == "NOR")
                        {
                            A = Math.Abs(minvi * mcosp);
                            B = Convert.ToDecimal(rhis["hrp_cantidad"].ToString()) * Convert.ToDecimal(rhis["hrp_costo"].ToString());
                            C = Math.Abs(minvi) + Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            D = (C == 0) ? Convert.ToDecimal(rhis["hrp_costo"].ToString()) : ((A + B) / C);
                            minvi = C;
                            mcosp = Math.Abs(D);
                            //Actualizacion de historico
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_existencia = '" + minvi.ToString() + "', hrp_costo_prom = '" + mcosp.ToString() + "'" +
                                " WHERE hrp_tipo_recepcion = '" + mtip + "' AND alm_clave = '" + malm + "' AND hrp_recibo = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                        }
                        if (rhis["hrp_tipo_recepcion"].ToString() == "ENT" && rhis["hrp_situacion"].ToString() == "DEV")
                        {
                            minvi = minvi + Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            mcant = Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            //Actualizacion de historico
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_existencia = '" + minvi.ToString() + "', hrp_costo_prom = '" + mcosp.ToString() + "', hrp_costo = '" + mcosp.ToString() + "'" +
                                " WHERE hrp_tipo_recepcion = '" + mtip + "' AND alm_clave = '" + malm + "' AND hrp_recibo = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();

                            decimal cto = 0;
                            cto = mcosp * mcant;
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "UPDATE tb_det_ent_temporal SET entd_costo = '" + cto.ToString() + "'" +
                                " WHERE alm_clave = '" + malm + "' AND ent_folio = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                        }
                        if (rhis["hrp_tipo_recepcion"].ToString() == "SAL")
                        {
                            minvi = minvi - Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            //Actualizacion de historico
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_existencia = '" + minvi.ToString() + "', hrp_costo_prom = '" + mcosp.ToString() + "', hrp_costo = '" + mcosp.ToString() + "'" +
                                " WHERE hrp_tipo_recepcion = '" + mtip + "' AND alm_clave = '" + malm + "' AND hrp_recibo = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();

                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "UPDATE tb_det_salidas_temporal SET sald_costo = '" + mcosp.ToString() + "'" +
                                " WHERE alm_clave = '" + malm + "' AND sal_folio = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                        }

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + mcosp.ToString() + "', emp_existencia = '" + minvi + "'" +
                            " WHERE emp_clave = '" + mcod + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                    }
                    if (hay == false)
                    {
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_cat_empaques SET emp_existencia = '" + minvi + "', emp_costo = '" + mcosp + "'" +
                            " WHERE emp_clave = '" + mcod + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();
                    }

                }
                //thisConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.4", ex.ToString(), "SISEMP");
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + ex.ToString());
            }



        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            if (operacion == "consulta")
            {
                if (lblAguilares.Text == "AGUILARES")
                {
                    MessageBox.Show("La entrada fue registrada desde AGUILARES no puede ser modificada desde GAB, solicitar movimiento a AGUILARES, por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                cancelar();
            }
            else
                MessageBox.Show("Para cancelar debe consultar la entrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
        }

        //public string CheckPrinterConfiguration(string printerIP, string printerName)
        //{
        //    var server = new PrintServer();
        //    var queues = server.GetPrintQueues(new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections });
        //    string fulllName = queues.Where(q => q.Name == printerName &&
        //    q.QueuePort.Name == printerIP).Select(q => q.FullName).FirstOrDefault();
        //    return fulllName;
        //}

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            //imprimir();
            //printDocument1.Print();

            //string printerFullName = CheckPrinterConfiguration("192.168.123.159", "EPSON LX-300+");

            sum_importe_print = 0;
            sum_gtos_im_print = 0;

            if (lblPedimento.Text == "1")
            {
                foreach(DataGridViewRow gr in dtgDetalle2.Rows)
                {
                    sum_importe_print += Convert.ToDecimal(gr.Cells["import"].Value.ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text);
                }
                sum_gtos_im_print = Convert.ToDecimal(txt_gto_aduanal.Text) + Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text);
            }

            if (File.Exists(@"c:\\Reportes\document.pdf"))
            {
                File.Delete(@"c:\\Reportes\document.pdf");
                iTextSharp.text.Document doc = new iTextSharp.text.Document();
                iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new FileStream(@"c:\\Reportes\document.pdf", FileMode.Create));
                doc.Open();
                iTextSharp.text.Paragraph title = new iTextSharp.text.Paragraph();
                title.Add("");
                doc.Add(title);
                doc.Add(new iTextSharp.text.Paragraph(""));
                doc.Close();
            }
            else
            {
                iTextSharp.text.Document doc = new iTextSharp.text.Document();
                iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new FileStream(@"c:\\Reportes\document.pdf", FileMode.Create));
                doc.Open();
                iTextSharp.text.Paragraph title = new iTextSharp.text.Paragraph();
                title.Add("");
                doc.Add(title);
                doc.Add(new iTextSharp.text.Paragraph(""));
                doc.Close();
            }

            using (PrintDocument pd = new PrintDocument())
            {
                
                //PaperSize pagesize = new PaperSize();
                //pagesize

                //if (glo_var_chr_alm_clave == "03")
                //{
                //    pd.PrinterSettings.PrinterName = "EPSON LX-300+";
                //}

                pd.PrintPage += (this.printDocument1_PrintPage);

                //

                pd.Print();
                //pd.Dispose();

                //pd.EndPrint += new System.Drawing.Printing.PrintEventHandler(this.printDocument1_EndPrint); 
            }

            //foreach (Process proceso in Process.GetProcesses())
            //{
            //    if (proceso.ProcessName == "splwow64")
            //    {
            //        proceso.Kill();
            //    }
            //}

            //printDocument1.Dispose();
            
            
            //printDocument2.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
            //printDocument2.Print();
            //printDocument2.Dispose();
            //string str1 = "c:\\Reportes\\document.pdf";
            //string str2 = "";
            //if (this.DDLAlmacen1.SelectedIndex == 2)
            //    str2 = "COM";
            //if (this.DDLAlmacen1.SelectedIndex == 3)
            //    str2 = "GUA";
            //string str3;
            //if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
            //    str3 = "c:\\Reportes\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
            //else
            //    str3 = "\\\\gabira1\\Entradas\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
            //if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
            //{
            //    string str4 = "c:\\Reportes\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
            //    if (System.IO.File.Exists(str4))
            //        System.IO.File.Delete(str4);
            //    System.IO.File.Copy(str1, str4, true);
            //}
            //else
            //{
            //    if (System.IO.File.Exists(str3))
            //        System.IO.File.Delete(str3);
            //    System.IO.File.Copy(str1, str3, true);
            //}

            renglonant = 0;
            ren = 0;

            string str2 = "";
            if (this.DDLAlmacen1.SelectedIndex == 2)
                str2 = "COM";
            if (this.DDLAlmacen1.SelectedIndex == 3)
                str2 = "GUA";
            string ruta_entrada_validacion = @"\\gabira1\Entradas\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
            string file_name = "Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
            genera_pdf(file_name);
            
            if (System.IO.File.Exists(ruta_entrada_validacion))
            {
                System.IO.File.Delete(ruta_entrada_validacion);


                string ROUTE_PDF = "c:\\Reportes\\" + file_name;
                //string COM_GUA = "";
                //if (this.DDLAlmacen1.SelectedIndex == 2)
                //    COM_GUA = "COM";
                //if (this.DDLAlmacen1.SelectedIndex == 3)
                //    COM_GUA = "GUA";
                //string str3;
                if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
                {
                }
                else
                {
                    string ROUTE_SAVED = @"\\gabira1\Entradas\" + file_name;
                    System.IO.File.Copy(ROUTE_PDF, ROUTE_SAVED, true);
                    //System.IO.File.Delete(ROUTE_PDF);
                }
                    //str3 = "c:\\Reportes\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + COM_GUA + ".pdf";
                //else
                //    str3 = "\\\\gabira1\\Entradas\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + COM_GUA + ".pdf";
                //if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
                //{
                //    string str4 = "c:\\Reportes\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + COM_GUA + ".pdf";
                //    System.IO.File.Copy(str1, str4, true);
                //}
                //else
                //{
                //    System.IO.File.Copy(str1, str3, true);
                //}

                //imprmir_entrada_servidor();
            }
            else
            {
                string ROUTE_PDF = "c:\\Reportes\\" + file_name;
                if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
                {
                }
                else
                {
                    string ROUTE_SAVED = @"\\gabira1\Entradas\" + file_name;
                    System.IO.File.Copy(ROUTE_PDF, ROUTE_SAVED, true);
                    //System.IO.File.Delete(ROUTE_PDF);
                }
            }
                //imprmir_entrada_servidor();


        }

        public void imprmir_entrada_servidor()
        {
            try
            {
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + "Entra EMAIL");
                

                using (PrintDocument pd = new PrintDocument())
                {
                    //printDocument2.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                    //printDocument2.Print();
                    //printDocument2.Dispose();
                    pd.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                    pd.PrintPage += (this.printDocument2_PrintPage);
                    pd.Print();
                    //pd.Dispose();
                }

                //foreach (Process proceso in Process.GetProcesses())
                //{
                //    if (proceso.ProcessName == "splwow64")
                //    {
                //        proceso.Kill();
                //    }
                //}

                string str1 = "c:\\Reportes\\document.pdf";
                string str2 = "";
                if (this.DDLAlmacen1.SelectedIndex == 2)
                    str2 = "COM";
                if (this.DDLAlmacen1.SelectedIndex == 3)
                    str2 = "GUA";
                string str3;
                if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
                    str3 = "c:\\Reportes\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
                else
                    str3 = "\\\\gabira1\\Entradas\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
                if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
                {
                    string str4 = "c:\\Reportes\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
                    //if (System.IO.File.Exists(str4))
                        //System.IO.File.Delete(str4);
                    System.IO.File.Copy(str1, str4, true);
                }
                else
                {
                    //if (System.IO.File.Exists(str3))
                        //System.IO.File.Delete(str3);
                    System.IO.File.Copy(str1, str3, true);
                }

                renglonant = 0;
                ren = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de catch " + ex.ToString());
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + ex.ToString());
                if (ex.ToString().Contains("El proceso no puede obtener acceso al archivo"))
                {
                    form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + "Entra a catch");
                    //printDocument2.Dispose();
                    //printDocument2.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                    //printDocument2.Print();
                    //printDocument2.Dispose();
                    using (PrintDocument pd = new PrintDocument())
                    {
                        pd.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                        pd.PrintPage += (this.printDocument2_PrintPage);
                        pd.Print();
                        //pd.Dispose();
                    }
                    //foreach (Process proceso in Process.GetProcesses())
                    //{
                    //    if (proceso.ProcessName == "splwow64")
                    //    {
                    //        proceso.Kill();
                    //    }
                    //}
                    string str1 = "c:\\Reportes\\document.pdf";
                    string str2 = "";
                    if (this.DDLAlmacen1.SelectedIndex == 2)
                        str2 = "COM";
                    if (this.DDLAlmacen1.SelectedIndex == 3)
                        str2 = "GUA";
                    string str3;
                    if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
                        str3 = "c:\\Reportes\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
                    else
                        str3 = "\\\\gabira1\\Entradas\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
                    if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
                    {
                        string str4 = "c:\\Reportes\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
                        //if (System.IO.File.Exists(str4))
                        //    System.IO.File.Delete(str4);
                        System.IO.File.Copy(str1, str4, true);
                    }
                    else
                    {
                        //if (System.IO.File.Exists(str3))
                        //    System.IO.File.Delete(str3);
                        System.IO.File.Copy(str1, str3, true);
                    }

                    renglonant = 0;
                    ren = 0;
                }
            }
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            txtFolio.Text = "";
            DDL_ent_procedencia.Text = "";
            txt_prov_clave.Text = "";
            DDLProveedor.SelectedIndex = -1;
            dtpFecha.Text = DateTime.Now.ToShortDateString();
            txt_ent_factura.Text = "";
            txt_ent_pedimento.Text = "";
            txt_gto_aduanal.Text = "0";
            txtFlete.Text = "0";
            txt_sis_tipo_cambio.Text = "1";
            //DDLElaboro.Text = "";
            //DDLAutorizo.Text = "";
            DDLElaboro.SelectedIndex = -1;
            DDLAutorizo.SelectedIndex = -1;
            txt_ent_iva.Text = "0";
            txt_ent_ieps.Text = "0";
            txt_subtotal1.Text = "0";
            txt_subtotal2.Text = "0";
            txt_adv.Text = "0";
            txt_dta.Text = "0";
            txt_iva.Text = "0";
            txt_total.Text = "0";

            glo_var_int_registro = 0;
            //glo_var_chr_alm_clave = "";
            var_int_sis_folios = "";
            filtroprov.Clear();
            provnotfound = "";
            tmgrdpro.Clear();
            dtempoc.Clear();

            iva1 = 0;
            subtotal1 = 0;
            totaladv = 0;
            totaldta = 0;
            totalieps = 0;
            subtotal2 = 0;
            iva2 = 0;
            total = 0;

            facturas.Clear();
            pedimentos.Clear();

            //usuarionombre.Clear();

            //iepsporc = 0;

            ren = 0;
            renglonant = 0;

            nuevo_costo = 0;
            nueva_existencia = 0;
            nueva_costoprom = 0;

            operacion = "";

            dtordenes.Clear();
            dtgOrdenes.Rows.Clear();
            tmgrdpro.Clear();
            dtgDetalle.Rows.Clear();
            dtgDetalle2.Rows.Clear();
            dtdetalles.Clear();
            dtoriginales.Clear();

            //DDLAlmacen1.SelectedIndex = -1;
            //DDLAlmacen1.Text = "";

            txtObservaciones.Text = "";

            lstOC.Items.Clear();

            lblEstatus.Text = "-";

            btnBuscar.Enabled = true;
            btnCancelar.Enabled = false;
            btnComedor.Enabled = false;
            btnEditar.Enabled = false;
            btnGuardar.Enabled = true;
            btnImprimir.Enabled = false;
            btnLimpiar.Enabled = true;

            if (DDLAlmacen1.SelectedIndex > -1)
            {
                DDLAlmacen1_SelectionChangeCommitted(null, null);
            }

            txtFleteImp.Text = "0";
            chkFactura.Checked = false;

            DDL_ent_procedencia.SelectedIndex = 0;
            DDL_ent_procedencia_SelectionChangeCommitted(null, null);

            pbxCalendarUlEn.Enabled = true;

            lblXmlFile.Text = "-";
            txtFolioFiscal.Text = "";
            lblRFC.Text = "-";
            lblValidaXML.Text = "0";
            lblPathXml.Text = "-";
            lblPortal.Text = "-";

            lblPorcentajeDesc.Text = "0";
            chkDescuento.Checked = false;
            btnValidar.Enabled = true;
            txtAcuse.Text = "";
            txtFechaFact.Text = "";
            lblXmlProv.Text = "0";
            lblTCFact.Text = "1";
            lblToc.Text = "0";
            lblFactNal.Text = "0";
            lblFechaEnt.Text = "-";
            lblTotal.Text = "0";

            
            btnCambioFactura.Enabled = true;

            txtDesc.Text = "0";
            txtRetIva.Text = "0";
            txtIsr.Text = "0";
            txtCedular.Text = "0";

            lblAguilares.Text = "-";

            txt_prov_clave.ReadOnly = false;
            DDLProveedor.Enabled = true;
            DDL_ent_procedencia.Enabled = true;

            lblImporteDescuento.Text = "0";

            lblPedimento.Text = "0";

            btnSubirArchivos.Visible = false;
            btnSubirArchivos.Update();

            btnReajustar.Visible = false;

        }

        private void btnComedor_Click(object sender, EventArgs e)
        {
            empaques dlg;
            if (glo_var_chr_alm_clave != "")
            {
                dlg = new empaques(glo_var_chr_alm_clave);
                dlg.ShowDialog();
            }
                
        }


        decimal sum_importe_print = 0;
        decimal sum_gtos_im_print = 0;
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                thisConnection.Open();
                e.Graphics.PageUnit = GraphicsUnit.Millimeter;

                //if (glo_var_chr_alm_clave == "03")
                //{
                //    e.PageSettings.PaperSize = new System.Drawing.Printing.PaperSize("Carta", 216, 141);
                //}
                
                Graphics g = e.Graphics;
                PageSettings ps = e.PageSettings;
                SolidBrush brsh = new SolidBrush(Color.Black);

                Font fuente1 = new System.Drawing.Font("Courier New", 12);
                //g.DrawString(txtFolio.Text, fuente1, brsh, 160, 4);
                g.DrawString(txtFolio.Text, fuente1, brsh, 160, 2);

                //g.DrawString("NORMAL", fuente1, brsh, 160, 8);
                g.DrawString("NORMAL", fuente1, brsh, 160, 6);

                Font fuente = new System.Drawing.Font("Courier New", 9);
                g.DrawString(dtpFecha.Text, fuente, brsh, 140, 14);

                //if (DDLProveedor.SelectedItem.ToString().Length > 40)
                //    g.DrawString(DDLProveedor.SelectedItem.ToString().Substring(0, 39).ToString(), fuente, brsh, 5, 27);
                //else
                //    g.DrawString(DDLProveedor.SelectedItem.ToString(), fuente, brsh, 5, 27);
                if (DDLProveedor.SelectedItem.ToString().Length > 40)
                    g.DrawString(DDLProveedor.SelectedItem.ToString().Substring(0, 39).ToString(), fuente, brsh, 5, 23);
                else
                    g.DrawString(DDLProveedor.SelectedItem.ToString(), fuente, brsh, 5, 23);


                //g.DrawString(txt_ent_factura.Text, fuente, brsh, 105, 27);
                //g.DrawString(lstOC.Items[0].ToString(), fuente, brsh, 170, 27);
                ////g.DrawString(ordencomprax, fuente, brsh, 170, 23);
                g.DrawString(txt_ent_factura.Text, fuente, brsh, 105, 23);
                g.DrawString(lstOC.Items[0].ToString(), fuente, brsh, 170, 23);

                
                int renglon = 33;

                g.DrawString("Lote Rastreabilidad", fuente, brsh, 108, renglon);

                string prm_ent_entregox = "";
                string prm_ent_autorizox = "";
                if (DDLElaboro.SelectedItem.ToString().Length > 25)
                    prm_ent_entregox = DDLElaboro.SelectedItem.ToString().Substring(0, 24);//prm_ent_entregox = DDLElaboro.SelectedItem.ToString().Substring(0, 27);
                else
                    prm_ent_entregox = DDLElaboro.SelectedItem.ToString();
                if (DDLAutorizo.SelectedItem.ToString().Length > 25)
                    prm_ent_autorizox = DDLAutorizo.SelectedItem.ToString().Substring(0, 24);//prm_ent_autorizox = DDLAutorizo.SelectedItem.ToString().Substring(0, 27);
                else
                    prm_ent_autorizox = DDLAutorizo.SelectedItem.ToString();
                ////g.DrawString(txtObservaciones.Text, fuente, brsh, 2, 116);
                ////g.DrawString("F-500-16", fuente, brsh, 10, 70);
                //g.DrawString(prm_ent_entregox, fuente, brsh, 50, 130);
                //g.DrawString(prm_ent_autorizox, fuente, brsh, 5, 130);
                g.DrawString(prm_ent_entregox, fuente, brsh, 50, 127);
                g.DrawString(prm_ent_autorizox, fuente, brsh, 5, 127);
                ////g.DrawString("REV: 00", fuente, brsh, 10, 70);

                //bool entra = false;
                //g.DrawString("Subtotal", fuente, brsh, 150, 112);
                //string total_iepsx = ""; string ivax = ""; string subtotalx = ""; string totalx = "";
                //total_iepsx = (txt_ent_ieps.Text == "") ? "0.00" : txt_ent_ieps.Text;
                //ivax = txt_iva.Text;
                //subtotalx = txt_subtotal2.Text;
                //totalx = txt_total.Text;
                //if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) == 0)
                //{
                //}
                //else if (Convert.ToDecimal(total_iepsx) != 0 && Convert.ToDecimal(ivax) == 0)
                //{
                //    g.DrawString("I.E.P.S.", fuente, brsh, 150, 115);
                //}
                //else if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) != 0)
                //{
                //    g.DrawString("I.V.A.", fuente, brsh, 150, 115);
                //}
                //else
                //{
                //    g.DrawString("I.E.P.S.", fuente, brsh, 150, 115);
                //    g.DrawString("I.V.A.", fuente, brsh, 150, 118);
                //    g.DrawString("Total", fuente, brsh, 150, 121);
                //    entra = true;
                //}
                //if (entra == false)
                //    g.DrawString("Total", fuente, brsh, 150, 118);

                //g.DrawString(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00")), 112);
                //if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) == 0)
                //{
                //}
                //else if (Convert.ToDecimal(total_iepsx) != 0 && Convert.ToDecimal(ivax) == 0)
                //{
                //    g.DrawString(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00")), 115);
                //}
                //else if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) != 0)
                //{
                //    g.DrawString(Convert.ToDecimal(ivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(ivax).ToString("###,###,##0.00")), 115);
                //}
                //else
                //{
                //    g.DrawString(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00")), 115);
                //    g.DrawString(Convert.ToDecimal(ivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(ivax).ToString("###,###,##0.00")), 118);
                //    g.DrawString(Convert.ToDecimal(totalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(totalx).ToString("###,###,##0.00")), 121);
                //}

                //if (entra == false)
                //    g.DrawString(Convert.ToDecimal(totalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(totalx).ToString("###,###,##0.00")), 118);

                //g.DrawString("F-500-11", fuente, brsh, 177, 127);
                //g.DrawString("REV: 00", fuente, brsh, 177, 130);
                g.DrawString("F-500-11", fuente, brsh, 177, 124);
                g.DrawString("REV: 00", fuente, brsh, 177, 127);

                //renglon = renglon + 5;
                renglon = renglon + 3;
                string var_chr_um_clave = "";



                if (renglonant == 0)
                {
                    

                    for (int u = 0; u < tmgrdpro.Rows.Count; u++)
                    {
                        if (tmgrdpro.Rows[u]["borrar"].ToString() == "1")
                            continue;
                        form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + "Entra EMAIL");
                        if (Convert.ToDecimal(tmgrdpro.Rows[u]["cantidad"].ToString()) > 0 && Convert.ToDecimal(tmgrdpro.Rows[u]["unitario"].ToString()) > 0)
                        {
                            g.DrawString(Convert.ToDecimal(tmgrdpro.Rows[u]["cantidad"].ToString()).ToString("###,###,##0.00"), fuente, brsh, posicion(Convert.ToDecimal(tmgrdpro.Rows[u]["cantidad"].ToString()).ToString("###,###,##0.00")), renglon);
                            cmnd2 = thisConnection.CreateCommand();
                            cmnd2.CommandText = "SELECT um_clave FROM tb_cat_empaques WHERE emp_clave = '" + tmgrdpro.Rows[u]["clave"].ToString().Trim() + "'";
                            reader2 = cmnd2.ExecuteReader();
                            if (reader2.HasRows)
                            {
                                reader2.Read();
                                var_chr_um_clave = reader2.GetValue(0).ToString();
                            }
                            reader2.Close();

                            g.DrawString(var_chr_um_clave.Substring(0, 3).ToString(), fuente, brsh, 25, renglon);
                            //g.DrawString(tmgrdpro.Rows[u][0].ToString().Substring(0, 4).ToString() + " " + tmgrdpro.Rows[u][1].ToString(), fuente, brsh, 40, renglon);
                            g.DrawString(tmgrdpro.Rows[u]["clave"].ToString() + " " + tmgrdpro.Rows[u]["nombre"].ToString(), fuente, brsh, 40, renglon);
                            g.DrawString(tmgrdpro.Rows[u]["lote"].ToString(), fuente, brsh, 130, renglon);

                            //decimal cos_sin_gto_imp = Math.Round(((txt_sis_tipo_cambio.Text != "1") ? (Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) : Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString())) / Convert.ToDecimal(tmgrdpro.Rows[u]["cantidad"].ToString()), 4);
                            //if (lblPedimento.Text == "1")
                            //{
                            //    g.DrawString(Convert.ToDecimal(tmgrdpro.Rows[u]["unitario"].ToString()).ToString("###,###,##0.00"), fuente, brsh, posicionprecio(Convert.ToDecimal(tmgrdpro.Rows[u]["unitario"].ToString()).ToString("###,###,##0.00")), renglon);//precio 
                            //}
                            //else
                            //{
                            //    g.DrawString(Convert.ToDecimal(tmgrdpro.Rows[u]["unitario"].ToString()).ToString("###,###,##0.00"), fuente, brsh, posicionprecio(Convert.ToDecimal(tmgrdpro.Rows[u]["unitario"].ToString()).ToString("###,###,##0.00")), renglon);//precio 
                            //}

                            decimal imp_print = (Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                            decimal cos_print = imp_print / Convert.ToDecimal(tmgrdpro.Rows[u]["cantidad"].ToString());

                            if (lblPedimento.Text == "1")
                            {
                                g.DrawString(cos_print.ToString("###,###,##0.00"), fuente, brsh, posicionprecio(cos_print.ToString("###,###,##0.00")), renglon);//precio 
                            }
                            else
                            {
                                g.DrawString(Convert.ToDecimal(tmgrdpro.Rows[u]["unitario"].ToString()).ToString("###,###,##0.00"), fuente, brsh, posicionprecio(Convert.ToDecimal(tmgrdpro.Rows[u]["unitario"].ToString()).ToString("###,###,##0.00")), renglon);//precio 
                            }

                            //sum_gtos_im_print = sum_gtos_im_print + (Convert.ToDecimal(tmgrdpro.Rows[u]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[u]["flete"].ToString()));
                            //sum_importe_print = (Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));

                            
                            
                            if (tmgrdpro.Rows[u]["costo"].ToString() != "" && Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) > 0)
                            {
                                
                                decimal totprod = (Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + (Convert.ToDecimal(tmgrdpro.Rows[u]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[u]["flete"].ToString()));
                                if (lblPedimento.Text == "1")
                                {
                                    g.DrawString(imp_print.ToString("###,###,##0.00"), fuente, brsh, posicionimporte(imp_print.ToString("###,###,##0.00")), renglon);//importe
                                }
                                else
                                {
                                    g.DrawString(totprod.ToString("###,###,##0.00"), fuente, brsh, posicionimporte(totprod.ToString("###,###,##0.00")), renglon);//importe
                                }
                                
                                //if (DDL_ent_procedencia.SelectedIndex == 0)
                                //{
                                //    decimal totprod = Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) + (Convert.ToDecimal(tmgrdpro.Rows[u]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[u]["flete"].ToString()));
                                //    g.DrawString(totprod.ToString("###,###,##0.00"), fuente, brsh, posicionimporte(totprod.ToString("###,###,##0.00")), renglon);//importe
                                //}
                                //else
                                //{
                                    
                                //}
                            }
                            //envio a otra pagina
                            renglon = renglon + 3;
                            renglonant++;
                            ren++;
                            if (ren == 20)
                            {
                                ren = 0;
                                break;
                            }
                        }

                    }
                    if (renglonant >= tmgrdpro.Rows.Count)
                    {
                        string ocs = "O.C. ";
                        for (int i = 1; i < lstOC.Items.Count; i++)
                        {
                            ocs += lstOC.Items[i].ToString() + ", "; 
                        }
                        if (lstOC.Items.Count > 1)
                            g.DrawString(txtObservaciones.Text + " " + ocs.TrimEnd(','), fuente, brsh, 2, 116);//g.DrawString(txtObservaciones.Text + " " + ocs.TrimEnd(','), fuente, brsh, 2, 122);
                        else
                            g.DrawString(txtObservaciones.Text, fuente, brsh, 2, 116); //g.DrawString(txtObservaciones.Text, fuente, brsh, 2, 122);

                        g.DrawString("Gtos. Imp.", fuente, brsh, 150, 91);
                        g.DrawString(Convert.ToDecimal(sum_gtos_im_print).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(sum_gtos_im_print).ToString("###,###,##0.00")), 91);

                        g.DrawString("I.E.P.S.", fuente, brsh, 150, 94);
                        string total_iepsx = "";
                        //total_iepsx = (txt_ent_ieps.Text == "") ? "0.00" : txt_ent_ieps.Text;
                        total_iepsx = txt_ent_ieps.Text;// (txt_ent_ieps.Text == "") ? "0.00" : txt_ent_ieps.Text;

                        g.DrawString("Subtotal", fuente, brsh, 150, 97);
                        string subtotalx = "";
                        subtotalx = txt_subtotal2.Text;
                        if (lblPedimento.Text == "1")
                        {
                            g.DrawString(Convert.ToDecimal(sum_importe_print).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(sum_importe_print).ToString("###,###,##0.00")), 97);
                        }
                        else
                        {
                            g.DrawString(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00")), 97);
                        }
                        //g.DrawString("Subtotal", fuente, brsh, 150, 97);
                        //string subtotalx = "";
                        //subtotalx = txt_subtotal2.Text;
                        //g.DrawString(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00")), 97);

                        g.DrawString("Descuento", fuente, brsh, 150, 100);
                        string descuentox = "";
                        descuentox = txtDesc.Text;
                        g.DrawString(Convert.ToDecimal(descuentox).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(descuentox).ToString("###,###,##0.00")), 100);

                        g.DrawString("I.V.A.", fuente, brsh, 150, 103);
                        string ivax = "";
                        ivax = txt_iva.Text;
                        g.DrawString(Convert.ToDecimal(ivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(ivax).ToString("###,###,##0.00")), 103);

                        g.DrawString("Ret IVA", fuente, brsh, 150, 106);
                        string retivax = "";
                        retivax = txtRetIva.Text;
                        g.DrawString(Convert.ToDecimal(retivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(retivax).ToString("###,###,##0.00")), 106);

                        g.DrawString("Ret ISR", fuente, brsh, 150, 109);
                        string retisrx = "";
                        retisrx = txtIsr.Text;
                        g.DrawString(Convert.ToDecimal(retisrx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(retisrx).ToString("###,###,##0.00")), 109);

                        g.DrawString("Cedular", fuente, brsh, 150, 112);
                        string cedularx = "";
                        cedularx = txtCedular.Text;
                        g.DrawString(Convert.ToDecimal(cedularx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(cedularx).ToString("###,###,##0.00")), 112);

                        //g.DrawString("Total " + ((DDL_ent_procedencia.SelectedIndex == 0) ? "MN" : "USD"), fuente, brsh, 150, 115);
                        g.DrawString("Total", fuente, brsh, 150, 115);
                        string totalx = "";
                        totalx = txt_total.Text;
                        g.DrawString(Convert.ToDecimal(totalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(totalx).ToString("###,###,##0.00")), 115);

                        g.DrawString("MN", fuente, brsh, 195, 115);

                        if (Convert.ToDecimal(lblImporteDescuento.Text) > 0)
                        {
                            decimal desc_pp = 0;
                            desc_pp = Convert.ToDecimal(lblImporteDescuento.Text) * Convert.ToDecimal(txt_sis_tipo_cambio.Text);
                            g.DrawString("Descuento por pronto pago: ", fuente, brsh, 40, 112);
                            totalx = txt_total.Text;
                            g.DrawString("$" + desc_pp.ToString("###,###,##0.00"), fuente, brsh, 95, 112);
                        }
                        

                        //bool entra = false;
                        //g.DrawString("Subtotal", fuente, brsh, 150, 94);//112
                        //g.DrawString("Subtotal", fuente, brsh, 150, 92);//112
                        //string total_iepsx = ""; string ivax = ""; string subtotalx = ""; string totalx = "";
                        //total_iepsx = (txt_ent_ieps.Text == "") ? "0.00" : txt_ent_ieps.Text;
                        //ivax = txt_iva.Text;
                        //subtotalx = txt_subtotal2.Text;
                        //totalx = txt_total.Text;
                        //if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) == 0)
                        //{
                        //}
                        //else if (Convert.ToDecimal(total_iepsx) != 0 && Convert.ToDecimal(ivax) == 0)
                        //{
                        //    g.DrawString("I.E.P.S.", fuente, brsh, 150, 95); //g.DrawString("I.E.P.S.", fuente, brsh, 150, 97);//115
                        //}
                        //else if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) != 0)
                        //{
                        //    g.DrawString("I.V.A.", fuente, brsh, 150, 95); //g.DrawString("I.V.A.", fuente, brsh, 150, 97);//115
                        //}
                        //else
                        //{
                        //    //g.DrawString("I.E.P.S.", fuente, brsh, 150, 97);//115
                        //    //g.DrawString("I.V.A.", fuente, brsh, 150, 100);//118
                        //    //g.DrawString("Total", fuente, brsh, 150, 103);//121
                        //    g.DrawString("I.E.P.S.", fuente, brsh, 150, 95);//115
                        //    g.DrawString("I.V.A.", fuente, brsh, 150, 98);//118
                        //    g.DrawString("Total", fuente, brsh, 150, 101);//121
                        //    entra = true;
                        //}
                        //if (entra == false)
                        //    g.DrawString("Total", fuente, brsh, 150, 98);//g.DrawString("Total", fuente, brsh, 150, 100);//98

                        ////g.DrawString(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00")), 94);//112
                        //g.DrawString(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00")), 92);//112
                        //if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) == 0)
                        //{
                        //}
                        //else if (Convert.ToDecimal(total_iepsx) != 0 && Convert.ToDecimal(ivax) == 0)
                        //{
                        //    g.DrawString(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00")), 95); //g.DrawString(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00")), 97);//115
                        //}
                        //else if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) != 0)
                        //{
                        //    g.DrawString(Convert.ToDecimal(ivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(ivax).ToString("###,###,##0.00")), 95); //g.DrawString(Convert.ToDecimal(ivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(ivax).ToString("###,###,##0.00")), 97);//115
                        //}
                        //else
                        //{
                        //    //g.DrawString(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00")), 97);//115
                        //    //g.DrawString(Convert.ToDecimal(ivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(ivax).ToString("###,###,##0.00")), 100);//118
                        //    //g.DrawString(Convert.ToDecimal(totalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(totalx).ToString("###,###,##0.00")), 103);//121
                        //    g.DrawString(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00")), 95);//115
                        //    g.DrawString(Convert.ToDecimal(ivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(ivax).ToString("###,###,##0.00")), 98);//118
                        //    g.DrawString(Convert.ToDecimal(totalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(totalx).ToString("###,###,##0.00")), 101);//121
                        //}
                        //if (entra == false)
                        //    g.DrawString(Convert.ToDecimal(totalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(totalx).ToString("###,###,##0.00")), 98);//118 g.DrawString(Convert.ToDecimal(totalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(totalx).ToString("###,###,##0.00")), 100);//118


                        e.HasMorePages = false;
                    }
                    if (renglonant < tmgrdpro.Rows.Count)
                    {
                        //g.DrawString("CONTINUA EN LA SIGUIENTE HOJA.....", fuente, brsh, 2, 122);
                        g.DrawString("CONTINUA EN LA SIGUIENTE HOJA.....", fuente, brsh, 2, 116);
                        e.HasMorePages = true;
                    }
                }
                else
                {
                    //renglon = renglon + 3;
                    for (int u = renglonant; u < tmgrdpro.Rows.Count; u++)
                    {
                        if (tmgrdpro.Rows[u]["borrar"].ToString() == "1")
                            continue;
                        g.DrawString(Convert.ToDecimal(tmgrdpro.Rows[u]["cantidad"].ToString()).ToString("###,###,##0.00"), fuente, brsh, posicion(Convert.ToDecimal(tmgrdpro.Rows[u]["cantidad"].ToString()).ToString("###,###,##0.00")), renglon);
                        if (Convert.ToDecimal(tmgrdpro.Rows[u]["cantidad"].ToString()) > 0 && Convert.ToDecimal(tmgrdpro.Rows[u]["unitario"].ToString()) > 0)
                        {
                            cmnd2 = thisConnection.CreateCommand();
                            cmnd2.CommandText = "SELECT um_clave FROM tb_cat_empaques WHERE emp_clave = '" + tmgrdpro.Rows[u]["clave"].ToString().Trim() + "'";
                            reader2 = cmnd2.ExecuteReader();
                            if (reader2.HasRows)
                            {
                                reader2.Read();
                                var_chr_um_clave = reader2.GetValue(0).ToString();
                            }
                            reader2.Close();
                        }
                        g.DrawString(var_chr_um_clave.Substring(0, 3).ToString(), fuente, brsh, 25, renglon);
                        //g.DrawString(tmgrdpro.Rows[u][0].ToString().Substring(0, 4).ToString() + " " + tmgrdpro.Rows[u][1].ToString(), fuente, brsh, 40, renglon);
                        g.DrawString(tmgrdpro.Rows[u]["clave"].ToString() + " " + tmgrdpro.Rows[u]["nombre"].ToString(), fuente, brsh, 40, renglon);
                        g.DrawString(tmgrdpro.Rows[u]["lote"].ToString(), fuente, brsh, 130, renglon);

                        decimal imp_print = (Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                        decimal cos_print = imp_print / Convert.ToDecimal(tmgrdpro.Rows[u]["cantidad"].ToString());

                        if (lblPedimento.Text == "1")
                        {
                            g.DrawString(cos_print.ToString("###,###,##0.00"), fuente, brsh, posicionprecio(cos_print.ToString("###,###,##0.00")), renglon);//precio 
                        }
                        else
                        {
                            g.DrawString(Convert.ToDecimal(tmgrdpro.Rows[u]["unitario"].ToString()).ToString("###,###,##0.00"), fuente, brsh, posicionprecio(Convert.ToDecimal(tmgrdpro.Rows[u]["unitario"].ToString()).ToString("###,###,##0.00")), renglon);
                        }

                        //sum_gtos_im_print = sum_gtos_im_print + (Convert.ToDecimal(tmgrdpro.Rows[u]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[u]["flete"].ToString()));
                        //sum_importe_print = (Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));

                        
                        if (tmgrdpro.Rows[u]["costo"].ToString() != "" && Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) > 0)
                        {
                            decimal totprod = (Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + (Convert.ToDecimal(tmgrdpro.Rows[u]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[u]["flete"].ToString()));
                            if (lblPedimento.Text == "1")
                            {
                                g.DrawString(cos_print.ToString("###,###,##0.00"), fuente, brsh, posicionimporte(cos_print.ToString("###,###,##0.00")), renglon);//importe
                            }
                            else
                            {
                                g.DrawString(totprod.ToString("###,###,##0.00"), fuente, brsh, posicionimporte(totprod.ToString("###,###,##0.00")), renglon);//importe
                            }
                            //if (DDL_ent_procedencia.SelectedIndex == 0)
                            //{
                            //    decimal totprod = Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) + (Convert.ToDecimal(tmgrdpro.Rows[u]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[u]["flete"].ToString()));
                            //    g.DrawString(totprod.ToString("###,###,##0.00"), fuente, brsh, posicionimporte(totprod.ToString("###,###,##0.00")), renglon);//importe
                            //}
                            //else
                            //{
                            //    decimal totprod = (Convert.ToDecimal(tmgrdpro.Rows[u]["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + (Convert.ToDecimal(tmgrdpro.Rows[u]["gastoaduanal"].ToString()) + Convert.ToDecimal(tmgrdpro.Rows[u]["flete"].ToString()));
                            //    g.DrawString(totprod.ToString("###,###,##0.00"), fuente, brsh, posicionimporte(totprod.ToString("###,###,##0.00")), renglon);//importe
                            //}
                        }
                        //envio a otra pagina
                        renglon = renglon + 3;
                        renglonant++;
                        ren++;
                        if (ren == 20)
                        {
                            ren = 0;
                            break;
                        }
                    }
                    if (renglonant >= tmgrdpro.Rows.Count)
                    {
                        string ocs = "O.C. ";
                        for (int i = 1; i < lstOC.Items.Count; i++)
                        {
                            ocs += lstOC.Items[i].ToString() + ", ";
                        }
                        if (lstOC.Items.Count > 1)
                            g.DrawString(txtObservaciones.Text + " " + ocs.TrimEnd(','), fuente, brsh, 2, 116);//g.DrawString(txtObservaciones.Text + " " + ocs.TrimEnd(','), fuente, brsh, 2, 122);
                        else
                            g.DrawString(txtObservaciones.Text, fuente, brsh, 2, 116); //g.DrawString(txtObservaciones.Text, fuente, brsh, 2, 122);
                        //g.DrawString(txtObservaciones.Text + ocs, fuente, brsh, 2, 116);

                        g.DrawString("Gtos. Imp.", fuente, brsh, 150, 91);
                        g.DrawString(Convert.ToDecimal(sum_gtos_im_print).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(sum_gtos_im_print).ToString("###,###,##0.00")), 91);
                        
                        
                        g.DrawString("I.E.P.S.", fuente, brsh, 150, 94);
                        string total_iepsx = "";
                        total_iepsx = txt_ent_ieps.Text;// (txt_ent_ieps.Text == "") ? "0.00" : txt_ent_ieps.Text;

                        g.DrawString("Subtotal", fuente, brsh, 150, 97);
                        string subtotalx = "";
                        subtotalx = txt_subtotal2.Text;
                        if (lblPedimento.Text == "1")
                        {
                            g.DrawString(Convert.ToDecimal(sum_importe_print).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(sum_importe_print).ToString("###,###,##0.00")), 97);
                        }
                        else
                        {
                            g.DrawString(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00")), 97);
                        }
                        

                        g.DrawString("Descuento", fuente, brsh, 150, 100);
                        string descuentox = "";
                        descuentox = txtDesc.Text;
                        g.DrawString(Convert.ToDecimal(descuentox).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(descuentox).ToString("###,###,##0.00")), 100);

                        g.DrawString("I.V.A.", fuente, brsh, 150, 103);
                        string ivax = "";
                        ivax = txt_iva.Text;
                        g.DrawString(Convert.ToDecimal(ivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(ivax).ToString("###,###,##0.00")), 103);

                        g.DrawString("Ret IVA", fuente, brsh, 150, 106);
                        string retivax = "";
                        retivax = txtRetIva.Text;
                        g.DrawString(Convert.ToDecimal(retivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(retivax).ToString("###,###,##0.00")), 106);
                        
                        g.DrawString("Ret ISR", fuente, brsh, 150, 109);
                        string retisrx = "";
                        retisrx = txtIsr.Text;
                        g.DrawString(Convert.ToDecimal(retisrx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(retisrx).ToString("###,###,##0.00")), 109);

                        g.DrawString("Cedular", fuente, brsh, 150, 112);
                        string cedularx = "";
                        cedularx = txtCedular.Text;
                        g.DrawString(Convert.ToDecimal(cedularx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(cedularx).ToString("###,###,##0.00")), 112);

                        g.DrawString("Total", fuente, brsh, 150, 115);//((DDL_ent_procedencia.SelectedIndex == 0) ? "MN" : "USD")
                        string totalx = "";
                        totalx = txt_total.Text;
                        g.DrawString(Convert.ToDecimal(totalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(totalx).ToString("###,###,##0.00")), 115);

                        g.DrawString("MN", fuente, brsh, 195, 115);

                        //bool entra = false;
                        //g.DrawString("Subtotal", fuente, brsh, 150, 92);//g.DrawString("Subtotal", fuente, brsh, 150, 94);//112
                        //string total_iepsx = ""; string ivax = ""; string subtotalx = ""; string totalx = "";
                        //total_iepsx = (txt_ent_ieps.Text == "") ? "0.00" : txt_ent_ieps.Text;
                        //ivax = txt_iva.Text;
                        //subtotalx = txt_subtotal2.Text;
                        //totalx = txt_total.Text;
                        //if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) == 0)
                        //{
                        //}
                        //else if (Convert.ToDecimal(total_iepsx) != 0 && Convert.ToDecimal(ivax) == 0)
                        //{
                        //    g.DrawString("I.E.P.S.", fuente, brsh, 150, 95);//g.DrawString("I.E.P.S.", fuente, brsh, 150, 97);//115
                        //}
                        //else if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) != 0)
                        //{
                        //    g.DrawString("I.V.A.", fuente, brsh, 150, 95); //g.DrawString("I.V.A.", fuente, brsh, 150, 97);//115
                        //}
                        //else
                        //{
                        //    //g.DrawString("I.E.P.S.", fuente, brsh, 150, 97);//115
                        //    //g.DrawString("I.V.A.", fuente, brsh, 150, 100);//118
                        //    //g.DrawString("Total", fuente, brsh, 150, 103);//121
                        //    g.DrawString("I.E.P.S.", fuente, brsh, 150, 95);//115
                        //    g.DrawString("I.V.A.", fuente, brsh, 150, 98);//118
                        //    g.DrawString("Total", fuente, brsh, 150, 101);//121
                        //    entra = true;
                        //}
                        //if (entra == false)
                        //    g.DrawString("Total", fuente, brsh, 150, 98);//98

                        //g.DrawString(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(subtotalx).ToString("###,###,##0.00")), 92);//112
                        //if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) == 0)
                        //{
                        //}
                        //else if (Convert.ToDecimal(total_iepsx) != 0 && Convert.ToDecimal(ivax) == 0)
                        //{
                        //    g.DrawString(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00")), 95);//115
                        //}
                        //else if (Convert.ToDecimal(total_iepsx) == 0 && Convert.ToDecimal(ivax) != 0)
                        //{
                        //    g.DrawString(Convert.ToDecimal(ivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(ivax).ToString("###,###,##0.00")), 95);//115
                        //}
                        //else
                        //{
                        //    g.DrawString(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(total_iepsx).ToString("###,###,##0.00")), 95);//115
                        //    g.DrawString(Convert.ToDecimal(ivax).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(ivax).ToString("###,###,##0.00")), 98);//118
                        //    g.DrawString(Convert.ToDecimal(totalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(totalx).ToString("###,###,##0.00")), 101);//121
                        //}
                        //if (entra == false)
                        //    g.DrawString(Convert.ToDecimal(totalx).ToString("###,###,##0.00"), fuente, brsh, posicion2(Convert.ToDecimal(totalx).ToString("###,###,##0.00")), 98);//118


                        e.HasMorePages = false;
                    }
                    if (renglonant < tmgrdpro.Rows.Count)
                    {
                        g.DrawString("CONTINUA EN LA SIGUIENTE HOJA.....", fuente, brsh, 2, 116);
                        e.HasMorePages = true;
                        /*else
                            e.HasMorePages = true;*/
                    }
                }

                //ren = ren + 1;
                thisConnection.Close();
            }
            catch (SystemException ex)
            {
                thisConnection.Close();
                MessageBox.Show(ex.ToString());
            }
            
        }
        
        private int posicion(string val)
        {
            int x = 0;
            if (val.Length == 4)
            {
                x = 16;
            }
            if (val.Length == 5)
            {
                x = 14;
            }
            if (val.Length == 6)
            {
                x = 12;
            }
            if (val.Length == 7)
            {
                x = 10;
            }
            if (val.Length == 8)
            {
                x = 8;
            }
            if (val.Length == 9)
            {
                x = 6;
            }
            if (val.Length == 10)
            {
                x = 4;
            }
            if (val.Length == 11)
            {
                x = 2;
            }
            if (val.Length == 12)
            {
                x = 0;
            }
            if (val.Length == 13)
            {
                x = 0;
            }
            if (val.Length == 14)
            {
                x = 0;
            }
            return x;
        }

        private float posicion_pdf(string val)
        {
            float num = 0.0f;
            if (val.Length == 4)
                num = float.Parse("25.8");
            if (val.Length == 5)
                num = float.Parse("24.3");
            if (val.Length == 6)
                num = 23f;
            if (val.Length == 7)
                num = 10f;
            if (val.Length == 8)
                num = 20f;
            if (val.Length == 9)
                num = float.Parse("18.5");
            if (val.Length == 10)
                num = float.Parse("17");
            if (val.Length == 11)
                num = 2f;
            if (val.Length == 12)
                num = float.Parse("14");
            if (val.Length == 13)
                num = 0.0f;
            if (val.Length == 14)
                num = 0.0f;
            return num;
        }

        private int posicion2(string val)
        {
            int x = 0;
            if (val.Length == 4)
            {
                x = 187;
            }
            if (val.Length == 5)
            {
                x = 185;
            }
            if (val.Length == 6)
            {
                x = 183;
            }
            if (val.Length == 7)
            {
                x = 181;
            }
            if (val.Length == 8)
            {
                x = 179;
            }
            if (val.Length == 9)
            {
                x = 177;
            }
            if (val.Length == 10)
            {
                x = 175;
            }
            if (val.Length == 11)
            {
                x = 173;
            }
            if (val.Length == 12)
            {
                x = 171;
            }
            if (val.Length == 13)
            {
                x = 169;
            }
            if (val.Length == 14)
            {
                x = 167;
            }
            if (val.Length == 15)
            {
                x = 165;
            }
            return x;
        }

        private float posicion2_pdf(string val)
        {
            float num = 0.0f;
            if (val.Length == 4)
                num = 183f;
            if (val.Length == 5)
                num = float.Parse("200.5");
            if (val.Length == 6)
                num = float.Parse("199");
            if (val.Length == 7)
                num = 181f;
            if (val.Length == 8)
                num = 196f;
            if (val.Length == 9)
                num = 194f;
            if (val.Length == 10)
                num = float.Parse("192.3");
            if (val.Length == 11)
                num = 173f;
            if (val.Length == 12)
                num = float.Parse("189.4");
            if (val.Length == 13)
                num = 169f;
            if (val.Length == 14)
                num = 167f;
            if (val.Length == 15)
                num = 165f;
            return num;
        }

        private int posicionprecio(string val)
        {
            int x = 0;
            if (val.Length == 4)
            {
                x = 155;
            }
            if (val.Length == 5)
            {
                x = 154;
            }
            if (val.Length == 6)
            {
                x = 153;
            }
            if (val.Length == 7)
            {
                x = 152;
            }
            if (val.Length == 8)
            {
                x = 151;
            }
            if (val.Length == 9)
            {
                x = 150;
            }
            if (val.Length == 10)
            {
                x = 149;
            }
            return x;
        }

        private float posicionprecio_pdf(string val)
        {
            float num = 0.0f;
            if (val.Length == 4)
                num = 172f;
            if (val.Length == 5)
                num = float.Parse("170.5");
            if (val.Length == 6)
                num = 169f;
            if (val.Length == 7)
                num = 152f;
            if (val.Length == 8)
                num = 151f;
            if (val.Length == 9)
                num = 150f;
            if (val.Length == 10)
                num = 149f;
            return num;
        }

        private int posicionimporte(string val)
        {
            int x = 0;
            if (val.Length == 4)
            {
                x = 186;
            }
            if (val.Length == 5)
            {
                x = 184;
            }
            if (val.Length == 6)
            {
                x = 182;
            }
            if (val.Length == 7)
            {
                x = 180;
            }
            if (val.Length == 8)
            {
                x = 178;
            }
            if (val.Length == 9)
            {
                x = 176;
            }
            if (val.Length == 10)
            {
                x = 174;
            }
            if (val.Length == 11)
            {
                x = 172;
            }
            if (val.Length == 12)
            {
                x = 170;
            }
            if (val.Length == 13)
            {
                x = 168;
            }
            if (val.Length == 14)
            {
                x = 166;
            }
            if (val.Length == 15)
            {
                x = 164;
            }
            return x;
        }

        private float posicionimporte_pdf(string val)
        {
            float num = 0.0f;
            if (val.Length == 4)
                num = 200f;
            if (val.Length == 5)
                num = float.Parse("199.5");
            if (val.Length == 6)
                num = 198f;
            if (val.Length == 7)
                num = 180f;
            if (val.Length == 8)
                num = 195f;
            if (val.Length == 9)
                num = float.Parse("193.5");
            if (val.Length == 10)
                num = float.Parse("192");
            if (val.Length == 11)
                num = 172f;
            if (val.Length == 12)
                num = 170f;
            if (val.Length == 13)
                num = 168f;
            if (val.Length == 14)
                num = 166f;
            if (val.Length == 15)
                num = 164f;
            return num;
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            try
            {
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                thisConnection.Open();
                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "SELECT TOP 1 inicio_sesion, usu_login FROM tb_cat_historial_dia where nombre_maquina = '" + Environment.MachineName + "' AND sistema = 'SISEMP' ORDER BY inicio_sesion desc";
                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    Utilerias.Class1.Inicio_sesion = reader2.GetSqlDateTime(0).Value;
                    Utilerias.Class1.Usu_login = reader2.GetSqlString(1).ToString();
                    Utilerias.Class1.Nombre_equipo = Environment.MachineName;
                }
                reader2.Close();

                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "update tb_cat_historial_dia set formulario = ' ' where nombre_maquina ='" + Utilerias.Class1.Nombre_equipo + "' and usu_login = '" + Utilerias.Class1.Usu_login + "' and inicio_sesion = '" + Utilerias.Class1.Inicio_sesion.ToString("s") + "' AND sistema = 'SISEMP'";
                reader2 = cmnd2.ExecuteReader();
                reader2.Close();
                thisConnection.Close();

                string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                using (StreamWriter sw = File.AppendText(filelog))
                {
                    sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Salida 2.4 Entradas de Almacen");
                    sw.Close();
                }

                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.5", ex.ToString(), "SISEMP");
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + ex.ToString());
            }
        }

        private void txt_sis_tipo_cambio_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (tmgrdpro.Rows.Count > 0)
                {
                    
                    decimal a = 0;
                    decimal b = 0;
                    foreach(DataRow rw in tmgrdpro.Rows)
                    {
                        if (rw["borrar"].ToString() == "1")
                            continue;
                        a = (Convert.ToDecimal(rw["costo"].ToString()) / Convert.ToDecimal(rw["cantidad"].ToString())) * Convert.ToDecimal(txt_sis_tipo_cambio.Text);
                        rw["unitario"] = a.ToString();

                        //verificar si existe ieps
                        if (Convert.ToDecimal(rw["iepsimporte"].ToString()) > 0)
                        {
                            b = (Convert.ToDecimal(rw["costo"].ToString()) * iepsporc) * Convert.ToDecimal(txt_sis_tipo_cambio.Text);
                            rw["iepsimporte"] = b.ToString();
                            rw["ieps"] = true;
                        }
                    }

                    if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
                    {
                        txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                    }
                    if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
                    {
                        txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                    }
                    calculatotales();
                    dtgDetalle.Rows.Clear();
                    foreach (DataRow rr in tmgrdpro.Rows)
                    {
                        if (rr["borrar"].ToString() == "1")
                            continue;
                        //dtgDetalle.Rows.Add(rr.ItemArray);
                        dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
                            Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
                            Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                            Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                            Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                            rr["borrar"].ToString(), rr["oper"].ToString());
                    }
                }
                else
                {

                    if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
                    {
                        txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                    }
                    if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
                    {
                        txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                    }
                    calculatotales();
                    dtgDetalle.Rows.Clear();
                    foreach (DataRow rr in tmgrdpro.Rows)
                    {
                        if (rr["borrar"].ToString() == "1")
                            continue;
                        //dtgDetalle.Rows.Add(rr.ItemArray);
                        dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
                            Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
                            Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                            Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                            Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                            rr["borrar"].ToString(), rr["oper"].ToString());
                    }
                }
            }
        }

        private void traer_costopromedio(string folio, string cve_empaque, string fecha_ent)
        {
            string query = "";
            int bandera = 0;//CAMBIA A 1 SI EL EMPAQUE ESTA EN HISTORICO
            int bandera1 = 0;//CAMBIA A 1 SI EL EMPAQUE ESTA EN INVENTARIO_EMP
            decimal invemp_inicial = 0;
            decimal invemp_costopro = 0;
            try
            {
                string f1 = "";
                //string f = Convert.ToDateTime(fecha_ent).AddDays(-1).ToShortDateString();
                //cmnd4 = thisConnection.CreateCommand();
                //query = "SELECT (hrp_cantidad * hrp_costo) as total, hrp_cantidad FROM tb_historico_recepcion WHERE hrp_folio = '" + folio + "' AND emp_clave = '" + cve_empaque + "' " +
                //    "AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR' AND hrp_fecha > '" + f + "'";
                //reader4 = cmnd4.ExecuteReader();
                //if (reader1.HasRows)
                //{
                //    while (reader4.Read())
                //    {
                //        nuevo_costo = nuevo_costo + Convert.ToDecimal(reader4.GetValue(0).ToString().Trim());
                //        nueva_existencia = nueva_existencia + Convert.ToDecimal(reader4.GetValue(1).ToString().Trim());
                //    }
                //}
                //nueva_costoprom = nuevo_costo / nueva_existencia;


                //FECHA ULTIMO CIERRE DE MES
                cmnd4 = thisConnection.CreateCommand();
                cmnd4.CommandText = "SELECT TOP 1 invemp_fecha FROM tb_mstr_inventario_emp ORDER BY invemp_fecha DESC";
                reader4 = cmnd4.ExecuteReader();
                if (reader4.HasRows)
                {
                    reader4.Read();
                    f1 = Convert.ToDateTime(reader4.GetValue(0).ToString().Trim()).ToShortDateString();
                }
                cmnd4.Dispose();

                //TOTAL DE ENTRADAS Y SALIDAS DE EMPAQUE EN HISTORICO
                decimal totsal = 0;
                decimal totent = 0;
                cmnd4 = thisConnection.CreateCommand();
                cmnd4.CommandText = "SELECT hrp_recibo, hrp_cantidad, hrp_tipo_recepcion, hrp_situacion FROM tb_historico_recepcion WHERE hrp_estatus <> 'C'" +  
                    " AND emp_clave = '" + cve_empaque + "' AND hrp_fecha >= '" + f1 + "' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
                reader4 = cmnd4.ExecuteReader();
                if (reader4.HasRows)
                {
                    while (reader4.Read())
                    {
                        if (reader4.GetValue(2).ToString().Trim() == "ENT" && reader4.GetValue(3).ToString().Trim() == "NOR")
                            totent = totent + Convert.ToDecimal(reader4.GetValue(1).ToString().Trim());
                        if (reader4.GetValue(2).ToString().Trim() == "ENT" && reader4.GetValue(3).ToString().Trim() == "DEV")
                            totent = totent + Convert.ToDecimal(reader4.GetValue(1).ToString().Trim());
                        if (reader4.GetValue(2).ToString().Trim() == "SAL")
                            totsal = totsal + Convert.ToDecimal(reader4.GetValue(1).ToString().Trim());
                    }
                }
                reader4.Close();
                reader4.Dispose();
                cmnd4.Dispose();

                //INICIAL DE EMPAQUE
                decimal inicial = 0;
                decimal totexistencia = 0;
                cmnd4 = thisConnection.CreateCommand();
                cmnd4.CommandText = "SELECT TOP 1 invemp_inicial, invemp_costopro FROM tb_mstr_inventario_emp WHERE emp_clave = '" + cve_empaque + "' AND invemp_fecha = '" + f1 + "' ORDER BY invemp_fecha DESC";
                reader4 = cmnd4.ExecuteReader();
                if (reader4.HasRows)
                {
                    reader4.Read();
                    inicial = Convert.ToDecimal(reader4.GetValue(0).ToString().Trim());
                    invemp_inicial = Convert.ToDecimal(reader4.GetValue(0).ToString().Trim());
                    invemp_costopro = Convert.ToDecimal(reader4.GetValue(1).ToString().Trim());
                }
                cmnd4.Dispose();

                //TOTAL DE EXISTENCIAS
                totexistencia = (inicial + totent) - totsal;

                //COSTO PROMEDIO
                cmnd4 = thisConnection.CreateCommand();
                //query = "SELECT (hrp_cantidad * hrp_costo) as total, hrp_cantidad FROM tb_historico_recepcion WHERE hrp_recibo = '" + folio + "' AND emp_clave = '" + cve_empaque + "' " +
                //    "AND hrp_tipo_recepcion ='ENT' AND hrp_situacion = 'NOR' AND hrp_fecha >= '" + f1 + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND hrp_estatus <> 'C'";
                query = "SELECT (hrp_cantidad * hrp_costo) as total, hrp_cantidad FROM tb_historico_recepcion WHERE emp_clave = '" + cve_empaque + "' " +
                    "AND hrp_tipo_recepcion ='ENT' AND hrp_situacion = 'NOR' AND hrp_fecha >= '" + f1 + "' AND alm_clave = '" + glo_var_chr_alm_clave + "' AND hrp_estatus <> 'C'";
                cmnd4.CommandText = query;
                reader4 = cmnd4.ExecuteReader();
                if (reader4.HasRows)
                {
                    while (reader4.Read())
                    {
                        nuevo_costo = nuevo_costo + Convert.ToDecimal(reader4.GetValue(0).ToString().Trim());
                        nueva_existencia = nueva_existencia + Convert.ToDecimal(reader4.GetValue(1).ToString().Trim());
                    }
                    bandera = 1;
                }
                reader4.Close();
                reader4.Dispose();
                cmnd4.Dispose();
                if (nueva_existencia > 0)
                {
                    nueva_costoprom = nuevo_costo / nueva_existencia;

                    //ACTUALIZAR TB_CAT_EMPAQUES (EMP_COSTO, EMP_EXISTENCIA)
                    cmnd4 = thisConnection.CreateCommand();
                    cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + nueva_costoprom + "', emp_existencia = '" + totexistencia + "' WHERE emp_clave = '" + cve_empaque + "'";
                    cmnd4.ExecuteNonQuery();
                    cmnd4.Dispose();
                }
                else
                {
                    //ACTUALIZAR TB_CAT_EMPAQUES (EMP_COSTO, EMP_EXISTENCIA)
                    cmnd4 = thisConnection.CreateCommand();
                    cmnd4.CommandText = "UPDATE tb_cat_empaques SET emp_costo = '" + invemp_costopro + "', emp_existencia = '" + invemp_inicial + "' WHERE emp_clave = '" + cve_empaque + "'";
                    cmnd4.ExecuteNonQuery();
                    cmnd4.Dispose();
                }
                
            }
            catch(SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de modificar la entrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.5", txtFolio.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }
        }

        private void btnMinimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void txt_ent_factura_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (DDL_ent_procedencia.SelectedIndex == 1)
                {
                    txt_ent_pedimento.Focus();
                }
                else
                {
                    txtFlete.Focus();
                }
            }
        }

        private void DDLElaboro_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (DDLElaboro.SelectedIndex > -1)
            {
                //
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT estatus FROM tb_cat_usuarios_emp WHERE usu_nombre = '" + DDLElaboro.SelectedItem + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    
                    if (reader1["estatus"].ToString().Trim() == "B")
                    {
                        MessageBox.Show("Usuario dado de baja, no puede ser seleccionado", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();
                        thisConnection.Close();
                        DDLElaboro.SelectedIndex = -1;
                        return;
                    }
                }
                thisConnection.Close();
                DDLAutorizo.Focus();
                DDLAutorizo.SelectedIndex = 0;
                DDLAutorizo.DroppedDown = true;
            }
        }

        private void pbxCalendarUlEn_Click(object sender, EventArgs e)
        {
            calendario dlg = new calendario();
            dlg.ShowDialog();
            if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                dtpFecha.Text = form_entradas_ocs.calendario.SharedData.Polino.fecha;
            }
        }

        private void chkBorrados_CheckStateChanged(object sender, EventArgs e)
        {
            if (chkBorrados.Checked == true)
            {
                foreach (DataRow rw in tmgrdpro.Select("borrar = '1'"))
                {
                    rw["borrar"] = "";
                }
                foreach (DataRow rr in tmgrdpro.Rows)
                {
                    //dtgDetalle.Rows.Add(rr.ItemArray);
                    dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.0000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                        Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                        rr["borrar"].ToString(), rr["oper"].ToString());
                }
                if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
                {
                    txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
                {
                    txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                calculatotales();
                chkBorrados.Checked = false;
            }
            else
            {
                if (tmgrdpro.Rows.Count == 0)
                {
                    MessageBox.Show("No hay datoa para filtrar", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //using (ZipFile zip = ZipFile.Read(@"\\gabira1\entradasfact\ent\Ent_079957_Materiales.zip"))
            //{
            //    zip.AddFile(@"C:\Users\Angel Escamilla\Pictures\3018625_640px.jpg", "");
            //    zip.Save();
            //}
            //////imprimirtxt();
            //////guarda_genera_pdf_servidor();
            
            //////recalculo_isr_descuento();

            ////foreach (DataGridViewRow gr in dtgDetalle.Rows)
            ////{
            ////    if (gr.Cells[0].Value == null)
            ////        continue;
            ////    string e_m_p = gr.Cells[0].Value.ToString();
                
            ////    //recalculo_costos(e_m_p);
            ////    //recalculo_existencia(e_m_p);
            ////    modificar_ultima_entrada(e_m_p);
            ////}

            //bool fndDiff = VerificarOrdenCompraEntrada(dtdetalles);
            ModificaOrdenCompraEntrada(dtdetalles);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(operacion != "consulta")
                dtgDetalle.Rows.Clear();
            tmgrdpro.Clear();
            bool fnd = false;
            foreach (DataRow r1 in dtdetalles.Rows)
            {
                foreach (DataRow r2 in tmgrdpro.Select("clave = '" + r1["emp_clave"].ToString() + "'"))
                {
                    fnd = true;
                    r2["cantidad"] = (Convert.ToDecimal(r2["cantidad"].ToString()) + Convert.ToDecimal(r1["cantidad"].ToString())).ToString();
                    r2["costo"] = (Convert.ToDecimal(r2["costo"].ToString()) + Convert.ToDecimal(r1["importe"].ToString())).ToString();
                }

                if (fnd == false)
                {
                    DataRow rw;
                    //foreach (DataRow r2 in dtdetalles.Rows)
                    //{
                    rw = tmgrdpro.NewRow();
                    rw["clave"] = r1["emp_clave"].ToString();
                    rw["nombre"] = nombreempaque(r1["emp_clave"].ToString());
                    rw["cantidad"] = r1["cantidad"].ToString();
                    string costo = "0.000";
                    if (Convert.ToDecimal(r1["importe"].ToString()) == 0 || Convert.ToDecimal(r1["cantidad"].ToString()) == 0)
                        costo = "0.000";
                    else
                        costo = (Convert.ToDecimal(r1["importe"].ToString()) / Convert.ToDecimal(r1["cantidad"].ToString())).ToString();

                    if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) > 1)
                    {
                        rw["unitario"] = (Convert.ToDecimal(costo) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                    }
                    else
                    {
                        rw["unitario"] = costo;
                    }
                    rw["costo"] = r1["importe"].ToString();
                    rw["adv"] = "0";
                    rw["dta"] = "0";
                    rw["lote"] = r1["lote"].ToString();

                    if (fnd_ieps == true)
                    {
                        rw["ieps"] = true;
                        rw["iepsimporte"] = Math.Round(Convert.ToDecimal(r1["importe"].ToString()) * Convert.ToDecimal("0.08"), 4);
                    }
                    else
                    {
                        if (Convert.ToDecimal(r1["ieps"].ToString()) > 0)
                        {
                            rw["ieps"] = true;
                            rw["iepsimporte"] = Convert.ToDecimal(r1["ieps"].ToString());
                        }
                        else
                        {
                            rw["ieps"] = false;
                            rw["iepsimporte"] = Convert.ToDecimal(r1["ieps"].ToString());
                        }
                    }
                    
                    rw["gastoaduanal"] = "0";
                    rw["flete"] = "0";

                    if(DDL_ent_procedencia.SelectedIndex == 1)
                    {
                        if (r1["iva"].ToString() != "")
                            rw["iva"] = (Convert.ToDecimal(r1["iva"].ToString()) > 0) ? (((Convert.ToDecimal(r1["importe"].ToString()) + Convert.ToDecimal(r1["ieps"].ToString())) * Convert.ToDecimal("0.16")) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString() : "0.0000";
                        else
                            rw["iva"] = "0.0000";
                    }
                    else
                    {
                        if (r1["iva"].ToString() != "")
                            rw["iva"] = (Convert.ToDecimal(r1["iva"].ToString()) > 0) ? ((Convert.ToDecimal(r1["importe"].ToString()) + Convert.ToDecimal(r1["ieps"].ToString())) * Convert.ToDecimal("0.16")).ToString() : "0.0000";
                        else
                            rw["iva"] = "0.0000";
                    }
                    
                    tmgrdpro.Rows.Add(rw);

                    //}
                }
                fnd = false;
            }

            if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
            {
                txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
            }

            if (Convert.ToDecimal(txtFleteImp.Text) > 0)
                txtFleteImp_KeyPress(sender, new KeyPressEventArgs((char)Keys.Enter));

            if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
            {
                txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
            }
            calculatotales();
            
           dtgDetalle.Rows.Clear();
            foreach (DataRow rr in tmgrdpro.Rows)
            {
                dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
                    Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.0000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
                    Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                    Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                    Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                    rr["borrar"].ToString(), rr["oper"].ToString(), "", Convert.ToDecimal(rr["iva"].ToString()).ToString("###,###,##0.000"));
            }
            
        }

        private string nombreempaque(string clave)
        {
            string nome = "";
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT emp_nombre FROM tb_cat_empaques WHERE emp_clave = '" + clave + "'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                nome = reader1.GetValue(0).ToString().Trim();
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();

            return nome;
        }

        private void dtgDetalle2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                if (dtgDetalle2.Rows[e.RowIndex].Cells[2].Value == null)
                    return;

                string emp_clave = dtgDetalle2.Rows[e.RowIndex].Cells[0].Value.ToString();
                string costo = dtgDetalle2.Rows[e.RowIndex].Cells[5].Value.ToString();

                string cant_original = dtgDetalle2.Rows[e.RowIndex].Cells["cant_orig"].Value.ToString();
                if (Convert.ToDecimal(cant_original) < Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells[2].Value.ToString()))
                {
                    if (MessageBox.Show("La cantidad ingresada producira una Orden de Compra con saldo negativo, ¿Desea continuar?", "SISEMP",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        if (MessageBox.Show("¿Está seguro del movimiento?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            dtdetalles.Rows[e.RowIndex]["resp"] = "1";
                            dtgDetalle2.Rows[e.RowIndex].Cells["resp"].Value = "1";
                        }
                        else
                        {
                            dtdetalles.Rows[e.RowIndex]["cantidad"] = dtdetalles.Rows[e.RowIndex]["cant_orig"];
                            dtgDetalle2.Rows[e.RowIndex].Cells[2].Value = Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells["cant_orig"].Value.ToString()).ToString("###,###,##0.000");
                            dtdetalles.Rows[e.RowIndex]["resp"] = "0";
                            dtgDetalle2.Rows[e.RowIndex].Cells["resp"].Value = "0";
                            return;
                        }
                    }
                    else
                    {
                        dtdetalles.Rows[e.RowIndex]["cantidad"] = dtdetalles.Rows[e.RowIndex]["cant_orig"];
                        dtgDetalle2.Rows[e.RowIndex].Cells[2].Value = Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells["cant_orig"].Value.ToString()).ToString("###,###,##0.000");
                        return;
                    }
                }
                else
                {
                    dtdetalles.Rows[e.RowIndex]["resp"] = "0";
                    dtgDetalle2.Rows[e.RowIndex].Cells["resp"].Value = "0";
                }

                dtdetalles.Rows[e.RowIndex]["cantidad"] = dtgDetalle2.Rows[e.RowIndex].Cells[2].Value.ToString();
                dtgDetalle2.Rows[e.RowIndex].Cells[2].Value = Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells[2].Value.ToString()).ToString("###,###,##0.000");

                dtdetalles.Rows[e.RowIndex]["importe"] = Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells[2].Value.ToString()) * Convert.ToDecimal(costo);
                dtdetalles.Rows[e.RowIndex]["importe_org"] = Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells[2].Value.ToString()) * Convert.ToDecimal(costo);
                dtdetalles.Rows[e.RowIndex]["importe_orig"] = Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells[2].Value.ToString()) * Convert.ToDecimal(costo);//
                dtgDetalle2.Rows[e.RowIndex].Cells[3].Value = (Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells[2].Value.ToString()) * Convert.ToDecimal(costo)).ToString("###,###,##0.000");
                dtgDetalle2.Rows[e.RowIndex].Cells["imp_orig"].Value = (Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells[2].Value.ToString()) * Convert.ToDecimal(costo)).ToString("###,###,##0.000");

                //ieps
                if (dtgDetalle2.Rows[e.RowIndex].Cells["ieps1"].Value.ToString() != "")
                {
                    if (Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells["ieps1"].Value.ToString()) > 0)
                    {
                        dtdetalles.Rows[e.RowIndex]["ieps"] = Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["importe"]) * (Convert.ToDecimal("8") / 100);
                        dtgDetalle2.Rows[e.RowIndex].Cells["ieps1"].Value = (Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["importe"]) * (Convert.ToDecimal("8") / 100)).ToString("###,###,##0.000");
                    }
                }
                

                //iva
                if (dtgDetalle2.Rows[e.RowIndex].Cells["checkiva"].Value.ToString() == "1")
                {
                    dtdetalles.Rows[e.RowIndex]["iva"] = (Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["importe"]) + Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["ieps"])) * (Convert.ToDecimal("16") / 100);
                    dtgDetalle2.Rows[e.RowIndex].Cells["iva_det"].Value = ((Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["importe"]) + Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["ieps"])) * (Convert.ToDecimal("16") / 100)).ToString("###,###,##0.000");
                }
                

                string no_oc = dtgDetalle2.Rows[e.RowIndex].Cells[1].Value.ToString();
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT ieps_oc FROM tb_det_ordenescompra WHERE numero_oc = '" + no_oc + "' AND prod_clave = '" + emp_clave + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    if (Convert.ToDecimal(reader1["ieps_oc"].ToString()) > 0)
                        fnd_ieps = true;
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();

                dtgDetalle.Rows.Clear();
                button2_Click(null, null);
                //Cargar datos al datagrid

                if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
                {
                    txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                if (txtFleteImp.Text != "" || Convert.ToDecimal(txtFleteImp.Text) > 0)
                {
                    txtFleteImp_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
                {
                    txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                calculatotales();

                recalculo_isr_descuento();

                var imp_desc = dtdetalles.AsEnumerable().Sum(x => Convert.ToDecimal(x["importe"]));
                var imp_orig = dtdetalles.AsEnumerable().Sum(x => Convert.ToDecimal(x["importe_orig"]));

                lblImporteDescuento.Text = Math.Round((Convert.ToDecimal(imp_orig) - Convert.ToDecimal(imp_desc)), 4).ToString();
            }
            if (e.ColumnIndex == 3)
            {
                if (dtgDetalle2.Rows[e.RowIndex].Cells[3].Value == null)
                    return;
                string emp_clave = dtgDetalle2.Rows[e.RowIndex].Cells[0].Value.ToString();
                dtdetalles.Rows[e.RowIndex]["importe"] = dtgDetalle2.Rows[e.RowIndex].Cells[3].Value.ToString();
                dtdetalles.Rows[e.RowIndex]["importe_org"] = dtgDetalle2.Rows[e.RowIndex].Cells[3].Value.ToString();
                dtdetalles.Rows[e.RowIndex]["importe_orig"] = dtgDetalle2.Rows[e.RowIndex].Cells[3].Value.ToString();
                dtgDetalle2.Rows[e.RowIndex].Cells[3].Value = Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells[3].Value.ToString()).ToString("###,###,##0.000");
                dtgDetalle2.Rows[e.RowIndex].Cells["imp_orig"].Value = Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells[3].Value.ToString()).ToString("###,###,##0.000");
                //decimal val1 = Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells["cant"].Value);
                //decimal val2 = Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells["import"].Value);
                //decimal val3 = val2 / val1;
                //dtgDetalle2.Rows[e.RowIndex].Cells["nvo_costo"].Value = val3.ToString("###,##0.00000");

                //ieps
                //ieps
                if (dtgDetalle2.Rows[e.RowIndex].Cells["ieps1"].Value.ToString() != "")
                {
                    if (Convert.ToDecimal(dtgDetalle2.Rows[e.RowIndex].Cells["ieps1"].Value.ToString()) > 0)
                    {
                        dtdetalles.Rows[e.RowIndex]["ieps"] = Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["importe"]) * (Convert.ToDecimal("8") / 100);
                        dtgDetalle2.Rows[e.RowIndex].Cells["ieps1"].Value = (Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["importe"]) * (Convert.ToDecimal("8") / 100)).ToString("###,###,##0.000");
                    }
                }

                //iva
                if (dtgDetalle2.Rows[e.RowIndex].Cells["checkiva"].Value.ToString() == "1")
                {
                    //dtdetalles.Rows[e.RowIndex]["iva"] = Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["importe"]) * (Convert.ToDecimal("16") / 100);
                    //dtgDetalle2.Rows[e.RowIndex].Cells["iva_det"].Value = (Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["importe"]) * (Convert.ToDecimal("16") / 100)).ToString("###,###,##0.000");
                    dtdetalles.Rows[e.RowIndex]["iva"] = (Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["importe"]) + Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["ieps"])) * (Convert.ToDecimal("16") / 100);
                    dtgDetalle2.Rows[e.RowIndex].Cells["iva_det"].Value = ((Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["importe"]) + Convert.ToDecimal(dtdetalles.Rows[e.RowIndex]["ieps"])) * (Convert.ToDecimal("16") / 100)).ToString("###,###,##0.000");
                }

                dtgDetalle.Rows.Clear();
                button2_Click(null, null);
                //Cargar datos al datagrid

                if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
                {
                    txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                if (txtFleteImp.Text != "" || Convert.ToDecimal(txtFleteImp.Text) > 0)
                {
                    txtFleteImp_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
                {
                    txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                calculatotales();

                recalculo_isr_descuento();

                var imp_desc = dtdetalles.AsEnumerable().Sum(x => Convert.ToDecimal(x["importe"]));
                var imp_orig = dtdetalles.AsEnumerable().Sum(x => Convert.ToDecimal(x["importe_orig"]));

                lblImporteDescuento.Text = Math.Round((Convert.ToDecimal(imp_orig) - Convert.ToDecimal(imp_desc)), 4).ToString();
            }

            if (e.ColumnIndex == 11)
            {
                if (dtgDetalle2.Rows[e.RowIndex].Cells[11].Value == null)
                    return;
                string emp_clave = dtgDetalle2.Rows[e.RowIndex].Cells[0].Value.ToString();
                string ordencompra = dtgDetalle2.Rows[e.RowIndex].Cells["ordencompra"].Value.ToString();
                foreach (DataRow rw in tmgrdpro.Select("clave = '" + emp_clave + "'"))
                {
                    rw["lote"] = dtgDetalle2.Rows[e.RowIndex].Cells[11].Value.ToString();
                }
                foreach (DataRow rx in dtdetalles.Rows)
                {
                    if (rx["emp_clave"].ToString() == emp_clave && rx["orden_compra"].ToString() == ordencompra)
                    {
                        rx["lote"] = dtgDetalle2.Rows[e.RowIndex].Cells[11].Value.ToString();
                    }
                }
                foreach (DataGridViewRow rw in dtgDetalle.Rows)
                {
                    if (rw.Cells["clave"].Value == null)
                        break;
                    if(rw.Cells["clave"].Value.ToString() == emp_clave)
                        rw.Cells["lote"].Value = dtgDetalle2.Rows[e.RowIndex].Cells[11].Value.ToString();
                }
            }
        }

        private void dtgDetalle2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                if (MessageBox.Show("¿Quitar empaque?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    return;
                string emp_clave = dtgDetalle2.Rows[e.RowIndex].Cells[0].Value.ToString();
                string oc_clave = dtgDetalle2.Rows[e.RowIndex].Cells[1].Value.ToString();

                string filtro = "";
                if (oc_clave == "")
                    filtro = "emp_clave = '" + emp_clave + "'";
                else
                    filtro = "emp_clave = '" + emp_clave + "' AND orden_compra = '" + oc_clave + "'";

                foreach (DataRow r1 in dtdetalles.Select(filtro))
                {
                    r1.Delete();
                }

                foreach (DataRow rw in tmgrdpro.Select("clave = '" + emp_clave + "'"))
                {
                    rw.Delete();
                }

                button2_Click(null, null);
                dtgDetalle2.Rows.Clear();
                //Cargar datos al datagrid
                foreach (DataRow rr in dtdetalles.Rows)
                {
                    ////dtgDetalle.Rows.Add(rr.ItemArray);
                    ////dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"), 
                    ////    Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"), 
                    ////    Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                    ////    Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                    ////    Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                    ////    rr["borrar"].ToString(), rr["oper"].ToString());
                    //dtgDetalle2.Rows.Add(rr["emp_clave"].ToString(), rr["orden_compra"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["importe"].ToString()).ToString("###,###,##0.000"),
                    //    rr["requisicion"].ToString(), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.00000"), rr["borrar"].ToString(), rr["operacion"].ToString(), "",
                    //    Convert.ToDecimal(rr["ieps"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["iva"].ToString()).ToString("###,###,##0.000"), rr["tiva"].ToString());
                    dtgDetalle2.Rows.Add(rr["emp_clave"].ToString(), rr["orden_compra"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["importe"].ToString()).ToString("###,###,##0.000"),
                        rr["requisicion"].ToString(), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.00000"), rr["borrar"].ToString(), rr["operacion"].ToString(), "",
                        Convert.ToDecimal(rr["ieps"].ToString()).ToString("###,###,##0.000"), "", rr["lote"].ToString(), rr["iva"].ToString(), rr["tiva"].ToString());
                }
                dtgDetalle2.Update();
                if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
                {
                    txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
                {
                    txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                calculatotales();
                dtgDetalle.Rows.Clear();
                foreach (DataRow rr in tmgrdpro.Rows)
                {
                    if (rr["borrar"].ToString() == "1")
                        continue;
                    //dtgDetalle.Rows.Add(rr.ItemArray);
                    dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.0000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                        Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                        rr["borrar"].ToString(), rr["oper"].ToString());
                }

                lstOC.Items.Clear();
                var x = (from row in dtdetalles.AsEnumerable()
                             select row["orden_compra"]).Distinct();
                foreach (var item in x)
                {
                    lstOC.Items.Add(item);
                }

                recalculo_isr_descuento();
            }

            //APARTADO PARA AGREGAR NUMEROS DE SERIE DE LOS MOTORES
            if (e.ColumnIndex == 1)
            {
                
            }
        }

        private void cargaempaque(string clave)
        {
            if (DDL_ent_procedencia.SelectedIndex == -1)
            {
                MessageBox.Show("Debe seleccionar si es nacional o importación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (tmgrdpro.Rows.Count > 0)
            {
                //Validar que no haya sido cargado ya el empaque
                //string clv2 = txtempaque.Text;
                foreach (DataRow rr in tmgrdpro.Rows)
                {
                    if (clave == rr["clave"].ToString())
                    {
                        MessageBox.Show("El empaque ya fué dado cargado", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dtgDetalle.Rows.Clear();
                        //foreach (DataRow rr2 in tmgrdpro.Rows)
                        //{
                        //    dtgDetalle.Rows.Add(rr2.ItemArray);
                        //}
                        if (dtgDetalle.Rows.Count > 0)
                        {
                            foreach (DataRow rr2 in tmgrdpro.Rows)
                            {
                                if (rr["borrar"].ToString() != "1")
                                {
                                    dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
                                        Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.0000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
                                        Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                                        Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                                        Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                                        rr["borrar"].ToString(), rr["oper"].ToString(), "0.0000");
                                }
                                //dtgDetalle.Rows.Add(rr.ItemArray);
                            }
                        }

                        return;
                    }
                }
            }


            dtgDetalle.Rows.Clear();
            dtgDetalle2.Rows.Clear();

            //cargar datos de la oc
            DataRow rw = tmgrdpro.NewRow();
            rw["clave"] = clave;
            foreach (DataRow rr in dtempaques.Select("emp_clave = '" + clave + "'"))
            {
                rw["nombre"] = rr["emp_nombre"];
            }
            rw["cantidad"] = "0.000";
            rw["unitario"] = "0.0000";
            rw["costo"] = "0.000";
            rw["adv"] = "0.000";
            rw["dta"] = "0.000";
            rw["lote"] = "";
            rw["ieps"] = false;
            rw["iepsimporte"] = "0.000";
            rw["gastoaduanal"] = "0.000";
            rw["flete"] = "0.000";
            rw["oc"] = "";
            rw["req"] = "";
            rw["surtido"] = "0.000";
            if (operacion == "consulta")
                rw["oper"] = "nueva";
            else
                rw["oper"] = "";
            rw["iva"] = "0.000";
            tmgrdpro.Rows.Add(rw);

            DataRow rx = dtdetalles.NewRow();
            rx["emp_clave"] = clave;
            rx["orden_compra"] = "";
            rx["cantidad"] = "0.000";
            rx["importe"] = "0.000";
            rx["requisicion"] = "";
            rx["costo"] = "0.00000";
            rx["borrar"] = "";
            rx["operacion"] = "";
            rx["renglon"] = "";
            rx["ieps"] = "0.000";
            rx["iva"] = "0.000";
            dtdetalles.Rows.Add(rx);
            
            foreach (DataRow rr in tmgrdpro.Rows)
            {
                if (rr["borrar"].ToString() != "1")
                {
                    dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.0000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.000"), rr["lote"].ToString(),
                        Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.000"),
                        Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                        rr["borrar"].ToString(), rr["oper"].ToString(), rr["iva"].ToString());
                }
            }
            foreach (DataRow rr in dtdetalles.Rows)
            {
                dtgDetalle2.Rows.Add(rr["emp_clave"].ToString(), rr["orden_compra"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["importe"].ToString()).ToString("###,###,##0.000"),
                        rr["requisicion"].ToString(), Convert.ToDecimal(rr["ieps"].ToString()).ToString("###,###,##0.00000"), rr["borrar"].ToString(), rr["operacion"].ToString(), "",
                        Convert.ToDecimal(rr["ieps"].ToString()).ToString("###,###,##0.000"));
            }
            if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
            {
                txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
            }
            if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
            {
                txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
            }

            dtgDetalle2.Focus();
            //dtgDetalle2.Rows[dtgDetalle2.Rows.Count-1].Cells[2].Selected = true;
            //this.dtgDetalle2.CurrentCell = this.dtgDetalle2[3, 0];       
            this.dtgDetalle2.CurrentCell = this.dtgDetalle2[2, dtgDetalle2.Rows.Count - 1];
                
        }

        public void empaques()
        {
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT emp_clave, emp_nombre FROM tb_cat_empaques WHERE alm_clave = '" + glo_var_chr_alm_clave + "' ORDER BY emp_clave";
            reader1 = cmnd1.ExecuteReader();
            DataRow rw;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    rw = dtempaques.NewRow();
                    rw["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                    rw["emp_nombre"] = reader1.GetValue(1).ToString().Trim();
                    dtempaques.Rows.Add(rw);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();
        }

        private void dtgDetalle_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                DataGridViewTextBoxEditingControl tb = e.Control as DataGridViewTextBoxEditingControl;
                tb.KeyPress -= dtgDetalle_KeyPress;
                tb.KeyPress += new KeyPressEventHandler(dtgDetalle_KeyPress);
            }
        }

        private void dtgDetalle_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.ToString() == e.KeyChar.ToString().ToUpper())
            {
                return;
            }
            e.Handled = true;
            SendKeys.Send(e.KeyChar.ToString().ToUpper());
        }

        private void txt_oc_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (e.KeyChar == (char)Keys.Enter)
            //{
            //    txt_oc_KeyPress(null, null);
            //}
            //else
            //{
            for (int y = 0; y < dtgOrdenes.Rows.Count; y++)
            {
                if (dtgOrdenes.Rows[y].Cells["numero_oc"].Value.ToString().StartsWith(txt_oc.Text))
                {
                    //dtgProductos.Rows[y].Selected = true;
                    dtgOrdenes.FirstDisplayedScrollingRowIndex = y;
                    break;
                }
            }
            //}
            
        }

        private bool validagrid_N_datatable()
        {
            bool bandera = false;
            bool fnd1 = false;
            foreach (DataRow rw in dtdetalles.Rows)
            {
                if (chkFactura.Checked == true)
                {
                    if (Convert.ToDecimal(rw["cantidad"]) <= 0)
                    {
                        bandera = true;
                        fnd1 = true;
                        break;
                    }
                }
                else
                {
                    if (Convert.ToDecimal(rw["cantidad"]) <= 0 || Convert.ToDecimal(rw["importe"]) <= 0)
                    {
                        bandera = true;
                        fnd1 = true;
                        break;
                    }
                }
                
            }

            if (fnd1 == false)
            {
                for (int i = 0; i < dtgDetalle2.Rows.Count; i++)
                {
                    if (chkFactura.Checked == true)
                    {
                        if (Convert.ToDecimal(dtgDetalle2.Rows[i].Cells["cant"].Value.ToString()) <= 0)
                        {
                            bandera = true;
                            break;
                        }
                    }
                    else
                    {
                        if (Convert.ToDecimal(dtgDetalle2.Rows[i].Cells["cant"].Value.ToString()) <= 0 || Convert.ToDecimal(dtgDetalle2.Rows[i].Cells["import"].Value.ToString()) <= 0)
                        {
                            bandera = true;
                            break;
                        }
                    }
                }
            }

            return bandera;
        }

        private bool validar_empaque_lote()
        {
            bool validado = false;
            thisConnection.Open();
            string falta = "Falta ingresar lote a los siguientes empaques:\n";
            foreach (DataRow rx in tmgrdpro.Select("lote = ''"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT emp_lote FROM tb_cat_empaques WHERE emp_clave = '" + rx["clave"].ToString() + "'";
                reader1 = cmnd1.ExecuteReader();

                if (reader1.HasRows)
                {
                    reader1.Read();
                    if (reader1.GetValue(0).ToString().Trim() == "1")
                    {
                        falta += rx["clave"].ToString() + " " + rx["nombre"].ToString() + "\n";
                        validado = true;
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
            }
            falta += "Favor de ingresar los valores";
            thisConnection.Close();
            if (validado == true)
            {
                MessageBox.Show(falta, "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return validado;
        }

        private void btnCargar_Click(object sender, EventArgs e)
        {
            if (lblFactNal.Text == "1" && (lblToc.Text == "D" || lblToc.Text == "E") && txtFolioFiscal.Text == "")
            {
                MessageBox.Show("Debe ingresar primeramente el XML", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }

            int num2 = Convert.ToDateTime(this.txtFechaFact.Text).Day;
            string str1 = num2.ToString();
            num2 = Convert.ToDateTime(this.txtFechaFact.Text).Month;
            string str2 = num2.ToString();
            DateTime dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
            num2 = dateTime.Year;
            string str3 = num2.ToString();
            this.thisConnection.Open();
            this.cmnd1 = this.thisConnection.CreateCommand();
            this.cmnd1.CommandText = "SELECT valor FROM tb_cat_tipocambio WHERE dia = '" + str1 + "' AND mes = '" + str2 + "' AND año = '" + str3 + "'";
            this.reader1 = this.cmnd1.ExecuteReader();
            bool flag1 = false;
            if (this.reader1.HasRows)
            {
                this.reader1.Read();
                this.lblTCFact.Text = this.reader1["valor"].ToString().Trim();
                flag1 = true;
            }
            this.reader1.Close();
            this.reader1.Dispose();
            this.cmnd1.Dispose();
            if (!flag1)
            {
                dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
                dateTime = dateTime.AddDays(-1.0);
                dateTime.ToShortDateString();
                dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
                num2 = dateTime.Day;
                string str4 = num2.ToString();
                dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
                num2 = dateTime.Month;
                string str5 = num2.ToString();
                dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
                num2 = dateTime.Year;
                string str6 = num2.ToString();
                this.thisConnection.Open();
                this.cmnd1 = this.thisConnection.CreateCommand();
                this.cmnd1.CommandText = "SELECT valor FROM tb_cat_tipocambio WHERE dia = '" + str4 + "' AND mes = '" + str5 + "' AND año = '" + str6 + "'";
                this.reader1 = this.cmnd1.ExecuteReader();
                if (this.reader1.HasRows)
                {
                    this.reader1.Read();
                    this.lblTCFact.Text = this.reader1["valor"].ToString().Trim();
                }
                this.reader1.Close();
                this.reader1.Dispose();
                this.cmnd1.Dispose();
            }
            this.thisConnection.Close();

            string oc = txt_oc.Text;//dtgOrdenes.CurrentRow.Cells[0].Value.ToString();
            string fecha_oc = "";

            //
            bool found = false;
            foreach(DataGridViewRow row in dtgOrdenes.Rows)
            {
                if (oc == row.Cells["numero_oc"].Value.ToString())
                {
                    found = true;
                    fecha_oc = row.Cells["nom"].Value.ToString();
                    break;
                }
            }

            if (found == false)
            {
                MessageBox.Show("El número de Orden de compra no existe, favor de verificarlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txt_oc.Text = "";
                return;
            }

            //----------02/05/2017----------//
            //VALIDACION DE FECHA DE ENTRADA CONTRA LA FECHA DE LA ORDEN DE COMPRA
            if (found == true)
            {
                if (Convert.ToDateTime(dtpFecha.Text) < Convert.ToDateTime(fecha_oc))
                {
                    MessageBox.Show("La fecha de la entrada no debe ser menor a la fecha de la orden de compra", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txt_oc.Text = "";
                    fecha_oc = "";
                    return;
                }
                else
                {
                    pbxCalendarUlEn.Enabled = false;
                }
            }
            //FIN VALIDACION DE FECHA DE ENTRADA CONTRA LA FECHA DE LA ORDEN DE COMPRA


            bool entro = false;
            if (DDL_ent_procedencia.SelectedIndex == 1)
            {
                if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) == 1)
                {
                    MessageBox.Show("Las ordenes son de importación, debe ingresar el tipo de cambio");
                    return;
                }
            }

            if (MessageBox.Show("¿Desea agregar esta orden de compra a la entrada?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                dtgDetalle2.Rows.Clear();
                //cargar datos de la oc
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT A.subtotal_oc, A.status_oc, A.tipo_oc, A.req_oc, B.cantidad_oc, B.unidad_oc, B.prod_clave, B.nomprod_oc, B.orden_oc, B.cargo_oc, " +
                    "B.precio_oc, B.importe_oc, B.ieps_oc, B.surtido_oc, C.urgente_oc FROM tb_mstr_ordencompra A, tb_det_ordenescompra B, tb_oc_cfdi C " +
                    "WHERE A.numero_oc = B.numero_oc AND A.numero_oc = C.numero_oc AND A.numero_oc = '" + oc + "'";
                reader1 = cmnd1.ExecuteReader();
                DataRow rw;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        decimal num_1 = Convert.ToDecimal(reader1.GetValue(4).ToString().Trim());
                        decimal num_2 = Convert.ToDecimal(reader1.GetValue(13).ToString().Trim());

                        if (num_1 <= num_2)
                        {
                            continue;
                        }

                        if (num_2 < 0)
                        {
                            continue;
                        }

                        //VALIDAR EMPAQUES YA CARGADOS
                        bool enc = false;
                        if (dtdetalles.Rows.Count > 0)
                        {
                            foreach (DataRow rx in dtdetalles.Select("emp_clave = '" + reader1.GetValue(6).ToString().Trim() + "' AND orden_compra = '" + oc + "'"))
                            {
                                enc = true;
                            }

                            if (enc == true)
                            {
                                enc = false;
                                continue;
                            }
                        }

                        rw = dtdetalles.NewRow();

                        //Se evalua que el empaque sea del almacen seleccionado
                        if (product_almacen(reader1.GetValue(6).ToString().Trim()) == true)
                        {
                            rw["emp_clave"] = reader1.GetValue(6).ToString().Trim();
                            decimal surt = 0;
                            surt = Convert.ToDecimal(reader1.GetValue(4).ToString().Trim()) - Convert.ToDecimal(reader1.GetValue(13).ToString().Trim());
                            if (surt > 0)
                            {
                                rw["cantidad"] = surt;
                                rw["importe"] = (surt * Convert.ToDecimal(reader1.GetValue(10).ToString().Trim()));
                            }

                            else
                            {
                                rw["cantidad"] = reader1.GetValue(4).ToString().Trim();
                                rw["importe"] = reader1.GetValue(11).ToString().Trim();
                            }
                            rw["ieps"] = reader1.GetValue(12).ToString().Trim();
                            rw["orden_compra"] = oc;
                            rw["requisicion"] = reader1.GetValue(3).ToString().Trim();
                            rw["costo"] = reader1.GetValue(10).ToString().Trim();
                            if (operacion == "consulta")
                            {
                                //rw["oper"] = "nueva";
                            }
                            rw["importe_org"] = rw["importe"];
                            rw["costo_org"] = rw["costo"];
                            entro = true;
                            dtdetalles.Rows.Add(rw);
                        }
                        else
                        {
                            MessageBox.Show("Los productos de esa orden de compra no corresponden a el almacen seleccionado", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }

                }

                if (dtdetalles.Rows.Count == 0)
                {
                    MessageBox.Show("La Orden de Compra ya fué surtida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (thisConnection.State == ConnectionState.Open)
                        thisConnection.Close();
                    return;
                }
                thisConnection.Close();

                //Validamos que los productos de la orden de compra no hayan sido ya surtidos
                button2_Click(null, null);
                foreach (DataRow rr in dtdetalles.Rows)
                {
                    dtgDetalle2.Rows.Add(rr["emp_clave"].ToString(), rr["orden_compra"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rr["importe"].ToString()).ToString("###,###,##0.000"),
                        rr["requisicion"].ToString(), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.00000"), rr["borrar"].ToString(), rr["operacion"].ToString(), "",
                        Convert.ToDecimal(rr["ieps"].ToString()).ToString("###,###,##0.000"), "", "", "", "", "", rr["urgente_oc"].ToString());
                }

                txt_oc.Text = "";
                if (entro == true)
                {
                    bool fnd = false;
                    foreach (var item in lstOC.Items)
                    {
                        if (item.ToString() == oc)
                        {
                            fnd = true;
                            break;
                        }
                    }
                    if (fnd == false)
                    {
                        lstOC.Items.Add(oc);
                    }
                }
                entro = false;

            }

            txt_oc.Text = "";
        }

        private void txt_prov_clave_Leave(object sender, EventArgs e)
        {
            //txt_prov_clave.LostFocus += new EventHandler(txt_prov_clave_LostFocus);
            //if (txt_prov_clave_LostFocus(null, null) == false)
            //{
            //    MessageBox.Show("No ha seleccionado al proveedor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    txt_prov_clave.Focus();
            //    return;
            //}
        }

        //private bool txt_prov_clave_LostFocus(object sender, EventArgs e)
        //{
        //    //BUSQUEDA DEL PROVEEDOR AL DEJAR EL CONTROL
        //    //bool fnd = false;
        //    //for (int i = 0; i < provclave.Count; i++)
        //    //{
        //    //    if (txt_prov_clave.Text == provclave[i].ToString())
        //    //    {
        //    //        DDLProveedor.SelectedIndex = i;
        //    //        fnd = true;
        //    //        break;
        //    //    }
        //    //}
        //    //return fnd;
        //}

        private void txtFleteImp_KeyPress(object sender, KeyPressEventArgs e)
        {
            //realizar calculo para asignar el flete a cada producto(prorrateo)
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (txtFleteImp.Text == "")
                {
                    MessageBox.Show("Ingrese un valor al flete", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (validarnumero(txtFleteImp.Text) == false)
                {
                    MessageBox.Show("El valor de flete no es numérico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                decimal totflete = 0;
                foreach (DataRow rw in tmgrdpro.Rows)
                {
                    if (rw["borrar"].ToString() != "1")
                        totflete = totflete + Convert.ToDecimal(rw["costo"].ToString());
                }
                if (totflete > 0)
                {
                    foreach (DataRow rw in tmgrdpro.Rows)
                    {
                        if (rw["borrar"].ToString() != "1")
                        {
                            if (DDL_ent_procedencia.SelectedIndex == 0)
                            {
                                rw["flete"] = ((Convert.ToDecimal(rw["costo"]) / totflete) * (Convert.ToDecimal(txtFleteImp.Text) + Convert.ToDecimal(txtFlete.Text))).ToString();
                                if (Convert.ToDecimal(rw["cantidad"]) > 0)
                                {
                                    if (chkFactura.Checked == true)
                                        rw["unitario"] = (Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                    else
                                        rw["unitario"] = (Convert.ToDecimal(rw["costo"]) + Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                    //rw["unitario"] = (Convert.ToDecimal(rw["unitario"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString("###,###,##0.000");
                                }
                            }
                            else
                            {
                                rw["flete"] = ((Convert.ToDecimal(rw["costo"]) / totflete) * (Convert.ToDecimal(txtFleteImp.Text) + Convert.ToDecimal(txtFlete.Text))).ToString();//flete en pesos
                                if (Convert.ToDecimal(rw["cantidad"]) > 0)
                                {
                                    if (chkFactura.Checked == true)
                                        rw["unitario"] = (Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                    else
                                        rw["unitario"] = ((Convert.ToDecimal(rw["costo"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + Convert.ToDecimal(rw["gastoaduanal"]) + Convert.ToDecimal(rw["flete"]) + Convert.ToDecimal(rw["adv"]) + Convert.ToDecimal(rw["dta"])) / Convert.ToDecimal(rw["cantidad"]);
                                    //rw["unitario"] = (Convert.ToDecimal(rw["unitario"]) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString("###,###,##0.000");
                                }
                            }



                        }



                        ////else
                        ////    rw["flete"] = ((Convert.ToDecimal(rw["costo"]) / totflete) * Convert.ToDecimal(txtFlete.Text) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();

                    }
                }

                //Cargar datos al datagrid
                dtgDetalle.Rows.Clear();
                foreach (DataRow rr in tmgrdpro.Rows)
                {
                    if (rr["borrar"].ToString() != "1")
                    {
                        dtgDetalle.Rows.Add(rr["clave"].ToString(), rr["nombre"].ToString(), Convert.ToDecimal(rr["cantidad"].ToString()).ToString("###,###,##0.0000"),
                            Convert.ToDecimal(rr["unitario"].ToString()).ToString("###,###,##0.0000"), Convert.ToDecimal(rr["costo"].ToString()).ToString("###,###,##0.0000"),
                            Convert.ToDecimal(rr["adv"].ToString()).ToString("###,###,##0.0000"), Convert.ToDecimal(rr["dta"].ToString()).ToString("###,###,##0.0000"), rr["lote"].ToString(),
                            Convert.ToBoolean(rr["ieps"].ToString()), Convert.ToDecimal(rr["iepsimporte"].ToString()).ToString("###,###,##0.0000"), Convert.ToDecimal(rr["gastoaduanal"].ToString()).ToString("###,###,##0.0000"),
                            Convert.ToDecimal(rr["flete"].ToString()).ToString("###,###,##0.0000"), rr["oc"].ToString(), rr["req"].ToString(), rr["surtido"].ToString(), rr["cantorig"].ToString(), rr["unitorig"].ToString(),
                            rr["borrar"].ToString(), rr["oper"].ToString(), "", Convert.ToDecimal(rr["iva"].ToString()).ToString("###,###,##0.000"));
                    }
                    //dtgDetalle.Rows.Add(rr.ItemArray);
                }
                calculatotales();
            }
        }

        private void txtFleteImp2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                txtFleteImp.Text = (Convert.ToDecimal(txtFleteImp2.Text) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString();
                txtFleteImp_KeyPress(sender, new KeyPressEventArgs((char)Keys.Enter));
                txtFleteImp2.Text = "0";
            }
        }

        private void chkFactura_CheckStateChanged(object sender, EventArgs e)
        {
            //if (operacion == "consulta")
            //    return;
            if (chkFactura.Checked == true)
            {
                if (operacion != "consulta")
                {
                    if (DDL_ent_procedencia.SelectedIndex < 1)
                    {
                        MessageBox.Show("La entrada debe ser de importación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        chkFactura.Checked = false;
                        return;
                    }
                    if (Convert.ToDecimal(txt_sis_tipo_cambio.Text) <= 1)
                    {
                        MessageBox.Show("Debe ingresar el tipo de cambio", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        chkFactura.Checked = false;
                        return;
                    }
                    if (Convert.ToDecimal(txt_gto_aduanal.Text) <= 0)
                    {
                        MessageBox.Show("Debe ingresar el gasto aduanal", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        chkFactura.Checked = false;
                        return;
                    }
                    if (Convert.ToDecimal(txtFlete.Text) <= 0)
                    {
                        MessageBox.Show("Debe ingresar el flete nacional", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        chkFactura.Checked = false;
                        return;
                    }
                    if (tmgrdpro.Rows.Count == 0)
                    {
                        MessageBox.Show("Debe ingresar un producto al listado", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        chkFactura.Checked = false;
                        return;
                    }
                    else
                    {
                        if (Convert.ToDecimal(dtgDetalle2.Rows[0].Cells[2].Value.ToString()) <= 0)
                        {
                            MessageBox.Show("Debe ingresar la cantidad de producto", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            chkFactura.Checked = false;
                            return;
                        }
                    }
                }

                foreach(DataGridViewRow row in dtgDetalle.Rows)
                {
                    if (row.Cells[0].Value != null)
                    {
                        string str2 = row.Cells[0].Value.ToString();
                        string str3 = row.Cells[10].Value.ToString();
                        string str4 = row.Cells[11].Value.ToString();
                        string str5 = row.Cells[2].Value.ToString();
                        Decimal num = (Convert.ToDecimal(str3) + Convert.ToDecimal(str4)) / Convert.ToDecimal(str5);
                        string str6 = num.ToString("0.000");
                        row.Cells[3].Value = (object)str6;
                        foreach (DataRow dataRow1 in this.tmgrdpro.Select("clave = '" + str2 + "'"))
                        {
                            DataRow dataRow2 = dataRow1;
                            num = Convert.ToDecimal(str3);
                            string str7 = num.ToString();
                            dataRow2["gastoaduanal"] = (object)str7;
                            DataRow dataRow3 = dataRow1;
                            num = Convert.ToDecimal(str4);
                            string str8 = num.ToString();
                            dataRow3["flete"] = (object)str8;
                            dataRow1["unitario"] = (object)str6;
                        }
                        foreach (DataRow dataRow in this.dtdetalles.Select("emp_clave = '" + str2 + "'"))
                            dataRow["costo"] = (object)str6;
                    }
                }

                Decimal num1 = new Decimal(0);
                Decimal num2 = new Decimal(0);
                foreach (DataGridViewRow row in (IEnumerable)this.dtgDetalle.Rows)
                {
                    if (row.Cells[0].Value != null)
                    {
                        num1 += Convert.ToDecimal(row.Cells[10].Value.ToString());
                        num2 += Convert.ToDecimal(row.Cells[11].Value.ToString());
                    }
                }

                decimal SUMA_DET_IVA = 0;
                foreach (DataGridViewRow row in (IEnumerable)this.dtgDetalle2.Rows)
                {
                    SUMA_DET_IVA = SUMA_DET_IVA + Convert.ToDecimal(row.Cells["iva_det"].Value.ToString());
                }

                Decimal num3 = num1 + num2;
                //Decimal num4 = num3 * Convert.ToDecimal("0.16");
                Decimal num5 = num3 + SUMA_DET_IVA;//Convert.ToDecimal(num4.ToString("0.000"));
                this.txt_subtotal1.Text = num3.ToString("###,###,###,##0.0000");
                this.txt_subtotal2.Text = num3.ToString("###,###,###,##0.0000");
                this.txt_iva.Text = SUMA_DET_IVA.ToString("###,###,###,##0.0000");// num4.ToString("###,###,###,##0.0000");
                this.txt_total.Text = num5.ToString("###,###,###,##0.0000");

                //decimal imp = 0;
                //decimal cos = 0;
                //decimal sub = 0;
                //decimal sub2 = 0;
                //decimal ivay = 0;
                //decimal toty = 0;

                //imp = (Convert.ToDecimal(txt_gto_aduanal.Text) + Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text)) / Convert.ToDecimal(txt_sis_tipo_cambio.Text);//IMPORTE EN DOLARES
                //cos = (imp / Convert.ToDecimal(tmgrdpro.Rows[0]["cantidad"].ToString())) * Convert.ToDecimal(txt_sis_tipo_cambio.Text);//COSTO DE EMPAQUE EN PESOS
                //sub = Convert.ToDecimal(txt_gto_aduanal.Text) + Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text);
                //sub2 = Convert.ToDecimal(txt_gto_aduanal.Text) + Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text);
                //ivay = (Convert.ToDecimal(txt_gto_aduanal.Text) + Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text)) * (Convert.ToDecimal(txt_ent_iva.Text) / 100);
                //toty = sub2 + ivay;

                //foreach(DataRow r1 in tmgrdpro.Rows)
                //{
                //    ////tmgrdpro.Rows[0]["costo"] = imp.ToString("0.0000");//IMPORTE
                //    //tmgrdpro.Rows[0]["unitario"] = cos.ToString("0.0000");//UNITARIO
                //    //tmgrdpro.Rows[0]["gastoaduanal"] = txt_gto_aduanal.Text;
                //    //tmgrdpro.Rows[0]["flete"] = (Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text)).ToString();
                //    r1["unitario"] = cos.ToString("0.0000");
                //    r1["gastoaduanal"] = cos.ToString("0.0000");
                //    r1["flete"] = cos.ToString("0.0000");
                //}

                //foreach (DataRow r2 in dtdetalles.Rows)
                //{
                //    ////dtdetalles.Rows[0]["importe"] = imp.ToString("0.0000");//IMPORTE
                //    //dtdetalles.Rows[0]["costo"] = cos.ToString("0.0000");//IMPORTE
                //    r2["costo"] = cos.ToString("0.0000");
                //}

                

                //txt_subtotal1.Text = sub.ToString("###,###,###,##0.0000");
                //txt_subtotal2.Text = sub2.ToString("###,###,###,##0.0000");
                //txt_iva.Text = ivay.ToString("###,###,###,##0.0000");
                //txt_total.Text = toty.ToString("###,###,###,##0.0000");

                //foreach (DataGridViewRow r3 in dtgDetalle2.Rows)
                //{
                //    if (r3.Cells[0].Value == null)
                //        break;
                //    ////dtgDetalle2.Rows[0].Cells[3].Value = cos.ToString("###,###,##0.0000");//UNITARIO
                //    //dtgDetalle2.Rows[0].Cells[5].Value = cos.ToString("###,###,##0.0000");//UNITARIO
                //    ////dtgDetalle2.Rows[0].Cells[10].Value = "0.0000";
                //    ////dtgDetalle2.Rows[0].Cells[11].Value = "0.0000";
                //    r3.Cells[5].Value = cos.ToString("###,###,##0.0000");
                //}

                //foreach(DataGridViewRow r4 in dtgDetalle.Rows)
                //{
                //    if (r4.Cells[0].Value == null)
                //        break;
                //    //dtgDetalle.Rows[0].Cells[3].Value = cos.ToString("###,###,##0.0000");//UNITARIO
                //    ////dtgDetalle.Rows[0].Cells[4].Value = imp.ToString("###,###,##0.0000");//UNITARIO
                //    ////dtgDetalle.Rows[0].Cells[10].Value = "0.0000";
                //    ////dtgDetalle.Rows[0].Cells[11].Value = "0.0000";
                //    //dtgDetalle.Rows[0].Cells[10].Value = txt_gto_aduanal.Text;
                //    //dtgDetalle.Rows[0].Cells[11].Value = (Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text)).ToString("###,###,##0.0000");
                //    r4.Cells[3].Value = cos.ToString("###,###,##0.0000");
                //    r4.Cells[10].Value = txt_gto_aduanal.Text;
                //    r4.Cells[11].Value = (Convert.ToDecimal(txtFlete.Text) + Convert.ToDecimal(txtFleteImp.Text)).ToString("###,###,##0.0000");
                //}
                
            }
            else
            {
                if (this.operacion != "consulta")
                {
                    if (this.DDL_ent_procedencia.SelectedIndex < 1)
                    {
                        int num = (int)MessageBox.Show("La entrada debe ser de importación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.chkFactura.Checked = false;
                        return;
                    }
                    if (Convert.ToDecimal(this.txt_sis_tipo_cambio.Text) <= new Decimal(1))
                    {
                        int num = (int)MessageBox.Show("Debe ingresar el tipo de cambio", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.chkFactura.Checked = false;
                        return;
                    }
                    if (Convert.ToDecimal(this.txt_gto_aduanal.Text) <= new Decimal(0))
                    {
                        int num = (int)MessageBox.Show("Debe ingresar el gasto aduanal", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.chkFactura.Checked = false;
                        return;
                    }
                    if (Convert.ToDecimal(this.txtFlete.Text) <= new Decimal(0))
                    {
                        int num = (int)MessageBox.Show("Debe ingresar el flete nacional", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.chkFactura.Checked = false;
                        return;
                    }
                    if (this.tmgrdpro.Rows.Count == 0)
                    {
                        int num = (int)MessageBox.Show("Debe ingresar un producto al listado", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.chkFactura.Checked = false;
                        return;
                    }
                    if (Convert.ToDecimal(this.dtgDetalle2.Rows[0].Cells[2].Value.ToString()) <= new Decimal(0))
                    {
                        int num = (int)MessageBox.Show("Debe ingresar la cantidad de producto", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.chkFactura.Checked = false;
                        return;
                    }
                }
                foreach (DataGridViewRow row in (IEnumerable)this.dtgDetalle.Rows)
                {
                    if (row.Cells[0].Value != null)
                    {
                        string str2 = row.Cells[0].Value.ToString();
                        string str3 = row.Cells[10].Value.ToString();
                        string str4 = row.Cells[11].Value.ToString();
                        string str5 = row.Cells[2].Value.ToString();
                        string str6 = row.Cells[4].Value.ToString();
                        Decimal num;
                        if (this.DDL_ent_procedencia.SelectedIndex == 1)
                        {
                            num = Convert.ToDecimal(str6) * Convert.ToDecimal(this.txt_sis_tipo_cambio.Text);
                            str6 = num.ToString("0.000");
                        }
                        num = (Convert.ToDecimal(str6) + Convert.ToDecimal(str3) + Convert.ToDecimal(str4)) / Convert.ToDecimal(str5);
                        string str7 = num.ToString("0.000");
                        row.Cells[3].Value = (object)str7;
                        foreach (DataRow dataRow1 in this.tmgrdpro.Select("clave = '" + str2 + "'"))
                        {
                            DataRow dataRow2 = dataRow1;
                            num = Convert.ToDecimal(str3);
                            string str8 = num.ToString();
                            dataRow2["gastoaduanal"] = (object)str8;
                            DataRow dataRow3 = dataRow1;
                            num = Convert.ToDecimal(str4);
                            string str9 = num.ToString();
                            dataRow3["flete"] = (object)str9;
                            dataRow1["unitario"] = (object)str7;
                        }
                        foreach (DataRow dataRow in this.dtdetalles.Select("emp_clave = '" + str2 + "'"))
                            dataRow["costo"] = (object)str7;
                    }
                }
                Decimal num1 = new Decimal(0);
                Decimal num2 = new Decimal(0);
                Decimal num3 = new Decimal(0);

                decimal SUMA_DET_IVA = 0;
                foreach (DataGridViewRow row in (IEnumerable)this.dtgDetalle2.Rows)
                {
                    SUMA_DET_IVA = SUMA_DET_IVA + Convert.ToDecimal(row.Cells["iva_det"].Value.ToString());
                }

                foreach (DataGridViewRow row in (IEnumerable)this.dtgDetalle.Rows)
                {
                    if (row.Cells[0].Value != null)
                    {
                        num1 += Convert.ToDecimal(row.Cells[10].Value.ToString());
                        num2 += Convert.ToDecimal(row.Cells[11].Value.ToString());
                        num3 += Convert.ToDecimal(row.Cells[4].Value.ToString());
                    }
                }
                Decimal num4 = num3 + num1 + num2;
                //Decimal num5 = num4 * Convert.ToDecimal("0.16");
                Decimal num6 = num4 + SUMA_DET_IVA - Convert.ToDecimal(txtDesc.Text);//Convert.ToDecimal(num5.ToString("0.000"));
                this.txt_subtotal1.Text = num4.ToString("###,###,###,##0.0000");
                this.txt_subtotal2.Text = num4.ToString("###,###,###,##0.0000");
                this.txt_iva.Text = SUMA_DET_IVA.ToString("###,###,###,##0.0000"); //num5.ToString("###,###,###,##0.0000");
                this.txt_total.Text = num6.ToString("###,###,###,##0.0000");
            }
        }

        private void N(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnXml_Click(object sender, EventArgs e)
        {
            if (operacion == "consulta" && txtFolioFiscal.Text != "")
            {
                if (lblAguilares.Text == "AGUILARES")
                {
                    MessageBox.Show("La entrada fue registrada desde AGUILARES no puede ser modificada desde GAB, solicitar movimiento a AGUILARES, por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                MessageBox.Show("Para reemplazar la factura favor de hacerlo desde el botón a un lado del numero de la factura en la parte media izquierda", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (lblAguilares.Text == "AGUILARES")
            {
                MessageBox.Show("La entrada fue registrada desde AGUILARES no puede ser modificada desde GAB, solicitar movimiento a AGUILARES, por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (lblPedimento.Text == "0")
            {
                #region factura
                if (DDLAlmacen1.SelectedIndex == 0)
                    openFileDialog1.InitialDirectory = "c:\\MATERIALES";
                if (DDLAlmacen1.SelectedIndex == 1)
                    openFileDialog1.InitialDirectory = "c:\\EMPAQUES";
                if (DDLAlmacen1.SelectedIndex == 2 || DDLAlmacen1.SelectedIndex == 3)
                    openFileDialog1.InitialDirectory = "c:\\COMEDOR";
                if (DDLAlmacen1.SelectedIndex == 4)
                    openFileDialog1.InitialDirectory = "c:\\ENFERMERIA";
                if (DDLAlmacen1.SelectedIndex == 5)
                    openFileDialog1.InitialDirectory = "c:\\PAPELERIA";
                if (DDLAlmacen1.SelectedIndex == 6)
                    openFileDialog1.InitialDirectory = "c:\\DENTAL";
                if (DDLAlmacen1.SelectedIndex == 7)
                    openFileDialog1.InitialDirectory = "c:\\COMBUSTIBLE";
                if (DDLAlmacen1.SelectedIndex == 8)
                    openFileDialog1.InitialDirectory = "c:\\SISTEMAS";
                if (DDLAlmacen1.SelectedIndex == 9)
                    openFileDialog1.InitialDirectory = "c:\\GUARDERIA";
                if (DDLAlmacen1.SelectedIndex == 10)
                    openFileDialog1.InitialDirectory = "c:\\GUARDERIACOM";
                if (DDLAlmacen1.SelectedIndex == 11)
                    openFileDialog1.InitialDirectory = "c:\\EVENTOS";
                if (DDLAlmacen1.SelectedIndex == 12)
                    openFileDialog1.InitialDirectory = "c:\\MOTORES";
                openFileDialog1.Filter = "xml files (*.xml)|*.xml";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;
                string FOL_FIS = "";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    lblPathXml.Text = openFileDialog1.FileName;

                    string pdf = lblPathXml.Text.Replace(".xml", ".pdf");

                    FileInfo file_copy = new FileInfo(@pdf);
                    if (file_copy.Exists == false)
                    {
                        MessageBox.Show("El PDF del archivo XML no fue localizado, el archivo debe tener el mismo nombre y estar en la misma carpeta", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        valida_carga_xml = false;
                        return;
                    }

                    XmlDocument xDoc = new XmlDocument();

                    try
                    {
                        xDoc.Load(openFileDialog1.FileName);
                    }
                    catch (XmlException ex)
                    {
                        MessageBox.Show("Se detectó un error en el archivo, se enviará a el departamento de sistemas para resolverlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        MailMessage message = new MailMessage();
                        SmtpClient smtpClient = new SmtpClient();
                        Attachment attachment = new Attachment(this.openFileDialog1.FileName);
                        message.From = new MailAddress("aescamilla@mrlucky.com.mx");
                        message.To.Add(new MailAddress("aescamilla@mrlucky.com.mx"));
                        message.Subject = "Error en archivo XML";
                        message.SubjectEncoding = Encoding.UTF8;
                        string str2 = "<table border='1'><tr><td align='center'><b>XML</b></td></tr><tr><td>Archivo dañado</td></tr></table>";
                        message.Body = str2;
                        message.BodyEncoding = Encoding.UTF8;
                        message.IsBodyHtml = true;
                        message.Attachments.Add(attachment);
                        message.From = new MailAddress("aescamilla@mrlucky.com.mx");
                        smtpClient.Credentials = (ICredentialsByHost)new NetworkCredential("aescamilla", "atrejo");
                        smtpClient.Port = 587;
                        smtpClient.Host = "mail1.mrlucky.com.mx";
                        smtpClient.EnableSsl = true;
                        smtpClient.Send(message);
                        return;
                    }


                    XmlNodeList lista = xDoc.GetElementsByTagName("cfdi:Comprobante");

                    string Folio = "";
                    string Serie = "";
                    string Fecha = "";
                    string RFC = "";
                    string nom_prov = "";
                    string Total = "";

                    string UsoCFDI = "";

                    foreach (XmlElement nodo in lista)
                    {
                        Folio = nodo.GetAttribute("Folio");
                        if (Folio == "")
                            Folio = nodo.GetAttribute("folio");
                        Serie = nodo.GetAttribute("Serie");
                        if (Serie == "")
                            Serie = nodo.GetAttribute("serie");
                        Fecha = nodo.GetAttribute("Fecha");
                        if (Fecha == "")
                            Fecha = nodo.GetAttribute("fecha");
                        Total = nodo.GetAttribute("Total");
                        if (Total == "")
                            Total = nodo.GetAttribute("total");
                    }

                    XmlNodeList lista2 = ((XmlElement)lista[0]).GetElementsByTagName("cfdi:Complemento");
                    XmlNodeList lista3 = ((XmlElement)lista2[0]).GetElementsByTagName("tfd:TimbreFiscalDigital");
                    XmlNodeList lista4 = ((XmlElement)lista[0]).GetElementsByTagName("cfdi:Receptor");

                    foreach (XmlElement nodo in lista3)
                    {
                        FOL_FIS = nodo.GetAttribute("UUID");
                    }

                    foreach (XmlElement nodo in lista4)
                    {
                        UsoCFDI = nodo.GetAttribute("UsoCFDI");
                    }

                    if (UsoCFDI.Contains("CP") || UsoCFDI.Contains("CN"))
                    {
                        MessageBox.Show("El tipo de documento que intenta subir no es valido. Este es un Comprobante de Pago (REP). Seleccione una factura por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }



                    ////23/03/2021
                    ////validacion de uuid de xml contra uuid de PDF
                    //string pathito = "";//System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
                    //pathito = @pdf;
                    //bool validacion_xml_pdf = ReadPdfFile(pathito, FOL_FIS);

                    //if (validacion_xml_pdf == false)
                    //{
                    //    //Se cierra carga de xml por discordancia de UUID
                    //    txtFechaFact.Text = "";
                    //    txtFolioFiscal.Text = "";
                    //    lblRFC.Text = "-";
                    //    lblPathXml.Text = "-";
                    //    valida_carga_xml = false;
                    //    return;
                    //}
                    ////validacion de uuid contra uuid de PDF
                    ////23/03/2021

                    //if (txtFolioFiscal.Text != "")
                    //{
                    //    if (txtFolioFiscal.Text.ToUpper() != FOL_FIS.ToUpper())
                    //    {
                    //        MessageBox.Show("El folio fiscal de la factura seleccionada no corresponde al generado por medio de la insercion del acuse, favor de seleccionar la factura correcta", "SISMEP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //        valida_carga_xml = false;
                    //        return;
                    //    }
                    //}

                    //Validar que la factura no haya sido ingresada anteriormente
                    string noc = "";
                    string fct = "";
                    string fef = "";
                    string fer = "";
                    string tpo = "";
                    if (FOL_FIS != "")
                    {
                        thisConnection.Open();
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "SELECT numero_oc, factura, fecha_factura, fecha_registro, tipo FROM tb_det_oc_facturas WHERE folio_fiscal = '" + FOL_FIS + "' AND estatus <> 'C'";
                        reader1 = cmnd1.ExecuteReader();
                        bool fnd = false;

                        if (reader1.HasRows)
                        {
                            reader1.Read();
                            noc = reader1["numero_oc"].ToString().Trim();
                            fct = reader1["factura"].ToString().Trim();
                            fef = Convert.ToDateTime(reader1["fecha_factura"].ToString().Trim()).ToShortDateString();
                            fer = Convert.ToDateTime(reader1["fecha_registro"].ToString().Trim()).ToShortDateString();
                            tpo = reader1["tipo"].ToString().Trim();
                            fnd = true;
                        }
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();
                        thisConnection.Close();



                        if (fnd == true)
                        {
                            //if (txt_prov_clave.Text.Trim() == "404" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "833" || txt_prov_clave.Text.Trim() == "1550")
                            string importacion = "0";
                            thisConnection.Open();
                            SqlCommand cmndimportacion = thisConnection.CreateCommand();
                            cmndimportacion.CommandText = "SELECT prov_imp FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text.Trim() + "'";
                            try
                            {
                                importacion = Convert.ToString(cmndimportacion.ExecuteScalar().ToString().Trim());
                            }
                            catch
                            {
                                importacion = "";
                            }

                            cmndimportacion.Dispose();
                            thisConnection.Close();

                            if (importacion == "1")
                            {

                                /*if (txt_prov_clave.Text == "81" || txt_prov_clave.Text == "404" || txt_prov_clave.Text == "11" || txt_prov_clave.Text == "8851" || txt_prov_clave.Text.Trim() == "318" || txt_prov_clave.Text.Trim() == "1269" || txt_prov_clave.Text.Trim() == "694" || txt_prov_clave.Text.Trim() == "9957" || txt_prov_clave.Text.Trim() == "833" || txt_prov_clave.Text.Trim() == "1550" || txt_prov_clave.Text.Trim() == "1761" || txt_prov_clave.Text.Trim() == "935" || txt_prov_clave.Text.Trim() == "6854" || txt_prov_clave.Text == "354" || txt_prov_clave.Text == "1250" || txt_prov_clave.Text == "1536" || txt_prov_clave.Text == "9758" || txt_prov_clave.Text == "1799" || txt_prov_clave.Text == "1244" || txt_prov_clave.Text == "191" || txt_prov_clave.Text == "812" || txt_prov_clave.Text == "8014" || txt_prov_clave.Text == "305" || txt_prov_clave.Text == "106")
                                {*/
                                MessageBox.Show("*****PROVEEDOR VALIDO PARA GUARDAR MISMA FACTURA*****\nEl folio fiscal ya fue registrado.\nEntrada: " + noc + "\nFactura: " + fct + "\nFecha Factura: " + fef + "\nFecha Registro: " + fer + "\nAlmacen: " + tpo, "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                valida_carga_xml = true;
                            }
                            else
                            {
                                MessageBox.Show("El folio fiscal ya fue registrado.\nEntrada: " + noc + "\nFactura: " + fct + "\nFecha Factura: " + fef + "\nFecha Registro: " + fer + "\nAlmacen: " + tpo, "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                valida_carga_xml = false;
                                return;
                            }

                        }
                        //else {
                        //    if (txt_prov_clave.Text.Trim() != "404" && txt_prov_clave.Text.Trim() != "1269" && txt_prov_clave.Text.Trim() != "694")
                        //    {
                        //        Utilerias.Class1.registrar_movimiento3(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "MF", "2.5", txtFolio.Text.Trim(), "FACTURA ENCONTRADA", "SISEMP", "", "FACTURA REGISTRADA Entrada: " + noc + " Factura: " + fct + " Fecha Factura: " + fef + " Fecha Registro: " + fer + " Almacen: " + tpo);
                        //        //valida_carga_xml = false;
                        //        //return;
                        //    }
                        //}
                    }

                    txtFolioFiscal.Text = FOL_FIS;
                    txt_ent_factura.Text = (Serie.Trim() + Folio.Trim()).Replace("/", "");
                    txtFechaFact.Text = Convert.ToDateTime(Fecha).ToShortDateString();
                    lblTotal.Text = Total;

                    XmlNodeList lista6 = ((XmlElement)lista[0]).GetElementsByTagName("cfdi:Emisor");
                    foreach (XmlElement nodo in lista6)
                    {
                        string rfc2 = nodo.GetAttribute("Rfc");
                        if (rfc2 == "")
                            rfc2 = nodo.GetAttribute("rfc");
                        RFC = rfc2;
                        nom_prov = nodo.GetAttribute("Nombre");
                        if (nom_prov == "")
                            nom_prov = nodo.GetAttribute("nombre");
                    }
                    lblRFC.Text = RFC;
                    string shortDateString;
                    if ((this.lblToc.Text == "D" || this.lblToc.Text == "E") && this.DDL_ent_procedencia.SelectedIndex == 0)
                    {
                        string str2 = Convert.ToDateTime(this.txtFechaFact.Text).Day.ToString();
                        string str9 = Convert.ToDateTime(this.txtFechaFact.Text).Month.ToString();
                        string str10 = Convert.ToDateTime(this.txtFechaFact.Text).Year.ToString();
                        this.thisConnection.Open();
                        this.cmnd1 = this.thisConnection.CreateCommand();
                        this.cmnd1.CommandText = "SELECT valor FROM tb_cat_tipocambio WHERE dia = '" + str2 + "' AND mes = '" + str9 + "' AND año = '" + str10 + "'";
                        this.reader1 = this.cmnd1.ExecuteReader();
                        bool flag = false;
                        if (this.reader1.HasRows)
                        {
                            this.reader1.Read();
                            this.lblTCFact.Text = this.reader1["valor"].ToString().Trim();
                            flag = true;
                        }
                        this.reader1.Close();
                        this.reader1.Dispose();
                        this.cmnd1.Dispose();
                        if (!flag)
                        {
                            DateTime dateTime = Convert.ToDateTime(this.txtFechaFact.Text).AddDays(-1.0);
                            shortDateString = dateTime.ToShortDateString();
                            dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
                            string str11 = dateTime.Day.ToString();
                            dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
                            string str12 = dateTime.Month.ToString();
                            dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
                            string str13 = dateTime.Year.ToString();
                            this.thisConnection.Open();
                            this.cmnd1 = this.thisConnection.CreateCommand();
                            this.cmnd1.CommandText = "SELECT valor FROM tb_cat_tipocambio WHERE dia = '" + str11 + "' AND mes = '" + str12 + "' AND año = '" + str13 + "'";
                            this.reader1 = this.cmnd1.ExecuteReader();
                            if (this.reader1.HasRows)
                            {
                                this.reader1.Read();
                                this.lblTCFact.Text = this.reader1["valor"].ToString().Trim();
                            }
                            this.reader1.Close();
                            this.reader1.Dispose();
                            this.cmnd1.Dispose();
                        }
                        this.thisConnection.Close();
                    }
                    if (this.DDL_ent_procedencia.SelectedIndex == 1)
                    {
                        string str2 = Convert.ToDateTime(this.txtFechaFact.Text).Day.ToString();
                        string str9 = Convert.ToDateTime(this.txtFechaFact.Text).Month.ToString();
                        string str10 = Convert.ToDateTime(this.txtFechaFact.Text).Year.ToString();
                        this.thisConnection.Open();
                        this.cmnd1 = this.thisConnection.CreateCommand();
                        this.cmnd1.CommandText = "SELECT valor FROM tb_cat_tipocambio WHERE dia = '" + str2 + "' AND mes = '" + str9 + "' AND año = '" + str10 + "'";
                        this.reader1 = this.cmnd1.ExecuteReader();
                        bool flag = false;
                        if (this.reader1.HasRows)
                        {
                            this.reader1.Read();
                            this.txt_sis_tipo_cambio.Text = this.reader1["valor"].ToString().Trim();
                            flag = true;
                        }
                        this.reader1.Close();
                        this.reader1.Dispose();
                        this.cmnd1.Dispose();
                        if (!flag)
                        {
                            DateTime dateTime = Convert.ToDateTime(this.txtFechaFact.Text).AddDays(-1.0);
                            shortDateString = dateTime.ToShortDateString();
                            dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
                            string str11 = dateTime.Day.ToString();
                            dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
                            string str12 = dateTime.Month.ToString();
                            dateTime = Convert.ToDateTime(this.txtFechaFact.Text);
                            string str13 = dateTime.Year.ToString();
                            //this.thisConnection.Open();
                            this.cmnd1 = this.thisConnection.CreateCommand();
                            this.cmnd1.CommandText = "SELECT valor FROM tb_cat_tipocambio WHERE dia = '" + str11 + "' AND mes = '" + str12 + "' AND año = '" + str13 + "'";
                            this.reader1 = this.cmnd1.ExecuteReader();
                            if (this.reader1.HasRows)
                            {
                                this.reader1.Read();
                                this.txt_sis_tipo_cambio.Text = this.reader1["valor"].ToString().Trim();
                            }
                            this.reader1.Close();
                            this.reader1.Dispose();
                            this.cmnd1.Dispose();
                        }
                        this.thisConnection.Close();
                    }

                    //VERIFICACION DE RFC DE PROVEEDOR CON EL DE LA FACTURA
                    string ProvRfc = "";
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT prov_rfc FROM tb_cat_proveedor WHERE prov_clave = '" + txt_prov_clave.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        ProvRfc = reader1["prov_rfc"].ToString().Trim();
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    if (ProvRfc != lblRFC.Text)
                    {
                        if (MessageBox.Show("El XML corresponde al proveedor:\n" + nom_prov + "\n\n¿Desea continuar?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == System.Windows.Forms.DialogResult.No)
                        {
                            //MessageBox.Show("LOS RFC NO CORRESPONDEN, FAVOR DE VALIDAR LOS RFC Y CORREGIRLO EN SISTEMA DE SER NECESARIO.\n RFC - XML:" + ProvRfc + "\n RFC - SISTEMA:" + lblRFC.Text + "\n", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            txtFechaFact.Text = "";
                            txtFolioFiscal.Text = "";
                            lblRFC.Text = "-";
                            return;
                        }
                        else
                        {
                            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + "Proveedor XML: " + nom_prov + ", RFC XML: " + lblRFC.Text + " y RFC BD: " + ProvRfc);
                        }
                    }

                    bool ping_con = ping_portal();
                    if (ping_con == true)
                    {
                        conn.Open();
                        comm = conn.CreateCommand();

                        string cfd_timfis = txtFolioFiscal.Text;
                        comm.CommandText = "SELECT cfd_keyusu, cfd_keyupl, cfd_subtot, cfd_imptra, cfd_totalv FROM CF_REGISTRO_WEB where cfd_timfis = '" + cfd_timfis + "'";
                        redr = comm.ExecuteReader();
                        if (redr.HasRows)
                        {
                            redr.Read();
                            lblPortal.Text = redr["cfd_keyusu"].ToString().Trim() + "-" + redr["cfd_keyupl"].ToString().Trim();
                        }
                        redr.Close();
                        redr.Dispose();
                        comm.Dispose();

                        conn.Close();

                    }
                    valida_carga_xml = false;

                }
                else
                    valida_carga_xml = true;
                #endregion
            }
            else
            {
                panel4.Visible = true;
            }

        }

        //public bool ReadPdfFile(object Filename, string uuid)
        //{
        //    PdfReader reader2 = new PdfReader((string)Filename);
        //    string strText = string.Empty;

        //    for (int page = 1; page <= reader2.NumberOfPages; page++)
        //    {
        //        ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
        //        PdfReader reader = new PdfReader((string)Filename);
        //        String s = PdfTextExtractor.GetTextFromPage(reader, page, its);

        //        s = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(s)));
        //        strText = strText + s;
        //        reader.Close();
        //    }

        //    if (strText.Contains(uuid))
        //        return true;
        //    else
        //        return false;
        //}

        public bool ping_portal()
        {
            Ping pings = new Ping();
            int timeout = 10;
            if (pings.Send("grupou-mrlucky.com.mx", timeout).Status == IPStatus.Success)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void chkDescuento_CheckedChanged(object sender, EventArgs e)
        {
            

            //if (this.chkDescuento.Checked)
            //{
            //    foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
            //    {
            //        Decimal num = Convert.ToDecimal(row["importe_org"]) * (Convert.ToDecimal(this.lblPorcentajeDesc.Text) / new Decimal(100));
            //        row["importe"] = (object)(Convert.ToDecimal(row["importe_org"]) - num).ToString();
            //        row["costo"] = (object)(Convert.ToDecimal(row["importe"]) / Convert.ToDecimal(row["cantidad"])).ToString();
            //    }
            //    this.button2_Click((object)null, (EventArgs)null);
            //    this.dtgDetalle2.Rows.Clear();
            //    foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
            //    {
            //        DataGridViewRowCollection rows = this.dtgDetalle2.Rows;

            //        object[] objArray1 = new object[12];
            //        objArray1[0] = (object)row["emp_clave"].ToString();
            //        objArray1[1] = (object)row["orden_compra"].ToString();
            //        object[] objArray2 = objArray1;
            //        Decimal num = Convert.ToDecimal(row["cantidad"].ToString());
            //        string str1 = num.ToString("###,###,##0.000");
            //        objArray2[2] = (object)str1;
            //        object[] objArray3 = objArray1;
            //        num = Convert.ToDecimal(row["importe"].ToString());
            //        string str2 = num.ToString("###,###,##0.000");
            //        objArray3[3] = (object)str2;
            //        objArray1[4] = (object)row["requisicion"].ToString();
            //        object[] objArray4 = objArray1;
            //        num = Convert.ToDecimal(row["costo"].ToString());
            //        string str3 = num.ToString("###,###,##0.00000");
            //        objArray4[5] = (object)str3;
            //        objArray1[6] = (object)row["borrar"].ToString();
            //        objArray1[7] = (object)row["operacion"].ToString();
            //        objArray1[8] = (object)"";
            //        object[] objArray5 = objArray1;
            //        num = Convert.ToDecimal(row["ieps"].ToString());
            //        string str4 = num.ToString("###,###,##0.000");
            //        objArray5[9] = (object)str4;
            //        object[] objArray6 = objArray1;
                    

            //        rows.Add(objArray6);
            //    }
            //}
            //if (this.chkDescuento.Checked)
            //    return;
            //foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
            //{
            //    row["importe"] = row["importe_org"];
            //    row["costo"] = row["costo_org"];
            //}
            //this.button2_Click((object)null, (EventArgs)null);
            //this.dtgDetalle2.Rows.Clear();
            //foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
            //{
            //    DataGridViewRowCollection rows = this.dtgDetalle2.Rows;
            //    object[] objArray1 = new object[10];
            //    objArray1[0] = (object)row["emp_clave"].ToString();
            //    objArray1[1] = (object)row["orden_compra"].ToString();
            //    object[] objArray2 = objArray1;
            //    Decimal num = Convert.ToDecimal(row["cantidad"].ToString());
            //    string str1 = num.ToString("###,###,##0.000");
            //    objArray2[2] = (object)str1;
            //    object[] objArray3 = objArray1;
            //    num = Convert.ToDecimal(row["importe"].ToString());
            //    string str2 = num.ToString("###,###,##0.000");
            //    objArray3[3] = (object)str2;
            //    objArray1[4] = (object)row["requisicion"].ToString();
            //    object[] objArray4 = objArray1;
            //    num = Convert.ToDecimal(row["costo"].ToString());
            //    string str3 = num.ToString("###,###,##0.00000");
            //    objArray4[5] = (object)str3;
            //    objArray1[6] = (object)row["borrar"].ToString();
            //    objArray1[7] = (object)row["operacion"].ToString();
            //    objArray1[8] = (object)"";
            //    object[] objArray5 = objArray1;
            //    num = Convert.ToDecimal(row["ieps"].ToString());
            //    string str4 = num.ToString("###,###,##0.000");
            //    objArray5[9] = (object)str4;
            //    object[] objArray6 = objArray1;
            //    rows.Add(objArray6);
            //}
        }

        public void aplica_descuento_consulta()
        {
            if (Convert.ToDecimal(lblPorcentajeDesc.Text) > 0)
            {
                foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
                {
                    Decimal num = Convert.ToDecimal(row["importe_org"]) * (Convert.ToDecimal(this.lblPorcentajeDesc.Text) / new Decimal(100));
                    row["importe"] = (object)(Convert.ToDecimal(row["importe_org"]) - num).ToString();
                    row["costo"] = (object)(Convert.ToDecimal(row["importe"]) / Convert.ToDecimal(row["cantidad"])).ToString();
                }
                this.button2_Click((object)null, (EventArgs)null);
                this.dtgDetalle2.Rows.Clear();
                foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
                {
                    DataGridViewRowCollection rows = this.dtgDetalle2.Rows;
                    object[] objArray1 = new object[10];
                    objArray1[0] = (object)row["emp_clave"].ToString();
                    objArray1[1] = (object)row["orden_compra"].ToString();
                    object[] objArray2 = objArray1;
                    Decimal num = Convert.ToDecimal(row["cantidad"].ToString());
                    string str1 = num.ToString("###,###,##0.000");
                    objArray2[2] = (object)str1;
                    object[] objArray3 = objArray1;
                    num = Convert.ToDecimal(row["importe"].ToString());
                    string str2 = num.ToString("###,###,##0.000");
                    objArray3[3] = (object)str2;
                    objArray1[4] = (object)row["requisicion"].ToString();
                    object[] objArray4 = objArray1;
                    num = Convert.ToDecimal(row["costo"].ToString());
                    string str3 = num.ToString("###,###,##0.00000");
                    objArray4[5] = (object)str3;
                    objArray1[6] = (object)row["borrar"].ToString();
                    objArray1[7] = (object)row["operacion"].ToString();
                    objArray1[8] = (object)"";
                    object[] objArray5 = objArray1;
                    num = Convert.ToDecimal(row["ieps"].ToString());
                    string str4 = num.ToString("###,###,##0.000");
                    objArray5[9] = (object)str4;
                    object[] objArray6 = objArray1;
                    rows.Add(objArray6);
                }
            }
            else
            { 
            }
            //if (this.chkDescuento.Checked)
            //{
            //    foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
            //    {
            //        Decimal num = Convert.ToDecimal(row["importe_org"]) * (Convert.ToDecimal(this.lblPorcentajeDesc.Text) / new Decimal(100));
            //        row["importe"] = (object)(Convert.ToDecimal(row["importe_org"]) - num).ToString();
            //        row["costo"] = (object)(Convert.ToDecimal(row["importe"]) / Convert.ToDecimal(row["cantidad"])).ToString();
            //    }
            //    this.button2_Click((object)null, (EventArgs)null);
            //    this.dtgDetalle2.Rows.Clear();
            //    foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
            //    {
            //        DataGridViewRowCollection rows = this.dtgDetalle2.Rows;
            //        object[] objArray1 = new object[10];
            //        objArray1[0] = (object)row["emp_clave"].ToString();
            //        objArray1[1] = (object)row["orden_compra"].ToString();
            //        object[] objArray2 = objArray1;
            //        Decimal num = Convert.ToDecimal(row["cantidad"].ToString());
            //        string str1 = num.ToString("###,###,##0.000");
            //        objArray2[2] = (object)str1;
            //        object[] objArray3 = objArray1;
            //        num = Convert.ToDecimal(row["importe"].ToString());
            //        string str2 = num.ToString("###,###,##0.000");
            //        objArray3[3] = (object)str2;
            //        objArray1[4] = (object)row["requisicion"].ToString();
            //        object[] objArray4 = objArray1;
            //        num = Convert.ToDecimal(row["costo"].ToString());
            //        string str3 = num.ToString("###,###,##0.00000");
            //        objArray4[5] = (object)str3;
            //        objArray1[6] = (object)row["borrar"].ToString();
            //        objArray1[7] = (object)row["operacion"].ToString();
            //        objArray1[8] = (object)"";
            //        object[] objArray5 = objArray1;
            //        num = Convert.ToDecimal(row["ieps"].ToString());
            //        string str4 = num.ToString("###,###,##0.000");
            //        objArray5[9] = (object)str4;
            //        object[] objArray6 = objArray1;
            //        rows.Add(objArray6);
            //    }
            //}
            //if (this.chkDescuento.Checked)
            //    return;
            //foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
            //{
            //    row["importe"] = row["importe_org"];
            //    row["costo"] = row["costo_org"];
            //}
            //this.button2_Click((object)null, (EventArgs)null);
            //this.dtgDetalle2.Rows.Clear();
            //foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
            //{
            //    DataGridViewRowCollection rows = this.dtgDetalle2.Rows;
            //    object[] objArray1 = new object[10];
            //    objArray1[0] = (object)row["emp_clave"].ToString();
            //    objArray1[1] = (object)row["orden_compra"].ToString();
            //    object[] objArray2 = objArray1;
            //    Decimal num = Convert.ToDecimal(row["cantidad"].ToString());
            //    string str1 = num.ToString("###,###,##0.000");
            //    objArray2[2] = (object)str1;
            //    object[] objArray3 = objArray1;
            //    num = Convert.ToDecimal(row["importe"].ToString());
            //    string str2 = num.ToString("###,###,##0.000");
            //    objArray3[3] = (object)str2;
            //    objArray1[4] = (object)row["requisicion"].ToString();
            //    object[] objArray4 = objArray1;
            //    num = Convert.ToDecimal(row["costo"].ToString());
            //    string str3 = num.ToString("###,###,##0.00000");
            //    objArray4[5] = (object)str3;
            //    objArray1[6] = (object)row["borrar"].ToString();
            //    objArray1[7] = (object)row["operacion"].ToString();
            //    objArray1[8] = (object)"";
            //    object[] objArray5 = objArray1;
            //    num = Convert.ToDecimal(row["ieps"].ToString());
            //    string str4 = num.ToString("###,###,##0.000");
            //    objArray5[9] = (object)str4;
            //    object[] objArray6 = objArray1;
            //    rows.Add(objArray6);
            //}
        }

        private void printDocument2_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                this.renglonant = 0;
                this.thisConnection.Open();
                e.Graphics.PageUnit = GraphicsUnit.Millimeter;
                Graphics graphics1 = e.Graphics;
                
                SolidBrush solidBrush1 = new SolidBrush(Color.Black);
                Font font1 = new Font("Courier New", 10f);
                
                //SizeF sz = new System.Drawing.SizeF(((float)216), ((float)279));

                using (Bitmap img = new Bitmap((Image)Resources.NOTA_DE_ENTRADA_011))
                {
                    graphics1.DrawImage(img, 0, 0, ((float)216), ((float)279));
                }

                try 
                {
                    
                    
                }
                catch(Exception ex)
                {
                    //if (ex.ToString().Contains("Error genérico en GDI+."))
                    //{
                    //    using (Bitmap img = new Bitmap((Image)Resources.NOTA_DE_ENTRADA_011))
                    //    {
                    //        graphics1.DrawImage(img, 0, 0, ((float)216), ((float)279));
                    //    }
                    //}
                }
                
                
                
                //graphics1.DrawImage(img, 0, 0);
                //Rectangle r1 = new Rectangle(0, 0, );
                //TextureBrush tb = new TextureBrush(img);
                //graphics1.FillRectangle(img, );
                //Bitmap img = new Bitmap((Image)Resources.NOTA_DE_ENTRADA_011, 216, 279);
                ////Graphics graphicsImg = Graphics.FromImage(img);
                ////graphics1.DrawImage(img, 0, 0, ((float)216), ((float)279));
                //graphics1.DrawImage(img, 0, 0);
                //img.Dispose();
                
                graphics1.DrawString(this.txtFolio.Text, font1, (Brush)solidBrush1, 180f, 6f);
                graphics1.DrawString("NORMAL", font1, (Brush)solidBrush1, 180f, 9f);
                Font font2 = new Font("Courier New", 7f);
                graphics1.DrawString(this.dtpFecha.Text, font1, (Brush)solidBrush1, 140f, 16f);
                if (this.DDLProveedor.SelectedItem.ToString().Length > 40)
                    graphics1.DrawString(this.DDLProveedor.SelectedItem.ToString().Substring(0, 39).ToString(), font2, (Brush)solidBrush1, 8f, 26f);
                else
                    graphics1.DrawString(this.DDLProveedor.SelectedItem.ToString(), font2, (Brush)solidBrush1, 8f, 26f);
                graphics1.DrawString(this.txt_ent_factura.Text, font2, (Brush)solidBrush1, 115f, 26f);
                graphics1.DrawString(this.lstOC.Items[0].ToString(), font2, (Brush)solidBrush1, 180f, 26f);
                int num1 = 35;
                graphics1.DrawString("Lote Rastreabilidad", font2, (Brush)solidBrush1, 125f, (float)num1);
                string s1 = this.DDLElaboro.SelectedItem.ToString().Length <= 25 ? this.DDLElaboro.SelectedItem.ToString() : this.DDLElaboro.SelectedItem.ToString().Substring(0, 24);
                string s2 = this.DDLAutorizo.SelectedItem.ToString().Length <= 25 ? this.DDLAutorizo.SelectedItem.ToString() : this.DDLAutorizo.SelectedItem.ToString().Substring(0, 24);
                graphics1.DrawString(s1, font2, (Brush)solidBrush1, 58f, 270f);
                graphics1.DrawString(s2, font2, (Brush)solidBrush1, 8f, 270f);
                graphics1.DrawString("F-500-11", font2, (Brush)solidBrush1, 177f, 267f);
                graphics1.DrawString("REV: 00", font2, (Brush)solidBrush1, 177f, 270f);
                graphics1.DrawString(this.lblFechaEnt.Text, font2, (Brush)solidBrush1, 140f, 275f);
                int num2 = num1 + 5;
                string str1 = "";
                string str2;
                string str3;
                string str4;
                string str5;
                if (this.renglonant == 0)
                {
                    Decimal num3;
                    for (int index = 0; index < this.tmgrdpro.Rows.Count; ++index)
                    {
                        if (!(this.tmgrdpro.Rows[index]["borrar"].ToString() == "1") && (Convert.ToDecimal(this.tmgrdpro.Rows[index]["cantidad"].ToString()) > new Decimal(0) && Convert.ToDecimal(this.tmgrdpro.Rows[index]["unitario"].ToString()) > new Decimal(0)))
                        {
                            Graphics graphics2 = graphics1;
                            num3 = Convert.ToDecimal(this.tmgrdpro.Rows[index]["cantidad"].ToString());
                            string s3 = num3.ToString("###,###,##0.00");
                            Font font3 = font2;
                            SolidBrush solidBrush2 = solidBrush1;
                            num3 = Convert.ToDecimal(this.tmgrdpro.Rows[index]["cantidad"].ToString());
                            double num4 = (double)this.posicion_pdf(num3.ToString("###,###,##0.00"));
                            double num5 = (double)num2;
                            graphics2.DrawString(s3, font3, (Brush)solidBrush2, (float)num4, (float)num5);
                            this.cmnd2 = this.thisConnection.CreateCommand();
                            this.cmnd2.CommandText = "SELECT um_clave FROM tb_cat_empaques WHERE emp_clave = '" + this.tmgrdpro.Rows[index]["clave"].ToString().Trim() + "'";
                            this.reader2 = this.cmnd2.ExecuteReader();
                            if (this.reader2.HasRows)
                            {
                                this.reader2.Read();
                                str1 = this.reader2.GetValue(0).ToString();
                            }
                            this.reader2.Close();
                            graphics1.DrawString(str1.Substring(0, 3).ToString(), font2, (Brush)solidBrush1, 38f, (float)num2);
                            graphics1.DrawString(this.tmgrdpro.Rows[index]["clave"].ToString() + " " + this.tmgrdpro.Rows[index]["nombre"].ToString(), font2, (Brush)solidBrush1, 50f, (float)num2);
                            graphics1.DrawString(this.tmgrdpro.Rows[index]["lote"].ToString(), font2, (Brush)solidBrush1, 130f, (float)num2);
                            Graphics graphics3 = graphics1;
                            num3 = Convert.ToDecimal(this.tmgrdpro.Rows[index]["unitario"].ToString());
                            string s4 = num3.ToString("###,###,##0.00");
                            Font font4 = font2;
                            SolidBrush solidBrush3 = solidBrush1;
                            num3 = Convert.ToDecimal(this.tmgrdpro.Rows[index]["unitario"].ToString());
                            double num6 = (double)this.posicionprecio_pdf(num3.ToString("###,###,##0.00"));
                            double num7 = (double)num2;
                            graphics3.DrawString(s4, font4, (Brush)solidBrush3, (float)num6, (float)num7);
                            Decimal num8;
                            if (this.tmgrdpro.Rows[index]["costo"].ToString() != "" && Convert.ToDecimal(this.tmgrdpro.Rows[index]["costo"].ToString()) > new Decimal(0))
                            {
                                if (this.DDL_ent_procedencia.SelectedIndex == 0)
                                {
                                    num8 = Convert.ToDecimal(this.tmgrdpro.Rows[index]["costo"].ToString());
                                    Decimal num9 = Convert.ToDecimal(this.tmgrdpro.Rows[index]["gastoaduanal"].ToString());
                                    Decimal num10 = Convert.ToDecimal(this.tmgrdpro.Rows[index]["flete"].ToString());
                                    num8 = num8 + num9 + num10;
                                    graphics1.DrawString(num8.ToString("###,###,##0.00"), font2, (Brush)solidBrush1, this.posicionimporte_pdf(num8.ToString("###,###,##0.00")), (float)num2);
                                }
                                else
                                {
                                    num8 = Convert.ToDecimal(this.tmgrdpro.Rows[index]["costo"].ToString()) * Convert.ToDecimal(this.txt_sis_tipo_cambio.Text);
                                    Decimal num9 = Convert.ToDecimal(this.tmgrdpro.Rows[index]["gastoaduanal"].ToString());
                                    Decimal num10 = Convert.ToDecimal(this.tmgrdpro.Rows[index]["flete"].ToString());
                                    num8 = num8 + num9 + num10;
                                    graphics1.DrawString(num8.ToString("###,###,##0.00"), font2, (Brush)solidBrush1, this.posicionimporte_pdf(num8.ToString("###,###,##0.00")), (float)num2);
                                }
                            }
                            num2 += 3;
                            ++this.renglonant;
                            ++this.ren;
                            if (this.ren == 60)
                            {
                                this.ren = 0;
                                break;
                            }
                        }
                    }
                    if (this.renglonant >= this.tmgrdpro.Rows.Count)
                    {
                        string str6 = "O.C. ";
                        for (int index = 1; index < this.lstOC.Items.Count; ++index)
                            str6 = str6 + this.lstOC.Items[index].ToString() + ", ";
                        if (this.lstOC.Items.Count > 1)
                            graphics1.DrawString(this.txtObservaciones.Text + " " + str6.TrimEnd(','), font2, (Brush)solidBrush1, 8f, 260f);
                        else
                            graphics1.DrawString(this.txtObservaciones.Text, font2, (Brush)solidBrush1, 8f, 260f);
                        bool flag = false;
                        graphics1.DrawString("Subtotal", font2, (Brush)solidBrush1, 160f, 230f);
                        str2 = "";
                        str3 = "";
                        str4 = "";
                        str5 = "";
                        string str7 = this.txt_ent_ieps.Text == "" ? "0.00" : this.txt_ent_ieps.Text;
                        string text1 = this.txt_iva.Text;
                        string text2 = this.txt_subtotal2.Text;
                        string text3 = this.txt_total.Text;
                        if (!(Convert.ToDecimal(str7) == new Decimal(0)) || !(Convert.ToDecimal(text1) == new Decimal(0)))
                        {
                            if (Convert.ToDecimal(str7) != new Decimal(0) && Convert.ToDecimal(text1) == new Decimal(0))
                                graphics1.DrawString("I.E.P.S.", font2, (Brush)solidBrush1, 160f, 233f);
                            else if (Convert.ToDecimal(str7) == new Decimal(0) && Convert.ToDecimal(text1) != new Decimal(0))
                            {
                                graphics1.DrawString("I.V.A.", font2, (Brush)solidBrush1, 160f, 233f);
                            }
                            else
                            {
                                graphics1.DrawString("I.E.P.S.", font2, (Brush)solidBrush1, 160f, 233f);
                                graphics1.DrawString("I.V.A.", font2, (Brush)solidBrush1, 160f, 236f);
                                graphics1.DrawString("Total", font2, (Brush)solidBrush1, 160f, 239f);
                                flag = true;
                            }
                        }
                        if (!flag)
                            graphics1.DrawString("Total", font2, (Brush)solidBrush1, 160f, 236f);
                        Graphics graphics2 = graphics1;
                        num3 = Convert.ToDecimal(text2);
                        string s3 = num3.ToString("###,###,##0.00");
                        Font font3 = font2;
                        SolidBrush solidBrush2 = solidBrush1;
                        num3 = Convert.ToDecimal(text2);
                        double num4 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                        graphics2.DrawString(s3, font3, (Brush)solidBrush2, (float)num4, 230f);
                        if (!(Convert.ToDecimal(str7) == new Decimal(0)) || !(Convert.ToDecimal(text1) == new Decimal(0)))
                        {
                            if (Convert.ToDecimal(str7) != new Decimal(0) && Convert.ToDecimal(text1) == new Decimal(0))
                            {
                                Graphics graphics3 = graphics1;
                                num3 = Convert.ToDecimal(str7);
                                string s4 = num3.ToString("###,###,##0.00");
                                Font font4 = font2;
                                SolidBrush solidBrush3 = solidBrush1;
                                num3 = Convert.ToDecimal(str7);
                                double num5 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                                graphics3.DrawString(s4, font4, (Brush)solidBrush3, (float)num5, 233f);
                            }
                            else if (Convert.ToDecimal(str7) == new Decimal(0) && Convert.ToDecimal(text1) != new Decimal(0))
                            {
                                Graphics graphics3 = graphics1;
                                num3 = Convert.ToDecimal(text1);
                                string s4 = num3.ToString("###,###,##0.00");
                                Font font4 = font2;
                                SolidBrush solidBrush3 = solidBrush1;
                                num3 = Convert.ToDecimal(text1);
                                double num5 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                                graphics3.DrawString(s4, font4, (Brush)solidBrush3, (float)num5, 233f);
                            }
                            else
                            {
                                Graphics graphics3 = graphics1;
                                num3 = Convert.ToDecimal(str7);
                                string s4 = num3.ToString("###,###,##0.00");
                                Font font4 = font2;
                                SolidBrush solidBrush3 = solidBrush1;
                                num3 = Convert.ToDecimal(str7);
                                double num5 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                                graphics3.DrawString(s4, font4, (Brush)solidBrush3, (float)num5, 233f);
                                Graphics graphics4 = graphics1;
                                num3 = Convert.ToDecimal(text1);
                                string s5 = num3.ToString("###,###,##0.00");
                                Font font5 = font2;
                                SolidBrush solidBrush4 = solidBrush1;
                                num3 = Convert.ToDecimal(text1);
                                double num6 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                                graphics4.DrawString(s5, font5, (Brush)solidBrush4, (float)num6, 236f);
                                Graphics graphics5 = graphics1;
                                num3 = Convert.ToDecimal(text3);
                                string s6 = num3.ToString("###,###,##0.00");
                                Font font6 = font2;
                                SolidBrush solidBrush5 = solidBrush1;
                                num3 = Convert.ToDecimal(text3);
                                double num7 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                                graphics5.DrawString(s6, font6, (Brush)solidBrush5, (float)num7, 239f);
                            }
                        }
                        if (!flag)
                        {
                            Graphics graphics3 = graphics1;
                            num3 = Convert.ToDecimal(text3);
                            string s4 = num3.ToString("###,###,##0.00");
                            Font font4 = font2;
                            SolidBrush solidBrush3 = solidBrush1;
                            num3 = Convert.ToDecimal(text3);
                            double num5 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                            graphics3.DrawString(s4, font4, (Brush)solidBrush3, (float)num5, 239f);
                        }
                        e.HasMorePages = false;
                    }
                    if (this.renglonant < this.tmgrdpro.Rows.Count)
                    {
                        graphics1.DrawString("CONTINUA EN LA SIGUIENTE HOJA.....", font2, (Brush)solidBrush1, 2f, 116f);
                        e.HasMorePages = true;
                    }
                }
                else
                {
                    for (int renglonant = this.renglonant; renglonant < this.tmgrdpro.Rows.Count; ++renglonant)
                    {
                        if (!(this.tmgrdpro.Rows[renglonant]["borrar"].ToString() == "1"))
                        {
                            graphics1.DrawString(Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["cantidad"].ToString()).ToString("###,###,##0.00"), font2, (Brush)solidBrush1, (float)this.posicion(Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["cantidad"].ToString()).ToString("###,###,##0.00")), (float)num2);
                            if (Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["cantidad"].ToString()) > new Decimal(0) && Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["unitario"].ToString()) > new Decimal(0))
                            {
                                this.cmnd2 = this.thisConnection.CreateCommand();
                                this.cmnd2.CommandText = "SELECT um_clave FROM tb_cat_empaques WHERE emp_clave = '" + this.tmgrdpro.Rows[renglonant]["clave"].ToString().Trim() + "'";
                                this.reader2 = this.cmnd2.ExecuteReader();
                                if (this.reader2.HasRows)
                                {
                                    this.reader2.Read();
                                    str1 = this.reader2.GetValue(0).ToString();
                                }
                                this.reader2.Close();
                            }
                            graphics1.DrawString(str1.Substring(0, 3).ToString(), font2, (Brush)solidBrush1, 38f, (float)num2);
                            graphics1.DrawString(this.tmgrdpro.Rows[renglonant]["clave"].ToString() + " " + this.tmgrdpro.Rows[renglonant]["nombre"].ToString(), font2, (Brush)solidBrush1, 50f, (float)num2);
                            graphics1.DrawString(this.tmgrdpro.Rows[renglonant]["lote"].ToString(), font2, (Brush)solidBrush1, 130f, (float)num2);
                            graphics1.DrawString(Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["unitario"].ToString()).ToString("###,###,##0.00"), font2, (Brush)solidBrush1, (float)this.posicionprecio(Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["unitario"].ToString()).ToString("###,###,##0.00")), (float)num2);
                            Decimal num3;
                            if (this.tmgrdpro.Rows[renglonant]["costo"].ToString() != "" && Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["costo"].ToString()) > new Decimal(0))
                            {
                                if (this.DDL_ent_procedencia.SelectedIndex == 0)
                                {
                                    num3 = Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["costo"].ToString());
                                    Decimal num4 = Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["gastoaduanal"].ToString());
                                    Decimal num5 = Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["flete"].ToString());
                                    num3 = num3 + num4 + num5;
                                    graphics1.DrawString(num3.ToString("###,###,##0.00"), font2, (Brush)solidBrush1, (float)this.posicionimporte(num3.ToString("###,###,##0.00")), (float)num2);
                                }
                                else
                                {
                                    num3 = Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["costo"].ToString()) * Convert.ToDecimal(this.txt_sis_tipo_cambio.Text);
                                    Decimal num4 = Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["gastoaduanal"].ToString());
                                    Decimal num5 = Convert.ToDecimal(this.tmgrdpro.Rows[renglonant]["flete"].ToString());
                                    num3 = num3 + num4 + num5;
                                    graphics1.DrawString(num3.ToString("###,###,##0.00"), font2, (Brush)solidBrush1, (float)this.posicionimporte(num3.ToString("###,###,##0.00")), (float)num2);
                                }
                            }
                            num2 += 3;
                            ++this.renglonant;
                            ++this.ren;
                            if (this.ren == 60)
                            {
                                this.ren = 0;
                                break;
                            }
                        }
                    }
                    if (this.renglonant >= this.tmgrdpro.Rows.Count)
                    {
                        string str6 = "O.C. ";
                        for (int index = 1; index < this.lstOC.Items.Count; ++index)
                            str6 = str6 + this.lstOC.Items[index].ToString() + ", ";
                        if (this.lstOC.Items.Count > 1)
                            graphics1.DrawString(this.txtObservaciones.Text + " " + str6.TrimEnd(','), font2, (Brush)solidBrush1, 2f, 116f);
                        else
                            graphics1.DrawString(this.txtObservaciones.Text, font2, (Brush)solidBrush1, 2f, 116f);
                        bool flag = false;
                        graphics1.DrawString("Subtotal", font2, (Brush)solidBrush1, 160f, 230f);
                        str2 = "";
                        str3 = "";
                        str4 = "";
                        str5 = "";
                        string str7 = this.txt_ent_ieps.Text == "" ? "0.00" : this.txt_ent_ieps.Text;
                        string text1 = this.txt_iva.Text;
                        string text2 = this.txt_subtotal2.Text;
                        string text3 = this.txt_total.Text;
                        if (!(Convert.ToDecimal(str7) == new Decimal(0)) || !(Convert.ToDecimal(text1) == new Decimal(0)))
                        {
                            if (Convert.ToDecimal(str7) != new Decimal(0) && Convert.ToDecimal(text1) == new Decimal(0))
                                graphics1.DrawString("I.E.P.S.", font2, (Brush)solidBrush1, 160f, 233f);
                            else if (Convert.ToDecimal(str7) == new Decimal(0) && Convert.ToDecimal(text1) != new Decimal(0))
                            {
                                graphics1.DrawString("I.V.A.", font2, (Brush)solidBrush1, 160f, 233f);
                            }
                            else
                            {
                                graphics1.DrawString("I.E.P.S.", font2, (Brush)solidBrush1, 160f, 233f);
                                graphics1.DrawString("I.V.A.", font2, (Brush)solidBrush1, 160f, 236f);
                                graphics1.DrawString("Total", font2, (Brush)solidBrush1, 160f, 239f);
                                flag = true;
                            }
                        }
                        if (!flag)
                            graphics1.DrawString("Total", font2, (Brush)solidBrush1, 160f, 236f);
                        Graphics graphics2 = graphics1;
                        Decimal num3 = Convert.ToDecimal(text2);
                        string s3 = num3.ToString("###,###,##0.00");
                        Font font3 = font2;
                        SolidBrush solidBrush2 = solidBrush1;
                        num3 = Convert.ToDecimal(text2);
                        double num4 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                        graphics2.DrawString(s3, font3, (Brush)solidBrush2, (float)num4, 230f);
                        if (!(Convert.ToDecimal(str7) == new Decimal(0)) || !(Convert.ToDecimal(text1) == new Decimal(0)))
                        {
                            if (Convert.ToDecimal(str7) != new Decimal(0) && Convert.ToDecimal(text1) == new Decimal(0))
                            {
                                Graphics graphics3 = graphics1;
                                num3 = Convert.ToDecimal(str7);
                                string s4 = num3.ToString("###,###,##0.00");
                                Font font4 = font2;
                                SolidBrush solidBrush3 = solidBrush1;
                                num3 = Convert.ToDecimal(str7);
                                double num5 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                                graphics3.DrawString(s4, font4, (Brush)solidBrush3, (float)num5, 233f);
                            }
                            else if (Convert.ToDecimal(str7) == new Decimal(0) && Convert.ToDecimal(text1) != new Decimal(0))
                            {
                                Graphics graphics3 = graphics1;
                                num3 = Convert.ToDecimal(text1);
                                string s4 = num3.ToString("###,###,##0.00");
                                Font font4 = font2;
                                SolidBrush solidBrush3 = solidBrush1;
                                num3 = Convert.ToDecimal(text1);
                                double num5 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                                graphics3.DrawString(s4, font4, (Brush)solidBrush3, (float)num5, 233f);
                            }
                            else
                            {
                                Graphics graphics3 = graphics1;
                                num3 = Convert.ToDecimal(str7);
                                string s4 = num3.ToString("###,###,##0.00");
                                Font font4 = font2;
                                SolidBrush solidBrush3 = solidBrush1;
                                num3 = Convert.ToDecimal(str7);
                                double num5 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                                graphics3.DrawString(s4, font4, (Brush)solidBrush3, (float)num5, 233f);
                                Graphics graphics4 = graphics1;
                                num3 = Convert.ToDecimal(text1);
                                string s5 = num3.ToString("###,###,##0.00");
                                Font font5 = font2;
                                SolidBrush solidBrush4 = solidBrush1;
                                num3 = Convert.ToDecimal(text1);
                                double num6 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                                graphics4.DrawString(s5, font5, (Brush)solidBrush4, (float)num6, 236f);
                                Graphics graphics5 = graphics1;
                                num3 = Convert.ToDecimal(text3);
                                string s6 = num3.ToString("###,###,##0.00");
                                Font font6 = font2;
                                SolidBrush solidBrush5 = solidBrush1;
                                num3 = Convert.ToDecimal(text3);
                                double num7 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                                graphics5.DrawString(s6, font6, (Brush)solidBrush5, (float)num7, 239f);
                            }
                        }
                        if (!flag)
                        {
                            Graphics graphics3 = graphics1;
                            num3 = Convert.ToDecimal(text3);
                            string s4 = num3.ToString("###,###,##0.00");
                            Font font4 = font2;
                            SolidBrush solidBrush3 = solidBrush1;
                            num3 = Convert.ToDecimal(text3);
                            double num5 = (double)this.posicion2_pdf(num3.ToString("###,###,##0.00"));
                            graphics3.DrawString(s4, font4, (Brush)solidBrush3, (float)num5, 236f);
                        }
                        e.HasMorePages = false;
                    }
                    if (this.renglonant < this.tmgrdpro.Rows.Count)
                    {
                        graphics1.DrawString("CONTINUA EN LA SIGUIENTE HOJA.....", font2, (Brush)solidBrush1, 2f, 116f);
                        e.HasMorePages = true;
                    }
                }
                this.thisConnection.Close();
            }
            catch (SystemException ex)
            {
                
                thisConnection.Close();
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + ex.ToString());
                int num = (int)MessageBox.Show(ex.ToString());
            }
        }

        private void btnValidar_Click(object sender, EventArgs e)
        {
            if (!this.txtAcuse.Text.Contains<char>('-'))
            {
                MessageBox.Show("El acuse ingresado es erroneo, vuelva a ingresarlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.txtAcuse.Text = "";
            }
            else
            {
                this.conn.Open();
                string[] strArray = this.txtAcuse.Text.Split('-');
                this.comm = this.conn.CreateCommand();
                this.comm.CommandText = "SELECT cfd_keyfec, cfd_timfis, cfd_totalv, cfd_sercfd, cfd_folioc, cfd_keyems FROM CF_REGISTRO_WEB WHERE cfd_keyusu = '" + strArray[0].ToString() + "' AND cfd_keyupl = '" + strArray[1] + "'";
                this.redr = this.comm.ExecuteReader();
                if (this.redr.HasRows)
                {
                    while (this.redr.Read())
                    {
                        DataRow row = this.dtPortal.NewRow();
                        row["cfd_keyfec"] = (object)this.redr["cfd_keyfec"].ToString().Trim();
                        row["cfd_timfis"] = (object)this.redr["cfd_timfis"].ToString().Trim();
                        row["cfd_totalv"] = (object)this.redr["cfd_totalv"].ToString().Trim();
                        row["cfd_sercfd"] = (object)this.redr["cfd_sercfd"].ToString().Trim();
                        row["cfd_folioc"] = (object)this.redr["cfd_folioc"].ToString().Trim();
                        row["cfd_keyems"] = (object)this.redr["cfd_keyems"].ToString().Trim();
                        this.dtPortal.Rows.Add(row);
                        this.lblRfcPortal.Text = this.redr["cfd_keyems"].ToString().Trim();
                    }
                }
                this.redr.Close();
                this.redr.Dispose();
                this.comm.Dispose();
                this.conn.Close();
                if (this.dtPortal.Rows.Count == 1)
                {
                    this.lblRfcPortal.Text = this.dtPortal.Rows[0]["cfd_keyems"].ToString().Trim();
                    this.txt_ent_factura.Text = this.dtPortal.Rows[0]["cfd_sercfd"].ToString().Trim() + this.dtPortal.Rows[0]["cfd_folioc"].ToString().Trim();
                    this.txtFechaFact.Text = Convert.ToDateTime(this.dtPortal.Rows[0]["cfd_keyfec"].ToString().Trim()).ToShortDateString();
                    this.lblTotal.Text = this.dtPortal.Rows[0]["cfd_totalv"].ToString().Trim();
                    this.txtFolioFiscal.Text = this.dtPortal.Rows[0]["cfd_timfis"].ToString().Trim();
                }
                else
                {
                    foreach (DataRow row in (InternalDataCollectionBase)this.dtPortal.Rows)
                        this.dtgPortal.Rows.Add((object)row["cfd_sercfd"].ToString(), (object)row["cfd_folioc"].ToString(), (object)Convert.ToDateTime(row["cfd_keyfec"].ToString()).ToShortDateString(), (object)row["cfd_timfis"].ToString(), (object)Convert.ToDecimal(row["cfd_totalv"].ToString()).ToString("$###,###,##0.00"), (object)row["cfd_keyems"].ToString());
                    this.panel3.Visible = true;
                }
            }
        }

        private void btnCierra_Click(object sender, EventArgs e)
        {
            this.dtPortal.Clear();
            this.dtgPortal.Rows.Clear();
            this.panel3.Visible = false;
        }

        private void dtgPortal_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.lblRfcPortal.Text = this.dtPortal.Rows[this.dtgPortal.CurrentRow.Index]["cfd_keyems"].ToString().Trim();
            this.txt_ent_factura.Text = this.dtPortal.Rows[this.dtgPortal.CurrentRow.Index]["cfd_sercfd"].ToString().Trim() + this.dtPortal.Rows[this.dtgPortal.CurrentRow.Index]["cfd_folioc"].ToString().Trim();
            this.txtFechaFact.Text = Convert.ToDateTime(this.dtPortal.Rows[this.dtgPortal.CurrentRow.Index]["cfd_keyfec"].ToString().Trim()).ToShortDateString();
            this.lblTotal.Text = this.dtPortal.Rows[this.dtgPortal.CurrentRow.Index]["cfd_totalv"].ToString().Trim();
            this.txtFolioFiscal.Text = this.dtPortal.Rows[this.dtgPortal.CurrentRow.Index]["cfd_timfis"].ToString().Trim();
            this.dtgPortal.Rows.Clear();
            this.dtPortal.Clear();
            this.panel3.Visible = false;
        }

        private void btnCambioFactura_Click(object sender, EventArgs e)
        {
            if (operacion == "consulta")
            {
                if (lblAguilares.Text == "AGUILARES")
                {
                    MessageBox.Show("La entrada fue registrada desde AGUILARES no puede ser modificada desde GAB, solicitar movimiento a AGUILARES, por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                factura dlg = new factura(txtFolio.Text, glo_var_chr_alm_clave, DDLAlmacen1.SelectedItem.ToString(), txt_prov_clave.Text, DDLProveedor.SelectedItem.ToString(), txt_ent_factura.Text, txtFolioFiscal.Text);
                dlg.ShowDialog();
            }   
        }

        private void DDLAutorizo_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (DDLAutorizo.SelectedIndex > -1)
            {
                //
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT estatus FROM tb_cat_usuarios_emp WHERE usu_nombre = '" + DDLAutorizo.SelectedItem + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    if (reader1["estatus"].ToString().Trim() == "B")
                    {
                        MessageBox.Show("Usuario dado de baja, no puede ser seleccionado", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();
                        thisConnection.Close();
                        DDLAutorizo.SelectedIndex = -1;
                        return;
                    }
                }
                thisConnection.Close();
            }
        }

        public string recalculo_costos_empaque(string clave, string fech)
        { 
            string mcod = "";
            decimal minvi = 0;
            decimal mcosp = 0;

            string mrec = "";
            string mtip = "";
            string malm = "";
            decimal mcant = 0;

            decimal A = 0, B = 0, C = 0, D = 0, C2 = 0;

            DataTable tbhist = new DataTable();
            tbhist.Columns.Add("hrp_recibo", typeof(string));
            tbhist.Columns.Add("hrp_fecha", typeof(string));
            tbhist.Columns.Add("emp_clave", typeof(string));
            tbhist.Columns.Add("hrp_cantidad", typeof(string));
            tbhist.Columns.Add("hrp_tipo_recepcion", typeof(string));
            tbhist.Columns.Add("alm_clave", typeof(string));
            tbhist.Columns.Add("hrp_situacion", typeof(string));
            tbhist.Columns.Add("hrp_costo", typeof(string));
            tbhist.Columns.Add("hrp_regimen", typeof(string));
            tbhist.Columns.Add("hrp_procedencia", typeof(string));
            tbhist.Columns.Add("hrp_costo_prom", typeof(string));
            tbhist.Columns.Add("hrp_tipo", typeof(string));

            DataTable tbinv = new DataTable();
            tbinv.Columns.Add("emp_clave", typeof(string));
            tbinv.Columns.Add("invemp_inicial", typeof(string));
            tbinv.Columns.Add("invemp_salidas", typeof(string));
            tbinv.Columns.Add("invemp_entradas", typeof(string));
            tbinv.Columns.Add("alm_clave", typeof(string));
            tbinv.Columns.Add("invemp_costopro", typeof(string));

            //cat_empaques ya esa cargado en dtempaques

            DataTable tbent = new DataTable();
            tbent.Columns.Add("ent_folio", typeof(string));
            tbent.Columns.Add("emp_clave", typeof(string));
            tbent.Columns.Add("alm_clave", typeof(string));
            tbent.Columns.Add("entd_cantidad", typeof(string));
            tbent.Columns.Add("entd_costo", typeof(string));
            tbent.Columns.Add("ent_tipo", typeof(string));

            DataTable tbsal = new DataTable();
            tbsal.Columns.Add("sal_folio", typeof(string));
            tbsal.Columns.Add("emp_clave", typeof(string));
            tbsal.Columns.Add("sald_cantidad", typeof(string));
            tbsal.Columns.Add("sald_costo", typeof(string));
            tbsal.Columns.Add("alm_clave", typeof(string));
            tbsal.Columns.Add("sal_tipo", typeof(string));

            DataTable tbinv2 = new DataTable();
            tbinv2.Columns.Add("emp_clave", typeof(string));
            tbinv2.Columns.Add("invemp_inicial", typeof(string));
            tbinv2.Columns.Add("invemp_salidas", typeof(string));
            tbinv2.Columns.Add("invemp_entradas", typeof(string));
            tbinv2.Columns.Add("alm_clave", typeof(string));
            tbinv2.Columns.Add("invemp_costopro", typeof(string));

            //DateTime fecha = new DateTime();
            //fecha = Convert.ToDateTime(dtpFecha.Text);

            string fecha_inicio = "";
            string fecha_fin = "";

            //fecha_inicio = fecha.AddMonths(-1).ToShortDateString();
            fecha_fin = Convert.ToDateTime(fech).ToShortDateString();//fecha.ToShortDateString();//fecha.AddDays(-1).ToShortDateString();

            string query = "";
            try
            {
                thisConnection.Open();

                //SACAR FECHA ULTIMO CIERRE
                string var_date_fecha = "";
                cmnd1 = thisConnection.CreateCommand();
                query = "SELECT TOP 1 invemp_fecha FROM tb_mstr_inventario_emp WHERE invemp_fecha <= '" + DateTime.Now.ToShortDateString() + "' ORDER BY invemp_fecha DESC";
                cmnd1.CommandText = query;
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    var_date_fecha = reader1.GetValue(0).ToString().Trim();

                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                fecha_inicio = var_date_fecha;

                cmnd1 = thisConnection.CreateCommand();
                query = "SELECT hrp_recibo, hrp_fecha, emp_clave, hrp_cantidad, hrp_tipo_recepcion, alm_clave, hrp_situacion, hrp_costo, hrp_regimen, hrp_procedencia, hrp_costo_prom, hrp_tipo" +
                    " FROM tb_historico_recepcion WHERE hrp_estatus <> 'C' AND (hrp_fecha >= '" + Convert.ToDateTime(fecha_inicio).ToShortDateString() + "' AND hrp_fecha <= '" + fecha_fin + "')" +
                    " AND alm_clave = '" + glo_var_chr_alm_clave + "' and emp_clave = '" + clave + "'" +
                    //" and emp_clave = 'B0001'" +//and emp_clave = 'AB088'
                    " ORDER BY emp_clave, hrp_fecha, hrp_tipo_recepcion";
                cmnd1.CommandText = query;

                reader1 = cmnd1.ExecuteReader();
                DataRow rw;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rw = tbhist.NewRow();
                        rw["hrp_recibo"] = reader1.GetValue(0).ToString().Trim();
                        rw["hrp_fecha"] = reader1.GetValue(1).ToString().Trim();
                        rw["emp_clave"] = reader1.GetValue(2).ToString().Trim();
                        rw["hrp_cantidad"] = reader1.GetValue(3).ToString().Trim();
                        rw["hrp_tipo_recepcion"] = reader1.GetValue(4).ToString().Trim();
                        rw["alm_clave"] = reader1.GetValue(5).ToString().Trim();
                        rw["hrp_situacion"] = reader1.GetValue(6).ToString().Trim();
                        rw["hrp_costo"] = reader1.GetValue(7).ToString().Trim();
                        rw["hrp_regimen"] = reader1.GetValue(8).ToString().Trim();
                        rw["hrp_procedencia"] = reader1.GetValue(9).ToString().Trim();
                        rw["hrp_costo_prom"] = reader1.GetValue(10).ToString().Trim();
                        rw["hrp_tipo"] = reader1.GetValue(11).ToString().Trim();
                        tbhist.Rows.Add(rw);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                query = "SELECT emp_clave, alm_clave FROM tb_cat_empaques WHERE emp_nombre <> '' AND alm_clave = '" + glo_var_chr_alm_clave + "' and emp_clave = '" + clave + "' ORDER BY emp_clave";//
                cmnd1.CommandText = query;
                DataRow r1;
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        r1 = tbinv.NewRow();
                        r1["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                        r1["invemp_inicial"] = "0";
                        r1["invemp_salidas"] = "0";
                        r1["invemp_entradas"] = "0";
                        r1["alm_clave"] = reader1.GetValue(1).ToString().Trim();
                        r1["invemp_costopro"] = "0";
                        tbinv.Rows.Add(r1);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                query = "SELECT emp_clave, invemp_inicial, invemp_salidas, invemp_entradas, alm_clave, invemp_costopro FROM tb_mstr_inventario_emp " +
                     "WHERE invemp_fecha = '" + Convert.ToDateTime(fecha_inicio).ToShortDateString() + "' and emp_clave = '" + clave + "' ORDER BY emp_clave";//--and emp_clave = 'B0001'
                cmnd1.CommandText = query;
                reader1 = cmnd1.ExecuteReader();
                DataRow r2;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        r2 = tbinv2.NewRow();
                        r2["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                        r2["invemp_inicial"] = reader1.GetValue(1).ToString().Trim();
                        r2["invemp_salidas"] = reader1.GetValue(2).ToString().Trim();
                        r2["invemp_entradas"] = reader1.GetValue(3).ToString().Trim();
                        r2["alm_clave"] = reader1.GetValue(4).ToString().Trim();
                        r2["invemp_costopro"] = reader1.GetValue(5).ToString().Trim();
                        tbinv2.Rows.Add(r2);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();

                foreach (DataRow rz in tbinv.Rows)
                {
                    foreach (DataRow ry in tbinv2.Select("emp_clave = '" + rz["emp_clave"] + "'"))
                    {
                        rz["invemp_inicial"] = ry["invemp_inicial"];
                        rz["invemp_costopro"] = ry["invemp_costopro"];
                    }
                }

                thisConnection.Open();
                bool hay = false;
                //int y = 0;
                foreach (DataRow rinv in tbinv.Rows)
                {
                    mcod = rinv[0].ToString();
                    minvi = Convert.ToDecimal(rinv[1].ToString());
                    mcosp = Convert.ToDecimal(rinv[5].ToString());

                    ////validar si ya fue generado el cierre para ese empaque
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "SELECT emp_clave FROM tb_empaque_cierre WHERE emp_clave = '" + mcod + "'";
                    //reader1 = cmnd1.ExecuteReader();
                    //if (reader1.HasRows)
                    //{
                    //    y++;
                    //    label1.Text = (y).ToString() + " DE " + dtempaques.Rows.Count.ToString() + " " + mcod;
                    //    label1.Update();
                    //    continue;
                    //}
                        

                    foreach (DataRow rhis in tbhist.Select("emp_clave = '" + mcod + "'"))
                    {
                        hay = true;
                        //mesh = Convert.ToDateTime(rhis[1].ToString()).Month.ToString();
                        mrec = rhis["hrp_recibo"].ToString();
                        mtip = rhis["hrp_tipo_recepcion"].ToString();
                        malm = rhis["alm_clave"].ToString();
                        if (rhis["hrp_tipo_recepcion"].ToString() == "ENT" && rhis["hrp_situacion"].ToString() == "NOR")
                        {
                            decimal hrp_cantidad = Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            decimal hrp_costo = Convert.ToDecimal(rhis["hrp_costo"].ToString());
                            A = Math.Abs(minvi * mcosp);
                            B = hrp_cantidad * hrp_costo;
                            C = minvi + hrp_cantidad;
                            C2 = Math.Abs(minvi) + hrp_cantidad;
                            D = (C2 == 0) ? hrp_costo : ((A + B) / C2);
                            minvi = C;
                            mcosp = Math.Abs(D);
                            //Actualizacion de historico
                            //cmnd1 = thisConnection.CreateCommand();
                            //cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_existencia = '" + minvi.ToString() + "', hrp_costo_prom = '" + mcosp.ToString() + "'" +
                            //    " WHERE hrp_tipo_recepcion = '" + mtip + "' AND alm_clave = '" + malm + "' AND hrp_recibo = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            ////cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();
                        }
                        if (rhis["hrp_tipo_recepcion"].ToString() == "ENT" && rhis["hrp_situacion"].ToString() == "DEV")
                        {
                            minvi = minvi + Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            mcant = Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            //Actualizacion de historico
                            //cmnd1 = thisConnection.CreateCommand();
                            //cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_existencia = '" + minvi.ToString() + "', hrp_costo_prom = '" + mcosp.ToString() + "', hrp_costo = '" + mcosp.ToString() + "'" +
                            //    " WHERE hrp_tipo_recepcion = '" + mtip + "' AND alm_clave = '" + malm + "' AND hrp_recibo = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();

                            decimal cto = 0;
                            cto = mcosp * mcant;
                            //cmnd1 = thisConnection.CreateCommand();
                            //cmnd1.CommandText = "UPDATE tb_det_ent_temporal SET entd_costo = '" + cto.ToString() + "'" +
                            //    " WHERE alm_clave = '" + malm + "' AND ent_folio = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();
                        }
                        if (rhis["hrp_tipo_recepcion"].ToString() == "SAL")
                        {
                            minvi = minvi - Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            ////Actualizacion de historico
                            //cmnd1 = thisConnection.CreateCommand();
                            //cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_existencia = '" + minvi.ToString() + "', hrp_costo_prom = '" + mcosp.ToString() + "', hrp_costo = '" + mcosp.ToString() + "'" +
                            //    " WHERE hrp_tipo_recepcion = '" + mtip + "' AND alm_clave = '" + malm + "' AND hrp_recibo = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();

                            //cmnd1 = thisConnection.CreateCommand();
                            //cmnd1.CommandText = "UPDATE tb_det_salidas_temporal SET sald_costo = '" + mcosp.ToString() + "'" +
                            //    " WHERE alm_clave = '" + malm + "' AND sal_folio = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();
                        }

                        cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = "UPDATE tb_cat_empaques_prueba SET emp_costo = '" + mcosp.ToString() + "', emp_existencia = '" + minvi + "'" +
                        //    " WHERE emp_clave = '" + mcod + "'";
                        query = "UPDATE tb_cat_empaques SET emp_costo = '" + Math.Round(mcosp, 3).ToString() + "'" +
                            " WHERE emp_clave = '" + mcod + "'";
                        cmnd1.CommandText = query;
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        try
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            query = "UPDATE tb_cat_empaques SET emp_costo = '" + Math.Round(mcosp, 3).ToString() + "' " +
                                " WHERE emp_clave = '" + mcod + "' and alm_clave = '" + malm + "'";
                            cmnd1.CommandText = query;
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                        }
                        catch (SqlException ex)
                        {
                            if (ex.ToString().Contains("La transacción (id. de proceso"))
                            {
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + ex.ToString() + "1RA QUERY interbloqueo: " + query);

                                cmnd1 = thisConnection.CreateCommand();
                                query = "UPDATE tb_cat_empaques SET emp_costo = '" + Math.Round(mcosp, 3).ToString() + "' " +
                                    " WHERE emp_clave = '" + mcod + "' and alm_clave = '" + malm + "'";
                                cmnd1.CommandText = query;
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();
                            }
                        }


                    }
                    if (hay == false)
                    {
                        try
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            query = "UPDATE tb_cat_empaques SET emp_costo = '" + Math.Round(mcosp, 3).ToString() + "' " +
                                " WHERE emp_clave = '" + mcod + "' and alm_clave = '" + malm + "'";
                            cmnd1.CommandText = query;
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                        }
                        catch (SqlException ex)
                        {
                            if (ex.ToString().Contains("La transacción (id. de proceso"))
                            {
                                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + ex.ToString() + "2DA QUERY interbloqueo: " + query);

                                cmnd1 = thisConnection.CreateCommand();
                                query = "UPDATE tb_cat_empaques SET emp_costo = '" + Math.Round(mcosp, 3).ToString() + "' " +
                                    " WHERE emp_clave = '" + mcod + "' and alm_clave = '" + malm + "'";
                                cmnd1.CommandText = query;
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();
                            }
                        }
                        //query = "UPDATE tb_cat_empaques SET emp_costo = '" + Math.Round(mcosp, 3).ToString() + "' " +
                        //    " WHERE emp_clave = '" + mcod + "' and alm_clave = '" + malm + "'";
                        //cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = query;
                        //cmnd1.ExecuteNonQuery();
                        //cmnd1.Dispose();

                        //cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = "UPDATE tb_cat_empaques_prueba SET emp_existencia = '" + minvi + "'" +
                        //    " WHERE emp_clave = '" + mcod + "'";
                        //cmnd1.ExecuteNonQuery();
                        //cmnd1.Dispose();
                    }
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "INSERT INTO tb_empaque_cierre(emp_clave) VALUES('" + mcod + "')";
                    //cmnd1.ExecuteNonQuery();
                    //cmnd1.Dispose();

                    //y++;
                    //label1.Text = (y).ToString() + " DE " + dtempaques.Rows.Count.ToString() + " " + mcod;
                    //label1.Update();
                }

                cmnd1 = thisConnection.CreateCommand();
                query = "UPDATE tb_cat_empaques SET emp_entrada_ultima = '" + fech + "' WHERE emp_clave = '" + clave + "' AND emp_entrada_ultima < '" + fech + "'";
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();

                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + query);

                thisConnection.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.1", ex.ToString(), "SISEMP");
                form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + ex.ToString() + "QUERY: " + query);
            }
            return mcosp.ToString();
        }

        public string recalculo_existencia_empaque(string clave)
        {
            dtempaque.Clear();
            dtinv.Clear();
            dtconent.Clear();
            dtconsal.Clear();

            thisConnection.Open();

            //CARGAR TODOS LOS EMPAQUES
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT emp_clave, emp_nombre FROM tb_cat_empaques WHERE emp_clave = '" + clave + "' ORDER BY emp_clave";
            reader1 = cmnd1.ExecuteReader();
            DataRow rw;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    rw = dtempaque.NewRow();
                    rw["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                    rw["emp_nombre"] = reader1.GetValue(1).ToString().Trim();
                    rw["invemp_inicial"] = "0";
                    rw["tent"] = "0";
                    rw["tsal"] = "0";
                    rw["inv_actual"] = "0";
                    dtempaque.Rows.Add(rw);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            //SACAR FECHA ULTIMO CIERRE
            string var_date_fecha = "";
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT TOP 1 invemp_fecha FROM tb_mstr_inventario_emp WHERE invemp_fecha <= '" + DateTime.Now.ToShortDateString() + "' ORDER BY invemp_fecha DESC";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                var_date_fecha = reader1.GetValue(0).ToString().Trim();

            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            string actual = "0";


            cmnd1.CommandText = "SELECT I.emp_clave, I.invemp_inicial FROM tb_mstr_inventario_emp I" +
                " WHERE I.invemp_fecha = '" + Convert.ToDateTime(var_date_fecha).ToShortDateString() + "' AND emp_clave = '" + clave + "' ORDER BY I.emp_clave";
            reader1 = cmnd1.ExecuteReader();
            DataRow dtr;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    dtr = dtinv.NewRow();
                    dtr["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                    dtr["inicial"] = reader1.GetValue(1).ToString().Trim();
                    dtinv.Rows.Add(dtr);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            foreach (DataRow r1 in dtempaque.Rows)
            {
                foreach (DataRow r2 in dtinv.Select("emp_clave = '" + r1["emp_clave"] + "'"))
                {
                    r1["invemp_inicial"] = r2["inicial"].ToString();
                }
            }

            cmnd1 = thisConnection.CreateCommand();
            //cmnd1.CommandText = "select emp_clave, sum(hrp_cantidad) FROM tb_historico_recepcion " +
            //    "WHERE hrp_fecha >= '" + Convert.ToDateTime(fecha).ToShortDateString() + "' and hrp_tipo_recepcion = 'SAL' and hrp_estatus <> 'C' " +
            //    "group by emp_clave " +
            //    "order by emp_clave";
            cmnd1.CommandText = "select emp_clave, sum(hrp_cantidad) FROM tb_historico_recepcion " +
                "WHERE hrp_fecha >= '" + Convert.ToDateTime(var_date_fecha).ToShortDateString() + "' and hrp_fecha <= '" + DateTime.Now.ToShortDateString() + "' " +
                " and hrp_tipo_recepcion = 'SAL' and hrp_estatus <> 'C' AND emp_clave = '" + clave + "' " +
                "group by emp_clave " +
                "order by emp_clave";
            reader1 = cmnd1.ExecuteReader();
            DataRow r_1;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    r_1 = dtconsal.NewRow();
                    r_1["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                    r_1["total"] = reader1.GetValue(1).ToString().Trim();
                    dtconsal.Rows.Add(r_1);
                }
            }
            cmnd1.Dispose();
            cmnd1 = thisConnection.CreateCommand();
            //cmnd1.CommandText = "select emp_clave, sum(hrp_cantidad) FROM tb_historico_recepcion " +
            //    "WHERE hrp_fecha >= '" + Convert.ToDateTime(fecha).ToShortDateString() + "' and hrp_tipo_recepcion = 'ENT' AND hrp_situacion in ('NOR','DEV') and hrp_estatus <> 'C' " +
            //    "group by emp_clave " +
            //    "order by emp_clave";
            cmnd1.CommandText = "select emp_clave, sum(hrp_cantidad) FROM tb_historico_recepcion " +
                "WHERE hrp_fecha >= '" + Convert.ToDateTime(var_date_fecha).ToShortDateString() + "' and hrp_fecha <= '" + DateTime.Now.ToShortDateString() + "' " +
                "and hrp_tipo_recepcion = 'ENT' AND hrp_situacion in ('NOR','DEV', 'AJT') and hrp_estatus <> 'C' AND emp_clave = '" + clave + "' " +
                "group by emp_clave " +
                "order by emp_clave";
            reader1 = cmnd1.ExecuteReader();
            DataRow r_2;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    r_2 = dtconent.NewRow();
                    r_2["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                    r_2["total"] = reader1.GetValue(1).ToString().Trim();
                    dtconent.Rows.Add(r_2);
                }
            }
            cmnd1.Dispose();


            foreach (DataRow rxx in dtempaque.Rows)
            {
                foreach (DataRow rx1 in dtconent.Select("emp_clave = '" + rxx["emp_clave"].ToString() + "'"))
                {
                    rxx["tent"] = rx1["total"].ToString();
                    rxx["inv_actual"] = ((Convert.ToDecimal(rxx["invemp_inicial"].ToString()) + Convert.ToDecimal(rxx["tent"].ToString())) - Convert.ToDecimal(rxx["tsal"].ToString())).ToString();
                }
                foreach (DataRow rx2 in dtconsal.Select("emp_clave = '" + rxx["emp_clave"].ToString() + "'"))
                {
                    rxx["tsal"] = rx2["total"].ToString();
                    rxx["inv_actual"] = ((Convert.ToDecimal(rxx["invemp_inicial"].ToString()) + Convert.ToDecimal(rxx["tent"].ToString())) - Convert.ToDecimal(rxx["tsal"].ToString())).ToString();
                }
            }

            foreach (DataRow rxx in dtempaque.Rows)
            {
                rxx["inv_actual"] = ((Convert.ToDecimal(rxx["invemp_inicial"].ToString()) + Convert.ToDecimal(rxx["tent"].ToString())) - Convert.ToDecimal(rxx["tsal"].ToString())).ToString();
            }



            actual = dtempaque.Rows[0]["inv_actual"].ToString();

            cmnd1 = thisConnection.CreateCommand();
            string query = "update tb_cat_empaques set emp_existencia = '" + actual + "' WHERE emp_clave = '" + clave + "'";
            cmnd1.CommandText = query;                
            cmnd1.ExecuteNonQuery();
            cmnd1.Dispose();
            

            thisConnection.Close();

            form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + query);

            return actual;
        }

        private void printDocument1_EndPrint(object sender, PrintEventArgs e)
        {
            MessageBox.Show("Se termino de imprimir");
            //Dispose(true);
            
        }

        public void genera_pdf(string arquivo_nome)
        {
            iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER, 0, 0, 0, 0);
            iTextSharp.text.pdf.PdfWriter write = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, new FileStream(@"c:\Reportes\" + arquivo_nome, FileMode.Create));
            doc.Open();

            //iTextSharp.text.Paragraph title = new iTextSharp.text.Paragraph();
            Image imagen = form_entradas_ocs.Properties.Resources.NOTA_DE_ENTRADA_011;
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imagen, iTextSharp.text.BaseColor.WHITE);
            jpg.ScaleToFit(doc.PageSize.Width, doc.PageSize.Height);
            jpg.SetAbsolutePosition(0, 0);
            doc.Add(jpg);
            //title.Font = iTextSharp.text.FontFactory.GetFont(iTextSharp.text.FontFactory.TIMES, 18f, iTextSharp.text.BaseColor.BLUE);
            //title.Add("Hola Mundo!!");
            //doc.Add(title);
            iTextSharp.text.pdf.PdfContentByte cb = write.DirectContent;
            iTextSharp.text.pdf.BaseFont bfTimes = iTextSharp.text.pdf.BaseFont.CreateFont(iTextSharp.text.pdf.BaseFont.COURIER, iTextSharp.text.pdf.BaseFont.CP1252, false);
            iTextSharp.text.Font fuente1 = new iTextSharp.text.Font(bfTimes, 10);

            cb.BeginText();
            cb.SetFontAndSize(bfTimes, 10);
            //

            cb.SetTextMatrix(510, 767);//TIPO DE ENTRADA
            cb.ShowText(txtFolio.Text);

            cb.SetTextMatrix(510, 760);//CLAVE DE LA ENTRADA
            cb.ShowText("NORMAL");

            cb.SetTextMatrix(380, 740);//FECHA DE LA ENTRADA
            cb.ShowText(dtpFecha.Text);

            cb.SetFontAndSize(bfTimes, 8);

            cb.SetTextMatrix(25, 710);//PROVEEDOR
            cb.ShowText(DDLProveedor.Text);

            cb.SetTextMatrix(330, 710);//FACTURA
            cb.ShowText(txt_ent_factura.Text);

            List<string> lista_oc = new List<string>();

            foreach (string val in lstOC.Items)
            {
                lista_oc.Add(val);
            }
            //thisConnection.Open();
            //cmnd1 = thisConnection.CreateCommand();
            //cmnd1.CommandText = "SELECT distinct numero_oc FROM tb_det_entrada_oc WHERE ent_folio = '" + txtFolio.Text + "' and alm_clave = '" + glo_var_chr_alm_clave + "'";
            //reader1 = cmnd1.ExecuteReader();
            //if (reader1.HasRows)
            //{
            //    while (reader1.Read())
            //    {
            //        lista_oc.Add(reader1["numero_oc"].ToString().Trim());
            //    }
            //}
            //reader1.Close();
            //reader1.Dispose();
            //cmnd1.Dispose();
            //thisConnection.Close();
            cb.SetTextMatrix(510, 710);//ORDEN COMPRA lstOC.Items[0].ToString()
            cb.ShowText(lista_oc[0].ToString());

            cb.SetTextMatrix(25, 50);//OBSERVACIONES
            cb.ShowText(txtObservaciones.Text);

            cb.SetTextMatrix(25, 25);//RECIBIDO
            cb.ShowText(DDLElaboro.Text);

            cb.SetTextMatrix(165, 25);//CAPTURADO
            cb.ShowText(DDLAutorizo.Text);

            cb.SetTextMatrix(305, 25);//ENTREGADO
            cb.ShowText("");

            cb.SetTextMatrix(480, 27);//REVISION
            cb.ShowText("F-500-11");
            cb.SetTextMatrix(480, 20);//REVISION
            cb.ShowText("REV: 00");

            cb.SetFontAndSize(bfTimes, 7);

            cb.SetTextMatrix(365, 689);//TRAZABILIDAD
            cb.ShowText("Lote Rastreabilidad");

            decimal sum_importe = 0;
            decimal sum_gtos_im = 0;

            float row_count = 682;
            foreach (DataRow indx in tmgrdpro.Rows)
            {
                //CONSULTAR MEDIDA
                //thisConnection.Open();
                if (thisConnection.State == ConnectionState.Closed)
                    thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT um_clave FROM tb_cat_empaques WHERE emp_clave = '" + indx["clave"].ToString() + "'";
                reader1 = cmnd1.ExecuteReader();
                string medida = "";
                if (reader1.HasRows)
                {
                    reader1.Read();
                    medida = reader1["um_clave"].ToString().Trim();
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                //thisConnection.Close();
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();

                cb.SetTextMatrix(cantidad_pdf_gab(Convert.ToDecimal(indx["cantidad"].ToString()).ToString("###,###,###,##0.00")), row_count);//CANTIDAD
                cb.ShowText(Convert.ToDecimal(indx["cantidad"].ToString()).ToString("###,###,###,##0.00"));

                cb.SetTextMatrix(110, row_count);//MEDIDA
                cb.ShowText(medida);

                //CLAVE Y NOMBRE
                cb.SetTextMatrix(143, row_count);//CLAVE Y NOMBRE PRODUCTO
                cb.ShowText(indx["clave"].ToString() + " " + indx["nombre"].ToString());

                cb.SetTextMatrix(390, row_count);//TRAZABILIDAD
                cb.ShowText(indx["lote"].ToString());

                decimal cos_sin_gto_imp = Math.Round(((txt_sis_tipo_cambio.Text != "1") ? (Convert.ToDecimal(indx["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) : Convert.ToDecimal(indx["costo"].ToString())) / Convert.ToDecimal(indx["cantidad"].ToString()), 4);

                if (lblPedimento.Text == "1")
                {
                    cb.SetTextMatrix(precio_pdf_gab(cos_sin_gto_imp.ToString("###,###,###,##0.00")), row_count);//PRECIO
                    cb.ShowText(cos_sin_gto_imp.ToString("###,###,###,##0.00"));
                }
                else
                {
                    cb.SetTextMatrix(precio_pdf_gab(Convert.ToDecimal(indx["unitario"].ToString()).ToString("###,###,###,##0.00")), row_count);//PRECIO
                    cb.ShowText(Convert.ToDecimal(indx["unitario"].ToString()).ToString("###,###,###,##0.00"));
                }
                

                decimal importe_pesos = (txt_sis_tipo_cambio.Text != "1") ? ((Convert.ToDecimal(indx["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) + Convert.ToDecimal(indx["gastoaduanal"].ToString()) + Convert.ToDecimal(indx["flete"].ToString())) : Convert.ToDecimal(indx["costo"].ToString());

                decimal imp_sin_gto_imp = ((txt_sis_tipo_cambio.Text != "1") ? (Convert.ToDecimal(indx["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) : Convert.ToDecimal(indx["costo"].ToString()));
                
                sum_importe = sum_importe + ((txt_sis_tipo_cambio.Text != "1") ? (Convert.ToDecimal(indx["costo"].ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) : Convert.ToDecimal(indx["costo"].ToString()));
                sum_gtos_im = sum_gtos_im + (Convert.ToDecimal(indx["gastoaduanal"].ToString()) + Convert.ToDecimal(indx["flete"].ToString()));

                if (lblPedimento.Text == "1")
                {
                    cb.SetTextMatrix(importe_pdf_gab(imp_sin_gto_imp.ToString("###,###,###,##0.00")), row_count);//IMPORTE
                    cb.ShowText(imp_sin_gto_imp.ToString("###,###,###,##0.00"));
                }
                else
                {
                    cb.SetTextMatrix(importe_pdf_gab(importe_pesos.ToString("###,###,###,##0.00")), row_count);//IMPORTE
                    cb.ShowText(importe_pesos.ToString("###,###,###,##0.00"));
                }
                

                row_count = row_count - 7;
            }
            //foreach (DataGridViewRow rw in dtgDetalle.Rows)
            //{
            //    if (rw.Cells["clave"].Value == null)
            //        continue;

            //    cb.SetTextMatrix(cantidad_pdf_gab(Convert.ToDecimal(rw.Cells["cantidad"].Value.ToString()).ToString("###,###,###,##0.00")), row_count);//CANTIDAD
            //    cb.ShowText(Convert.ToDecimal(rw.Cells["cantidad"].Value.ToString()).ToString("###,###,###,##0.00"));

            //    //CONSULTAR MEDIDA
            //    thisConnection.Open();
            //    cmnd1 = thisConnection.CreateCommand();
            //    cmnd1.CommandText = "SELECT um_clave FROM tb_cat_empaques WHERE emp_clave = '" + rw.Cells["clave"].Value.ToString() + "'";
            //    reader1 = cmnd1.ExecuteReader();
            //    string medida = "";
            //    if (reader1.HasRows)
            //    {
            //        reader1.Read();
            //        medida = reader1["um_clave"].ToString().Trim();
            //    }
            //    reader1.Close();
            //    reader1.Dispose();
            //    cmnd1.Dispose();
            //    thisConnection.Close();

            //    cb.SetTextMatrix(110, row_count);//MEDIDA
            //    cb.ShowText(medida);

            //    cb.SetTextMatrix(143, row_count);//CLAVE Y NOMBRE PRODUCTO
            //    cb.ShowText(rw.Cells["clave"].Value.ToString() + " " + rw.Cells["nombre"].Value.ToString());

            //    cb.SetTextMatrix(390, row_count);//TRAZABILIDAD
            //    cb.ShowText(rw.Cells["lote"].Value.ToString());

            //    cb.SetTextMatrix(precio_pdf_gab(Convert.ToDecimal(rw.Cells["unitario"].Value.ToString()).ToString("###,###,###,##0.00")), row_count);//PRECIO
            //    cb.ShowText(Convert.ToDecimal(rw.Cells["unitario"].Value.ToString()).ToString("###,###,###,##0.00"));

            //    decimal importe_pesos = (txt_sis_tipo_cambio.Text != "1") ? (Convert.ToDecimal(rw.Cells["costo"].Value.ToString()) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) : Convert.ToDecimal(rw.Cells["costo"].Value.ToString());
            //    cb.SetTextMatrix(importe_pdf_gab(importe_pesos.ToString("###,###,###,##0.00")), row_count);//IMPORTE
            //    cb.ShowText(importe_pesos.ToString("###,###,###,##0.00"));

            //    row_count = row_count - 7;

            //}
            cb.SetTextMatrix(460, 150);//I.E.P.S.
            cb.ShowText("Gto. Import.");
            cb.SetTextMatrix(importe_pdf_gab(sum_gtos_im.ToString("###,###,###,##0.00")), 150);
            cb.ShowText(sum_gtos_im.ToString("###,###,###,##0.00"));
            

            cb.SetTextMatrix(460, 140);//I.E.P.S.
            cb.ShowText("I.E.P.S.");
            cb.SetTextMatrix(importe_pdf_gab(Convert.ToDecimal(txt_ent_ieps.Text).ToString("###,###,###,##0.00")), 140);
            cb.ShowText(Convert.ToDecimal(txt_ent_ieps.Text).ToString("###,###,###,##0.00"));

            cb.SetTextMatrix(460, 130);//SUBTOTAL
            cb.ShowText("Subtotal");
            if (lblPedimento.Text == "1")
            {
                cb.SetTextMatrix(importe_pdf_gab(sum_importe.ToString("###,###,###,##0.00")), 130);
                cb.ShowText(sum_importe.ToString("###,###,###,##0.00"));
            }
            else
            {
                cb.SetTextMatrix(importe_pdf_gab(Convert.ToDecimal(txt_subtotal2.Text).ToString("###,###,###,##0.00")), 130);
                cb.ShowText(Convert.ToDecimal(txt_subtotal2.Text).ToString("###,###,###,##0.00"));
            }
            

            cb.SetTextMatrix(460, 120);//SUBTOTAL
            cb.ShowText("Descuento");
            cb.SetTextMatrix(importe_pdf_gab(Convert.ToDecimal(txtDesc.Text).ToString("###,###,###,##0.00")), 120);
            cb.ShowText(Convert.ToDecimal(txtDesc.Text).ToString("###,###,###,##0.00"));

            cb.SetTextMatrix(460, 110);//I.V.A.
            cb.ShowText("I.V.A.");
            cb.SetTextMatrix(importe_pdf_gab(Convert.ToDecimal(txt_iva.Text).ToString("###,###,###,##0.00")), 110);
            cb.ShowText(Convert.ToDecimal(txt_iva.Text).ToString("###,###,###,##0.00"));

            cb.SetTextMatrix(460, 100);//Ret IVA
            cb.ShowText("Ret IVA");
            cb.SetTextMatrix(importe_pdf_gab(Convert.ToDecimal(txtRetIva.Text).ToString("###,###,###,##0.00")), 100);
            cb.ShowText(Convert.ToDecimal(txtRetIva.Text).ToString("###,###,###,##0.00"));

            cb.SetTextMatrix(460, 90);//Ret ISR
            cb.ShowText("Ret ISR");
            cb.SetTextMatrix(importe_pdf_gab(Convert.ToDecimal(txtIsr.Text).ToString("###,###,###,##0.00")), 90);
            cb.ShowText(Convert.ToDecimal(txtIsr.Text).ToString("###,###,###,##0.00"));

            cb.SetTextMatrix(460, 80);//Cedular
            cb.ShowText("Cedular");
            cb.SetTextMatrix(importe_pdf_gab(Convert.ToDecimal(txtCedular.Text).ToString("###,###,###,##0.00")), 80);
            cb.ShowText(Convert.ToDecimal(txtCedular.Text).ToString("###,###,###,##0.00"));

            cb.SetTextMatrix(460, 70);//Total
            //cb.ShowText("Total " + ((DDL_ent_procedencia.SelectedIndex == 0) ? "MN" : "USD"));
            cb.ShowText("Total MN");
            cb.SetTextMatrix(importe_pdf_gab(Convert.ToDecimal(txt_total.Text).ToString("###,###,###,##0.00")), 70);
            cb.ShowText(Convert.ToDecimal(txt_total.Text).ToString("###,###,###,##0.00"));

            if (Convert.ToDecimal(lblImporteDescuento.Text) > 0)
            {
                decimal desc_pp = Convert.ToDecimal(lblImporteDescuento.Text) * Convert.ToDecimal(txt_sis_tipo_cambio.Text);
                cb.SetTextMatrix(143, 70);//Descuento or pronto pago
                //cb.ShowText("Total " + ((DDL_ent_procedencia.SelectedIndex == 0) ? "MN" : "USD"));
                cb.ShowText("Descuento por pronto pago: ");

                cb.SetTextMatrix(260, 70);
                cb.ShowText("$" + desc_pp.ToString("###,###,###,##0.00"));
            }
            

            if (lista_oc.Count > 1)
            {
                string cadena_oc = "";
                
                foreach (string val in lista_oc)
                {
                        
                    //cadena_oc + "-" + val;
                }
                cb.SetTextMatrix(143, 50);//ordenens de compra
                cb.ShowText("Órdenes de Compra: ");
                cb.SetTextMatrix(importe_pdf_gab(Convert.ToDecimal(txt_total.Text).ToString("###,###,###,##0.00")), 70);
                cb.ShowText(Convert.ToDecimal(txt_total.Text).ToString("###,###,###,##0.00"));
            }

            cb.EndText();

            doc.Close();
            //GUARDAR ARCHIVO EN SERVIDOR


            //BORRAR ARCHIVO DE REPORTES
        }

        public float cantidad_pdf_gab(string cant)
        {
            float pos = 0;
            if (cant.Length == 12)
                pos = 42;
            if (cant.Length == 10)
                pos = 50;
            if (cant.Length == 9)
                pos = 54;
            if (cant.Length == 8)
                pos = 58;
            if (cant.Length == 6)
                pos = 67;
            if (cant.Length == 5)
                pos = 71;
            if (cant.Length == 4)
                pos = 75;
            return pos;
        }

        public float precio_pdf_gab(string cant)
        {
            float pos = 0;
            if (cant.Length == 10)
                pos = 464;
            if (cant.Length == 9)
                pos = 468;
            if (cant.Length == 8)
                pos = 472;
            if (cant.Length == 6)
                pos = 482;
            if (cant.Length == 5)
                pos = 486;
            if (cant.Length == 4)
                pos = 490;
            return pos;
        }

        public float importe_pdf_gab(string cant)
        {
            float pos = 0;
            if (cant.Length == 12)
                pos = 542;
            if (cant.Length == 10)
                pos = 550;
            if (cant.Length == 9)
                pos = 554;
            if (cant.Length == 8)
                pos = 558;
            if (cant.Length == 6)
                pos = 567;
            if (cant.Length == 5)
                pos = 572;
            if (cant.Length == 4)
                pos = 576;
            return pos;
        }

        private void btnEtiquetas_Click(object sender, EventArgs e)
        {
            if (operacion == "consulta")
            {
                etiquetas dlg = new etiquetas(txtFolio.Text, glo_var_chr_alm_clave);
                dlg.ShowDialog();
            }
        }

        public void guarda_genera_pdf_servidor()
        {
            string str2 = "";
            if (this.DDLAlmacen1.SelectedIndex == 2)
                str2 = "COM";
            if (this.DDLAlmacen1.SelectedIndex == 3)
                str2 = "GUA";
            string ruta_entrada_validacion = @"\\gabira1\Entradas\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
            string file_name = "Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + str2 + ".pdf";
            genera_pdf(file_name);

            if (System.IO.File.Exists(ruta_entrada_validacion))
            {
                System.IO.File.Delete(ruta_entrada_validacion);


                string ROUTE_PDF = "c:\\Reportes\\" + file_name;
                //string COM_GUA = "";
                //if (this.DDLAlmacen1.SelectedIndex == 2)
                //    COM_GUA = "COM";
                //if (this.DDLAlmacen1.SelectedIndex == 3)
                //    COM_GUA = "GUA";
                //string str3;
                if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
                {
                }
                else
                {
                    string ROUTE_SAVED = @"\\gabira1\Entradas\" + file_name;
                    System.IO.File.Copy(ROUTE_PDF, ROUTE_SAVED, true);
                    System.IO.File.Delete(ROUTE_PDF);
                }
                //str3 = "c:\\Reportes\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + COM_GUA + ".pdf";
                //else
                //    str3 = "\\\\gabira1\\Entradas\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + COM_GUA + ".pdf";
                //if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
                //{
                //    string str4 = "c:\\Reportes\\Ent_" + this.txtFolio.Text + "_Alm_" + this.glo_var_chr_alm_clave + "_Sub_" + COM_GUA + ".pdf";
                //    System.IO.File.Copy(str1, str4, true);
                //}
                //else
                //{
                //    System.IO.File.Copy(str1, str3, true);
                //}

                //imprmir_entrada_servidor();
            }
            else
            {
                string ROUTE_PDF = "c:\\Reportes\\" + file_name;
                if (this.thisConnection.ConnectionString.Contains("189.206.160.206"))
                {
                }
                else
                {
                    string ROUTE_SAVED = @"\\gabira1\Entradas\" + file_name;
                    System.IO.File.Copy(ROUTE_PDF, ROUTE_SAVED, true);
                    System.IO.File.Delete(ROUTE_PDF);
                }
            }
            //imprmir_entrada_servidor();
        }

        public void recalculo_isr_descuento()
        {
            List<string> lista = new List<string>();
            string oc_act = "";
            string oc_ant = "";
            foreach (DataGridViewRow gr in dtgDetalle2.Rows)
            {
                oc_act = gr.Cells["ordencompra"].Value.ToString();
                if(lista.Count == 0)
                    lista.Add(oc_act);
                else
                {
                    bool fnd = false;
                    foreach (string si in lista)
                    {
                        if (si == oc_act)
                        {
                            fnd = true;
                        }     
                    }
                    if (fnd == false)
                    {
                        lista.Add(oc_act);
                    }
                }
                oc_ant = oc_act;
            }//listado de ordenes de compra

            //crear datatable con los empaques que estan en el datagrid detalle2
            //DataTable dtivas = new DataTable();
            //dtivas.Columns.Add("num_oc", typeof(string));
            //dtivas.Columns.Add("emp_oc", typeof(string));
            //dtivas.Columns.Add("imp_oc", typeof(string));
            //dtivas.Columns.Add("iva_oc", typeof(string));
            //foreach (DataGridViewRow gr in dtgDetalle2.Rows)
            //{
            //    string imp_det_iva = gr.Cells["ordencompra"] 
            //}

            //fin crear datatable con los empaques que estan en el datagrid detalle2

            thisConnection.Open();
            decimal subt_desc = 0;
            decimal subt_desc_oc = 0;
            foreach (string st in lista)//recorrer listado de ordenes de compra
            {
                //sacar el porcentaje de descuento
                string porcentaje_desc = "";
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT ISNULL(pordesc_oc, 0) as pordesc_oc FROM tb_mstr_ordencompra WHERE numero_oc = '" + st + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    //if (reader1["pordesc_oc"].ToString().Trim() == null)
                    //{
                    //    porcentaje_desc = "0";
                    //}
                    //else 
                        if (Convert.ToDecimal(reader1["pordesc_oc"].ToString().Trim()) == 0)
                    {
                        porcentaje_desc = "0";
                    }
                    else
                    {
                        porcentaje_desc = reader1["pordesc_oc"].ToString().Trim();
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                //fin de sacar el porcentaje de descuento

                if (porcentaje_desc != "0")
                {
                    //sumatoria de importes segun orden de compra * el tipo de cambio
                    foreach (DataGridViewRow gr in dtgDetalle2.Rows)
                    {
                        if (gr.Cells["ordencompra"].Value.ToString() == st)
                        {
                            decimal ieps1 = (gr.Cells["ieps1"].Value.ToString() != "") ? Convert.ToDecimal(gr.Cells["ieps1"].Value.ToString()) : 0; 
                            subt_desc_oc = (subt_desc_oc + ((Convert.ToDecimal(gr.Cells["import"].Value.ToString()) + ieps1) * Convert.ToDecimal(txt_sis_tipo_cambio.Text)));
                            //subt_desc = (subt_desc + Convert.ToDecimal(gr.Cells["import"].Value.ToString())) * Convert.ToDecimal(txt_sis_tipo_cambio.Text);
                        }
                    }

                    if (Convert.ToDecimal(porcentaje_desc) > 0)
                    {
                        subt_desc_oc = subt_desc_oc * (Convert.ToDecimal(porcentaje_desc) / Convert.ToDecimal("100"));
                        subt_desc = subt_desc + subt_desc_oc;
                    }
                }
                subt_desc_oc = 0;
            }

            txtDesc.Text = subt_desc.ToString("###,###,##0.00");

            //RECALCULAR IVA
            //txt_iva.Text = ((Convert.ToDecimal(txt_subtotal2.Text) - subt_desc) * Convert.ToDecimal("0.16")).ToString("###,###,##0.00");
            //txt_total.Text = ((Convert.ToDecimal(txt_subtotal2.Text) - subt_desc) + Convert.ToDecimal(txt_iva.Text)).ToString("###,###,##0.00");

            
            //consulta de porcentajes de retenciones y cedular
            string iva_por = "0";
            string isr_por = "0";
            string ced_por = "0";

            decimal ret_iva = 0;
            decimal ret_isr = 0;
            decimal cedular = 0;

            decimal sum_importe = 0;
            decimal sum_imp_tot = 0;
            decimal sum_imp_tot2 = 0;
            decimal sum_imp_iva = 0;

            decimal sum_ret_iva = 0;
            decimal sum_ret_isr = 0;
            decimal sum_cedular = 0;

            decimal sum_iva_total = 0;
            decimal sum_tot_total = 0;

            decimal sum_imp_limpio = 0;
            decimal sum_imp_limpio_desc = 0;

            //for (int i = 0; i < dtgDetalle2.Rows.Count; i++)
            //{
            //    if (dtgDetalle2.Rows[i].Cells["checkiva"].Value.ToString() == "1")
            //    {
            //        sum_iva_total = sum_iva_total + Convert.ToDecimal(dtgDetalle2.Rows[i].Cells["iva_det"].Value.ToString());
            //    }
            //}

            foreach (DataGridViewRow gr in dtgDetalle2.Rows)
            //foreach (DataRow gr in dtdetalles.Rows)
            {
                string porcentaje_desc = "";
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT ISNULL(pordesc_oc, 0) as pordesc_oc FROM tb_mstr_ordencompra WHERE numero_oc = '" + gr.Cells["ordencompra"].Value.ToString() + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    //if (reader1["pordesc_oc"].ToString().Trim() == null)
                    //{
                    //    porcentaje_desc = "0";
                    //}
                    //else 
                    if (Convert.ToDecimal(reader1["pordesc_oc"].ToString().Trim()) == 0)
                    {
                        porcentaje_desc = "0";
                    }
                    else
                    {
                        porcentaje_desc = reader1["pordesc_oc"].ToString().Trim();
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                //var vol_x = gr.Cells[""].Value;
                if (gr.Cells["checkiva"].Value.ToString() == "1")
                {
                    if (Convert.ToDecimal(porcentaje_desc) > 0)
                    {
                        decimal A = Convert.ToDecimal(gr.Cells["iva_det"].Value.ToString());
                        decimal B = Convert.ToDecimal(gr.Cells["iva_det"].Value.ToString()) * Math.Round(Convert.ToDecimal(porcentaje_desc) / 100, 4);
                        sum_iva_total = sum_iva_total + (A - B);
                        porcentaje_desc = "";
                    }
                    else
                    {
                        sum_iva_total = sum_iva_total + Convert.ToDecimal(gr.Cells["iva_det"].Value.ToString());
                        porcentaje_desc = "";
                    }
                    
                }
            }

            foreach (string st in lista)//recorrer listado de ordenes de compra
            {
                //sacar porcentajes de cada orden de compra 
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT retiva_oc, retiva_por, retisr_oc, retisr_por, cedular_oc, cedula_por FROM tb_oc_cfdi WHERE numero_oc = '" + st + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    iva_por = reader1["retiva_por"].ToString();
                    isr_por = reader1["retisr_por"].ToString();
                    ced_por = reader1["cedula_por"].ToString();
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                //cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = "SELECT  FROM tb_det_ordenescompra WHERE numero_oc = '" + st + "'";
                //reader1 = cmnd1.ExecuteReader();
                //if (reader1.HasRows)
                //{
                //    reader1.Read();
                //    sum_iva_total = sum_iva_total + Convert.ToDecimal(reader1["suma_iva"].ToString());
                //}
                //reader1.Close();
                //reader1.Dispose();
                //cmnd1.Dispose();
                


                string porcentaje_desc = "";
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT ISNULL(pordesc_oc, 0) as pordesc_oc FROM tb_mstr_ordencompra WHERE numero_oc = '" + st + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    //if (reader1["pordesc_oc"].ToString().Trim() == null)
                    //{
                    //    porcentaje_desc = "0";
                    //}
                    //else 
                    if (Convert.ToDecimal(reader1["pordesc_oc"].ToString().Trim()) == 0)
                    {
                        porcentaje_desc = "0";
                    }
                    else
                    {
                        porcentaje_desc = reader1["pordesc_oc"].ToString().Trim();
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                //sumatoria de importes en base a orden de compra
                foreach (DataGridViewRow gr in dtgDetalle2.Rows)
                {
                    if (gr.Cells["ordencompra"].Value.ToString() == st)
                    {
                        decimal ieps1 = (gr.Cells["ieps1"].Value.ToString() != "") ? Convert.ToDecimal(gr.Cells["ieps1"].Value.ToString()) : 0;
                        sum_importe = sum_importe + ((Convert.ToDecimal(gr.Cells["import"].Value.ToString()) + ieps1) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                        sum_imp_tot2 = sum_imp_tot2 + ((Convert.ToDecimal(gr.Cells["import"].Value.ToString()) + ieps1) * Convert.ToDecimal(txt_sis_tipo_cambio.Text));
                        sum_imp_limpio = sum_imp_limpio + Convert.ToDecimal(gr.Cells["import"].Value.ToString());
                    }                    
                }

                //si tiene descuento hace calculo
                decimal porcentaje_imp = 0;
                if (Convert.ToDecimal(porcentaje_desc) > 0)
                {
                    sum_importe = sum_importe - (sum_importe * (Convert.ToDecimal(porcentaje_desc) / Convert.ToDecimal("100")));
                    porcentaje_imp = (sum_imp_tot2 * (Convert.ToDecimal(porcentaje_desc) / Convert.ToDecimal("100")));
                }

                ////iva
                //if (Convert.ToDecimal(txt_ent_iva.Text) > 0)
                //{
                //    sum_imp_iva = sum_importe * (Convert.ToDecimal("0.16"));
                //    sum_iva_total = sum_iva_total + sum_imp_iva;
                //}
                //else
                //{
                //    //sum_imp_iva = 0;
                //}
                
                //importe 
                sum_imp_tot = sum_importe;// +sum_imp_iva;

                //sum_imp_limpio_desc = sum_imp_limpio - porcentaje_imp;
                sum_imp_limpio = sum_imp_limpio - porcentaje_imp;

                if (Convert.ToDecimal(iva_por) > 0)
                {
                    ret_iva = sum_imp_limpio * (Convert.ToDecimal(iva_por) / Convert.ToDecimal("100"));//ret_iva = sum_imp_tot * (Convert.ToDecimal(iva_por) / Convert.ToDecimal("100"));
                    sum_ret_iva = sum_ret_iva + ret_iva;//txtRetIva.Text = ret_iva.ToString("###,###,##0.00");
                }

                if (Convert.ToDecimal(isr_por) > 0)
                {
                    ret_isr = sum_imp_limpio * (Convert.ToDecimal(isr_por) / Convert.ToDecimal("100")); //ret_isr = sum_imp_tot * (Convert.ToDecimal(isr_por) / Convert.ToDecimal("100"));
                    sum_ret_isr = sum_ret_isr + ret_isr; //txtIsr.Text = ret_isr.ToString("###,###,##0.00");
                }

                if (Convert.ToDecimal(ced_por) > 0)
                {
                    cedular = sum_imp_limpio * (Convert.ToDecimal(ced_por) / Convert.ToDecimal("100")); //cedular = sum_imp_tot * (Convert.ToDecimal(ced_por) / Convert.ToDecimal("100"));
                    sum_cedular = sum_cedular + cedular; //txtCedular.Text = cedular.ToString("###,###,##0.00");
                }

                //sum_imp_tot = (sum_imp_tot2 + sum_iva_total) - (sum_ret_iva + sum_ret_isr + sum_cedular);
                sum_imp_tot = sum_imp_tot2 - (sum_ret_iva + sum_ret_isr + sum_cedular);

                sum_tot_total = sum_imp_tot - porcentaje_imp; // sum_tot_total + sum_imp_tot;

                sum_importe = 0;
                sum_imp_iva = 0;
                //sum_imp_tot = 0;
                ret_iva = 0;
                ret_isr = 0;
                cedular = 0;
            }



            sum_tot_total = sum_imp_tot + (sum_iva_total * Convert.ToDecimal(txt_sis_tipo_cambio.Text)) - subt_desc;

            decimal g_a = Convert.ToDecimal(txt_gto_aduanal.Text);
            decimal f_n = Convert.ToDecimal(txtFlete.Text);
            decimal f_i = Convert.ToDecimal(txtFleteImp.Text);

            sum_tot_total = sum_tot_total + (g_a + f_n + f_i);

            thisConnection.Close();
            //totales

            //30/04/2025 SE QUITO ESTE PORQUE EL SUBTOTAL YA INCLUYE EL IEPS
            //if (Convert.ToDecimal(txt_ent_ieps.Text) > 0)
            //    sum_tot_total = sum_tot_total + Convert.ToDecimal(txt_ent_ieps.Text);
            // FIN 30/04/2025

            txt_iva.Text = (sum_iva_total * Convert.ToDecimal(txt_sis_tipo_cambio.Text)).ToString("###,###,##0.00");
            txt_total.Text = sum_tot_total.ToString("###,###,##0.00");

            txtRetIva.Text = sum_ret_iva.ToString("###,###,##0.00");
            txtIsr.Text = sum_ret_isr.ToString("###,###,##0.00");
            txtCedular.Text = sum_cedular.ToString("###,###,##0.00");
        }

        private void btnFactExtra_Click(object sender, EventArgs e)
        {
            if (operacion == "consulta")
            {
                string e_n_t = txtFolio.Text.TrimStart('0');
                string conteo = "";
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT count(numero_oc) AS conteo FROM tb_det_oc_facturas WHERE numero_oc = '" + e_n_t + "' AND estatus = 'A' AND tipo = '" + DDLAlmacen1.Text.ToUpper() + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    conteo = reader1["conteo"].ToString().Trim();
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();

                if (Convert.ToInt32(conteo) >= 1 && (Utilerias.Class1.Usu_login.Trim() == "PEDRO" || Utilerias.Class1.Usu_login.Trim() == "N"))
                {
                    factura_extra dlg = new factura_extra(txtFolio.Text, glo_var_chr_alm_clave, DDLAlmacen1.SelectedItem.ToString(), txt_prov_clave.Text, DDLProveedor.SelectedItem.ToString(), Convert.ToDateTime(dtpFecha.Text).ToShortDateString());
                    dlg.ShowDialog();
                }
            }
        }

        private void btnFacturasExtra_Click(object sender, EventArgs e)
        {
            if (operacion == "consulta")
            {
                string e_n_t = txtFolio.Text.TrimStart('0');
                string conteo = "";
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT count(numero_oc) AS conteo FROM tb_det_oc_facturas WHERE numero_oc = '" + e_n_t + "' AND estatus = 'A' AND tipo = '" + DDLAlmacen1.Text.ToUpper() + "'";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    conteo = reader1["conteo"].ToString().Trim();
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();

                if (Convert.ToInt32(conteo) >= 1 && (Utilerias.Class1.Usu_login.Trim() == "PEDRO" || Utilerias.Class1.Usu_login.Trim() == "N"))
                {
                    fact_extras dlg = new fact_extras(txtFolio.Text, glo_var_chr_alm_clave, DDLAlmacen1.SelectedItem.ToString(), txt_prov_clave.Text, Convert.ToDateTime(dtpFecha.Text).ToShortDateString(), "0");
                    dlg.ShowDialog();
                }
            }
        }
        
        public string recalculo_costos(string clave)
        {
            string mcod = "";
            decimal minvi = 0;
            decimal mcosp = 0;

            string mrec = "";
            string mtip = "";
            string malm = "";
            decimal mcant = 0;

            decimal A = 0, B = 0, C = 0, D = 0, C2 = 0;

            DataTable tbhist = new DataTable();
            tbhist.Columns.Add("hrp_recibo", typeof(string));
            tbhist.Columns.Add("hrp_fecha", typeof(string));
            tbhist.Columns.Add("emp_clave", typeof(string));
            tbhist.Columns.Add("hrp_cantidad", typeof(string));
            tbhist.Columns.Add("hrp_tipo_recepcion", typeof(string));
            tbhist.Columns.Add("alm_clave", typeof(string));
            tbhist.Columns.Add("hrp_situacion", typeof(string));
            tbhist.Columns.Add("hrp_costo", typeof(string));
            tbhist.Columns.Add("hrp_regimen", typeof(string));
            tbhist.Columns.Add("hrp_procedencia", typeof(string));
            tbhist.Columns.Add("hrp_costo_prom", typeof(string));
            tbhist.Columns.Add("hrp_tipo", typeof(string));

            DataTable tbinv = new DataTable();
            tbinv.Columns.Add("emp_clave", typeof(string));
            tbinv.Columns.Add("invemp_inicial", typeof(string));
            tbinv.Columns.Add("invemp_salidas", typeof(string));
            tbinv.Columns.Add("invemp_entradas", typeof(string));
            tbinv.Columns.Add("alm_clave", typeof(string));
            tbinv.Columns.Add("invemp_costopro", typeof(string));

            //cat_empaques ya esa cargado en dtempaques

            DataTable tbent = new DataTable();
            tbent.Columns.Add("ent_folio", typeof(string));
            tbent.Columns.Add("emp_clave", typeof(string));
            tbent.Columns.Add("alm_clave", typeof(string));
            tbent.Columns.Add("entd_cantidad", typeof(string));
            tbent.Columns.Add("entd_costo", typeof(string));
            tbent.Columns.Add("ent_tipo", typeof(string));

            DataTable tbsal = new DataTable();
            tbsal.Columns.Add("sal_folio", typeof(string));
            tbsal.Columns.Add("emp_clave", typeof(string));
            tbsal.Columns.Add("sald_cantidad", typeof(string));
            tbsal.Columns.Add("sald_costo", typeof(string));
            tbsal.Columns.Add("alm_clave", typeof(string));
            tbsal.Columns.Add("sal_tipo", typeof(string));

            DataTable tbinv2 = new DataTable();
            tbinv2.Columns.Add("emp_clave", typeof(string));
            tbinv2.Columns.Add("invemp_inicial", typeof(string));
            tbinv2.Columns.Add("invemp_salidas", typeof(string));
            tbinv2.Columns.Add("invemp_entradas", typeof(string));
            tbinv2.Columns.Add("alm_clave", typeof(string));
            tbinv2.Columns.Add("invemp_costopro", typeof(string));

            //DateTime fecha = new DateTime();
            string fech1 = DateTime.Now.ToShortDateString();//Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();

            string fecha_inicio = "";
            string fecha_fin = "";

            //fecha_inicio = fecha.AddMonths(-1).ToShortDateString();
            fecha_fin = fech1;//fecha.ToShortDateString();//fecha.AddDays(-1).ToShortDateString();

            //string f_1 = "";
            //if (Convert.ToDateTime(dtpFecha1.Text).Year < Convert.ToDateTime(dtpFecha2.Text).Year)
            //{
            //    f_1 = "31/12/" + Convert.ToDateTime(dtpFecha1.Text).Year.ToString();
            //    fecha_fin = f_1;
            //}
            //else
            //{
            //    if (Convert.ToDateTime(dtpFecha1.Text).Month < Convert.ToDateTime(dtpFecha2.Text).Month)
            //    {
            //        string d_1 = DateTime.DaysInMonth(Convert.ToDateTime(dtpFecha1.Text).Year, Convert.ToDateTime(dtpFecha1.Text).Month).ToString();
            //        f_1 = d_1 + "/" + Convert.ToDateTime(dtpFecha1.Text).Month.ToString().PadLeft(2, '0') + "/" + Convert.ToDateTime(dtpFecha1.Text).Year.ToString();
            //        fecha_fin = f_1;
            //    }
            //}

            string query = "";
            string glo_var_chr_alm_clave = "";
            try
            {
                thisConnection.Open();

                //SACAR FECHA ULTIMO CIERRE
                string var_date_fecha = "";
                cmnd1 = thisConnection.CreateCommand();
                query = "SELECT TOP 1 invemp_fecha FROM tb_mstr_inventario_emp WHERE invemp_fecha <= '" + DateTime.Now.ToShortDateString() + "' ORDER BY invemp_fecha DESC";
                cmnd1.CommandText = query;
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    var_date_fecha = reader1.GetValue(0).ToString().Trim();

                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                //var_date_fecha = "01/09/2021";

                cmnd1 = thisConnection.CreateCommand();
                query = "SELECT alm_clave FROM tb_cat_empaques WHERE emp_clave = '" + clave + "'";
                cmnd1.CommandText = query;
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    reader1.Read();
                    glo_var_chr_alm_clave = reader1.GetValue(0).ToString().Trim();

                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                fecha_inicio = var_date_fecha;

                cmnd1 = thisConnection.CreateCommand();
                query = "SELECT hrp_recibo, hrp_fecha, emp_clave, hrp_cantidad, hrp_tipo_recepcion, alm_clave, hrp_situacion, hrp_costo, hrp_regimen, hrp_procedencia, hrp_costo_prom, hrp_tipo" +
                    " FROM tb_historico_recepcion WHERE hrp_estatus <> 'C' AND (hrp_fecha >= '" + Convert.ToDateTime(fecha_inicio).ToShortDateString() + "' AND hrp_fecha <= '" + fecha_fin + "')" +
                    " AND alm_clave = '" + glo_var_chr_alm_clave + "' and emp_clave = '" + clave + "'" +
                    //" and emp_clave = 'B0001'" +//and emp_clave = 'AB088'
                    " ORDER BY emp_clave, hrp_fecha, hrp_tipo_recepcion";
                cmnd1.CommandText = query;

                reader1 = cmnd1.ExecuteReader();
                DataRow rw;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rw = tbhist.NewRow();
                        rw["hrp_recibo"] = reader1.GetValue(0).ToString().Trim();
                        rw["hrp_fecha"] = reader1.GetValue(1).ToString().Trim();
                        rw["emp_clave"] = reader1.GetValue(2).ToString().Trim();
                        rw["hrp_cantidad"] = reader1.GetValue(3).ToString().Trim();
                        rw["hrp_tipo_recepcion"] = reader1.GetValue(4).ToString().Trim();
                        rw["alm_clave"] = reader1.GetValue(5).ToString().Trim();
                        rw["hrp_situacion"] = reader1.GetValue(6).ToString().Trim();
                        rw["hrp_costo"] = reader1.GetValue(7).ToString().Trim();
                        rw["hrp_regimen"] = reader1.GetValue(8).ToString().Trim();
                        rw["hrp_procedencia"] = reader1.GetValue(9).ToString().Trim();
                        rw["hrp_costo_prom"] = reader1.GetValue(10).ToString().Trim();
                        rw["hrp_tipo"] = reader1.GetValue(11).ToString().Trim();
                        tbhist.Rows.Add(rw);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                query = "SELECT emp_clave, alm_clave FROM tb_cat_empaques WHERE emp_nombre <> '' AND alm_clave = '" + glo_var_chr_alm_clave + "' and emp_clave = '" + clave + "' ORDER BY emp_clave";//
                cmnd1.CommandText = query;
                DataRow r1;
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        r1 = tbinv.NewRow();
                        r1["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                        r1["invemp_inicial"] = "0";
                        r1["invemp_salidas"] = "0";
                        r1["invemp_entradas"] = "0";
                        r1["alm_clave"] = reader1.GetValue(1).ToString().Trim();
                        r1["invemp_costopro"] = "0";
                        tbinv.Rows.Add(r1);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();


                query = "SELECT emp_clave, invemp_inicial, invemp_salidas, invemp_entradas, alm_clave, invemp_costopro FROM tb_mstr_inventario_emp " +
                     "WHERE invemp_fecha = '" + Convert.ToDateTime(fecha_inicio).ToShortDateString() + "' and emp_clave = '" + clave + "' ORDER BY emp_clave";//--and emp_clave = 'B0001'
                cmnd1.CommandText = query;
                reader1 = cmnd1.ExecuteReader();
                DataRow r2;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        r2 = tbinv2.NewRow();
                        r2["emp_clave"] = reader1.GetValue(0).ToString().Trim();//grid
                        r2["invemp_inicial"] = reader1.GetValue(1).ToString().Trim();//grid
                        r2["invemp_salidas"] = reader1.GetValue(2).ToString().Trim();//no
                        r2["invemp_entradas"] = reader1.GetValue(3).ToString().Trim();//no
                        r2["alm_clave"] = reader1.GetValue(4).ToString().Trim();//ya lo traigo
                        r2["invemp_costopro"] = reader1.GetValue(5).ToString().Trim();//grid
                        tbinv2.Rows.Add(r2);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();

                foreach (DataRow rz in tbinv.Rows)
                {
                    foreach (DataRow ry in tbinv2.Select("emp_clave = '" + rz["emp_clave"] + "'"))
                    {
                        rz["invemp_inicial"] = ry["invemp_inicial"];
                        rz["invemp_costopro"] = ry["invemp_costopro"];
                    }
                }

                thisConnection.Open();
                bool hay = false;
                //int y = 0;
                foreach (DataRow rinv in tbinv.Rows)
                {
                    mcod = rinv[0].ToString();
                    minvi = Convert.ToDecimal(rinv[1].ToString());
                    mcosp = Convert.ToDecimal(rinv[5].ToString());


                    foreach (DataRow rhis in tbhist.Select("emp_clave = '" + mcod + "'"))
                    {
                        hay = true;
                        //mesh = Convert.ToDateTime(rhis[1].ToString()).Month.ToString();
                        mrec = rhis["hrp_recibo"].ToString();
                        mtip = rhis["hrp_tipo_recepcion"].ToString();
                        malm = rhis["alm_clave"].ToString();
                        if (rhis["hrp_tipo_recepcion"].ToString() == "ENT" && rhis["hrp_situacion"].ToString() == "NOR")
                        {
                            decimal hrp_cantidad = Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            decimal hrp_costo = Convert.ToDecimal(rhis["hrp_costo"].ToString());
                            A = Math.Abs(minvi * mcosp);
                            B = hrp_cantidad * hrp_costo;
                            C = minvi + hrp_cantidad;
                            C2 = Math.Abs(minvi) + hrp_cantidad;
                            D = (C2 == 0) ? hrp_costo : ((A + B) / C2);
                            minvi = C;
                            mcosp = Math.Abs(D);
                            //Actualizacion de historico
                            //cmnd1 = thisConexion.CreateCommand();
                            //cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_existencia = '" + minvi.ToString() + "', hrp_costo_prom = '" + mcosp.ToString() + "'" +
                            //    " WHERE hrp_tipo_recepcion = '" + mtip + "' AND alm_clave = '" + malm + "' AND hrp_recibo = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            ////cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();
                        }
                        if (rhis["hrp_tipo_recepcion"].ToString() == "ENT" && rhis["hrp_situacion"].ToString() == "DEV")
                        {
                            minvi = minvi + Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            mcant = Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            //Actualizacion de historico
                            //cmnd1 = thisConnection.CreateCommand();
                            //cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_existencia = '" + minvi.ToString() + "', hrp_costo_prom = '" + mcosp.ToString() + "', hrp_costo = '" + mcosp.ToString() + "'" +
                            //    " WHERE hrp_tipo_recepcion = '" + mtip + "' AND alm_clave = '" + malm + "' AND hrp_recibo = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();

                            //decimal cto = 0;
                            //cto = mcosp * mcant;
                            //cmnd1 = thisConnection.CreateCommand();
                            //cmnd1.CommandText = "UPDATE tb_det_ent_temporal SET entd_costo = '" + cto.ToString() + "'" +
                            //    " WHERE alm_clave = '" + malm + "' AND ent_folio = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();
                        }
                        if (rhis["hrp_tipo_recepcion"].ToString() == "SAL")
                        {
                            minvi = minvi - Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            //Actualizacion de historico
                            //cmnd1 = thisConnection.CreateCommand();
                            //cmnd1.CommandText = "UPDATE tb_historico_recepcion SET hrp_existencia = '" + minvi.ToString() + "', hrp_costo_prom = '" + mcosp.ToString() + "', hrp_costo = '" + mcosp.ToString() + "'" +
                            //    " WHERE hrp_tipo_recepcion = '" + mtip + "' AND alm_clave = '" + malm + "' AND hrp_recibo = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();

                            //cmnd1 = thisConexion.CreateCommand();
                            //cmnd1.CommandText = "UPDATE tb_det_salidas_temporal SET sald_costo = '" + mcosp.ToString() + "'" +
                            //    " WHERE alm_clave = '" + malm + "' AND sal_folio = '" + mrec + "' AND emp_clave = '" + mcod + "'";
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();
                        }



                    }

                }

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "update tb_cat_empaques set emp_costo = '" + mcosp + "' WHERE emp_clave = '" + mcod + "'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();

                thisConnection.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                //Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "2.1", ex.ToString(), "SISEMP");
                //form_entradas_ocs.Data.GraphMailService.NotificarError(Environment.MachineName + " " + ex.ToString() + "QUERY: " + query);
            }
            return mcosp.ToString();
        }

        public void recalculo_existencia(string clave)
        {
            DataTable dtempaque = new DataTable();
            DataTable dtinv = new DataTable();
            DataTable dtconsal = new DataTable();
            DataTable dtconent = new DataTable();

            dtempaque.Columns.Add("emp_clave", typeof(string));
            dtempaque.Columns.Add("emp_nombre", typeof(string));
            dtempaque.Columns.Add("invemp_inicial", typeof(string));
            dtempaque.Columns.Add("tent", typeof(string));
            dtempaque.Columns.Add("tsal", typeof(string));
            dtempaque.Columns.Add("inv_actual", typeof(string));
            dtempaque.Columns.Add("costo", typeof(string));

            dtinv.Columns.Add("emp_clave", typeof(string));
            dtinv.Columns.Add("inicial", typeof(string));
            dtinv.Columns.Add("costo", typeof(string));

            dtconent.Columns.Add("emp_clave", typeof(string));
            dtconent.Columns.Add("total", typeof(string));

            dtconsal.Columns.Add("emp_clave", typeof(string));
            dtconsal.Columns.Add("total", typeof(string));

            dtempaque.Clear();
            dtinv.Clear();
            dtconent.Clear();
            dtconsal.Clear();

            thisConnection.Open();

            //CARGAR TODOS LOS EMPAQUES
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT emp_clave, emp_nombre FROM tb_cat_empaques WHERE emp_clave = '" + clave + "' ORDER BY emp_clave";
            reader1 = cmnd1.ExecuteReader();
            DataRow rw;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    rw = dtempaque.NewRow();
                    rw["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                    rw["emp_nombre"] = reader1.GetValue(1).ToString().Trim();
                    rw["invemp_inicial"] = "0";
                    rw["tent"] = "0";
                    rw["tsal"] = "0";
                    rw["inv_actual"] = "0";
                    dtempaque.Rows.Add(rw);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            string var_date_fecha = "";
            //SACAR FECHA ULTIMO CIERRE
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT TOP 1 invemp_fecha FROM tb_mstr_inventario_emp WHERE invemp_fecha <= '" + DateTime.Now.ToShortDateString() + "' ORDER BY invemp_fecha DESC";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                var_date_fecha = reader1.GetValue(0).ToString().Trim();

            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            //var_date_fecha = "01/09/2021";

            string actual = "0";

            cmnd1.CommandText = "SELECT I.emp_clave, I.invemp_inicial FROM tb_mstr_inventario_emp I" +
                " WHERE I.invemp_fecha = '" + Convert.ToDateTime(var_date_fecha).ToShortDateString() + "' AND emp_clave = '" + clave + "' ORDER BY I.emp_clave";
            reader1 = cmnd1.ExecuteReader();
            DataRow dtr;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    dtr = dtinv.NewRow();
                    dtr["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                    dtr["inicial"] = reader1.GetValue(1).ToString().Trim();
                    dtinv.Rows.Add(dtr);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();



            foreach (DataRow r1 in dtempaque.Rows)
            {
                foreach (DataRow r2 in dtinv.Select("emp_clave = '" + r1["emp_clave"] + "'"))
                {
                    r1["invemp_inicial"] = r2["inicial"].ToString();
                }
            }

             string f_1 = DateTime.Now.ToShortDateString();
            //if (Convert.ToDateTime(var_date_fecha).Year < Convert.ToDateTime(dtpFecha2.Text).Year)
            //{
            //    f_1 = "31/12/" + Convert.ToDateTime(var_date_fecha).Year.ToString();
            //}
            //else
            //{
            //    if (Convert.ToDateTime(var_date_fecha).Month < Convert.ToDateTime(dtpFecha2.Text).Month)
            //    {
            //        string d_1 = DateTime.DaysInMonth(Convert.ToDateTime(var_date_fecha).Year, Convert.ToDateTime(var_date_fecha).Month).ToString();
            //        f_1 = d_1 + "/" + Convert.ToDateTime(var_date_fecha).Month.ToString().PadLeft(2, '0') + "/" + Convert.ToDateTime(var_date_fecha).Year.ToString();
            //    }
            //}


            cmnd1 = thisConnection.CreateCommand();
            //cmnd1.CommandText = "select emp_clave, sum(hrp_cantidad) FROM tb_historico_recepcion " +
            //    "WHERE hrp_fecha >= '" + Convert.ToDateTime(fecha).ToShortDateString() + "' and hrp_tipo_recepcion = 'SAL' and hrp_estatus <> 'C' " +
            //    "group by emp_clave " +
            //    "order by emp_clave";
            //if (f_1 != "")
            //{
            cmnd1.CommandText = "select emp_clave, sum(hrp_cantidad) FROM tb_historico_recepcion " +
            "WHERE hrp_fecha >= '" + Convert.ToDateTime(var_date_fecha).ToShortDateString() + "' and hrp_fecha <= '" + Convert.ToDateTime(f_1).ToShortDateString() + "' " +
            " and hrp_tipo_recepcion = 'SAL' and hrp_estatus <> 'C' AND emp_clave = '" + clave + "' " +
            "group by emp_clave " +
            "order by emp_clave";
            //}
            //else
            //{
            //    cmnd1.CommandText = "select emp_clave, sum(hrp_cantidad) FROM tb_historico_recepcion " +
            //    "WHERE hrp_fecha >= '" + Convert.ToDateTime(var_date_fecha).ToShortDateString() + "' and hrp_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "' " +
            //    " and hrp_tipo_recepcion = 'SAL' and hrp_estatus <> 'C' AND emp_clave = '" + clave + "' " +
            //    "group by emp_clave " +
            //    "order by emp_clave";
            //}

            reader1 = cmnd1.ExecuteReader();
            DataRow r_1;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    r_1 = dtconsal.NewRow();
                    r_1["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                    r_1["total"] = reader1.GetValue(1).ToString().Trim();
                    dtconsal.Rows.Add(r_1);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            cmnd1 = thisConnection.CreateCommand();
            //cmnd1.CommandText = "select emp_clave, sum(hrp_cantidad) FROM tb_historico_recepcion " +
            //    "WHERE hrp_fecha >= '" + Convert.ToDateTime(fecha).ToShortDateString() + "' and hrp_tipo_recepcion = 'ENT' AND hrp_situacion in ('NOR','DEV') and hrp_estatus <> 'C' " +
            //    "group by emp_clave " +
            //    "order by emp_clave";
            //if (f_1 != "")
            //{
            cmnd1.CommandText = "select emp_clave, sum(hrp_cantidad) FROM tb_historico_recepcion " +
            "WHERE hrp_fecha >= '" + Convert.ToDateTime(var_date_fecha).ToShortDateString() + "' and hrp_fecha <= '" + Convert.ToDateTime(f_1).ToShortDateString() + "' " +
            "and hrp_tipo_recepcion = 'ENT' AND hrp_situacion in ('NOR','DEV', 'AJT') and hrp_estatus <> 'C' AND emp_clave = '" + clave + "' " +
            "group by emp_clave " +
            "order by emp_clave";
            //}
            //else
            //{
            //    cmnd1.CommandText = "select emp_clave, sum(hrp_cantidad) FROM tb_historico_recepcion " +
            //    "WHERE hrp_fecha >= '" + Convert.ToDateTime(var_date_fecha).ToShortDateString() + "' and hrp_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "' " +
            //    "and hrp_tipo_recepcion = 'ENT' AND hrp_situacion in ('NOR','DEV', 'AJT') and hrp_estatus <> 'C' AND emp_clave = '" + clave + "' " +
            //    "group by emp_clave " +
            //    "order by emp_clave";
            //}

            reader1 = cmnd1.ExecuteReader();
            DataRow r_2;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    r_2 = dtconent.NewRow();
                    r_2["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                    r_2["total"] = reader1.GetValue(1).ToString().Trim();
                    dtconent.Rows.Add(r_2);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();


            foreach (DataRow rxx in dtempaque.Rows)
            {
                foreach (DataRow rx1 in dtconent.Select("emp_clave = '" + rxx["emp_clave"].ToString() + "'"))
                {
                    rxx["tent"] = rx1["total"].ToString();
                    rxx["inv_actual"] = ((Convert.ToDecimal(rxx["invemp_inicial"].ToString()) + Convert.ToDecimal(rxx["tent"].ToString())) - Convert.ToDecimal(rxx["tsal"].ToString())).ToString();
                }
                foreach (DataRow rx2 in dtconsal.Select("emp_clave = '" + rxx["emp_clave"].ToString() + "'"))
                {
                    rxx["tsal"] = rx2["total"].ToString();
                    rxx["inv_actual"] = ((Convert.ToDecimal(rxx["invemp_inicial"].ToString()) + Convert.ToDecimal(rxx["tent"].ToString())) - Convert.ToDecimal(rxx["tsal"].ToString())).ToString();
                }
            }

            foreach (DataRow rxx in dtempaque.Rows)
            {
                rxx["inv_actual"] = ((Convert.ToDecimal(rxx["invemp_inicial"].ToString()) + Convert.ToDecimal(rxx["tent"].ToString())) - Convert.ToDecimal(rxx["tsal"].ToString())).ToString();
            }



            actual = dtempaque.Rows[0]["inv_actual"].ToString();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "update tb_cat_empaques set emp_existencia = '" + actual + "' WHERE emp_clave = '" + clave + "'";
            cmnd1.ExecuteNonQuery();
            cmnd1.Dispose();


            thisConnection.Close();

            //return actual;
        }

        public void modificar_ultima_entrada(string clave)
        {
            string fe = "";
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT MAX(hrp_fecha) AS fecha FROM tb_historico_recepcion WHERE emp_clave = '" + clave + "' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR' " +
                "AND hrp_estatus <> 'C' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
            reader1 = cmnd1.ExecuteReader();
            bool fnd = false;
            if (reader1.HasRows)
            {
                reader1.Read();
                if (reader1["fecha"].ToString() != null)
                {
                    if (reader1["fecha"].ToString().Trim() != "")
                    {
                        fe = Convert.ToDateTime(reader1["fecha"].ToString().Trim()).ToShortDateString();
                        fnd = true;
                    }

                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            if (fnd == true)
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_cat_empaques SET emp_entrada_ultima = '" + fe + "' WHERE emp_clave = '" + clave + "'";
                cmnd1.ExecuteNonQuery();
                thisConnection.Close();
            }
            else
                thisConnection.Close();

        }

        public void modificar_ultima_salida(string clave)
        {
            string fe = "";
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT MAX(hrp_fecha) AS fecha FROM tb_historico_recepcion WHERE emp_clave = '" + clave + "' AND hrp_tipo_recepcion = 'SAL' " +
                "AND hrp_estatus <> 'C' AND alm_clave = '" + glo_var_chr_alm_clave + "'";
            reader1 = cmnd1.ExecuteReader();

            bool fnd = false;

            if (reader1.HasRows)
            {
                reader1.Read();
                if (reader1["fecha"].ToString() != null)
                {
                    if (reader1["fecha"].ToString().Trim() != "")
                    {
                        fe = Convert.ToDateTime(reader1["fecha"].ToString().Trim()).ToShortDateString();
                        fnd = true;
                    }

                }

            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            if (fnd == true)
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_cat_empaques SET emp_salida_ultima = '" + fe + "' WHERE emp_clave = '" + clave + "'";
                cmnd1.ExecuteNonQuery();
                thisConnection.Close();
            }
            else
                thisConnection.Close();
        }

        public bool busca_orden_cerrada(string oc)
        {
            bool fnd = false;
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT cerrada FROM tb_mstr_ordencompra WHERE numero_oc = '" + oc + "'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                if (reader1["cerrada"].ToString().Trim() == "1")
                    fnd = true;
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();
            return fnd;
        }

        public void verificar_oc_datos_sin_clave()
        {
            try
            {
                thisConnection.Open();
                string o_c_ant = "";
                string o_c_act = "";
                decimal tot_oc = 0;
                decimal tot_fa = 0;
                SqlDataAdapter adap;
                DataSet set = new DataSet();
                List<string> sin_reg = new List<string>();

                //verificar ocs con registros sin clave
                foreach (DataRow rw in dtdetalles.Rows)
                {
                    o_c_act = rw["orden_compra"].ToString().Trim();
                    if (o_c_act != o_c_ant)
                    {
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "SELECT numero_oc FROM tb_det_ordenescompra WHERE prod_clave = '' AND numero_oc = '" + o_c_act + "'";
                        reader1 = cmnd1.ExecuteReader();
                        if (reader1.HasRows)
                        {
                            sin_reg.Add(o_c_act);
                        }
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();
                    }
                    o_c_ant = o_c_act;
                }

                

                o_c_ant = "";
                o_c_act = "";
                decimal sum_ocs = 0;
                decimal sum_ent = 0;
                foreach (string d in sin_reg)
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT total_oc FROM tb_mstr_ordencompra WHERE numero_oc = '" + d + "'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        sum_ocs = sum_ocs + Convert.ToDecimal(reader1["total_oc"].ToString());
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();

                    List<string> ent = new List<string>();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT distinct(ent_folio) FROM tb_det_entrada_oc WHERE numero_oc = '" + d + "' AND estatus = 'A'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            ent.Add(reader1["ent_folio"].ToString());
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();

                    string t = "";
                    switch (glo_var_chr_alm_clave)
                    {
                        case "01":
                            t = "MATERIALES";
                            break;
                        case "02":
                            t = "EMPAQUES";
                            break;
                        case "03":
                            t = "COMEDOR";
                            break;
                        case "04":
                            t = "ENFERMERIA";
                            break;
                        case "05":
                            t = "PAPELERIA";
                            break;
                        case "07":
                            t = "DENTAL";
                            break;
                        case "08":
                            t = "COMBUSTIBLE";
                            break;
                        case "09":
                            t = "SISTEMAS";
                            break;
                        case "10":
                            t = "GUARDERIAALM";
                            break;
                        case "11":
                            t = "GUARDERIACOM";
                            break;
                        case "12":
                            t = "EVENTOS";
                            break;
                        case "13":
                            t = "MOTORES";
                            break;
                    }
                    decimal sum_imp_fact = 0;
                    foreach (string r in ent)
                    {

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "SELECT monto FROM tb_det_oc_facturas WHERE numero_oc = '" + r + "' AND estatus = 'A' AND tipo = '" + t + "'";
                        reader1 = cmnd1.ExecuteReader();
                        if (reader1.HasRows)
                        {
                            while (reader1.Read())
                            {
                                sum_imp_fact = sum_imp_fact + Convert.ToDecimal(reader1["monto"].ToString());
                            }
                        }
                    }


                    Int32 abs_oc = Convert.ToInt32(sum_ocs);
                    Int32 abs_en = Convert.ToInt32(sum_imp_fact);

                    if ((abs_oc - abs_en) == 0)
                    {
                        //MessageBox.Show("Total ordenes: " + Math.Round(sum_ocs, 2).ToString() + "\nTotal factura: " + Math.Round(sum_imp_fact, 2));
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET surtido_oc = cantidad_oc WHERE prod_clave = '' AND numero_oc = '" + d + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "INSERT INTO tb_registro_movimientos(fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, " +
                            "detalle, sistema) VALUES('" + DateTime.Now.ToString() + "', '" + Environment.MachineName + "', '" + Utilerias.Class1.Usu_login.Trim() + "', " +
                            "'SPSC', '2.5', '" + d + "', 'SURTIDO DE PRODUCTO SIN CLAVE', 'SISEMP')";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                    }
                    thisConnection.Close();

                        
                }

                if (sin_reg.Count == 0)
                    thisConnection.Close();
            }
            catch (Exception ex)
            {
                thisConnection.Close();
            }
            
            ////foreach (DataRow rw in dtdetalles.Rows)
            ////{
            ////    o_c_act = rw["orden_compra"].ToString().Trim();
            ////    if (o_c_act != o_c_ant)
            ////    {
            ////        cmnd1 = thisConnection.CreateCommand();
            ////        cmnd1.CommandText = "SELECT total_oc FROM tb_mstr_ordencompra WHERE numero_oc = '" + o_c_act + "'";
            ////        reader1 = cmnd1.ExecuteReader();
            ////        if (reader1.HasRows)
            ////        {
            ////            reader1.Read();
            ////            sum_ocs = sum_ocs + Convert.ToDecimal(reader1["total_oc"].ToString());
            ////        }
            ////        reader1.Close();
            ////        reader1.Dispose();
            ////        cmnd1.Dispose();
            ////    }
            ////    o_c_ant = o_c_act;
            ////}

            //o_c_ant = "";
            //o_c_act = "";
            //List<string> ent = new List<string>();
            //foreach (string d in sin_reg)
            //{
            //    cmnd1 = thisConnection.CreateCommand();
            //    cmnd1.CommandText = "SELECT distinct(ent_folio) FROM tb_det_entrada_oc WHERE numero_oc = '" + d + "' AND estatus = 'A'";
            //    reader1 = cmnd1.ExecuteReader();
            //    if (reader1.HasRows)
            //    {
            //        while (reader1.Read())
            //        {
            //            ent.Add(reader1["ent_folio"].ToString());
            //        }
            //    }
            //    reader1.Close();
            //    reader1.Dispose();
            //    cmnd1.Dispose();
            //}
            ////foreach (DataRow rw in dtdetalles.Rows)
            ////{
            ////    o_c_act = rw["orden_compra"].ToString().Trim();
            ////    if (o_c_act != o_c_ant)
            ////    {
            ////        cmnd1 = thisConnection.CreateCommand();
            ////        cmnd1.CommandText = "SELECT distinct(ent_folio) FROM tb_det_entrada_oc WHERE numero_oc = '" + o_c_act + "' AND estatus = 'A'";
            ////        reader1 = cmnd1.ExecuteReader();
            ////        if (reader1.HasRows)
            ////        {
            ////            while (reader1.Read())
            ////            {
            ////                ent.Add(reader1["ent_folio"].ToString());
            ////            }
            ////        }
            ////        reader1.Close();
            ////        reader1.Dispose();
            ////        cmnd1.Dispose();
            ////    }
            ////    o_c_ant = o_c_act;
            ////}

            
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //verificar_oc_datos_sin_clave();
            CargarPedimentoFactura();
        }

        private void chkDescuento_Click(object sender, EventArgs e)
        {
            if (dtgDetalle2.Rows.Count == 0)
            {
                MessageBox.Show("Debe cargar información a la entrada para esta operación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                chkDescuento.Checked = false;
                return;
            }

            

            if (chkDescuento.Checked == true)
            {
                if (MessageBox.Show("Se recalcularan los costos, ¿Desea proseguir?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                {
                    chkDescuento.Checked = false;
                    return;
                }

                foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
                {
                    decimal num1 = Convert.ToDecimal(row["importe_orig"]);
                    decimal num2 = Convert.ToDecimal(this.lblPorcentajeDesc.Text);
                    decimal num3 = new Decimal(100);
                    Decimal num = num1 * (num2 / num3);
                    row["importe"] = (object)(Convert.ToDecimal(row["importe_orig"]) - num).ToString();
                    row["costo"] = (object)(Convert.ToDecimal(row["importe"]) / Convert.ToDecimal(row["cantidad"])).ToString();
                    num = Convert.ToDecimal(row["importe"]) * (Convert.ToDecimal(16) / new Decimal(100));
                    row["iva"] = (object)num.ToString();

                    string emp_clv = row["emp_clave"].ToString().Trim();
                    string ord_com = row["orden_compra"].ToString().Trim();
                    foreach (DataGridViewRow grow in dtgDetalle2.Rows)
                    { 
                        string e_c = grow.Cells["clave1"].Value.ToString();
                        string o_c = grow.Cells["ordencompra"].Value.ToString();
                        if (emp_clv == e_c && ord_com == o_c)
                        {
                            grow.Cells["import"].Value = (object)(Convert.ToDecimal(row["importe"])).ToString("###,###,##0.000");
                            grow.Cells["cost"].Value = (object)(Convert.ToDecimal(row["costo"])).ToString("###,###,##0.000");
                            grow.Cells["iva_det"].Value = (object)(Convert.ToDecimal(row["iva"]).ToString("###,###,##0.000"));
                            //dtgDetalle2.CurrentRow.Cells["import"].Value = (object)(Convert.ToDecimal(row["importe"])).ToString("###,###,##0.000");
                            //dtgDetalle2.CurrentRow.Cells["cost"].Value = (object)(Convert.ToDecimal(row["costo"])).ToString("###,###,##0.000");
                            //dtgDetalle2.CurrentRow.Cells["iva_det"].Value = (object)(Convert.ToDecimal(row["iva"]).ToString("###,###,##0.000"));
                        }
                    }

                    
                }

                dtgDetalle.Rows.Clear();
                button2_Click(null, null);
                //Cargar datos al datagrid

                if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
                {
                    txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                if (txtFleteImp.Text != "" || Convert.ToDecimal(txtFleteImp.Text) > 0)
                {
                    txtFleteImp_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
                {
                    txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                calculatotales();

                recalculo_isr_descuento();

                var imp_desc = dtdetalles.AsEnumerable().Sum(x => Convert.ToDecimal(x["importe"]));
                var imp_orig = dtdetalles.AsEnumerable().Sum(x => Convert.ToDecimal(x["importe_orig"]));

                lblImporteDescuento.Text = Math.Round((Convert.ToDecimal(imp_orig) - Convert.ToDecimal(imp_desc)), 4).ToString();
            }

            else if (chkDescuento.Checked == false)
            {
                if (MessageBox.Show("Se recalculará a los costos originales, ¿Desea proseguir?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No)
                {
                    chkDescuento.Checked = true;
                    return;
                }
                foreach (DataRow row in (InternalDataCollectionBase)this.dtdetalles.Rows)
                {
                    //Decimal num = Convert.ToDecimal(row["importe_orig"]) * (Convert.ToDecimal(this.lblPorcentajeDesc.Text) / new Decimal(100));
                    Decimal num = 0;
                    row["importe"] = (object)(Convert.ToDecimal(row["importe_orig"])).ToString();
                    row["costo"] = (object)(Convert.ToDecimal(row["importe"]) / Convert.ToDecimal(row["cantidad"])).ToString();
                    num = Convert.ToDecimal(row["importe"]) * (Convert.ToDecimal(16) / new Decimal(100));
                    row["iva"] = (object)num.ToString();

                    string emp_clv = row["emp_clave"].ToString().Trim();
                    string ord_com = row["orden_compra"].ToString().Trim();
                    foreach (DataGridViewRow grow in dtgDetalle2.Rows)
                    {
                        string e_c = grow.Cells["clave1"].Value.ToString();
                        string o_c = grow.Cells["ordencompra"].Value.ToString();
                        if (emp_clv == e_c && ord_com == o_c)
                        {
                            grow.Cells["import"].Value = (object)(Convert.ToDecimal(row["importe"])).ToString("###,###,##0.000");
                            grow.Cells["cost"].Value = (object)(Convert.ToDecimal(row["costo"])).ToString("###,###,##0.000");
                            grow.Cells["iva_det"].Value = (object)(Convert.ToDecimal(row["iva"]).ToString("###,###,##0.000"));
                        }
                    }

                    //dtgDetalle2.CurrentRow.Cells["import"].Value = (object)(Convert.ToDecimal(row["importe"])).ToString("###,###,##0.000");
                    //dtgDetalle2.CurrentRow.Cells["cost"].Value = (object)(Convert.ToDecimal(row["costo"])).ToString("###,###,##0.000");
                    //dtgDetalle2.CurrentRow.Cells["iva_det"].Value = (object)(Convert.ToDecimal(row["iva"]).ToString("###,###,##0.000"));
                }

                dtgDetalle.Rows.Clear();
                button2_Click(null, null);
                //Cargar datos al datagrid

                if (txtFlete.Text != "" || Convert.ToDecimal(txtFlete.Text) > 0)
                {
                    txtFlete_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                if (txtFleteImp.Text != "" || Convert.ToDecimal(txtFleteImp.Text) > 0)
                {
                    txtFleteImp_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                if (txt_gto_aduanal.Text != "" || Convert.ToDecimal(txt_gto_aduanal.Text) > 0)
                {
                    txt_gto_aduanal_KeyPress(this, new KeyPressEventArgs((char)Keys.Enter));
                }
                calculatotales();

                recalculo_isr_descuento();

                var imp_desc = dtdetalles.AsEnumerable().Sum(x => Convert.ToDecimal(x["importe"]));
                var imp_orig = dtdetalles.AsEnumerable().Sum(x => Convert.ToDecimal(x["importe_orig"]));

                lblImporteDescuento.Text = Math.Round((Convert.ToDecimal(imp_orig) - Convert.ToDecimal(imp_desc)), 4).ToString();
            }
        }

        public void correo_error(string cuerpo, string subj)
        {
            Dns.GetHostEntry(Dns.GetHostName());
            MailMessage message = new MailMessage();
            message.To.Add("aescamilla@mrlucky.com.mx");
            message.Subject = subj;
            message.SubjectEncoding = Encoding.UTF8;
            message.Body = cuerpo;
            message.BodyEncoding = Encoding.UTF8;
            message.IsBodyHtml = true;
            message.From = new MailAddress("pgonzalez@mrlucky.com.mx");
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Credentials = (ICredentialsByHost)new NetworkCredential("pgonzalez", "pcancino");
            smtpClient.Port = 587;
            smtpClient.EnableSsl = true;
            smtpClient.Host = "mail1.mrlucky.com.mx";
            try
            {
                smtpClient.Send(message);
            }
            catch (Exception ex)
            {
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (lsbPediFact.Items.Count == 2)
            {
                panel4.Visible = false;
            }
            else
            {
                if (MessageBox.Show("¿Desea cerrar sin cargar los archivos?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    lsbPediFact.Items.Clear();
                    panel4.Visible = false;
                    return;
                }
                if (lsbPediFact.Items.Count < 2)
                {
                    MessageBox.Show("Debe cargar los archivos correspondientes al Pedimento y la Factura", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }

        private void btnSelecciona_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "pdf files (*.pdf)|*.pdf";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (lsbPediFact.Items.Count > 0)
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
                    foreach (string item in lsbPediFact.Items)
                    {
                        if (item.Equals(textoABuscar, StringComparison.OrdinalIgnoreCase))
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
                }
                lsbPediFact.Items.Add(openFileDialog1.FileName);
            }
        }

        private void lsbPediFact_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (lsbPediFact.SelectedItems.Count > 0)
            {
                var item = lsbPediFact.SelectedItems[0];

                DialogResult result = MessageBox.Show(
                    "¿Deseas eliminar el elemento '{" + item.ToString() + "}'?",
                    "Confirmar eliminación",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (result == DialogResult.Yes)
                {
                    lsbPediFact.Items.Remove(item);
                }
            }
        }

        public void CargarPedimentoFactura()
        {
            string path_destino = "\\\\gabira1\\Pedimentos\\";
            foreach (string url in lsbPediFact.Items)
            {
                string dest = path_destino + Path.GetFileName(url);
                _consultaController.GuardarPedimentoFactura(txtFolio.Text, glo_var_chr_alm_clave, dest, "1");

                FileInfo archivo = new FileInfo(url);
                FileInfo doc_copy = new FileInfo(@dest);
                archivo.CopyTo(@dest, true);

                _consultaController.RegistroMovimiento(Environment.MachineName.Trim(), Utilerias.Class1.Usu_login.Trim(), "PEDIMENTO", "2.5",
                    txtFolio.Text, "Carga de archivos de pedimento almacen " + glo_var_chr_alm_clave, "SISEMP", dest);
            }
            lsbPediFact.Items.Clear();
            panel4.Visible = false;
        }

        private void btnPedFact_Click(object sender, EventArgs e)
        {
            if (lblPedimento.Text == "1")
            {
                fact_extras dlg = new fact_extras(txtFolio.Text, glo_var_chr_alm_clave, DDLAlmacen1.SelectedItem.ToString(), txt_prov_clave.Text, Convert.ToDateTime(dtpFecha.Text).ToShortDateString(), "1");
                dlg.ShowDialog();
            }
        }

        private void txt_sis_tipo_cambio_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnSubirArchivos_Click(object sender, EventArgs e)
        {
            CargarPedimentoFactura();
        }

        public void ModificaOrdenCompraEntrada(DataTable dtDatos)
        {
            //Validar modificaciones al 10 %
            bool ValidacionDiezPorciento = false;
            ValidacionDiezPorciento = VerificarOrdenCompraEntrada(dtDatos);
            if (ValidacionDiezPorciento == true)
            {
                MessageBox.Show("La Entrada tiene diferencias con la Orden de Compra a la cual hace referencia, debe ser autorizada para poder modificar la Orden");
                return;
            }

            thisConnection.Open();
            decimal suma_subtotal = 0;
            decimal suma_iva = 0;
            foreach (var items in lstOC.Items)
            {
                foreach (DataRow rw in dtDatos.Select("orden_compra = '" + items + "'"))
                {
                    cmnd1 = thisConnection.CreateCommand();
                    decimal iva_imp = 0;
                    decimal costo_p = 0;
                    decimal ieps_imp = 0;
                    costo_p = Math.Round((Convert.ToDecimal(rw["importe"].ToString()) / Convert.ToDecimal(rw["cantidad"].ToString())), 4);
                    if (Convert.ToDecimal(rw["ieps"].ToString()) > 0)
                    {
                        ieps_imp = Convert.ToDecimal(rw["importe"].ToString()) * (decimal)0.08;
                    }
                    if (Convert.ToDecimal(rw["iva"].ToString()) > 0)
                    {
                        iva_imp = (Convert.ToDecimal(rw["importe"].ToString()) + ieps_imp) * (decimal)0.16;
                    }
                    
                    cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET cantidad_oc = '" + rw["cantidad"].ToString() + "', precio_oc = '" + costo_p.ToString() + "', " +
                        "importe_oc = '" + rw["importe"].ToString() + "', iva_oc = '" + iva_imp.ToString() + "', ieps_oc = '" + ieps_imp.ToString() + "' " +
                        "WHERE numero_oc = '" + items + "' AND prod_clave = '" + rw["emp_clave"].ToString() + "'";
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();
                    //suma_subtotal = suma_subtotal + Convert.ToDecimal(rw["importe"].ToString());
                    //if (Convert.ToDecimal(rw["iva"].ToString()) > 0)
                    //{
                    //    suma_iva = suma_iva + Convert.ToDecimal(rw["iva"].ToString()); 
                    //}
                }

                adapter1 = new SqlDataAdapter("SELECT SUM(ISNULL(importe_oc, 0)) AS Subtotal, SUM(ISNULL(iva_oc, 0)) AS Iva, SUM(ISNULL(ieps_oc, 0)) AS Ieps " +
                "FROM tb_det_ordenescompra WHERE numero_oc = '" + items + "' GROUP BY numero_oc", thisConnection);
                DataTable dt = new DataTable();
                adapter1.Fill(dt);

                decimal subtotal_sin_ieps = Convert.ToDecimal(dt.Rows[0]["Subtotal"].ToString());
                decimal subtotal_con_ieps = 0;
                suma_subtotal = Convert.ToDecimal(dt.Rows[0]["Subtotal"].ToString()) + Convert.ToDecimal(dt.Rows[0]["Ieps"].ToString());
                subtotal_con_ieps = suma_subtotal;
                suma_iva = Convert.ToDecimal(dt.Rows[0]["Iva"].ToString());

                adapter1 = new SqlDataAdapter("SELECT pordesc_oc AS PorDescuento " +
                "FROM tb_mstr_ordencompra WHERE numero_oc = '" + items + "'", thisConnection);
                DataTable dta = new DataTable();
                adapter1.Fill(dta);

                decimal por_descuento = 0;
                decimal imp_descuento = 0;
                por_descuento = Convert.ToDecimal(dta.Rows[0]["PorDescuento"].ToString());
                imp_descuento = Math.Round((suma_subtotal * Math.Round((por_descuento / 100), 4)), 4);

                suma_subtotal = suma_subtotal - imp_descuento;

                if(imp_descuento > 0)
                    suma_iva = Math.Round((suma_subtotal * (decimal)0.16), 4);

                decimal por_retiva = 0;
                decimal por_retisr = 0;
                decimal por_cedula = 0;

                decimal imp_retiva = 0;
                decimal imp_retisr = 0;
                decimal imp_cedula = 0;

                adapter1 = new SqlDataAdapter("SELECT retiva_por AS PorIva, retisr_por AS PorIsr, cedula_por AS PorCedular " +
                "FROM tb_oc_cfdi WHERE numero_oc = '" + items + "'", thisConnection);
                DataTable dt2 = new DataTable();
                adapter1.Fill(dt2);

                por_retiva = Convert.ToDecimal(dt2.Rows[0]["PorIva"].ToString());
                por_retisr = Convert.ToDecimal(dt2.Rows[0]["PorIsr"].ToString());
                por_cedula = Convert.ToDecimal(dt2.Rows[0]["PorCedular"].ToString());

                imp_retiva = Math.Round((subtotal_sin_ieps - imp_descuento) * Math.Round((por_retiva / 100), 4), 4);
                imp_retisr = Math.Round((subtotal_sin_ieps - imp_descuento) * Math.Round((por_retisr / 100), 4), 4);
                imp_cedula = Math.Round((subtotal_sin_ieps - imp_descuento) * Math.Round((por_cedula / 100), 4), 4);

                decimal sum_subtotal_oc = subtotal_con_ieps;

                suma_subtotal = (subtotal_con_ieps + suma_iva) - (imp_retiva + imp_retisr + imp_cedula + imp_descuento);

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET subtotal_oc = '" + sum_subtotal_oc.ToString() + "', iva_oc = '" + suma_iva.ToString() + "', " +
                        "total_oc = '" + suma_subtotal.ToString() + "', desc_oc = '" + imp_descuento.ToString() + "' " +
                        "WHERE numero_oc = '" + items + "'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_oc_cfdi SET retiva_oc = '" + imp_retiva.ToString() + "', retisr_oc = '" + imp_retisr.ToString() + "', " +
                        "cedular_oc = '" + imp_cedula.ToString() + "', correo = 2 " +
                        "WHERE numero_oc = '" + items + "'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }
            thisConnection.Close();
        }

        public bool VerificarOrdenCompraEntrada(DataTable dtDatos)
        {
            bool fndDiff = false;
            string tbody = "";
            thisConnection.Open();

            DataTable dtElaboro = new DataTable();
            dtElaboro.Columns.Add("Elaboro", typeof(string));
            dtElaboro.Columns.Add("Mail", typeof(string));
            foreach (var items in lstOC.Items)
            {
                SqlDataAdapter adap = new SqlDataAdapter("SELECT RTRIM(A.elaboro_oc) AS Elaboro, RTRIM(B.usu_mail) AS Mail " +
                    "FROM tb_mstr_ordencompra A INNER JOIN tb_cat_usuarios_emp B ON A.elaboro_oc = B.usu_nombre " +
                    "WHERE A.numero_oc = '" + items + "'", thisConnection);
                DataTable dtDetElab = new DataTable();
                adap.Fill(dtDetElab);

                if (dtElaboro.Rows.Count > 0)
                {
                    string elaboro = dtDetElab.Rows[0]["Elaboro"].ToString();

                    bool fndElab = dtElaboro
                        .AsEnumerable()
                        .Any(r => r["Elaboro"].ToString() == elaboro);

                    if(fndElab == false)
                        dtElaboro.Rows.Add(dtDetElab.Rows[0]["Elaboro"], dtDetElab.Rows[0]["Mail"]);
                }
                else
                    dtElaboro.Rows.Add(dtDetElab.Rows[0]["Elaboro"], dtDetElab.Rows[0]["Mail"]);
            }

            foreach (var items in lstOC.Items)
            {
                SqlDataAdapter adap = new SqlDataAdapter("SELECT prod_clave, cantidad_oc, importe_oc FROM tb_det_ordenescompra WHERE numero_oc = '" + items + "'", thisConnection);
                DataTable dtDetOc = new DataTable();
                adap.Fill(dtDetOc);



                foreach (DataRow dr in dtDatos.Select("orden_compra = '" + items + "'"))
                {
                    string clve = dr["emp_clave"].ToString();
                    decimal impo = Convert.ToDecimal(dr["cantidad"].ToString());
                    decimal impo2 = Convert.ToDecimal(dr["importe"].ToString());
                    string nome = "";

                    SqlCommand cmnd = new SqlCommand();
                    cmnd = thisConnection.CreateCommand();
                    cmnd.CommandText = "SELECT RTRIM(emp_nombre) AS Nombre FROM tb_cat_empaques WHERE emp_clave = '" + clve + "'";
                    SqlDataReader readr;
                    readr = cmnd.ExecuteReader();
                    if (readr.HasRows)
                    {
                        readr.Read();
                        nome = readr["Nombre"].ToString().Trim();
                    }
                    readr.Close();
                    readr.Dispose();
                    cmnd.Dispose();

                    decimal impoDet = 0;
                    decimal impoDet2 = 0;
                    foreach (DataRow rx in dtDetOc.Select("prod_clave = '" + clve + "'"))
                    {
                        impoDet = Convert.ToDecimal(rx["cantidad_oc"].ToString());
                        impoDet2 = Convert.ToDecimal(rx["importe_oc"].ToString());
                    }

                    decimal diff = impo - impoDet; //diferencia puede ser positiva o negativa
                    decimal abso = Math.Abs(diff); //valor absoluto

                    decimal diffFactOC = impoDet - abso; //diferencia de orden de compra - abso

                    decimal porcen = Math.Round((diffFactOC * 100) / impoDet, 0);

                    decimal diffporcen = 100 - porcen;

                    
                    if (diffporcen > 10)
                    {
                        fndDiff = true;
                        //MessageBox.Show("El producto tiene más o menos del 10% de diferencia", "SISEMP");
                        tbody += "<tr>" +
                            "<td style='text-align:center;'>" + items + "</td>" +
                            "<td style='text-align:center;'>" + clve + "</td>" +
                            "<td>" + nome + "</td>" +
                            "<td style='text-align:center;'>" + impoDet.ToString("###,###.00") + "</td>" +
                            "<td style='text-align:right;'>" + impoDet2.ToString("###,###.00") + "</td>" +
                            "<td style='text-align:center; color:#d9534f;'>" + impo.ToString("###,###.00") + "</td>" +
                            "<td style='text-align:right; color:#d9534f;'>" + impo2.ToString("###,###.00") + "</td>" +
                        "</tr>";
                    }
                }
            }


            string almac = "";
            if (glo_var_chr_alm_clave == "03")
            {
                almac = "COMEDOR";
            }
            if (glo_var_chr_alm_clave == "11")
            {
                almac = "COMEDOR GUARDERIA";
            }

            SqlCommand cmndAut = new SqlCommand();
            cmndAut = thisConnection.CreateCommand();
            cmndAut.CommandText = "INSERT INTO tb_oc_autoriza_mod(numero_oc, fecha_sol, motivo, usuario, maquina, autorizacion, comedor, alm) " +
                "VALUES ('" + txtFolio.Text + "', getdate(), 'DIFERENCIA DE LO ENTREGADO CONTRA LO SOLICITADO EN " + almac + " ', 'COMPRAS', '" + Environment.MachineName + "', '1', '1', '" + glo_var_chr_alm_clave + "')";
            cmndAut.ExecuteNonQuery();
            cmndAut.Dispose();

            thisConnection.Close();

            if (fndDiff == true)
            {
                //Envio de correo
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();

                message.From = new MailAddress("comprasgab@mrlucky.com.mx");
                //message.To.Add(new MailAddress("musabiaga@mrlucky.com.mx"));
                message.To.Add(new MailAddress("aescamilla@mrlucky.com.mx"));
                if (glo_var_chr_alm_clave == "03")
                {
                    //message.CC.Add(new MailAddress("comedor@mrlucky.com.mx"));
                    //message.CC.Add(new MailAddress("compras@mrlucky.com.mx"));
                }
                if (glo_var_chr_alm_clave == "11")
                {
                    //message.CC.Add(new MailAddress("guarderia.directora@mrlucky.com.mx"));
                    //message.CC.Add(new MailAddress("compras@mrlucky.com.mx"));
                }
                //message.Bcc.Add(new MailAddress("aescamilla@mrlucky.com.mx"));
                string body = BodyCorreo(tbody);
                message.IsBodyHtml = true;
                message.SubjectEncoding = System.Text.Encoding.UTF8;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.Subject = "Solicitud de Ajuste de Orden de Compra del almacen de " + almac;
                message.Body = body;
                message.From = new MailAddress("comprasgab@mrlucky.com.mx", "Departament de Compras");
                smtp.Credentials = new System.Net.NetworkCredential("comprasgab", "OTM54A.up#c*01&qs&");
                smtp.Port = 587;
                smtp.Host = "mail1.mrlucky.com.mx";
                smtp.EnableSsl = true;

                smtp.Send(message);
            }
            return fndDiff;
        }

        public string BodyCorreo(string tbody)
        {
            string body = "<html>" +
                "<body style='font-family: Arial, sans-serif; font-size: 12pt; color: #333;'>" +
                    "<h2 style='color: #2a4d7f; text-align: center; margin-bottom: 20px;'>" +
                        "Informe para modificación de Orden de Compra" +
                    "</h2>" +
                    "<p style='margin: 10px 0;'>" +
                        "A quien corresponda," +
                    "</p>" +
                    "<p style='margin: 10px 0;'>" +
                        "Por este medio se informa a gerencia sobre las modificaciones registradas en la entrada" +
                    "</p>" +
                    "<p style='margin: 10px 0;'>" +
                        "<b style='color:#2a4d7f;'>Proveedor:</b> " + DDLProveedor.Text + "<br/>" +
                        //"<b style='color:#2a4d7f;'>Orden de Compra:</b> [NÚMERO DE ORDEN]<br/>" +
                        "<b style='color:#2a4d7f;'>Entrada:</b> " + txtFolio.Text +
                    "</p>" +
                    "<table border='1' cellpadding='6' cellspacing='0' " +
                           "style='border-collapse: collapse; width: 100%; margin-top: 15px; font-size: 11pt;'>" +
                        "<tr style='background-color:#f2f2f2; text-align: center;'>" +
                            "<th style='padding: 8px; color:#2a4d7f;'>Orden Compra</th>" +
                            "<th style='padding: 8px; color:#2a4d7f;'>Clave</th>" +
                            "<th style='padding: 8px; color:#2a4d7f;'>Producto</th>" +
                            "<th style='padding: 8px; color:#2a4d7f;'>Cantidad</th>" +
                            "<th style='padding: 8px; color:#2a4d7f;'>Importe</th>" +
                            "<th style='padding: 8px; color:#2a4d7f;'>Cantidad Nueva</th>" +
                            "<th style='padding: 8px; color:#2a4d7f;'>Importe Nuevo</th>" +
                        "</tr>" +
                        tbody +
                    "</table>" +
                    "<p style='margin: 15px 0;'>" +
                        "Se solicita autorización para proceder con el ajuste de la orden de compra conforme a la entrada registrada. " + txtObservaciones.Text + 
                    "</p>" +
                    "<p style='margin: 15px 0;'>" +
                        "En ocasiones el producto está compartido entre dos o más órdenes de compra." +
                    "</p>" +
                    "<p style='margin: 10px 0;'>" +
                        "Saludos cordiales,<br/>" +
                        "<span style='color:#2a4d7f; font-weight:bold;'>Departamento de Compras</span>" +
                    "</p>" +
                "</body>" +
                "</html>";
            return body;

            //"<tr>" +
            //                "<td style='text-align:center;'>[CLAVE]</td>" +
            //                "<td>[PRODUCTO]</td>" +
            //                "<td style='text-align:center;'>[CANTIDAD]</td>" +
            //                "<td style='text-align:right;'>[IMPORTE]</td>" +
            //                "<td style='text-align:center; color:#d9534f;'>[CANTIDAD NUEVA]</td>" +
            //                "<td style='text-align:right; color:#d9534f;'>[IMPORTE NUEVO]</td>" +
            //            "</tr>" +
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnReajustar_Click(object sender, EventArgs e)
        {
            ModificaOrdenCompraEntradaAutorizada(dtdetalles);

            MessageBox.Show("La orden de compra fue ajustada a la Entrada con éxito", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

            string almac = "";
            if (glo_var_chr_alm_clave == "03")
            {
                almac = "COMEDOR";
            }
            if (glo_var_chr_alm_clave == "11")
            {
                almac = "COMEDOR GUARDERIA";
            }

            MailMessage message = new MailMessage();
            SmtpClient smtp = new SmtpClient();

            message.From = new MailAddress("comprasgab@mrlucky.com.mx");
            message.To.Add(new MailAddress("aescamilla@mrlucky.com.mx"));
            //message.Bcc.Add(new MailAddress("aescamilla@mrlucky.com.mx"));
            string body = BodyCorreoModificado();
            message.IsBodyHtml = true;
            message.SubjectEncoding = System.Text.Encoding.UTF8;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.Subject = "Ajuste de Orden de Compra del almacen de " + almac;
            message.Body = body;
            message.From = new MailAddress("comprasgab@mrlucky.com.mx", "Departamento de Compras");
            smtp.Credentials = new System.Net.NetworkCredential("comprasgab", "OTM54A.up#c*01&qs&");
            smtp.Port = 587;
            smtp.Host = "mail1.mrlucky.com.mx";
            smtp.EnableSsl = true;

            smtp.Send(message);

            btnReajustar.Visible = false;
        }

        public void ModificaOrdenCompraEntradaAutorizada(DataTable dtDatos)
        {
            thisConnection.Open();
            decimal suma_subtotal = 0;
            decimal suma_iva = 0;
            foreach (var items in lstOC.Items)
            {
                foreach (DataRow rw in dtDatos.Select("orden_compra = '" + items + "'"))
                {
                    cmnd1 = thisConnection.CreateCommand();
                    decimal iva_imp = 0;
                    decimal costo_p = 0;
                    decimal ieps_imp = 0;
                    costo_p = Math.Round((Convert.ToDecimal(rw["importe"].ToString()) / Convert.ToDecimal(rw["cantidad"].ToString())), 4);
                    if (Convert.ToDecimal(rw["ieps"].ToString()) > 0)
                    {
                        ieps_imp = Convert.ToDecimal(rw["importe"].ToString()) * (decimal)0.08;
                    }
                    if (Convert.ToDecimal(rw["iva"].ToString()) > 0)
                    {
                        iva_imp = (Convert.ToDecimal(rw["importe"].ToString()) + ieps_imp) * (decimal)0.16;
                    }

                    cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET cantidad_oc = '" + rw["cantidad"].ToString() + "', precio_oc = '" + costo_p.ToString() + "', " +
                        "importe_oc = '" + rw["importe"].ToString() + "', iva_oc = '" + iva_imp.ToString() + "', ieps_oc = '" + ieps_imp.ToString() + "' " +
                        "WHERE numero_oc = '" + items + "' AND prod_clave = '" + rw["emp_clave"].ToString() + "'";
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();
                    //suma_subtotal = suma_subtotal + Convert.ToDecimal(rw["importe"].ToString());
                    //if (Convert.ToDecimal(rw["iva"].ToString()) > 0)
                    //{
                    //    suma_iva = suma_iva + Convert.ToDecimal(rw["iva"].ToString()); 
                    //}
                }

                adapter1 = new SqlDataAdapter("SELECT SUM(ISNULL(importe_oc, 0)) AS Subtotal, SUM(ISNULL(iva_oc, 0)) AS Iva, SUM(ISNULL(ieps_oc, 0)) AS Ieps " +
                "FROM tb_det_ordenescompra WHERE numero_oc = '" + items + "' GROUP BY numero_oc", thisConnection);
                DataTable dt = new DataTable();
                adapter1.Fill(dt);

                decimal subtotal_sin_ieps = Convert.ToDecimal(dt.Rows[0]["Subtotal"].ToString());
                decimal subtotal_con_ieps = 0;
                suma_subtotal = Convert.ToDecimal(dt.Rows[0]["Subtotal"].ToString()) + Convert.ToDecimal(dt.Rows[0]["Ieps"].ToString());
                subtotal_con_ieps = suma_subtotal;
                suma_iva = Convert.ToDecimal(dt.Rows[0]["Iva"].ToString());

                adapter1 = new SqlDataAdapter("SELECT pordesc_oc AS PorDescuento " +
                "FROM tb_mstr_ordencompra WHERE numero_oc = '" + items + "'", thisConnection);
                DataTable dta = new DataTable();
                adapter1.Fill(dta);

                decimal por_descuento = 0;
                decimal imp_descuento = 0;
                por_descuento = Convert.ToDecimal(dta.Rows[0]["PorDescuento"].ToString());
                imp_descuento = Math.Round((suma_subtotal * Math.Round((por_descuento / 100), 4)), 4);

                suma_subtotal = suma_subtotal - imp_descuento;

                if (imp_descuento > 0)
                    suma_iva = Math.Round((suma_subtotal * (decimal)0.16), 4);

                decimal por_retiva = 0;
                decimal por_retisr = 0;
                decimal por_cedula = 0;

                decimal imp_retiva = 0;
                decimal imp_retisr = 0;
                decimal imp_cedula = 0;

                adapter1 = new SqlDataAdapter("SELECT retiva_por AS PorIva, retisr_por AS PorIsr, cedula_por AS PorCedular " +
                "FROM tb_oc_cfdi WHERE numero_oc = '" + items + "'", thisConnection);
                DataTable dt2 = new DataTable();
                adapter1.Fill(dt2);

                por_retiva = Convert.ToDecimal(dt2.Rows[0]["PorIva"].ToString());
                por_retisr = Convert.ToDecimal(dt2.Rows[0]["PorIsr"].ToString());
                por_cedula = Convert.ToDecimal(dt2.Rows[0]["PorCedular"].ToString());

                imp_retiva = Math.Round((subtotal_sin_ieps - imp_descuento) * Math.Round((por_retiva / 100), 4), 4);
                imp_retisr = Math.Round((subtotal_sin_ieps - imp_descuento) * Math.Round((por_retisr / 100), 4), 4);
                imp_cedula = Math.Round((subtotal_sin_ieps - imp_descuento) * Math.Round((por_cedula / 100), 4), 4);

                decimal sum_subtotal_oc = subtotal_con_ieps;

                suma_subtotal = (subtotal_con_ieps + suma_iva) - (imp_retiva + imp_retisr + imp_cedula + imp_descuento);

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET subtotal_oc = '" + sum_subtotal_oc.ToString() + "', iva_oc = '" + suma_iva.ToString() + "', " +
                        "total_oc = '" + suma_subtotal.ToString() + "', desc_oc = '" + imp_descuento.ToString() + "' " +
                        "WHERE numero_oc = '" + items + "'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_oc_cfdi SET retiva_oc = '" + imp_retiva.ToString() + "', retisr_oc = '" + imp_retisr.ToString() + "', " +
                        "cedular_oc = '" + imp_cedula.ToString() + "', correo = 2 " +
                        "WHERE numero_oc = '" + items + "'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }
            thisConnection.Close();
        }

        public string BodyCorreoModificado()
        {
            //obtener las ordenes de compra
            string concatenados = string.Join(",",
                lstOC.Items.Cast<object>()
                .Select(item => item.ToString()));

            string body = "<html>" +
                "<body style='font-family: Arial, sans-serif; font-size: 12pt; color: #333;'>" +
                    "<h2 style='color: #2a4d7f; text-align: center; margin-bottom: 20px;'>" +
                        "Informe de modificación de Orden de Compra segun Factura entragada a almacen" +
                    "</h2>" +
                    "<p style='margin: 10px 0;'>" +
                        "A quien corresponda," +
                    "</p>" +
                    "<p style='margin: 10px 0;'>" +
                        "Por este medio se informa que la orden de compra ha sido reajustada a lo recibido" +
                    "</p>" +
                    "<p style='margin: 10px 0;'>" +
                        "<b style='color:#2a4d7f;'>Proveedor:</b> " + DDLProveedor.Text + "<br/>" +
                        "<b style='color:#2a4d7f;'>Orden de Compra:</b> " + concatenados + "<br/>" +
                        "<b style='color:#2a4d7f;'>Entrada:</b> " + txtFolio.Text +
                    "</p>" +
                    "<p style='margin: 10px 0;'>" +
                        "Saludos cordiales,<br/>" +
                        "<span style='color:#2a4d7f; font-weight:bold;'>Departamento de Compras</span>" +
                    "</p>" +
                "</body>" +
                "</html>";
            return body;
        }
    }
}
