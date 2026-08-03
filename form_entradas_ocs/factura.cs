using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Xml;
using System.Net.Mail;
using System.Net;
using System.Net.NetworkInformation;
using System.IO;
//using Ionic.Zip;

namespace form_entradas_ocs
{
    public partial class factura : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);//
        SqlDataReader reader1, reader2, reader3, reader4;
        SqlCommand cmnd4;
        SqlCommand cmnd3;
        SqlCommand cmnd2;
        SqlCommand cmnd1;
        SqlDataAdapter adapter1;

        //SqlConnection conn = new SqlConnection("server=grupou.net.,1433\\MSSQLSERVER2012; Initial Catalog = grupou_WMService; User Id = grupou_Cfdi; Password = Prv2017#i");
        SqlConnection conn = new SqlConnection("server=grupou-mrlucky.com.mx,15333; Initial Catalog = grupou_Service; User Id = mrlucky; Password = Cfdi79#!=");
        SqlCommand comm;
        SqlDataReader redr;

        string subtipo = "";

        string carga_solo_documentos = "0";

        public factura(string entrada, string cve_alm, string nom_alm, string cve_prv, string nom_prv, string fact_act, string fol_fis_act)
        {
            InitializeComponent();

            lblEntrada.Text = entrada;
            lblClaveAlm.Text = cve_alm;
            lblAlmacen.Text = nom_alm.ToUpper();
            lblClaveProveedor.Text = cve_prv;
            lblProveedor.Text = nom_prv;
            lblFactActual.Text = fact_act;
            lblFolioFiscalAct.Text = fol_fis_act;

            if (nom_alm.ToUpper() == "COMEDOR")
                subtipo = "COM";
            else if (nom_alm.ToUpper() == "GUARDERIA")
                subtipo = "COM";
            else
                subtipo = "";
        }

        private void btnCierra_Click(object sender, EventArgs e)
        {
            if (txtNomina.Text == "" || txtObservaciones.Text == "")
            {
                MessageBox.Show("Falta ingresar el número de nómina o el porque del cambio de factura", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            

            //BUSCAR NUMERO DE NOMINA
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT emp_clave FROM tb_cat_empleados WHERE emp_clave = '" + txtNomina.Text + "' AND [emp_nompuesto] like '%ALMACEN%' OR [emp_nompuesto] like '%COMEDOR%'";
            reader1 = cmnd1.ExecuteReader();
            if (!reader1.HasRows)
            {
                MessageBox.Show("Nómina no encontrada, favor de verificar", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();
                return;
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();

            openFileDialog1.InitialDirectory = "c:\\" + lblAlmacen;

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

                foreach (XmlElement nodo in lista3)
                {
                    FOL_FIS = nodo.GetAttribute("UUID");
                }

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
                    cmnd1.CommandText = "SELECT numero_oc, factura, fecha_factura, fecha_registro, tipo FROM tb_det_oc_facturas WHERE folio_fiscal = '" + FOL_FIS + "'";
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
                        if (lblClaveProveedor.Text.Trim() == "404" || lblClaveProveedor.Text.Trim() == "1269" || lblClaveProveedor.Text.Trim() == "694" || lblClaveProveedor.Text.Trim() == "318")
                        {
                            MessageBox.Show("El folio fiscal ya fue registrado.\nEntrada: " + noc + "\nFactura: " + fct + "\nFecha Factura: " + fef + "\nFecha Registro: " + fer + "\nAlmacen: " + tpo + "\nSolo se permitirá subir los documentos si la entrada registrada es igual a la entrada consultada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (noc.PadLeft(6, '0') != lblEntrada.Text)
                                carga_solo_documentos = "0";//return;
                            else
                                carga_solo_documentos = "1";
                        }
                        else
                        {
                            MessageBox.Show("El folio fiscal ya fue registrado.\nEntrada: " + noc + "\nFactura: " + fct + "\nFecha Factura: " + fef + "\nFecha Registro: " + fer + "\nAlmacen: " + tpo + "\nSolo se permitirá subir los documentos si la entrada registrada es igual a la entrada consultada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            if (noc.PadLeft(6, '0') != lblEntrada.Text)
                                return;
                            else
                                carga_solo_documentos = "1";
                        }   
                    }
                    //else
                    //{
                    //    if (fnd == true && (lblClaveProveedor.Text.Trim() != "404" || lblClaveProveedor.Text.Trim() != "1269" || lblClaveProveedor.Text.Trim() != "694"))
                    //    {
                    //        Utilerias.Class1.registrar_movimiento3(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "MF", "2.5", noc, "FACTURA ENCONTRADA", "SISEMP", "", "FACTURA REGISTRADA Entrada: " + noc + " Factura: " + fct + " Fecha Factura: " + fef + " Fecha Registro: " + fer + " Almacen: " + tpo);
                    //    }
                    //}
                    
                    
                    
                    
                    
                    
                    
                    /*if (fnd == true)
                    {
                        MessageBox.Show("El folio fiscal ya fue registrado.\nEntrada: " + noc + "\nFactura: " + fct + "\nFecha Factura: " + fef + "\nFecha Registro: " + fer + "\nAlmacen: " + tpo + "\nSolo se permitirá subir los documentos si la entrada registrada es igual a la entrada consultada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        if (noc.PadLeft(6, '0') != lblEntrada.Text)
                            return;
                        else
                            carga_solo_documentos = "1";
                    }*/
                }

                lblFolioFiscal.Text = FOL_FIS;
                lbl_ent_factura.Text = Serie + Folio;
                lblFechaFact.Text = Convert.ToDateTime(Fecha).ToShortDateString();
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
                

                //VERIFICACION DE RFC DE PROVEEDOR CON EL DE LA FACTURA
                string ProvRfc = ""; //PFC960325U48
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT prov_rfc FROM tb_cat_proveedor WHERE prov_clave = '" + lblClaveProveedor.Text + "'";
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
                    MessageBox.Show("LOS RFC NO CORRESPONDEN, FAVOR DE VALIDAR LOS RFC Y CORREGIRLO EN SISTEMA DE SER NECESARIO.\n RFC - XML:" + lblRFC.Text + "\n RFC - SISTEMA:" + ProvRfc + "\n", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    lblFechaFact.Text = "-";
                    lblFolioFiscal.Text = "-";
                    lblRFC.Text = "-";
                    return;

                    /*if (MessageBox.Show("El XML corresponde al proveedor:\n" + nom_prov + "\n¿Desea continuar?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        lblFechaFact.Text = "-";
                        lblFolioFiscal.Text = "-";
                        lblRFC.Text = "-";
                        return;
                    }*/
                }

                bool ping_con = ping_portal();
                try
                {
                    if (ping_con == true)
                    {
                        conn.Open();
                        comm = conn.CreateCommand();

                        string cfd_timfis = lblFolioFiscal.Text;
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
                }
                catch
                {
                }
                
                btnGuardar.Enabled = true;

            }
        }

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

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            string qry = "";
            try
            {
                if (carga_solo_documentos == "0")
                {
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    qry = "UPDATE tb_mstr_entradas SET ent_factura = '" + lbl_ent_factura.Text + "', ent_fechafact = '" + lblFechaFact.Text + "', " +
                       "ent_foliofiscal = '" + lblFolioFiscal.Text + "', ent_modif_fact = ent_modif_fact + 1, ent_fact_obser = CONCAT(ent_fact_obser, '" + txtObservaciones.Text + " fecha: " + DateTime.Now.ToLongDateString() + ";'), ent_nomina = CONCAT(ent_nomina, '" + txtNomina.Text + ";') " +
                       "WHERE ent_folio = '" + lblEntrada.Text + "' AND alm_clave = '" + lblClaveAlm.Text + "' AND ent_subtipo = '" + subtipo + "'";
                    cmnd1.CommandText = qry;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();

                    cmnd1 = thisConnection.CreateCommand();
                    string portal = "";
                    string sts_portal = "";
                    if (lblPortal.Text != "-")
                    {
                        sts_portal = "1";
                        portal = lblPortal.Text;
                    }
                    else
                    {
                        sts_portal = "0";
                        portal = "";
                    }

                    qry = "UPDATE tb_det_oc_facturas SET factura = '" + lbl_ent_factura.Text + "', fecha_factura = '" + lblFechaFact.Text + "', " +
                        "folio_fiscal = '" + lblFolioFiscal.Text + "', monto = '" + lblTotal.Text + "', portal = '" + portal + "', estatus_portal = '" + sts_portal + "' " +
                        "WHERE numero_oc = '" + (lblEntrada.Text.TrimStart('0').ToString()) + "' AND tipo = '" + lblAlmacen.Text + "'";
                    cmnd1.CommandText = qry;
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();

                    qry = "upload";
                    thisConnection.Close();
                }

                ////AGREGAR FACTURA A ARCHIVO ZIP
                //string carpeta_nueva = "";
                //if (lblClaveAlm.Text == "01")
                //    carpeta_nueva = "Materiales";
                //else if (lblClaveAlm.Text == "02")
                //    carpeta_nueva = "Empaques";
                //else if (lblClaveAlm.Text == "03")
                //    carpeta_nueva = "Comedor";
                //else if (lblClaveAlm.Text == "04")
                //    carpeta_nueva = "Enfermeria";
                //else if (lblClaveAlm.Text == "05")
                //    carpeta_nueva = "Papeleria";
                //else if (lblClaveAlm.Text == "07")
                //    carpeta_nueva = "Dental";
                //else if (lblClaveAlm.Text == "08")
                //    carpeta_nueva = "Combustible";

                FileInfo archivo = new FileInfo(@lblPathXml.Text);

                string file_name = lblClaveProveedor.Text + "_" + lbl_ent_factura.Text + "_" + lblEntrada.Text + ".xml";

                FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                if (liq_copy.Exists == true)
                {
                    liq_copy.Delete();
                }
                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                ////AGREGAR XML DE FACTURA A ARCHIVO ZIP
                //if (File.Exists(@"\\gabira1\entradasfact\ent\Ent_" + lblEntrada.Text + "_" + carpeta_nueva + ".zip") == true)
                //{
                //    using (ZipFile zip = ZipFile.Read(@"\\gabira1\entradasfact\ent\Ent_" + lblEntrada.Text + "_" + carpeta_nueva + ".zip"))
                //    {
                //        zip.AddFile(@"\\\\gabira1\\FacturasExternos\\" + file_name, "");
                //        zip.Save();
                //    }
                //}

                file_name = lblClaveProveedor.Text + "_" + lbl_ent_factura.Text + "_" + lblEntrada.Text + ".pdf";

                liq_copy = new FileInfo(@"\\\\gabira1\\FacturasExternos\\" + file_name);
                if (liq_copy.Exists == true)
                {
                    liq_copy.Delete();
                }
                archivo = new FileInfo(@lblPathXml.Text.Replace(".xml", ".pdf"));
                archivo.CopyTo(@"\\\\gabira1\\FacturasExternos\\" + file_name);

                ////AGREGAR PDF DE FACTURA A ARCHIVO ZIP
                //if (File.Exists(@"\\gabira1\entradasfact\ent\Ent_" + lblEntrada.Text + "_" + carpeta_nueva + ".zip") == true)
                //{
                //    using (ZipFile zip = ZipFile.Read(@"\\gabira1\entradasfact\ent\Ent_" + lblEntrada.Text + "_" + carpeta_nueva + ".zip"))
                //    {
                //        zip.AddFile(@"\\\\gabira1\\FacturasExternos\\" + file_name, "");
                //        zip.Save();
                //    }
                //}
                                
                //ELIMINAR LA FACTURA ANTERIOR
                if (lblFactActual.Text != lbl_ent_factura.Text)
                {
                    if (File.Exists(@"\\\\gabira1\\FacturasExternos\\" + lblClaveProveedor.Text + "_" + lblFactActual.Text + "_" + lblEntrada.Text + ".pdf"))
                    {
                        File.Delete(@"\\\\gabira1\\FacturasExternos\\" + lblClaveProveedor.Text + "_" + lblFactActual.Text + "_" + lblEntrada.Text + ".pdf");
                    }
                    if (File.Exists(@"\\\\gabira1\\FacturasExternos\\" + lblClaveProveedor.Text + "_" + lblFactActual.Text + "_" + lblEntrada.Text + ".xml"))
                    {
                        File.Delete(@"\\\\gabira1\\FacturasExternos\\" + lblClaveProveedor.Text + "_" + lblFactActual.Text + "_" + lblEntrada.Text + ".xml");
                    }

                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "UPDATE tb_det_oc_facturas SET estatus_entrega = '0' WHERE folio_fiscal = '" + lblFolioFiscalAct.Text + "' AND estatus = 'A'";
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();
                    thisConnection.Close();
                }
                

                ////ELIMINAR XML DE FACTURA DEL ARCHIVO ZIP SI EXISTE
                //string file_name2 = lblClaveProveedor.Text + "_" + lblFactActual.Text + "_" + lblEntrada.Text + ".pdf";
                //string file_name3 = lblClaveProveedor.Text + "_" + lblFactActual.Text + "_" + lblEntrada.Text + ".xml";
                //if (File.Exists(@"\\gabira1\entradasfact\ent\Ent_" + lblEntrada.Text + "_" + carpeta_nueva + ".zip") == true)
                //{
                //    using (ZipFile zip = ZipFile.Read(@"\\gabira1\entradasfact\ent\Ent_" + lblEntrada.Text + "_" + carpeta_nueva + ".zip"))
                //    {
                //        zip.RemoveEntry(file_name2);
                //        zip.RemoveEntry(file_name3);
                //        zip.Save();
                //    }
                //}

                //file_name2 = file_name.Replace(".pdf", ".xml");


                carga_solo_documentos = "0";

                //MessageBox.Show("Datos Guardados, se desactivará la autorización de descarga", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                MessageBox.Show("Datos Guardados", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //CANCELACION DE DESCARGA
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_mstr_entradas SET descarga = '0' WHERE ent_folio = '" + lblEntrada.Text + "' AND alm_clave = '" + lblClaveAlm.Text + "'";
                //cmnd1.ExecuteNonQuery();
                thisConnection.Close();

                //enviarcorreo_zip();
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error de sistema, no se termino de guardar la factura", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lblEntrada.Text + " " + ex.ToString() + " " + qry, "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + lblEntrada.Text + " " + ex.ToString() + " " + qry);
            }

        }

        public void enviarcorreo_zip()
        {
            string persona = "";
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT usu_nombre, usu_email, usu_pass FROM tb_cat_usuarios_emp WHERE usu_login = '" + Utilerias.Class1.Usu_login + "'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                persona = reader1["usu_nombre"].ToString().Trim();
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();
            
            //string host = Dns.GetHostName();
            //IPHostEntry ipEntry = Dns.GetHostEntry(host);
            string cuerpo = "<table border='1'><tr><td colspan='2' align='center'><h3>SE CANCELO LA DESCARGA DE ARCHIVO POR MOTIVO DE CAMBIO DE FACTURA A LA ENTRADA</h3></td></tr>" +
                "<tr><td>Liberado por:</td><td>" + persona + "</td></tr><tr><td>Fecha:</td><td>" + DateTime.Now.ToString() + "</td></tr>" +
                "<tr><td>Entrada / OC:</td><td>" + lblEntrada.Text + "</td></tr><tr><td>Almacen:</td><td>" + lblAlmacen.Text + "</td></tr>" +
                "<tr><td>Almacen:</td><td>FAVOR DE CONSIDERAR ESTA MODIFICACION SI YA REALIZO LA DESCARGA O RECIBIO LOS ARCHIVOS POR CORREO DE LOS DOCUMENTOS ANTES DE ESTE MOVIMIENTO.</td></tr></table>";


            System.Net.Mail.MailMessage mnsg = new System.Net.Mail.MailMessage();
            mnsg.To.Add("sistemas@mrlucky.com.mx");//usu_mail
            //mnsg.To.Add("cmoreno@mrlucky.com.mx");
            mnsg.CC.Add("pgonzalez@mrlucky.com.mx");
            //mnsg.CC.Add("compras@mrlucky.com.mx");
            //mnsg.CC.Add("meflores@mrlucky.com.mx");
            mnsg.CC.Add("avaladez@mrlucky.com.mx");
            mnsg.CC.Add("recepcion@mrlucky.com.mx");
            mnsg.CC.Add("almacen@mrlucky.com.mx");

            //if (des != "")
            //{
            mnsg.CC.Add("lupita@grupou.mx");
            mnsg.CC.Add("viridiana@grupou.mx");
            //}

            //mnsg.CC.Add("aescamilla@mrlucky.com.mx");
            mnsg.Bcc.Add("aescamilla@mrlucky.com.mx");
            //
            mnsg.Subject = "Cancelación de descarga de archivo por modificación de factura: " + lblEntrada.Text + ".";
            mnsg.SubjectEncoding = System.Text.Encoding.UTF8;
            mnsg.Body = cuerpo;
            mnsg.BodyEncoding = System.Text.Encoding.UTF8;
            mnsg.IsBodyHtml = true;
            //mnsg.From = new MailAddress("aescamilla@mrlucky.com.mx");
            mnsg.From = new MailAddress("sistemas@mrlucky.com.mx");

            SmtpClient cliente = new SmtpClient();
            cliente.Credentials = new System.Net.NetworkCredential("sistemas", "sisgab");
            //cliente.Credentials = new System.Net.NetworkCredential("aescamilla", "atrejo");
            cliente.Port = 587;
            cliente.EnableSsl = true;
            cliente.Host = "mail1.mrlucky.com.mx";

            try
            {
                cliente.Send(mnsg);
            }
            catch (Exception)
            {
            }
        }


    }
}
