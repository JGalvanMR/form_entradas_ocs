using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;

namespace form_entradas_ocs
{
    public partial class etiquetas : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader1;
        SqlCommand cmnd1;
        SqlDataAdapter adapter1;

        DataSet dtDatos = new DataSet();

        public etiquetas(string entrada, string almacen)
        {
            InitializeComponent();

            lblEntrada.Text = entrada;
            lblAlmacen.Text = almacen;

            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT alm_nombre from tb_cat_almacenes WHERE alm_clave = '" + almacen + "' AND alm_tipo <> 'GUA' AND alm_clave <> '06'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                lblNombre.Text = reader1["alm_nombre"].ToString().Trim();
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();

            adapter1 = new SqlDataAdapter("SELECT RTRIM(A.emp_clave) AS emp_clave, RTRIM(B.emp_nombre) AS emp_nombre, A.entd_cantidad " +
                "FROM tb_det_ent_temporal A JOIN tb_cat_empaques B ON A.emp_clave = B.emp_clave AND A.alm_clave = B.alm_clave " +
                "WHERE A.alm_clave = '" + almacen + "' AND A.ent_folio = '" + lblEntrada.Text + "'", thisConnection);
            adapter1.Fill(dtDatos, "tb_det_ent_temporal");
            adapter1.Dispose();

            foreach (DataRow rw in dtDatos.Tables["tb_det_ent_temporal"].Rows)
            {
                dtgEmpaques.Rows.Add(rw["emp_clave"].ToString(), rw["emp_nombre"].ToString(), Convert.ToDecimal(rw["entd_cantidad"].ToString()).ToString("###,###,##0.00"), "0");
            }
        }

        private void dtgEmpaques_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3)
            {
                string imp = dtgEmpaques.CurrentRow.Cells[3].Value.ToString();
                decimal imp2 = 0;
                try
                {
                    imp2 = Convert.ToDecimal(imp);
                }
                catch (Exception)
                {
                    MessageBox.Show("El valor ingresado debe ser numérico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    dtgEmpaques.CurrentRow.Cells[3].Value = "0";
                    return;
                }
            }
        }

        private void btnEtiquetas_Click(object sender, EventArgs e)
        {
            
            PrintDialog pd = new PrintDialog();
            pd.PrinterSettings = new PrinterSettings();
            if (DialogResult.OK == pd.ShowDialog(this))
            {
                foreach (DataGridViewRow gr in dtgEmpaques.Rows)
                {
                    if (gr.Cells[3].Value.ToString() == "0")
                        continue;
                    decimal cantidad = Math.Ceiling(Convert.ToDecimal(gr.Cells[3].Value.ToString()) / 2);
                    string cve_emp = gr.Cells[0].Value.ToString();
                    string nom_emp = gr.Cells[1].Value.ToString();
                    //for (int i = 0; i < cantidad; i++)
                    //{
                    string impr_blancas = "";

                    string elaboracion = DateTime.Now.ToString("dd/MMM/yy").ToUpper().Replace(".", "");
                    string caducidad = DateTime.Now.ToString("dd/MMM/yy").ToUpper().Replace(".", "");

                    impr_blancas += "^XA\n";
                    //impr_blancas += "~SD18\n";
                    //impr_blancas += "^PW832";
                    //impr_blancas += "^FO15,50^A0N,25,16^FDELABORACION^FS\n";
                    //impr_blancas += "^FO480,50^A0N,25,16^FDELABORACION^FS\n";
                    //impr_blancas += "^FO15,75^A0N,25,20^FD" + elaboracion + "^FS\n";
                    //impr_blancas += "^FO480,75^A0N,25,20^FD" + elaboracion + "^FS\n";
                    //impr_blancas += "^FO15,107^A0N,25,19^FDCADUCIDAD^FS\n";
                    //impr_blancas += "^FO480,107^A0N,25,19^FDCADUCIDAD^FS\n";
                    //impr_blancas += "^FO15,130^A0N,25,20^FD" + caducidad + "^FS\n";
                    //impr_blancas += "^FO480,130^A0N,25,20^FD" + caducidad + "^FS\n";
                    //impr_blancas += "^PQ" + cantidad.ToString() + ",1400,0,N";
                    //impr_blancas += "^XZ\n";

                    //impr_blancas += "^FO15,50^BY3^BCN,100,Y,N,N^FD>6CODE128^FS";
                    impr_blancas += "^FO50,20^BY3^BCN,100,Y,N,N^FD" + cve_emp + "^FS\n";
                    int cont_car = nom_emp.Length;
                    string cad1 = "";
                    string cad2 = "";
                    string cad3 = "";
                    string cad4 = "";
                    if (cont_car > 25)
                    {
                        cad1 = nom_emp.Substring(0, 25).ToString();
                        int resto = (nom_emp.Length - cad1.Length);
                        cad2 = nom_emp.Substring(25, resto).ToString();
                        if (cad2.Length > 25)
                        {
                            cad3 = cad2.Substring(0, 25).ToString();
                            int resto2 = (cad2.Length - cad3.Length);
                            cad4 = cad2.Substring(25, resto2).ToString();
                        }
                    }
                    else
                        cad1 = nom_emp;

                    impr_blancas += "^FO15,160^A0N,25,28^FD" + cad1 + "^FS\n";
                    if (cad2 != "")
                    {
                        if (cad3 != "")
                            impr_blancas += "^FO15,190^A0N,25,28^FD" + cad3 + "^FS\n";
                        else
                            impr_blancas += "^FO15,190^A0N,25,28^FD" + cad2 + "^FS\n";
                    }
                    if (cad4 != "")
                    {
                        impr_blancas += "^FO15,220^A0N,25,28^FD" + cad4 + "^FS\n";
                    }
                    impr_blancas += "^FO480,20^BY3^BCN,100,Y,N,N^FD" + cve_emp + "^FS\n";

                    impr_blancas += impr_blancas += "^FO430,160^A0N,25,28^FD" + cad1 + "^FS\n";
                    if (cad2 != "")
                    {
                        if(cad3 != "")
                            impr_blancas += "^FO430,190^A0N,25,28^FD" + cad3 + "^FS\n";
                        else
                            impr_blancas += "^FO430,190^A0N,25,28^FD" + cad2 + "^FS\n";
                    }
                    if (cad4 != "")
                    {
                        impr_blancas += "^FO430,220^A0N,25,28^FD" + cad4 + "^FS\n";
                    }
                    //impr_blancas += "^FO430,150^A0N,25,16^FD" + nom_emp + "^FS\n";
                    impr_blancas += "^PQ" + cantidad.ToString() + ",1400,0,N";
                    impr_blancas += "^XZ\n";
                    

                    RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, impr_blancas);
                    //}
                }
            }
        }

        public class RawPrinterHelper
        {
            // Structure and API declarions:
            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
            public class DOCINFOA
            {
                [MarshalAs(UnmanagedType.LPStr)]
                public string pDocName;
                [MarshalAs(UnmanagedType.LPStr)]
                public string pOutputFile;
                [MarshalAs(UnmanagedType.LPStr)]
                public string pDataType;
            }
            [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

            [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool ClosePrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

            [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool EndDocPrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool StartPagePrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool EndPagePrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

            // SendBytesToPrinter()
            // When the function is given a printer name and an unmanaged array
            // of bytes, the function sends those bytes to the print queue.
            // Returns true on success, false on failure.
            public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
            {
                Int32 dwError = 0, dwWritten = 0;
                IntPtr hPrinter = new IntPtr(0);
                DOCINFOA di = new DOCINFOA();
                bool bSuccess = false; // Assume failure unless you specifically succeed.

                di.pDocName = "My C#.NET RAW Document";
                di.pDataType = "RAW";

                // Open the printer.
                if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
                {
                    // Start a document.
                    if (StartDocPrinter(hPrinter, 1, di))
                    {
                        // Start a page.
                        if (StartPagePrinter(hPrinter))
                        {
                            // Write your bytes.
                            bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                            EndPagePrinter(hPrinter);
                        }
                        EndDocPrinter(hPrinter);
                    }
                    ClosePrinter(hPrinter);
                }
                // If you did not succeed, GetLastError may give more information
                // about why not.
                if (bSuccess == false)
                {
                    dwError = Marshal.GetLastWin32Error();
                }
                return bSuccess;
            }

            public static bool SendFileToPrinter(string szPrinterName, string szFileName)
            {
                // Open the file.
                FileStream fs = new FileStream(szFileName, FileMode.Open);
                // Create a BinaryReader on the file.
                BinaryReader br = new BinaryReader(fs);
                // Dim an array of bytes big enough to hold the file's contents.
                Byte[] bytes = new Byte[fs.Length];
                bool bSuccess = false;
                // Your unmanaged pointer.
                IntPtr pUnmanagedBytes = new IntPtr(0);
                int nLength;

                nLength = Convert.ToInt32(fs.Length);
                // Read the contents of the file into the array.
                bytes = br.ReadBytes(nLength);
                // Allocate some unmanaged memory for those bytes.
                pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
                // Copy the managed byte array into the unmanaged array.
                Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
                // Send the unmanaged bytes to the printer.
                bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);
                // Free the unmanaged memory that you allocated earlier.
                Marshal.FreeCoTaskMem(pUnmanagedBytes);
                return bSuccess;
            }
            public static bool SendStringToPrinter(string szPrinterName, string szString)
            {
                IntPtr pBytes;
                Int32 dwCount;
                // How many characters are in the string?
                dwCount = szString.Length;
                // Assume that the printer is expecting ANSI text, and then convert
                // the string to ANSI text.
                pBytes = Marshal.StringToCoTaskMemAnsi(szString);
                // Send the converted ANSI string to the printer.
                SendBytesToPrinter(szPrinterName, pBytes, dwCount);
                Marshal.FreeCoTaskMem(pBytes);
                return true;
            }
        }
    }
}
