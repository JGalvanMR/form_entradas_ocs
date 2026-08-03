using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data.SqlClient;

namespace form_entradas_ocs
{
    public partial class empaques : Form
    {
        private const int EM_SETCUEBANNER = 0x1501;
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern Int32 SendMessage(IntPtr hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)]string lParam);

        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader1;
        SqlCommand cmnd1;

        string almacen = "";
        DataTable dtempaques = new DataTable();

        public empaques(string alm)
        {
            InitializeComponent();

            SendMessage(txtempaque.Handle, EM_SETCUEBANNER, 0, "CLAVE");
            SendMessage(txtnomempaque.Handle, EM_SETCUEBANNER, 0, "NOMBRE");

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            dtempaques.Columns.Add("emp_clave", typeof(string));
            dtempaques.Columns.Add("emp_nombre", typeof(string));

            almacen = alm;

            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT emp_clave, emp_nombre FROM tb_cat_empaques WHERE alm_clave = '" + almacen + "' AND emp_nombre <> '' ORDER BY emp_clave";
            reader1 = cmnd1.ExecuteReader();
            DataRow r;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    r = dtempaques.NewRow();
                    r["emp_clave"] = reader1.GetValue(0).ToString().Trim();
                    r["emp_nombre"] = reader1.GetValue(1).ToString().Trim();
                    dtempaques.Rows.Add(r);
                }
            }
            thisConnection.Close();

            foreach (DataRow rw in dtempaques.Rows)
            {
                dtgEmpaques.Rows.Add(rw.ItemArray);
            }
        }

        private void chkClave_CheckStateChanged(object sender, EventArgs e)
        {
            if (chkClave.Checked == true)
            {
                dtgEmpaques.Rows.Clear();
                chknombre.Checked = false;
                DataView dw = dtempaques.DefaultView;
                dw.Sort = "emp_clave";
                dtempaques = dw.ToTable();

                foreach (DataRow rw in dtempaques.Rows)
                {
                    dtgEmpaques.Rows.Add(rw.ItemArray);
                }
            }
        }

        private void chknombre_CheckStateChanged(object sender, EventArgs e)
        {
            if (chknombre.Checked == true)
            {
                dtgEmpaques.Rows.Clear();
                chkClave.Checked = false;
                DataView dw = dtempaques.DefaultView;
                dw.Sort = "emp_nombre";
                dtempaques = dw.ToTable();

                foreach (DataRow rw in dtempaques.Rows)
                {
                    dtgEmpaques.Rows.Add(rw.ItemArray);
                }
            }
        }

        private void txtempaque_KeyPress(object sender, KeyPressEventArgs e)
        {
            for (int y = 0; y < dtempaques.Rows.Count; y++)
            {
                if (dtempaques.Rows[y][0].ToString().StartsWith(txtempaque.Text))
                {
                    //dtgProductos.Rows[y].Selected = true;
                    dtgEmpaques.FirstDisplayedScrollingRowIndex = y;
                    break;
                }
            }
        }

        private void txtnomempaque_KeyPress(object sender, KeyPressEventArgs e)
        {
            for (int y = 0; y < dtempaques.Rows.Count; y++)
            {
                if (dtempaques.Rows[y][1].ToString().StartsWith(txtnomempaque.Text))
                {
                    //dtgProductos.Rows[y].Selected = true;
                    dtgEmpaques.FirstDisplayedScrollingRowIndex = y;
                    break;
                }
            }
        }
    }
}
