using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace form_entradas_ocs
{
    public partial class calendario : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader1, reader2, reader3, reader4;
        SqlCommand cmnd4;
        SqlCommand cmnd3;
        SqlCommand cmnd2;
        SqlCommand cmnd1;

        public calendario()
        {
            InitializeComponent();

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            string fecha = "";
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT GETDATE() as fecha";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                fecha = Convert.ToDateTime(reader1["fecha"].ToString()).ToShortDateString();
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();

            DayOfWeek WK = Convert.ToDateTime(fecha).DayOfWeek;

            if (Utilerias.Class1.Usu_login != "PEDRO")
            {
                if (WK.ToString() == "Monday" || WK.ToString() == "Lunes")
                    monthCalendar1.MinDate = Convert.ToDateTime(fecha).AddDays(-2);
                else
                    monthCalendar1.MinDate = Convert.ToDateTime(fecha).AddDays(-3);
            }
            

            //monthCalendar1.MinDate = DateTime.Now.AddDays(-1);
            monthCalendar1.MaxDate = DateTime.Now;
        }

        public class datosval
        {
            private string _fecha;
            public string fecha
            {
                get { return _fecha; }
                set { _fecha = value; }
            }
        }
        public class SharedData
        {
            public static datosval Polino;
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            //MessageBox.Show(monthCalendar1.SelectionRange.Start.Date.ToShortDateString());
            datosval passdata = new datosval();
            passdata.fecha = monthCalendar1.SelectionRange.Start.Date.ToShortDateString();
            SharedData.Polino = passdata;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
