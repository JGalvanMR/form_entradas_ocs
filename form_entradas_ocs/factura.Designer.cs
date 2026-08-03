namespace form_entradas_ocs
{
    partial class factura
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblEntrada = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblClaveAlm = new System.Windows.Forms.Label();
            this.lblAlmacen = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtNomina = new System.Windows.Forms.TextBox();
            this.txtObservaciones = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.lblDestino = new System.Windows.Forms.Label();
            this.btnCierra = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.lblPathXml = new System.Windows.Forms.Label();
            this.lblFolioFiscal = new System.Windows.Forms.Label();
            this.lbl_ent_factura = new System.Windows.Forms.Label();
            this.lblFechaFact = new System.Windows.Forms.Label();
            this.lblTotal = new System.Windows.Forms.Label();
            this.lblRFC = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lblClaveProveedor = new System.Windows.Forms.Label();
            this.lblProveedor = new System.Windows.Forms.Label();
            this.lblPortal = new System.Windows.Forms.Label();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.lblFactActual = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.lblFolioFiscalAct = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(287, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "CAMBIAR FACTURA";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(9, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 15);
            this.label2.TabIndex = 1;
            this.label2.Text = "Entrada:";
            // 
            // lblEntrada
            // 
            this.lblEntrada.AutoSize = true;
            this.lblEntrada.BackColor = System.Drawing.Color.Transparent;
            this.lblEntrada.ForeColor = System.Drawing.Color.White;
            this.lblEntrada.Location = new System.Drawing.Point(66, 41);
            this.lblEntrada.Name = "lblEntrada";
            this.lblEntrada.Size = new System.Drawing.Size(13, 15);
            this.lblEntrada.TabIndex = 2;
            this.lblEntrada.Text = "-";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(161, 41);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(57, 15);
            this.label3.TabIndex = 3;
            this.label3.Text = "Almacen:";
            // 
            // lblClaveAlm
            // 
            this.lblClaveAlm.AutoSize = true;
            this.lblClaveAlm.BackColor = System.Drawing.Color.Transparent;
            this.lblClaveAlm.ForeColor = System.Drawing.Color.White;
            this.lblClaveAlm.Location = new System.Drawing.Point(224, 41);
            this.lblClaveAlm.Name = "lblClaveAlm";
            this.lblClaveAlm.Size = new System.Drawing.Size(13, 15);
            this.lblClaveAlm.TabIndex = 4;
            this.lblClaveAlm.Text = "-";
            // 
            // lblAlmacen
            // 
            this.lblAlmacen.AutoSize = true;
            this.lblAlmacen.BackColor = System.Drawing.Color.Transparent;
            this.lblAlmacen.ForeColor = System.Drawing.Color.White;
            this.lblAlmacen.Location = new System.Drawing.Point(259, 41);
            this.lblAlmacen.Name = "lblAlmacen";
            this.lblAlmacen.Size = new System.Drawing.Size(13, 15);
            this.lblAlmacen.TabIndex = 5;
            this.lblAlmacen.Text = "-";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(9, 80);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(52, 15);
            this.label4.TabIndex = 6;
            this.label4.Text = "Nómina:";
            // 
            // txtNomina
            // 
            this.txtNomina.Location = new System.Drawing.Point(67, 77);
            this.txtNomina.MaxLength = 4;
            this.txtNomina.Name = "txtNomina";
            this.txtNomina.Size = new System.Drawing.Size(65, 24);
            this.txtNomina.TabIndex = 7;
            // 
            // txtObservaciones
            // 
            this.txtObservaciones.Location = new System.Drawing.Point(12, 143);
            this.txtObservaciones.MaxLength = 100;
            this.txtObservaciones.Name = "txtObservaciones";
            this.txtObservaciones.Size = new System.Drawing.Size(660, 24);
            this.txtObservaciones.TabIndex = 8;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(9, 125);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(108, 15);
            this.label5.TabIndex = 9;
            this.label5.Text = "¿Por qué modifica?";
            // 
            // lblDestino
            // 
            this.lblDestino.AutoSize = true;
            this.lblDestino.BackColor = System.Drawing.Color.Transparent;
            this.lblDestino.Font = new System.Drawing.Font("Lucida Sans Unicode", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDestino.ForeColor = System.Drawing.Color.White;
            this.lblDestino.Location = new System.Drawing.Point(10, 268);
            this.lblDestino.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDestino.Name = "lblDestino";
            this.lblDestino.Size = new System.Drawing.Size(106, 11);
            this.lblDestino.TabIndex = 343;
            this.lblDestino.Text = "\\\\gabira1\\FacturasExternos\\";
            // 
            // btnCierra
            // 
            this.btnCierra.BackColor = System.Drawing.Color.Transparent;
            this.btnCierra.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnCierra.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnCierra.FlatAppearance.BorderSize = 2;
            this.btnCierra.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black;
            this.btnCierra.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnCierra.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCierra.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCierra.ForeColor = System.Drawing.Color.White;
            this.btnCierra.Location = new System.Drawing.Point(13, 181);
            this.btnCierra.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnCierra.Name = "btnCierra";
            this.btnCierra.Size = new System.Drawing.Size(658, 30);
            this.btnCierra.TabIndex = 354;
            this.btnCierra.Text = "Cargar XML";
            this.btnCierra.UseVisualStyleBackColor = false;
            this.btnCierra.Click += new System.EventHandler(this.btnCierra_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // lblPathXml
            // 
            this.lblPathXml.AutoSize = true;
            this.lblPathXml.BackColor = System.Drawing.Color.Transparent;
            this.lblPathXml.Font = new System.Drawing.Font("Lucida Sans Unicode", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPathXml.ForeColor = System.Drawing.Color.White;
            this.lblPathXml.Location = new System.Drawing.Point(13, 290);
            this.lblPathXml.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblPathXml.Name = "lblPathXml";
            this.lblPathXml.Size = new System.Drawing.Size(9, 11);
            this.lblPathXml.TabIndex = 355;
            this.lblPathXml.Text = "-";
            // 
            // lblFolioFiscal
            // 
            this.lblFolioFiscal.AutoSize = true;
            this.lblFolioFiscal.BackColor = System.Drawing.Color.Transparent;
            this.lblFolioFiscal.Font = new System.Drawing.Font("Lucida Sans Unicode", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFolioFiscal.ForeColor = System.Drawing.Color.White;
            this.lblFolioFiscal.Location = new System.Drawing.Point(13, 314);
            this.lblFolioFiscal.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblFolioFiscal.Name = "lblFolioFiscal";
            this.lblFolioFiscal.Size = new System.Drawing.Size(9, 11);
            this.lblFolioFiscal.TabIndex = 356;
            this.lblFolioFiscal.Text = "-";
            // 
            // lbl_ent_factura
            // 
            this.lbl_ent_factura.AutoSize = true;
            this.lbl_ent_factura.BackColor = System.Drawing.Color.Transparent;
            this.lbl_ent_factura.Font = new System.Drawing.Font("Lucida Sans Unicode", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_ent_factura.ForeColor = System.Drawing.Color.White;
            this.lbl_ent_factura.Location = new System.Drawing.Point(13, 336);
            this.lbl_ent_factura.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbl_ent_factura.Name = "lbl_ent_factura";
            this.lbl_ent_factura.Size = new System.Drawing.Size(9, 11);
            this.lbl_ent_factura.TabIndex = 357;
            this.lbl_ent_factura.Text = "-";
            // 
            // lblFechaFact
            // 
            this.lblFechaFact.AutoSize = true;
            this.lblFechaFact.BackColor = System.Drawing.Color.Transparent;
            this.lblFechaFact.Font = new System.Drawing.Font("Lucida Sans Unicode", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFechaFact.ForeColor = System.Drawing.Color.White;
            this.lblFechaFact.Location = new System.Drawing.Point(13, 360);
            this.lblFechaFact.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblFechaFact.Name = "lblFechaFact";
            this.lblFechaFact.Size = new System.Drawing.Size(9, 11);
            this.lblFechaFact.TabIndex = 358;
            this.lblFechaFact.Text = "-";
            // 
            // lblTotal
            // 
            this.lblTotal.AutoSize = true;
            this.lblTotal.BackColor = System.Drawing.Color.Transparent;
            this.lblTotal.Font = new System.Drawing.Font("Lucida Sans Unicode", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTotal.ForeColor = System.Drawing.Color.White;
            this.lblTotal.Location = new System.Drawing.Point(13, 385);
            this.lblTotal.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTotal.Name = "lblTotal";
            this.lblTotal.Size = new System.Drawing.Size(9, 11);
            this.lblTotal.TabIndex = 359;
            this.lblTotal.Text = "-";
            // 
            // lblRFC
            // 
            this.lblRFC.AutoSize = true;
            this.lblRFC.BackColor = System.Drawing.Color.Transparent;
            this.lblRFC.Font = new System.Drawing.Font("Lucida Sans Unicode", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRFC.ForeColor = System.Drawing.Color.White;
            this.lblRFC.Location = new System.Drawing.Point(13, 407);
            this.lblRFC.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblRFC.Name = "lblRFC";
            this.lblRFC.Size = new System.Drawing.Size(9, 11);
            this.lblRFC.TabIndex = 360;
            this.lblRFC.Text = "-";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(161, 80);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(65, 15);
            this.label6.TabIndex = 361;
            this.label6.Text = "Proveedor:";
            // 
            // lblClaveProveedor
            // 
            this.lblClaveProveedor.AutoSize = true;
            this.lblClaveProveedor.BackColor = System.Drawing.Color.Transparent;
            this.lblClaveProveedor.ForeColor = System.Drawing.Color.White;
            this.lblClaveProveedor.Location = new System.Drawing.Point(232, 80);
            this.lblClaveProveedor.Name = "lblClaveProveedor";
            this.lblClaveProveedor.Size = new System.Drawing.Size(13, 15);
            this.lblClaveProveedor.TabIndex = 362;
            this.lblClaveProveedor.Text = "-";
            // 
            // lblProveedor
            // 
            this.lblProveedor.AutoSize = true;
            this.lblProveedor.BackColor = System.Drawing.Color.Transparent;
            this.lblProveedor.ForeColor = System.Drawing.Color.White;
            this.lblProveedor.Location = new System.Drawing.Point(276, 80);
            this.lblProveedor.Name = "lblProveedor";
            this.lblProveedor.Size = new System.Drawing.Size(13, 15);
            this.lblProveedor.TabIndex = 363;
            this.lblProveedor.Text = "-";
            // 
            // lblPortal
            // 
            this.lblPortal.AutoSize = true;
            this.lblPortal.BackColor = System.Drawing.Color.Transparent;
            this.lblPortal.Font = new System.Drawing.Font("Lucida Sans Unicode", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPortal.ForeColor = System.Drawing.Color.White;
            this.lblPortal.Location = new System.Drawing.Point(147, 407);
            this.lblPortal.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblPortal.Name = "lblPortal";
            this.lblPortal.Size = new System.Drawing.Size(9, 11);
            this.lblPortal.TabIndex = 364;
            this.lblPortal.Text = "-";
            // 
            // btnGuardar
            // 
            this.btnGuardar.BackColor = System.Drawing.Color.Transparent;
            this.btnGuardar.BackgroundImage = global::form_entradas_ocs.Properties.Resources.appbar_save;
            this.btnGuardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnGuardar.Enabled = false;
            this.btnGuardar.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnGuardar.FlatAppearance.BorderSize = 2;
            this.btnGuardar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black;
            this.btnGuardar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGuardar.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGuardar.ForeColor = System.Drawing.Color.Black;
            this.btnGuardar.Location = new System.Drawing.Point(317, 217);
            this.btnGuardar.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(51, 49);
            this.btnGuardar.TabIndex = 365;
            this.btnGuardar.UseVisualStyleBackColor = false;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // lblFactActual
            // 
            this.lblFactActual.AutoSize = true;
            this.lblFactActual.BackColor = System.Drawing.Color.Transparent;
            this.lblFactActual.ForeColor = System.Drawing.Color.White;
            this.lblFactActual.Location = new System.Drawing.Point(555, 114);
            this.lblFactActual.Name = "lblFactActual";
            this.lblFactActual.Size = new System.Drawing.Size(13, 15);
            this.lblFactActual.TabIndex = 366;
            this.lblFactActual.Text = "-";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(461, 114);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(85, 15);
            this.label8.TabIndex = 367;
            this.label8.Text = "Factura actual:";
            // 
            // lblFolioFiscalAct
            // 
            this.lblFolioFiscalAct.AutoSize = true;
            this.lblFolioFiscalAct.BackColor = System.Drawing.Color.Transparent;
            this.lblFolioFiscalAct.Font = new System.Drawing.Font("Lucida Sans Unicode", 5.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFolioFiscalAct.ForeColor = System.Drawing.Color.White;
            this.lblFolioFiscalAct.Location = new System.Drawing.Point(556, 90);
            this.lblFolioFiscalAct.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblFolioFiscalAct.Name = "lblFolioFiscalAct";
            this.lblFolioFiscalAct.Size = new System.Drawing.Size(9, 11);
            this.lblFolioFiscalAct.TabIndex = 368;
            this.lblFolioFiscalAct.Text = "-";
            // 
            // factura
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkBlue;
            this.ClientSize = new System.Drawing.Size(684, 430);
            this.Controls.Add(this.lblFolioFiscalAct);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.lblFactActual);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.lblPortal);
            this.Controls.Add(this.lblProveedor);
            this.Controls.Add(this.lblClaveProveedor);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblRFC);
            this.Controls.Add(this.lblTotal);
            this.Controls.Add(this.lblFechaFact);
            this.Controls.Add(this.lbl_ent_factura);
            this.Controls.Add(this.lblFolioFiscal);
            this.Controls.Add(this.lblPathXml);
            this.Controls.Add(this.btnCierra);
            this.Controls.Add(this.lblDestino);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtObservaciones);
            this.Controls.Add(this.txtNomina);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblAlmacen);
            this.Controls.Add(this.lblClaveAlm);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblEntrada);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "factura";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Modificación de factura";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblEntrada;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblClaveAlm;
        private System.Windows.Forms.Label lblAlmacen;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtNomina;
        private System.Windows.Forms.TextBox txtObservaciones;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblDestino;
        private System.Windows.Forms.Button btnCierra;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label lblPathXml;
        private System.Windows.Forms.Label lblFolioFiscal;
        private System.Windows.Forms.Label lbl_ent_factura;
        private System.Windows.Forms.Label lblFechaFact;
        private System.Windows.Forms.Label lblTotal;
        private System.Windows.Forms.Label lblRFC;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblClaveProveedor;
        private System.Windows.Forms.Label lblProveedor;
        private System.Windows.Forms.Label lblPortal;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Label lblFactActual;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label lblFolioFiscalAct;
    }
}