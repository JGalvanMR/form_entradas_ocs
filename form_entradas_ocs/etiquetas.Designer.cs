namespace form_entradas_ocs
{
    partial class etiquetas
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dtgEmpaques = new System.Windows.Forms.DataGridView();
            this.emp_clave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.emp_nombre = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ent_cantidad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.imprimir = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnEtiquetas = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lblEntrada = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblAlmacen = new System.Windows.Forms.Label();
            this.lblNombre = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dtgEmpaques)).BeginInit();
            this.SuspendLayout();
            // 
            // dtgEmpaques
            // 
            this.dtgEmpaques.AllowUserToAddRows = false;
            this.dtgEmpaques.AllowUserToDeleteRows = false;
            this.dtgEmpaques.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dtgEmpaques.BackgroundColor = System.Drawing.Color.White;
            this.dtgEmpaques.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtgEmpaques.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.emp_clave,
            this.emp_nombre,
            this.ent_cantidad,
            this.imprimir});
            this.dtgEmpaques.GridColor = System.Drawing.Color.Black;
            this.dtgEmpaques.Location = new System.Drawing.Point(12, 63);
            this.dtgEmpaques.Name = "dtgEmpaques";
            this.dtgEmpaques.RowHeadersVisible = false;
            this.dtgEmpaques.Size = new System.Drawing.Size(790, 305);
            this.dtgEmpaques.TabIndex = 0;
            this.dtgEmpaques.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dtgEmpaques_CellEndEdit);
            // 
            // emp_clave
            // 
            this.emp_clave.HeaderText = "Clave";
            this.emp_clave.Name = "emp_clave";
            this.emp_clave.ReadOnly = true;
            this.emp_clave.Width = 61;
            // 
            // emp_nombre
            // 
            this.emp_nombre.HeaderText = "Nombre";
            this.emp_nombre.Name = "emp_nombre";
            this.emp_nombre.ReadOnly = true;
            this.emp_nombre.Width = 76;
            // 
            // ent_cantidad
            // 
            this.ent_cantidad.HeaderText = "Cantidad";
            this.ent_cantidad.Name = "ent_cantidad";
            this.ent_cantidad.ReadOnly = true;
            this.ent_cantidad.Width = 80;
            // 
            // imprimir
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.imprimir.DefaultCellStyle = dataGridViewCellStyle1;
            this.imprimir.HeaderText = "Imprimir";
            this.imprimir.Name = "imprimir";
            this.imprimir.Width = 80;
            // 
            // btnEtiquetas
            // 
            this.btnEtiquetas.BackColor = System.Drawing.Color.Transparent;
            this.btnEtiquetas.BackgroundImage = global::form_entradas_ocs.Properties.Resources.appbar_tag;
            this.btnEtiquetas.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnEtiquetas.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnEtiquetas.FlatAppearance.BorderSize = 2;
            this.btnEtiquetas.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black;
            this.btnEtiquetas.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnEtiquetas.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnEtiquetas.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEtiquetas.ForeColor = System.Drawing.Color.Black;
            this.btnEtiquetas.Location = new System.Drawing.Point(378, 382);
            this.btnEtiquetas.Margin = new System.Windows.Forms.Padding(5, 3, 5, 3);
            this.btnEtiquetas.Name = "btnEtiquetas";
            this.btnEtiquetas.Size = new System.Drawing.Size(59, 57);
            this.btnEtiquetas.TabIndex = 358;
            this.btnEtiquetas.UseVisualStyleBackColor = false;
            this.btnEtiquetas.Click += new System.EventHandler(this.btnEtiquetas_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(14, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 15);
            this.label1.TabIndex = 359;
            this.label1.Text = "Entrada:";
            // 
            // lblEntrada
            // 
            this.lblEntrada.AutoSize = true;
            this.lblEntrada.ForeColor = System.Drawing.Color.White;
            this.lblEntrada.Location = new System.Drawing.Point(71, 9);
            this.lblEntrada.Name = "lblEntrada";
            this.lblEntrada.Size = new System.Drawing.Size(13, 15);
            this.lblEntrada.TabIndex = 360;
            this.lblEntrada.Text = "-";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(14, 34);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 15);
            this.label2.TabIndex = 361;
            this.label2.Text = "Almacen:";
            // 
            // lblAlmacen
            // 
            this.lblAlmacen.AutoSize = true;
            this.lblAlmacen.ForeColor = System.Drawing.Color.White;
            this.lblAlmacen.Location = new System.Drawing.Point(77, 34);
            this.lblAlmacen.Name = "lblAlmacen";
            this.lblAlmacen.Size = new System.Drawing.Size(13, 15);
            this.lblAlmacen.TabIndex = 362;
            this.lblAlmacen.Text = "-";
            // 
            // lblNombre
            // 
            this.lblNombre.AutoSize = true;
            this.lblNombre.ForeColor = System.Drawing.Color.White;
            this.lblNombre.Location = new System.Drawing.Point(114, 34);
            this.lblNombre.Name = "lblNombre";
            this.lblNombre.Size = new System.Drawing.Size(13, 15);
            this.lblNombre.TabIndex = 363;
            this.lblNombre.Text = "-";
            // 
            // etiquetas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DarkBlue;
            this.ClientSize = new System.Drawing.Size(814, 451);
            this.Controls.Add(this.lblNombre);
            this.Controls.Add(this.lblAlmacen);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblEntrada);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnEtiquetas);
            this.Controls.Add(this.dtgEmpaques);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "etiquetas";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Generar Etiquetas para Productos";
            ((System.ComponentModel.ISupportInitialize)(this.dtgEmpaques)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dtgEmpaques;
        private System.Windows.Forms.Button btnEtiquetas;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblEntrada;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblAlmacen;
        private System.Windows.Forms.Label lblNombre;
        private System.Windows.Forms.DataGridViewTextBoxColumn emp_clave;
        private System.Windows.Forms.DataGridViewTextBoxColumn emp_nombre;
        private System.Windows.Forms.DataGridViewTextBoxColumn ent_cantidad;
        private System.Windows.Forms.DataGridViewTextBoxColumn imprimir;
    }
}