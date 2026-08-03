namespace form_entradas_ocs
{
    partial class empaques
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(empaques));
            this.dtgEmpaques = new System.Windows.Forms.DataGridView();
            this.emp_clave = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.emp_nombre = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chkClave = new System.Windows.Forms.CheckBox();
            this.chknombre = new System.Windows.Forms.CheckBox();
            this.txtnomempaque = new System.Windows.Forms.TextBox();
            this.txtempaque = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dtgEmpaques)).BeginInit();
            this.SuspendLayout();
            // 
            // dtgEmpaques
            // 
            this.dtgEmpaques.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dtgEmpaques.BackgroundColor = System.Drawing.Color.White;
            this.dtgEmpaques.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dtgEmpaques.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.emp_clave,
            this.emp_nombre});
            this.dtgEmpaques.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dtgEmpaques.Location = new System.Drawing.Point(0, 56);
            this.dtgEmpaques.Name = "dtgEmpaques";
            this.dtgEmpaques.RowHeadersVisible = false;
            this.dtgEmpaques.Size = new System.Drawing.Size(476, 449);
            this.dtgEmpaques.TabIndex = 0;
            // 
            // emp_clave
            // 
            this.emp_clave.HeaderText = "Clave";
            this.emp_clave.Name = "emp_clave";
            this.emp_clave.Width = 61;
            // 
            // emp_nombre
            // 
            this.emp_nombre.HeaderText = "Nombre";
            this.emp_nombre.Name = "emp_nombre";
            this.emp_nombre.Width = 76;
            // 
            // chkClave
            // 
            this.chkClave.AutoSize = true;
            this.chkClave.Location = new System.Drawing.Point(60, 22);
            this.chkClave.Name = "chkClave";
            this.chkClave.Size = new System.Drawing.Size(15, 14);
            this.chkClave.TabIndex = 340;
            this.chkClave.UseVisualStyleBackColor = true;
            this.chkClave.CheckStateChanged += new System.EventHandler(this.chkClave_CheckStateChanged);
            // 
            // chknombre
            // 
            this.chknombre.AutoSize = true;
            this.chknombre.Location = new System.Drawing.Point(207, 23);
            this.chknombre.Name = "chknombre";
            this.chknombre.Size = new System.Drawing.Size(15, 14);
            this.chknombre.TabIndex = 339;
            this.chknombre.UseVisualStyleBackColor = true;
            this.chknombre.CheckStateChanged += new System.EventHandler(this.chknombre_CheckStateChanged);
            // 
            // txtnomempaque
            // 
            this.txtnomempaque.BackColor = System.Drawing.SystemColors.Window;
            this.txtnomempaque.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtnomempaque.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtnomempaque.Font = new System.Drawing.Font("Lucida Sans Unicode", 6.5F);
            this.txtnomempaque.Location = new System.Drawing.Point(224, 18);
            this.txtnomempaque.Name = "txtnomempaque";
            this.txtnomempaque.Size = new System.Drawing.Size(202, 21);
            this.txtnomempaque.TabIndex = 338;
            this.txtnomempaque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtnomempaque_KeyPress);
            // 
            // txtempaque
            // 
            this.txtempaque.BackColor = System.Drawing.SystemColors.Window;
            this.txtempaque.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtempaque.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtempaque.Font = new System.Drawing.Font("Lucida Sans Unicode", 6.5F);
            this.txtempaque.Location = new System.Drawing.Point(81, 18);
            this.txtempaque.Name = "txtempaque";
            this.txtempaque.Size = new System.Drawing.Size(107, 21);
            this.txtempaque.TabIndex = 337;
            this.txtempaque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtempaque_KeyPress);
            // 
            // empaques
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.ClientSize = new System.Drawing.Size(476, 505);
            this.Controls.Add(this.chkClave);
            this.Controls.Add(this.chknombre);
            this.Controls.Add(this.txtnomempaque);
            this.Controls.Add(this.txtempaque);
            this.Controls.Add(this.dtgEmpaques);
            this.Font = new System.Drawing.Font("Lucida Sans Unicode", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "empaques";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Empaques";
            ((System.ComponentModel.ISupportInitialize)(this.dtgEmpaques)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dtgEmpaques;
        private System.Windows.Forms.DataGridViewTextBoxColumn emp_clave;
        private System.Windows.Forms.DataGridViewTextBoxColumn emp_nombre;
        private System.Windows.Forms.CheckBox chkClave;
        private System.Windows.Forms.CheckBox chknombre;
        private System.Windows.Forms.TextBox txtnomempaque;
        private System.Windows.Forms.TextBox txtempaque;
    }
}