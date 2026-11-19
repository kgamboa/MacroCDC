namespace TesoreriaCobranza
{
    partial class frmTesoreriaCobranza
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmTesoreriaCobranza));
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.bancosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.generaArchivoBancosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.categorizaBancosMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.ingresosPequeñosToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cobranzaToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.identificaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.identificaFacturasCLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.conciliaciónToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tesoreriaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cobranzaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.polizaSemanalToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cerrarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bancosToolStripMenuItem,
            this.cobranzaToolStripMenuItem1,
            this.conciliaciónToolStripMenuItem,
            this.polizaSemanalToolStripMenuItem,
            this.cerrarToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(197, 254);
            // 
            // bancosToolStripMenuItem
            // 
            this.bancosToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.generaArchivoBancosToolStripMenuItem,
            this.categorizaBancosMenuItem2,
            this.ingresosPequeñosToolStripMenuItem});
            this.bancosToolStripMenuItem.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.bancosToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.bancosToolStripMenuItem.Name = "bancosToolStripMenuItem";
            this.bancosToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.F1)));
            this.bancosToolStripMenuItem.ShowShortcutKeys = false;
            this.bancosToolStripMenuItem.Size = new System.Drawing.Size(196, 38);
            this.bancosToolStripMenuItem.Text = "Bancos";
            // 
            // generaArchivoBancosToolStripMenuItem
            // 
            this.generaArchivoBancosToolStripMenuItem.Enabled = false;
            this.generaArchivoBancosToolStripMenuItem.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.generaArchivoBancosToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.generaArchivoBancosToolStripMenuItem.Name = "generaArchivoBancosToolStripMenuItem";
            this.generaArchivoBancosToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.B)));
            this.generaArchivoBancosToolStripMenuItem.ShowShortcutKeys = false;
            this.generaArchivoBancosToolStripMenuItem.Size = new System.Drawing.Size(228, 38);
            this.generaArchivoBancosToolStripMenuItem.Text = "Generar Archivo Bancos";
            this.generaArchivoBancosToolStripMenuItem.Visible = false;
            this.generaArchivoBancosToolStripMenuItem.Click += new System.EventHandler(this.generaArchivoBancosToolStripMenuItem_Click);
            // 
            // categorizaBancosMenuItem2
            // 
            this.categorizaBancosMenuItem2.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.categorizaBancosMenuItem2.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.categorizaBancosMenuItem2.Name = "categorizaBancosMenuItem2";
            this.categorizaBancosMenuItem2.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.B)));
            this.categorizaBancosMenuItem2.ShowShortcutKeys = false;
            this.categorizaBancosMenuItem2.Size = new System.Drawing.Size(228, 38);
            this.categorizaBancosMenuItem2.Text = "Categorizar Archivo Bancos";
            this.categorizaBancosMenuItem2.Click += new System.EventHandler(this.categorizaBancosMenuItem2_Click);
            // 
            // ingresosPequeñosToolStripMenuItem
            // 
            this.ingresosPequeñosToolStripMenuItem.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.ingresosPequeñosToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.ingresosPequeñosToolStripMenuItem.Name = "ingresosPequeñosToolStripMenuItem";
            this.ingresosPequeñosToolStripMenuItem.Size = new System.Drawing.Size(228, 38);
            this.ingresosPequeñosToolStripMenuItem.Text = "Ingresos Pequeños";
            this.ingresosPequeñosToolStripMenuItem.Click += new System.EventHandler(this.ingresosPequeñosToolStripMenuItem_Click);
            // 
            // cobranzaToolStripMenuItem1
            // 
            this.cobranzaToolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.identificaToolStripMenuItem,
            this.identificaFacturasCLToolStripMenuItem });
            this.cobranzaToolStripMenuItem1.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.cobranzaToolStripMenuItem1.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.cobranzaToolStripMenuItem1.Name = "cobranzaToolStripMenuItem1";
            this.cobranzaToolStripMenuItem1.Size = new System.Drawing.Size(196, 38);
            this.cobranzaToolStripMenuItem1.Text = "Cobranza";
            // 
            // identificaToolStripMenuItem
            // 
            this.identificaToolStripMenuItem.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.identificaToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.identificaToolStripMenuItem.Name = "identificaToolStripMenuItem";
            this.identificaToolStripMenuItem.Size = new System.Drawing.Size(205, 38);
            this.identificaToolStripMenuItem.Text = "Identifica Facturas CC";
            this.identificaToolStripMenuItem.Click += new System.EventHandler(this.identificaToolStripMenuItem_Click);
            // 
            // identificaFacturasCLToolStripMenuItem
            // 
            this.identificaFacturasCLToolStripMenuItem.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.identificaFacturasCLToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.identificaFacturasCLToolStripMenuItem.Name = "identificaFacturasCLToolStripMenuItem";
            this.identificaFacturasCLToolStripMenuItem.Size = new System.Drawing.Size(205, 38);
            this.identificaFacturasCLToolStripMenuItem.Text = "Identifica Facturas CL";
            this.identificaFacturasCLToolStripMenuItem.Click += new System.EventHandler(this.identificaFacturasCLToolStripMenuItem_Click);
            //
            // 
            // conciliaciónToolStripMenuItem
            // 
            this.conciliaciónToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.contaToolStripMenuItem,
            this.tesoreriaToolStripMenuItem,
            this.cobranzaToolStripMenuItem});
            this.conciliaciónToolStripMenuItem.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.conciliaciónToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.conciliaciónToolStripMenuItem.Name = "conciliaciónToolStripMenuItem";
            this.conciliaciónToolStripMenuItem.Size = new System.Drawing.Size(196, 38);
            this.conciliaciónToolStripMenuItem.Text = "Conciliación";
            this.conciliaciónToolStripMenuItem.Visible = false;
            // 
            // contaToolStripMenuItem
            // 
            this.contaToolStripMenuItem.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.contaToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.contaToolStripMenuItem.Name = "contaToolStripMenuItem";
            this.contaToolStripMenuItem.Size = new System.Drawing.Size(196, 38);
            this.contaToolStripMenuItem.Text = "Contabilidad";
            this.contaToolStripMenuItem.Click += new System.EventHandler(this.contaToolStripMenuItem_Click);
            // 
            // tesoreriaToolStripMenuItem
            // 
            this.tesoreriaToolStripMenuItem.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.tesoreriaToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.tesoreriaToolStripMenuItem.Name = "tesoreriaToolStripMenuItem";
            this.tesoreriaToolStripMenuItem.Size = new System.Drawing.Size(196, 38);
            this.tesoreriaToolStripMenuItem.Text = "Tesoreria";
            this.tesoreriaToolStripMenuItem.Click += new System.EventHandler(this.tesoreriaToolStripMenuItem_Click);
            // 
            // cobranzaToolStripMenuItem
            // 
            this.cobranzaToolStripMenuItem.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.cobranzaToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.cobranzaToolStripMenuItem.Name = "cobranzaToolStripMenuItem";
            this.cobranzaToolStripMenuItem.Size = new System.Drawing.Size(196, 38);
            this.cobranzaToolStripMenuItem.Text = "Facturación";
            this.cobranzaToolStripMenuItem.Click += new System.EventHandler(this.cobranzaToolStripMenuItem_Click);
            // 
            // polizaSemanalToolStripMenuItem
            // 
            this.polizaSemanalToolStripMenuItem.Image = global::PTC.Properties.Resources._1486564397_menu_green_81507__1_;
            this.polizaSemanalToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.polizaSemanalToolStripMenuItem.Name = "polizaSemanalToolStripMenuItem";
            this.polizaSemanalToolStripMenuItem.Size = new System.Drawing.Size(196, 38);
            this.polizaSemanalToolStripMenuItem.Text = "Polizas";
            this.polizaSemanalToolStripMenuItem.ToolTipText = "Comisiones, ISR e Intereses";
            this.polizaSemanalToolStripMenuItem.Click += new System.EventHandler(this.polizaSemanalToolStripMenuItem_Click);
            // 
            // cerrarToolStripMenuItem
            // 
            this.cerrarToolStripMenuItem.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.cerrarToolStripMenuItem.Image = global::PTC.Properties.Resources.systemshutdown_103390;
            this.cerrarToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.cerrarToolStripMenuItem.Name = "cerrarToolStripMenuItem";
            this.cerrarToolStripMenuItem.Size = new System.Drawing.Size(196, 38);
            this.cerrarToolStripMenuItem.Text = "Cerrar";
            this.cerrarToolStripMenuItem.Click += new System.EventHandler(this.cerrarToolStripMenuItem_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(12, 201);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(149, 17);
            this.checkBox1.TabIndex = 4;
            this.checkBox1.Text = "Abrir despues de procesar";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.notifyIcon1.BalloonTipTitle = "Círculo Crédito";
            this.notifyIcon1.ContextMenuStrip = this.contextMenuStrip1;
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "Procesos Tesoreria y Cobranza";
            this.notifyIcon1.Visible = true;
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(32, 19);
            // 
            // frmTesoreriaCobranza
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(282, 240);
            this.Controls.Add(this.checkBox1);
            this.Enabled = false;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmTesoreriaCobranza";
            this.Opacity = 0D;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.frmBancoCategoria_Load);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.ToolStripMenuItem generaArchivoBancosToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cerrarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem conciliaciónToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem contaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tesoreriaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cobranzaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem bancosToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem categorizaBancosMenuItem2;
        //private System.Windows.Forms.ToolStripMenuItem polizasToolStripMenuItem;
        //private System.Windows.Forms.ToolStripMenuItem polizaBancosToolStripMenuItem;
        //private System.Windows.Forms.ToolStripMenuItem polizaCLBANCOSToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem identificaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cobranzaToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem identificaFacturasCLToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ingresosPequeñosToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem polizaSemanalToolStripMenuItem;
    }
}

