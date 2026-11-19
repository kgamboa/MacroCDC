using System;
using System.Globalization;

namespace PTC
{
    partial class frmGeneradorPolizas
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
        private System.Windows.Forms.MonthCalendar dateTimePicker1;
        private System.Windows.Forms.Label lblRango;

        private void InitializeComponent()
        {
            this.dateTimePicker1 = new System.Windows.Forms.MonthCalendar();
            this.lblRango = new System.Windows.Forms.Label();
            this.btnGenerarPolizas = new System.Windows.Forms.Button();
            this.chkListTipoPoliza = new System.Windows.Forms.CheckedListBox();
            this.chkTipoAgrupador = new System.Windows.Forms.CheckBox();
            this.chkTipoCalendario = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.AllowDrop = true;
            this.dateTimePicker1.Dock = System.Windows.Forms.DockStyle.Left;
            this.dateTimePicker1.FirstDayOfWeek = System.Windows.Forms.Day.Monday;
            this.dateTimePicker1.Location = new System.Drawing.Point(0, 0);
            this.dateTimePicker1.MaxSelectionCount = 1;
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.ScrollChange = 1;
            this.dateTimePicker1.ShowTodayCircle = false;
            this.dateTimePicker1.TabIndex = 0;
            this.dateTimePicker1.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.dateTimePicker1_DateChanged);
            this.dateTimePicker1.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.dateTimePicker1_DateSelected);
            this.dateTimePicker1.MouseCaptureChanged += new System.EventHandler(this.dateTimePicker1_MouseCaptureChanged);
            this.dateTimePicker1.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.dateTimePicker1_PreviewKeyDown);
            // 
            // lblRango
            // 
            this.lblRango.BackColor = System.Drawing.SystemColors.Menu;
            this.lblRango.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblRango.Location = new System.Drawing.Point(248, 135);
            this.lblRango.Name = "lblRango";
            this.lblRango.Size = new System.Drawing.Size(296, 25);
            this.lblRango.TabIndex = 1;
            this.lblRango.Text = "Semana";
            this.lblRango.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnGenerarPolizas
            // 
            this.btnGenerarPolizas.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnGenerarPolizas.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.btnGenerarPolizas.Enabled = false;
            this.btnGenerarPolizas.Location = new System.Drawing.Point(248, 160);
            this.btnGenerarPolizas.Name = "btnGenerarPolizas";
            this.btnGenerarPolizas.Size = new System.Drawing.Size(296, 38);
            this.btnGenerarPolizas.TabIndex = 2;
            this.btnGenerarPolizas.Text = "Generar Polizas";
            this.btnGenerarPolizas.UseVisualStyleBackColor = true;
            // 
            // chkListTipoPoliza
            // 
            this.chkListTipoPoliza.FormattingEnabled = true;
            this.chkListTipoPoliza.Items.AddRange(new object[] {
            "CCI- Comisiones, Compensaciones e Intereses",
            "IV- Ingresos Varios",
            "IOP- Ingresos OpenPay y Paypal",
            "IP- Ingresos Pequeños",
            "CM - Cierre Mensual"});
            this.chkListTipoPoliza.Location = new System.Drawing.Point(276, 13);
            this.chkListTipoPoliza.Name = "chkListTipoPoliza";
            this.chkListTipoPoliza.Size = new System.Drawing.Size(251, 79);
            this.chkListTipoPoliza.TabIndex = 3;
            this.chkListTipoPoliza.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.chkListTipoPoliza_ItemCheck);
            this.chkListTipoPoliza.SelectedIndexChanged += new System.EventHandler(this.chkListTipoPoliza_SelectedIndexChanged);
            // 
            // chkTipoAgrupador
            // 
            this.chkTipoAgrupador.AutoSize = true;
            this.chkTipoAgrupador.Checked = true;
            this.chkTipoAgrupador.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkTipoAgrupador.Location = new System.Drawing.Point(276, 106);
            this.chkTipoAgrupador.Name = "chkTipoAgrupador";
            this.chkTipoAgrupador.Size = new System.Drawing.Size(67, 17);
            this.chkTipoAgrupador.TabIndex = 4;
            this.chkTipoAgrupador.Text = "Empresa";
            this.chkTipoAgrupador.UseVisualStyleBackColor = true;
            this.chkTipoAgrupador.CheckedChanged += new System.EventHandler(this.chkTipoAgrupador_CheckedChanged);
            // 
            // chkTipoCalendario
            // 
            this.chkTipoCalendario.AutoSize = true;
            this.chkTipoCalendario.Checked = true;
            this.chkTipoCalendario.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkTipoCalendario.Location = new System.Drawing.Point(443, 106);
            this.chkTipoCalendario.Name = "chkTipoCalendario";
            this.chkTipoCalendario.Size = new System.Drawing.Size(84, 17);
            this.chkTipoCalendario.TabIndex = 4;
            this.chkTipoCalendario.Text = "Por Semana";
            this.chkTipoCalendario.ThreeState = true;
            this.chkTipoCalendario.UseVisualStyleBackColor = true;
            this.chkTipoCalendario.CheckStateChanged += new System.EventHandler(this.chkTipoCalendario_CheckStateChanged);
            // 
            // frmGeneradorPolizas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(544, 198);
            this.Controls.Add(this.lblRango);
            this.Controls.Add(this.chkTipoCalendario);
            this.Controls.Add(this.chkTipoAgrupador);
            this.Controls.Add(this.chkListTipoPoliza);
            this.Controls.Add(this.btnGenerarPolizas);
            this.Controls.Add(this.dateTimePicker1);
            this.MaximumSize = new System.Drawing.Size(560, 300);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(560, 237);
            this.Name = "frmGeneradorPolizas";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Generador de Polizas";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmGeneradorPolizas_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGenerarPolizas;
        private System.Windows.Forms.CheckedListBox chkListTipoPoliza;
        private System.Windows.Forms.CheckBox chkTipoAgrupador;
        private System.Windows.Forms.CheckBox chkTipoCalendario;
    }
}