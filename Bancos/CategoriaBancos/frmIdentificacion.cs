using Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TesoreriaCobranza;

namespace PTC
{
    public partial class frmIdentificacion : Form
    {
        private CancellationTokenSource _cancellationTokenSource;

        public string Empresa { get; private set; }
        public string PathLog { get; private set; }

        public frmIdentificacion(string pathLog, string empresa)
        {
            this.Empresa = empresa;
            this.PathLog = pathLog;

            InitializeComponent();

            btnStart.Text = "Identificación " + empresa;
            this.Text += " " + Empresa;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == openFileDialog1.ShowDialog())
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }


        private async void btnStart_Click(object sender, EventArgs e)
        {

            try
            {
                btnStart.Enabled = false;

                _cancellationTokenSource = new CancellationTokenSource();

                if (UtilityIO.ExisteArchivo(openFileDialog1.FileName))
                {
                    using (var wait = new PleaseWait(this.Location, this.PathLog, _cancellationTokenSource))
                    {
                        var fProceso2 = new IdentificaFacturas(openFileDialog1.FileName, Empresa, wait);
                        fProceso2.InicializaProceso();
                        // Iniciamos el proceso largo
                        await fProceso2.EjecutaProceso(_cancellationTokenSource.Token);
                    }
                    //MessageBox.Show("Proceso completado con éxito!");
                }

                if (checkBox1.Checked)
                {
                    Process.Start(openFileDialog1.FileName);
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Proceso cancelado por el usuario.");
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
                Process.Start(ErrorServices.singleton.ErrorLogPath);
            }
            finally
            {
                btnStart.Enabled = true;
                _cancellationTokenSource?.Dispose();
            }
        }

        private void frmIdentificacion_Load(object sender, EventArgs e)
        {

        }
    }
}
