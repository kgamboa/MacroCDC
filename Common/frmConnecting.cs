using System;
using System.Threading;
using System.Windows.Forms;

namespace Common
{
    public partial class frmConnecting : Form
    {
        private string textData;
        public string TextData
        {
            get
            {
                return textData;
            }
            set
            {
                try
                {
                    textData = value;
                    if (!this.IsDisposed)
                        Invoke(new UpdateTextBoxDelegate(InvokeUpdateTextBox), textData);
                }
                catch (InvalidOperationException ex)
                {
                    // Log or handle the exception as needed
                }
                finally { }
            }
        }

        private decimal progressValue;
        public decimal ProgressValue
        {
            get
            {
                return progressValue;
            }
            set
            {
                progressValue = value;
                if (!this.IsDisposed)
                    Invoke(new UpdateProgressBarDelegate(InvokeUpdatePogressBar), progressValue);
            }
        }

        public CancellationTokenSource CancellationTokenSource { get; internal set; }

        public delegate void UpdateTextBoxDelegate(string value);
        public delegate void UpdateProgressBarDelegate(decimal value);
        public void InvokeUpdateTextBox(string data)
        {
            textBox1.Text = data ?? string.Empty;
        }

        public void InvokeUpdatePogressBar(decimal data)
        {
            progressBar1.Value = (int)(data > 100 || data <0 ? 0 : data);
        }
        public frmConnecting()
        {
            InitializeComponent();
        }

        private void frmCambioEstatus_Shown(object sender, EventArgs e)
        {
            pictureBox2.Visible = true;
            pictureBox2.BringToFront();
            textBox1.Visible = true;
            textBox1.BringToFront();
        }

        private void frmCambioEstatus_Load(object sender, EventArgs e)
        {
            pictureBox2.Visible = true;
            pictureBox2.BringToFront();
            textBox1.Visible = true;
            textBox1.BringToFront();
        }

        private void frmConnecting_FormClosing(object sender, FormClosingEventArgs e)
        {
            CancellationTokenSource?.Cancel();
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            CancellationTokenSource?.Cancel();
        }
    }
}