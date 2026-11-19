using Common;
//using OCRAnalyzerApp;
using PTC;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;

namespace TesoreriaCobranza
{
    public partial class frmTesoreriaCobranza : Form
    {
        private CancellationTokenSource _cancellationTokenSource;

        public frmTesoreriaCobranza()
        {
            InitializeComponent();


            //UtilityIO.ExisteCreaDirectorio(@"C:\Tesoreria\Pago Proveedores\", DirCreateType.Create);
            var version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

            cerrarToolStripMenuItem.Text = "Cerrar (v." + version + ")";
            this.PathLog = ErrorServices.singleton.Path;
        }


        private void cerrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void generaArchivoBancosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    using (new PleaseWait(this.Location, this.PathLog))
            //    {
            //        GeneraBancos.GeneraArchivoBancos();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
            //}
        }
        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        public string PathLog { get; }

        private void frmBancoCategoria_Load(object sender, EventArgs e)
        {
        }



        private void contaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //ConciliacionContabilidad.EjecutaConciliacionContabilidad();
        }

        private void tesoreriaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //ConciliacionTesoreria.EjecutaConciliacionTesoreria();
        }

        private void cobranzaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //ConciliacionFacturacion.EjecutaConciliacionFacturacion();
        }


        private void categorizaBancosMenuItem2_Click(object sender, EventArgs e)
        {

            try
            {
                var openFileDialog1 = new OpenFileDialog();
                openFileDialog1.DefaultExt = "xlsx";
                openFileDialog1.Filter = "Excel | *.xlsx";
                if (DialogResult.OK == openFileDialog1.ShowDialog())
                {
                    if (UtilityIO.ExisteArchivo(openFileDialog1.FileName))
                    {
                        using (new PleaseWait(this.Location, this.PathLog))
                        {
                            var fProceso = new TesoreriaCobranza.ClasificacionBancos(openFileDialog1.FileName);
                            fProceso.InicializaProceso();
                            fProceso.EjecutaProceso();

                            Process.Start(openFileDialog1.FileName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
                Process.Start(ErrorServices.singleton.ErrorLogPath);
            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
           
        }

        private void polizaCLBANCOSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Excel | *.xlsx";
                openFileDialog1.DefaultExt = "xlsx";

                if (DialogResult.OK == openFileDialog1.ShowDialog())
                {
                    string selectedFileName = openFileDialog1.FileName;
                    _cancellationTokenSource = new CancellationTokenSource();
                    string tempFileName = UtilityIO.GeneraArchivoTemporal(selectedFileName);
                    DirectoryInfo dirInfo = new DirectoryInfo(new FileInfo(selectedFileName).Directory.FullName);
                    if (!string.IsNullOrEmpty(tempFileName))
                    {
                        using (var wait = new PleaseWait(this.Location, this.PathLog, _cancellationTokenSource))
                        {
                            var fProceso2 = new ProcesoPoliza(dirInfo, tempFileName, "CL", wait);
                            fProceso2.InicializaProceso();
                            // Iniciamos el proceso largo
                            fProceso2.EjecutaProceso();

                            var archivoFinal = UtilityIO.GeneraArchivoNuevo(selectedFileName, tempFileName, "CC");
                            if (checkBox1.Checked)
                            {
                                Process.Start(archivoFinal);
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
                Process.Start(ErrorServices.singleton.ErrorLogPath);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
            }
        }



        private void polizaBancosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Excel | *.xlsx";
                openFileDialog1.DefaultExt = "xlsx";

                if (DialogResult.OK == openFileDialog1.ShowDialog())
                {
                    string selectedFileName = openFileDialog1.FileName;
                    _cancellationTokenSource = new CancellationTokenSource();
                    string tempFileName = UtilityIO.GeneraArchivoTemporal(selectedFileName);
                    DirectoryInfo dirInfo = new DirectoryInfo(new FileInfo(selectedFileName).Directory.FullName);
                    if (!string.IsNullOrEmpty(tempFileName))
                    {
                        using (var wait = new PleaseWait(this.Location, this.PathLog, _cancellationTokenSource))
                        {
                            var fProceso2 = new ProcesoPoliza(dirInfo, tempFileName, "CC", wait);
                            fProceso2.InicializaProceso();
                            // Iniciamos el proceso largo
                            fProceso2.EjecutaProceso();

                            var archivoFinal = UtilityIO.GeneraArchivoNuevo(selectedFileName, tempFileName, "CC");
                            if (checkBox1.Checked)
                            {
                                Process.Start(archivoFinal);
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
                Process.Start(ErrorServices.singleton.ErrorLogPath);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
            }
        }

        private void identificaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var  openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Excel | *.xlsx";
                openFileDialog1.DefaultExt = "xlsx";

                if (DialogResult.OK == openFileDialog1.ShowDialog())
                {
                    string selectedFileName = openFileDialog1.FileName;
                    _cancellationTokenSource = new CancellationTokenSource();
                    string tempFileName = UtilityIO.GeneraArchivoTemporal(selectedFileName);
                    DirectoryInfo dirInfo = new DirectoryInfo(new FileInfo (selectedFileName).Directory.FullName);
                    if (!string.IsNullOrEmpty(tempFileName))
                    {
                        using (var wait = new PleaseWait(this.Location, this.PathLog, _cancellationTokenSource))
                        {
                            var fProceso2 = new IdentificaFacturas(dirInfo, tempFileName, "CC", wait);
                            fProceso2.InicializaProceso();
                            // Iniciamos el proceso largo
                            fProceso2.EjecutaProceso();

                            var archivoFinal = UtilityIO.GeneraArchivoNuevo(selectedFileName, tempFileName, "CC");
                            if (checkBox1.Checked)
                            {
                                Process.Start(archivoFinal);
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
                Process.Start(ErrorServices.singleton.ErrorLogPath);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
            }
        }


        private void identificaFacturasCLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Excel | *.xlsx";
                openFileDialog1.DefaultExt = "xlsx";
                if (DialogResult.OK == openFileDialog1.ShowDialog())
                {
                    string selectedFileName = openFileDialog1.FileName;
                    _cancellationTokenSource = new CancellationTokenSource();
                    string tempFileName = UtilityIO.GeneraArchivoTemporal(selectedFileName);
                    DirectoryInfo dirInfo = new DirectoryInfo(new FileInfo(selectedFileName).Directory.FullName);
                    if (!string.IsNullOrEmpty(tempFileName))
                    {
                        using (var wait = new PleaseWait(this.Location, this.PathLog, _cancellationTokenSource))
                        {
                            var fProceso2 = new IdentificaFacturas(dirInfo, tempFileName, "CL", wait);
                            fProceso2.InicializaProceso();
                            // Iniciamos el proceso largo
                            fProceso2.EjecutaProceso();


                            var archivoFinal = UtilityIO.GeneraArchivoNuevo(selectedFileName, tempFileName, "CL");
                            if (checkBox1.Checked)
                            {
                                Process.Start(archivoFinal);
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
                Process.Start(ErrorServices.singleton.ErrorLogPath);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
            }
        }

        private void ingresosPequeñosToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                var openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Excel | *.xlsx";
                openFileDialog1.DefaultExt = "xlsx";
                if (DialogResult.OK == openFileDialog1.ShowDialog())
                {
                    string selectedFileName = openFileDialog1.FileName;
                    _cancellationTokenSource = new CancellationTokenSource();
                    string tempFileName = UtilityIO.GeneraArchivoTemporal(selectedFileName);
                    DirectoryInfo dirInfo = new DirectoryInfo(new FileInfo(selectedFileName).Directory.FullName);
                    if (!string.IsNullOrEmpty(tempFileName))
                    {
                        using (var wait = new PleaseWait(this.Location, this.PathLog, _cancellationTokenSource))
                        {
                            var fProceso2 = new IngresosPequeños(dirInfo, tempFileName, wait);
                            fProceso2.InicializaProceso();
                            // Iniciamos el proceso largo
                            fProceso2.EjecutaProceso();


                            var archivoFinal = UtilityIO.GeneraArchivoNuevo(selectedFileName, tempFileName, "CL");
                            if (checkBox1.Checked)
                            {
                                Process.Start(archivoFinal);
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
                Process.Start(ErrorServices.singleton.ErrorLogPath);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
            }
        }

        private void polizaSemanalToolStripMenuItem_Click(object sender, EventArgs e)
        {

            frmGeneradorPolizas configPoliza = new frmGeneradorPolizas();
            DialogResult dr = configPoliza.ShowDialog();
            if (dr == DialogResult.OK)
            {
                try
                {
                    var openFileDialog1 = new OpenFileDialog();
                    openFileDialog1.Filter = "Excel | *.xlsx";
                    openFileDialog1.DefaultExt = "xlsx";
                    if (DialogResult.OK == openFileDialog1.ShowDialog())
                    {
                        string selectedFileName = openFileDialog1.FileName;
                        _cancellationTokenSource = new CancellationTokenSource();
                        string tempFileName = UtilityIO.GeneraArchivoTemporal(selectedFileName);
                        DirectoryInfo dirInfo = new DirectoryInfo(new FileInfo(selectedFileName).Directory.FullName);
                        if (!string.IsNullOrEmpty(tempFileName))
                        {
                            using (var wait = new PleaseWait(this.Location, this.PathLog, _cancellationTokenSource))
                            {
                                foreach (var key in configPoliza.TiposPolizasSeleccionados)
                                {
                                    if (key == (int)TiposPolizas.ComisionesIntereses) 
                                    {
                                        var fProceso2 = new ProcesoPoliza(dirInfo, tempFileName, wait, configPoliza.Dates, configPoliza.AgrupadoPor);
                                        fProceso2.InicializaProceso();
                                        fProceso2.EjecutaProceso();
                                    }
                                    else if (key == (int)TiposPolizas.IngresosVarios)
                                    {
                                        var fProceso2 = new ProcesoPolizaIngresos(dirInfo, tempFileName, wait, configPoliza.Dates, configPoliza.AgrupadoPor);
                                        fProceso2.InicializaProceso();
                                        fProceso2.EjecutaProceso();
                                    }
                                    else if (key == (int)TiposPolizas.IngresosOpenpayPaypal)
                                    {
                                        var fProceso2 = new ProcesoPolizaOP(dirInfo, tempFileName, wait, configPoliza.Dates, configPoliza.AgrupadoPor);
                                        fProceso2.InicializaProceso();
                                        fProceso2.EjecutaProceso();
                                    }
                                    else if (key == (int)TiposPolizas.IngresosPeq)
                                    {
                                        var fProceso2 = new ProcesoPolizaIngPeq(dirInfo, tempFileName, wait, configPoliza.Dates, configPoliza.AgrupadoPor);
                                        fProceso2.InicializaProceso();
                                        fProceso2.EjecutaProceso();
                                    }
                                    else if (key == (int)TiposPolizas.CierreMes)
                                    {
                                        var fProceso2 = new ProcesoPolizaCM(dirInfo, tempFileName, wait, configPoliza.Dates, configPoliza.AgrupadoPor);
                                        fProceso2.InicializaProceso();
                                        fProceso2.EjecutaProceso();
                                    }
                                }
                            }
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                }
                catch (Exception ex)
                {
                    Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
                    Process.Start(ErrorServices.singleton.ErrorLogPath);
                }
                finally
                {
                    _cancellationTokenSource?.Dispose();
                }
            }
        }

        //private void identificaPagosToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    var info = new frmOCRAnalyzer();
        //    info.ShowDialog();
        //}



        //private void nuevoToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    if (DialogResult.OK == openFileDialog1.ShowDialog())
        //    {
        //        if (UtilityIO.ExisteArchivo(openFileDialog1.FileName))
        //        {
        //            {
        //                DirectoryInfo dir = new DirectoryInfo(openFileDialog1.FileName);
        //                var list = dir.Parent.GetFiles("*.pdf").ToList();

        //                using (ExcelUtil.Utility excel = new ExcelUtil.Utility())
        //                {
        //                    excel.OpenFile(openFileDialog1.FileName, false);
        //                    var sheets = excel.GetSheetNames();
        //                    foreach (var item in sheets)
        //                    {

        //                        excel.SetActiveSheet(item);
        //                        var excelRange = excel.UsedRange;

        //                        int RowCount = excelRange.Rows.Count;
        //                        int ColumnCount = excelRange.Columns.Count;
        //                        for (int r = 2; r <= RowCount; r++)
        //                        {
        //                            var cell = (Range)excelRange.Cells[r, 2];
        //                            list.RemoveAll(c => c.Name == cell.value2);
        //                            //    var cell = (Range)excelRange.Cells[r, 1];
        //                            //    string curp = cell.value2.ToString().Substring(0, 10);
        //                            //    var max = 0d;
        //                            //    foreach (var i in list)
        //                            //    {
        //                            //        var sim = Common.UtilityString.findSimilarity(i.Name.Replace(i.Extension, "").Substring(0, 10), curp);
        //                            //        if (sim > max)
        //                            //        {
        //                            //            var cell2 = (Range)excelRange.Cells[r, 2];
        //                            //            cell2.Value2 = i.Name;
        //                            //            var cell3 = (Range)excelRange.Cells[r, 3];
        //                            //            cell3.Value2 = sim;
        //                            //            max = sim;
        //                            //        }
        //                            //    }
        //                        }

        //                        foreach (var i in list)
        //                        {
        //                            var cell = (Range)excelRange.Cells[RowCount++, 3];
        //                            cell.Value2 = i.Name;
        //                        }
        //                    }

        //                }
        //            }
        //        }
        //    }
        //}
    }

    //public class ScreenCapture
    //{
    //    [DllImport("user32.dll")]
    //    private static extern IntPtr GetForegroundWindow();

    //    [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
    //    public static extern IntPtr GetDesktopWindow();

    //    [StructLayout(LayoutKind.Sequential)]
    //    private struct Rect
    //    {
    //        public int Left;
    //        public int Top;
    //        public int Right;
    //        public int Bottom;
    //    }

    //    [DllImport("user32.dll")]
    //    private static extern IntPtr GetWindowRect(IntPtr hWnd, ref Rect rect);

    //    public static Image CaptureDesktop()
    //    {
    //        return CaptureWindow(GetDesktopWindow());
    //    }

    //    public static Bitmap CaptureActiveWindow()
    //    {
    //        return CaptureWindow(GetForegroundWindow());
    //    }

    //    public static Bitmap CaptureWindow(IntPtr handle)
    //    {
    //        var rect = new Rect();
    //        GetWindowRect(handle, ref rect);
    //        var bounds = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
    //        var result = new Bitmap(bounds.Width, bounds.Height);

    //        using (var graphics = Graphics.FromImage(result))
    //        {
    //            graphics.CopyFromScreen(new Point(bounds.Left, bounds.Top), Point.Empty, bounds.Size);
    //        }

    //        return result;
    //    }
    //}
}
