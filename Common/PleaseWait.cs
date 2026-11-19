using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using System.IO;
using System.Collections;

namespace Common
{
    public class PleaseWait : IDisposable
    {
        private readonly Point mLocation;
        private frmConnecting mSplash;
        private decimal progressValue;
        public decimal ProgressValue
        {
            get { return progressValue; }
            set
            {
                progressValue = value;
                if (mSplash != null)
                    mSplash.ProgressValue = value;
            }
        }

        private string title;
        public string Title
        {
            get { return title; }
            set
            {
                title = value;
                if (mSplash != null)
                    mSplash.Text = value;
            }
        }

        public CancellationTokenSource CancellationToken { get { return cancellationTokenSource; } }

        private FileSystemWatcher watcher;
        private CancellationTokenSource cancellationTokenSource;
        public PleaseWait(Point location, string pathLog, CancellationTokenSource _cancellationTokenSource = null)
        {
     
            mLocation = location;
            watcher = new FileSystemWatcher();
            watcher.Path = pathLog; //la carpeta que quiere controlar
            watcher.Filter = "*.log";
            watcher.Changed += OnChanged;
            watcher.EnableRaisingEvents = true;
            cancellationTokenSource = _cancellationTokenSource;

            var t = new Thread(workerThread);
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        public void Dispose()
        {
            if (mSplash != null && !mSplash.IsDisposed)
                mSplash.Invoke(new MethodInvoker(stopThread));
            if (watcher != null)
                watcher.Dispose();
        }

        private void stopThread()
        {
            mSplash.TextData= null;
            mSplash.Close();            
            watcher.Changed -= OnChanged;
        }

        private void workerThread()
        {
            mSplash = new frmConnecting(); // Substitute this with your own
            mSplash.StartPosition = FormStartPosition.CenterScreen;
            mSplash.Location = mLocation;
            mSplash.CancellationTokenSource = cancellationTokenSource;
            mSplash.TopMost = true;


            Application.Run(mSplash);
        }
        private void OnChanged(object source, FileSystemEventArgs e)
        {
            if (mSplash != null && !mSplash.IsDisposed)
            {
                var filePath = e.FullPath;
                ArrayList arrText = new ArrayList();

                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var objReader = new StreamReader(fs, Encoding.Default))
                {
                    string sLine = objReader.ReadLine();
                    while (sLine != null)
                    {
                        arrText.Add(sLine);
                        sLine = objReader.ReadLine();
                    }

                }
                arrText.Reverse();
                mSplash.TextData = string.Join("\r\n", arrText.ToArray());
            } 
        }
    }
}
