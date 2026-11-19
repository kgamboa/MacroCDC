using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace Common
{
    public class UtilityZIP
    {

        public static void CompressDirectory(String startPath)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(startPath);
            if (!dirInfo.Exists) dirInfo.Create();

            FileInfo[] files = dirInfo.GetFiles("*.*",SearchOption.AllDirectories);
            foreach (FileInfo file in files)
            {
                if (file.Extension == ".csv")
                {
                    //Compacta Archivos
                    Compress(file);
                }
            }
        }
        public static void Compress(FileInfo fi)
        {

            //Genera una carpeta aleteria (GUID)
            string startPath = "f:\\TEMP\\" + Guid.NewGuid(); 
            DirectoryInfo dir = new DirectoryInfo(startPath);
            if (!dir.Exists)
                dir.Create();

            //Mueve el archivo a la carpeta aleteria
            var fileName = startPath + "//" + fi.Name;
            fi.CopyTo(fileName);

            //Crea archivo zip
            string zipPath = fi.FullName.Replace(fi.Extension ,"") + ".zip";
            FileInfo zipFile = new FileInfo(zipPath);
            if (zipFile.Exists)
                zipFile.Delete();

            //Compacta carpeta
            ZipFile.CreateFromDirectory(startPath, zipPath);

            //Eliina carpeta aleatoria
            dir.Delete(true);

            fi.Delete();

            //zipFile.MoveTo( fi.FullName.Replace(fi.Extension, zipFile.Extension));
        }

        public static void Decompress(string fileName)
        {
            var fi = new FileInfo(fileName);
            if (fi.Exists)
            {
                string extractPath = fi.FullName.Replace(fi.Extension, "");
                ZipFile.ExtractToDirectory(fi.FullName, extractPath);
            }
        }
    }
    
}
