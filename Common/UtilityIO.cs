using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace Common
{
    public static class UtilityIO
    {
        private static readonly string macroPath = @"C:\MacroCDC_PTC\";
        public static string MacroPath => macroPath;
        public static string MacroLogPath => macroPath+@"/Log";
        public static string MacroTempPath => Path.Combine(MacroPath, "temp");
        public static IEnumerable<FileInfo> GetFilesByExtensions(List<DirectoryInfo> directoryList, SearchOption searchOption, params string[] extensions)
        {
            List<FileInfo> filesList = new List<FileInfo>();
            foreach (var dir in directoryList)
            {
                var files = GetFilesByExtensions(dir, searchOption, null, null, extensions);
                if (files != null)
                    filesList.AddRange(files);
            }

            return filesList;
        }

        public static IEnumerable<FileInfo> GetFilesByExtensions(string directorio, SearchOption searchOption, string withoutmark = null, string withmark = null, params string[] extensions)
        {
            var dir = ExisteCreaDirectorio(directorio);
            return GetFilesByExtensions(dir, searchOption, withoutmark, withmark, extensions);
        }

        public static IEnumerable<FileInfo> GetFilesByExtensions(this DirectoryInfo dir, SearchOption searchOption, string withoutmark = null, string withmark = null, params string[] extensions)
        {
            if (extensions == null)
                throw new ArgumentNullException("extensions");
            var exts = new List<string>(extensions.Select(t => t.ToLower()));

            IEnumerable<FileInfo> files = dir.EnumerateFiles("*.*", searchOption).Where(f => exts.Count == 0 || exts.Contains(f.Extension.ToLower()));
            if (!string.IsNullOrEmpty(withoutmark))
            {
                files = files.Where(f => !f.Name.ToUpper().Contains(withoutmark.ToUpper()));
            }
            if (!string.IsNullOrEmpty(withmark))
            {
                files = files.Where(f => f.Name.ToUpper().Contains(withmark.ToUpper()));
            }
            return files;
        }
        public static IEnumerable<FileInfo> GetFilesByExtensions2(this DirectoryInfo dir, SearchOption searchOption, string withoutmark = null, string withmark = null, params string[] extensions)
        {
            if (extensions == null)
                throw new ArgumentNullException("extensions");
            var exts = new List<string>(extensions.Select(t => t.ToLower()));

            IEnumerable<FileInfo> files = dir.EnumerateFiles("*.*", searchOption).Where(f => exts.Count == 0 || exts.Contains(f.Extension.ToLower()));
            if (!string.IsNullOrEmpty(withoutmark))
            {
                files = files.Where(f => !f.DirectoryName.ToUpper().Contains(withoutmark.ToUpper()));
            }
            if (!string.IsNullOrEmpty(withmark))
            {
                files = files.Where(f => f.Name.ToUpper().Contains(withmark.ToUpper()));
            }
            return files;
        }

        public static DirectoryInfo ExisteCreaDirectorio(string dir, DirCreateType tipo = DirCreateType.Create)
        {
            //Revisa que exista la Ruta
            var directorio = new DirectoryInfo(dir);
            if (tipo == DirCreateType.DeleteAndCreate || tipo == DirCreateType.Delete)
                if (directorio.Exists) RecursiveDelete(directorio);

            directorio = new DirectoryInfo(dir);
            if ((tipo == DirCreateType.Create || tipo == DirCreateType.DeleteAndCreate) && !directorio.Exists)
                directorio.Create();
            return directorio;
        }
        public static void RecursiveDelete(DirectoryInfo baseDir)
        {
            if (!baseDir.Exists)
                return;

            foreach (var dir in baseDir.EnumerateDirectories())
            {
                RecursiveDelete(dir);
            }
            baseDir.Delete(true);
        }
        public static bool ExisteArchivo(string fileName)
        {
            return new FileInfo(fileName).Exists;
        }
        public static string GeneraArchivoNuevo(string rutaArchivoOriginal, string archivoOriginal,string  proceso )
        {
            if (new FileInfo(archivoOriginal).Exists)
            {
                string rutaArchivoTemporal = Path.Combine(Path.GetTempPath(), Path.GetFileName(archivoOriginal));

                string nuevoNombreArchivo = $"{Regex.Replace(Path.GetFileNameWithoutExtension(archivoOriginal.Replace("_CC","").Replace("_CL","").Replace("_BANCOS","")), @"_\d{8}_\d{6}", "")}_{proceso}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                string rutaNuevoArchivo = Path.Combine(Path.GetDirectoryName(rutaArchivoOriginal), nuevoNombreArchivo); // Guarda en la misma carpeta que el original
                File.Copy(archivoOriginal, rutaNuevoArchivo, true);
                return rutaNuevoArchivo;
            }
            return null;
        }
        public static string GeneraArchivoTemporal(string rutaArchivoOriginal)
        {
            if (new FileInfo(rutaArchivoOriginal).Exists)
            {
                string rutaArchivoTemporal = Path.Combine(Path.GetTempPath(), Path.GetFileName(rutaArchivoOriginal));

                string nuevoNombreArchivo = $"{Regex.Replace(Path.GetFileNameWithoutExtension(rutaArchivoOriginal), @"_\d{8}_\d{6}", "")}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                DirectoryInfo directoryInfo = new DirectoryInfo(Path.Combine(MacroPath, "temp"));
                if (!directoryInfo.Exists)
                    directoryInfo.Create();
                string rutaNuevoArchivo = Path.Combine(directoryInfo.FullName, nuevoNombreArchivo); // Guarda en la misma carpeta que el original
                File.Copy(rutaArchivoOriginal, rutaNuevoArchivo, true);
                return rutaNuevoArchivo;
            }
            return null;
        }
        public static void MarcaArchivoProcesado(string fileName, string marca = "~")
        {
            FileInfo file = new FileInfo(fileName);
            if (file.Exists)
            {
                MarcaArchivoProcesado(file, marca);
            }
        }

        public static void MarcaArchivoProcesado(FileInfo file, string marca = "~")
        {
            try
            {
                if (!file.FullName.Contains(marca))
                {
                    //Se Marca Archivo
                    FileInfo infoName = new FileInfo(file.FullName);
                    //RenombraArhivo
                    var destinationFileName = RenombraArchivo(file.FullName, marca);

                    infoName.MoveTo(destinationFileName);
                }
            }
            catch
            {
                throw new Exception("No se logró marcar el archivo, por que ya estaba abierto en otro proceso");
            }
        }

        public static bool BorraArchivo(string filename, params string[] extesiones)
        {
            try
            {
                var exts = new List<string>();
                exts.AddRange(extesiones);
                if (exts.Count == 0) exts.Add("");

                foreach (string ext in exts)
                {
                    FileInfo file = new FileInfo(filename.Replace("\\\\", "\\") + ext);
                    if (file.Exists)
                    {
                        file.Delete();
                    }
                }
            }
            catch { return false; }
            return true;
        }

        public static string RenombraArchivo(string fullFileName, string marca = "(#)")
        {
            // Create a FileInfo object for easier manipulation of file path components.
            var file = new FileInfo(fullFileName);

            // Ensure the directory exists. If not, create it.
            // This is important before attempting to check for file existence or creating the file.
            if (file.Directory != null && !file.Directory.Exists)
            {
                file.Directory.Create();
            }

            // Initialize the destination filename with the original filename.
            // This will be used if no renaming is necessary (i.e., the original file doesn't exist).
            var destinationFileName = file.Name;

            // Check if the original file already exists.
            // If it does, we need to generate a new unique name.
            if (file.Exists)
            {
                // Extract the filename without its extension and the extension itself.
                // This allows us to insert the counter correctly.
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.Name);
                string extension = Path.GetExtension(file.Name);

                var count = 0; // Initialize the counter for renaming.
                string newProposedFileName; // Variable to hold the name with the appended counter.

                // Loop until a unique filename is found.
                do
                {
                    // Construct the new filename by appending the formatted 'marca' (e.g., "(0)", "(1)")
                    // before the original file extension.
                    newProposedFileName = $"{fileNameWithoutExtension} {marca.Replace("#", count.ToString())}{extension}";

                    count++; // Increment the counter for the next iteration.

                } while (new FileInfo(Path.Combine(file.Directory.FullName, newProposedFileName)).Exists);
                // Continue looping as long as a file with the new proposed name already exists in the directory.

                // Once a unique name is found, set it as the destination filename.
                destinationFileName = newProposedFileName;
            }

            // Return the full path of the new (or original) filename.
            // Path.Combine is used for robust path construction across different operating systems.
            return Path.Combine(file.Directory.FullName, destinationFileName);
        }

        public static string NormalizaNombre(string nombre)
        {
            if (!string.IsNullOrEmpty(nombre))
            {
                var index = 0;
                nombre = nombre.ToUpper().Replace(",", "").Replace(@"\", "-").Replace(".", "").Replace(@"/", "-").Replace(@":", "").Replace(@"*", "-").Replace(@"?", "-").Replace("\"", "-").Replace(@"<", "-").Replace(@">", "-").Replace(@"|", "-").Replace(@"'", "").Replace(@"´", "").Trim();


                if (nombre.TrimEnd().EndsWith("SA DE CV")) { index = nombre.LastIndexOf("SA DE CV"); }
                else if (nombre.TrimEnd().EndsWith("S A DE C V")) { index = nombre.LastIndexOf("S A DE C V"); }
                else if (nombre.TrimEnd().EndsWith("SA PI DE CV")) { index = nombre.LastIndexOf("SA PI DE CV"); }
                else if(nombre.TrimEnd().EndsWith(" SA")) { index = nombre.LastIndexOf(" SA"); }
                else if(nombre.TrimEnd().EndsWith(" AC")) { index = nombre.LastIndexOf(" AC"); }
                else if(nombre.TrimEnd().EndsWith(" SC")) { index = nombre.LastIndexOf(" SC"); }
                else if(nombre.Contains("SC DE")) { index = nombre.LastIndexOf("SC DE"); }
                else if (nombre.Contains("SA SOC")) { index = nombre.LastIndexOf("SA SOC"); }
                else if (nombre.Contains("SA D")) { index = nombre.LastIndexOf("SA D"); }
                else if (nombre.Contains("SADE ")) { index = nombre.LastIndexOf("SADE "); }
                else if (nombre.Contains("SAINSTITU")) { index = nombre.LastIndexOf("SAINSTITU"); }
                else if (nombre.Contains("SA INSTITU")) { index = nombre.LastIndexOf("SA INSTITU"); }
                else if (nombre.Contains("SAS ")) { index = nombre.LastIndexOf("SAS "); }
                else if (nombre.Contains("SAB ")) { index = nombre.LastIndexOf("SAB "); }
                else if (nombre.Contains("SAPI ")) { index = nombre.LastIndexOf("SAPI "); }
                else if (nombre.Contains("SAPIDE ")) { index = nombre.LastIndexOf("SAPIDE "); }
                else if (nombre.Contains("S DE RL")) { index = nombre.LastIndexOf("S DE RL"); }
                else if (nombre.Contains("S DERL")) { index = nombre.LastIndexOf("S DERL"); }
                else if (nombre.Contains("SDE RL")) { index = nombre.LastIndexOf("SDE RL"); }
                else if (nombre.Contains("SOCIEDAD ANONIMA")) { index = nombre.LastIndexOf("SOCIEDAD ANONIMA"); }
                else if (nombre.Contains("SPRDE ")) { index = nombre.LastIndexOf("SPRDE "); }
                else if (nombre.Contains("SPR DE")) { index = nombre.LastIndexOf("SPR DE"); }
                else if (nombre.Contains("AC ")) { index = nombre.LastIndexOf("AC "); }
                else if (nombre.Contains("SC DE AP DE RL DE CV")) { index = nombre.LastIndexOf("SC DE AP DE RL DE CV"); }
                else if (nombre.Contains("S A P I DE CV")) { index = nombre.LastIndexOf("S A P I DE CV"); }
                else if (nombre.Contains("S A B DE CV")) { index = nombre.LastIndexOf("S A B DE CV"); }

                else if (nombre.Contains("SUST ")) { index = nombre.IndexOf("SUST "); }
                else if (nombre.Contains("SUSTITUYE "))
                {
                    index = nombre.IndexOf("SUSTITUYE ");
                }
                else if (nombre.Contains("CXL INV")) { index = nombre.IndexOf("CXL INV"); }

                if (index > 0)
                {
                    nombre = nombre.Insert(index, " ");
                    var x = nombre.Substring(0, index+1);
                    if (x.Length == x.LastIndexOf(' ')+1 || nombre.EndsWith(" SA") || nombre.EndsWith(" AC") || nombre.EndsWith(" SC"))
                        nombre = nombre.Substring(0, index);
                }
                if (nombre.StartsWith("00Q") || nombre.StartsWith("SN"))
                {
                    index = nombre.IndexOf(" ");
                    nombre = nombre.Substring(index + 1);
                }


                nombre = Regex.Replace(nombre.ToLower().Normalize(NormalizationForm.FormD), @"[^a-zA-z0-9 ]+", "");
                return nombre.ToUpper().Trim();
            }
            return null;
        }

        public static void MarcaDirectorioProcesado(FileInfo file, string marca = "~")
        {
            try
            {
                var directory = file.Directory;

                DirectoryInfo infoName = new DirectoryInfo(directory.FullName);
                //RenombraDirectorio
                var destinationDirectoryName = RenombraDirectorio(directory.FullName, marca);
                file.MoveTo(destinationDirectoryName + @"\" + file.Name);
                if (infoName.GetFiles().Count() == 0)
                    infoName.Delete(true);
            }
            catch
            {
                throw new Exception("No se logró marcar el directorio, por que ya estaba abierto en otro proceso");
            }
        }

        private static string RenombraDirectorio(string fullName, string marca)
        {
            var dir = new DirectoryInfo(fullName);
            var newDir =new DirectoryInfo(dir.Parent.FullName + @"\" + marca + dir.Name);
            newDir.Create();
            return newDir.FullName;
        }


        // Función para calcular el hash SHA256 de un archivo
        private static string CalcularHashArchivo(string filePath)
        {
            using (var sha256 = SHA256.Create())
            {
                using (var stream = File.OpenRead(filePath))
                {
                    byte[] hashBytes = sha256.ComputeHash(stream);
                    return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                }
            }
        }
        public static void RemoveDuplicateFiles(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine($"El directorio '{directoryPath}' no existe.");
                return;
            }

            string[] archivos = Directory.GetFiles(directoryPath);
            if (archivos.Length <= 1)
            {
                Console.WriteLine("No hay suficientes archivos para buscar duplicados.");
                return;
            }

            // Diccionario para almacenar el hash del archivo como clave y la lista de rutas de archivo con ese hash como valor.
            Dictionary<string, List<string>> archivosPorHash = new Dictionary<string, List<string>>();

            Console.WriteLine("Analizando archivos en busca de duplicados...");

            foreach (string archivoPath in archivos.OrderBy(c => c.Length))
            {
                try
                {
                    string hash = CalcularHashArchivo(archivoPath);

                    if (archivosPorHash.ContainsKey(hash))
                    {
                        archivosPorHash[hash].Add(archivoPath);
                    }
                    else
                    {
                        archivosPorHash.Add(hash, new List<string> { archivoPath });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error al procesar el archivo '{archivoPath}': {ex.Message}");
                }
            }

            int archivosEliminados = 0;
            Console.WriteLine("\nEliminando archivos duplicados (manteniendo la primera instancia encontrada):");

            foreach (var parClaveValor in archivosPorHash)
            {
                if (parClaveValor.Value.Count > 1)
                {
                    // Mantener el primer archivo encontrado, eliminar los demás
                    for (int i = 1; i < parClaveValor.Value.Count; i++)
                    {
                        try
                        {
                            File.Delete(parClaveValor.Value[i]);
                            Console.WriteLine($"Eliminado duplicado: '{parClaveValor.Value[i]}'");
                            archivosEliminados++;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error al eliminar el archivo '{parClaveValor.Value[i]}': {ex.Message}");
                        }
                    }
                }
            }

            Console.WriteLine($"\nProceso completado. Se eliminaron {archivosEliminados} archivos duplicados.");
        }
    }
    public enum DirCreateType
    {
        Create,
        Delete,
        DeleteAndCreate
    }
}