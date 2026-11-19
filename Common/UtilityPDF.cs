using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;

namespace Common
{
   public class UtilityPDF
    {

        public static string parseUsingPDFBox(string input)
        {
            PDDocument doc = null;

            try
            {
                doc = PDDocument.load(input);
                PDFTextStripper stripper = new PDFTextStripper();
                return stripper.getText(doc);
            }
            catch (System.StackOverflowException ex)
            { return "Exception:" + ex.Message; }
            catch (Exception ex) { return "Exception:"+ex.Message; }
            finally
            {
                if (doc != null)
                {
                    doc.close();
                }
            }
        }

        public static string ObtieneInformacion(ref string data, string key, string nextKey = "\r\n")
        {

            var index = data.IndexOf(key);
            data = data.Substring(index + key.Length);

            index = data.IndexOf(nextKey);

            var indexEnter = data.IndexOf("\r\n");

            if (indexEnter < index || index == -1)
                index = indexEnter;
            if (index == 0)
                return string.Empty;
            if (index > 0)
                return data.Substring(0, index).Trim();
            return data;
        }

        public static string ObtieneInformacionNombre(ref string data, string nextKey,bool enter =true)
        {
            var index = 0;
            if (enter)
            {
                index = data.IndexOf("\r\n");
                data = data.Substring(index + "\r\n".Length);
                index = data.IndexOf(nextKey);
            }
            else
            {
                index = data.IndexOf(nextKey);
                if (index == -1)
                    index = data.IndexOf(nextKey.Trim());

                if (index > 0)
                {
                    var info = data.Substring(0, index).Trim();
                    data = data.Substring(index);

                    return info;
                }
            }

            if (index == 0)
                return string.Empty;
            if (index > 0)
                return data.Substring(0, index).Trim();
            return data;
        }

        public static string ObtieneAbrevRegimen(string regimen)
        {
            switch (regimen)
            {
                case "SOCIEDAD COOPERATIVA DE RESPONSABILIDAD LIMITADA DE CAPITAL VARIABLE":
                    return "SC DE RL DE CV";
                case "SOCIEDAD ANONIMA":
                    return "SA";
                case "SOCIEDAD ANONIMA DE CAPITAL VARIABLE (*)":
                case "SOCIEDAD ANONIMA DE CAPITAL VARIABLE":
                    return "SA DE CV";
                case "SOCIEDAD POR ACCIONES SIMPLIFICADA DE CAPITAL VARIABLE":
                    return "SAS DE CV";
                case "ASOCIACION CIVIL":
                    return "AC";
                case "SOCIEDAD CIVIL":
                    return "SC";
                case "SOCIEDAD ANONIMA PROMOTORA DE INVERSION DE CAPITAL VARIABLE":
                    return "SAPI DE CV";
                case "SOCIEDAD DE RESPONSABILIDAD LIMITADA DE CAPITAL VARIABLE":
                    return "S DE RL DE CV";
                case "SA PROMOTORA DE INVERSION DE CV, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE ENTIDAD NO REGULADA":
                case "SA PROMOTORA DE INVERSION DE CV, SOCIEDAD FINANCIERA DE OBJETO MULTIPLEENTIDAD NO REGULADA":
                    return "SAPI de CV SOFOM ENR";
                case "SOCIEDAD ANONIMA DE CAPITAL VARIABLE, SOCIEDADFINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA":
                case "SOCIEDAD ANONIMA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETO MULTIPLE, ENTIDAD NO REGULADA":
                case "SOCIEDAD ANONIMA DE CAPITAL VARIABLE, SOCIEDAD FINANCIERA DE OBJETOMULTIPLE, ENTIDAD NO REGULADA":
                    return "SA DE CV SOFOM ENR";
            }
            if (regimen == "")
                return "SA DE CV SFP";
            if (regimen == "")
                return "SAB DE CV";
            if (regimen == "")
                return "SC DE AP DE RL DE CV";
            if (regimen == "")
                return "S.A INSTITUCION DE BANCA MULTIPLE";


            return string.Empty;
        }

    }
}
