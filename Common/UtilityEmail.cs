using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;

namespace Common
{
    public class UtilityEmail
    {

        public static bool IsEmailValid(string emailaddress)
        {
            try
            {
                if (!string.IsNullOrEmpty(emailaddress))
                {
                    MailAddress m = new MailAddress(emailaddress);

                    return true;
                }
            }
            catch { }
            return false;
        }
    }
}
