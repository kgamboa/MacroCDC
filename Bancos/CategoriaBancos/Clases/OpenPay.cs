using System;
using System.Collections.Generic;

namespace TesoreriaCobranza
{
    public class OpenPayLiquidacion
    {
        public string FileName { get; internal set; }
        public decimal Amount { get; internal set; }
        public List<OpenPay> Charges { get; internal set; }
        public DateTime? PayDate { get; internal set; }

        public Dictionary<string, string> FeeAccount = new Dictionary<string, string>()
        {
            {"AURA", "5003500049" },
            { "IGH","5003500043"},
            { "RCE", "5003500023" },
            { "CALIFICATE", "5003500009" },
            { "AVISAME MUCHO MAS", "5003500075" },
        };
        public Dictionary<string, string> ChargeAccount = new Dictionary<string, string>()
        {
            {"AURA", "PEND" },
            { "IGH","PEND"},
            { "RCE", "4003500023" },
            { "CALIFICATE", "4003500009" },
            { "AVISAME MUCHO MAS", "4003500010" },
        };
    }
    public class OpenPay
    {
        public string TransactionId { get; internal set; }
        public string Category { get; internal set; }
        public string Method { get; internal set; }
        public string Status { get; internal set; }
        public string CreateDate { get; internal set; }
        public string Description { get; internal set; }
        public string CustomName { get; internal set; }
        public decimal Amount { get; internal set; }
        public decimal OriginalAmount { get; internal set; }
        public decimal Fee { get; internal set; }
        public decimal FeeTax { get; internal set; }
        public string FeeAura
        {
            get
            {
                if (Description.Contains("AURA")) return Fee.ToString();
                return string.Empty;
            }
        }
        public string FeeIGH
        {
            get
            {
                if (Description.Contains("IGH")) return Fee.ToString();
                return string.Empty;
            }
        }
        public string UserName { get; internal set; }
        public string Product { get; internal set; }
        internal Factura Invoice { get; set; }

        public override string ToString()
        {
            return string.Join(",", "Category", Category, "Method", Method, "Status", Status, "Description", Description, "CustomName", CustomName, "Amount", Amount, "OriginalAmount",OriginalAmount);
        }
    }
}