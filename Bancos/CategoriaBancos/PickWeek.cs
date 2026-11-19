using Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using TesoreriaCobranza;

namespace PTC
{
    public partial class frmGeneradorPolizas : Form
    {
        public frmGeneradorPolizas()
        {
            InitializeComponent();

            dateTimePicker1.BoldedDates = GetWeekWithSpecialDays(DateTime.Now);
        }

        public Dictionary<string, DateTime[]> Dates;
        private void dateTimePicker1_DateSelected(object sender, DateRangeEventArgs e)
        {
            if (chkTipoCalendario.CheckState == CheckState.Checked)
            {
                dateTimePicker1.BoldedDates = GetWeekWithSpecialDays(e.Start);
            }
            else if (chkTipoCalendario.CheckState == CheckState.Indeterminate)
            {
                dateTimePicker1.BoldedDates = GetMonthDays(e.Start);
            }
            else
            {
                var arrayList = dateTimePicker1.BoldedDates.ToArray();
                List<DateTime> list = new List<DateTime>(arrayList);
                list.Add(e.Start);
                dateTimePicker1.BoldedDates = list.ToArray();
            }
        }
        private DateTime[] GetWeekWithSpecialDays(DateTime date)
        {
            if (chkTipoCalendario.CheckState == CheckState.Checked)
            {
                var week = GetWeekOfMonth(date);
                lblRango.Text = "Semana " + week;
                lblRango.TextAlign = ContentAlignment.TopCenter;
                var weekDays = GetDaysOfWeek(date);
                return weekDays.ToArray();
            }
            else
            {
                var dates = new List<DateTime>() { date };
                return dates.ToArray();
            }
        }

        private List<DateTime> GetDaysOfWeek(DateTime date)
        {
            List<DateTime> dates = new List<DateTime>();
            var week = GetWeekOfMonth(date);
            for (DateTime day = new DateTime(date.Year, date.Month, 1); date.Month == day.Month; day = day.AddDays(1))
            {
                var curretnWeek = GetWeekOfMonth(day);
                if (week == curretnWeek)
                    dates.Add(day);
                if (curretnWeek > week)
                    break;

            }
            return dates;
        }

        private int GetWeekOfMonth(DateTime date)
        {
            var firstDate = new DateTime(date.Year, date.Month, 1);

            if (firstDate.DayOfWeek == DayOfWeek.Saturday || firstDate.DayOfWeek == DayOfWeek.Sunday)
            {
                firstDate = firstDate.AddDays(2);
            }
            var lastDate = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));
            if (lastDate.DayOfWeek == DayOfWeek.Monday && date == lastDate)
            {
                date = date.AddDays(-1);
            }

            // Calculate the week number
            CultureInfo ci = CultureInfo.CurrentCulture;
            int firstWeek = ci.Calendar.GetWeekOfYear(firstDate, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            int currentWeek = ci.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            var week = currentWeek - firstWeek + 1;
            return week == 0 ? 1 : week;
        }


        private void dateTimePicker1_DateChanged(object sender, DateRangeEventArgs e)
        {
            if (chkTipoCalendario.CheckState == CheckState.Checked)
                dateTimePicker1.BoldedDates = GetWeekWithSpecialDays(e.Start);
            else if (chkTipoCalendario.CheckState == CheckState.Indeterminate)
                dateTimePicker1.BoldedDates = GetMonthDays(e.Start);
            else
            {
                var arrayList = dateTimePicker1.BoldedDates.ToArray();
                List<DateTime> list = new List<DateTime>(arrayList);
                list.Add(e.Start);
                dateTimePicker1.BoldedDates = list.ToArray();
            }

        }



        private void chkTipoAgrupador_CheckedChanged(object sender, EventArgs e)
        {
            chkTipoAgrupador.Text = chkTipoAgrupador.Checked ? "Empresa" : "Por Cuenta Bancos";
        }


        private void chkTipoCalendario_CheckStateChanged(object sender, EventArgs e)
        {
            chkTipoCalendario.Text = chkTipoCalendario.CheckState == CheckState.Checked ? "Por Semana" : chkTipoCalendario.CheckState == CheckState.Unchecked ? "Por Rangos" : "Por Mes";
            lblRango.Text = chkTipoCalendario.CheckState == CheckState.Checked ? "Semana" : "Rango";
            dateTimePicker1.SelectionRange = new SelectionRange(dateTimePicker1.SelectionStart, dateTimePicker1.SelectionStart);
            dateTimePicker1.MaxSelectionCount = 1;

            if (chkTipoCalendario.CheckState == CheckState.Checked)
            {
                dateTimePicker1.BoldedDates = GetWeekWithSpecialDays(dateTimePicker1.SelectionStart);
            }
            else if (chkTipoCalendario.CheckState == CheckState.Indeterminate)
            {
                dateTimePicker1.BoldedDates = GetMonthDays(dateTimePicker1.SelectionStart);
            }
            else
            {
                dateTimePicker1.MaxSelectionCount = 31;
                dateTimePicker1.BoldedDates = new List<DateTime>() { dateTimePicker1.SelectionStart }.ToArray();
            }
        }
        private DateTime[] GetMonthDays(DateTime date)
        {
            List<DateTime> monthDays = new List<DateTime>();
            DateTime firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            DateTime lastDayOfMonth = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month));

            for (DateTime day = firstDayOfMonth; day <= lastDayOfMonth; day = day.AddDays(1))
            {
                monthDays.Add(day);
            }

            return monthDays.ToArray();
        }

        private void dateTimePicker1_MouseCaptureChanged(object sender, EventArgs e)
        {
            if (chkTipoCalendario.CheckState == CheckState.Unchecked)
            {
                if (isControl)
                {
                    dateTimePicker1.BoldedDates = new List<DateTime>(dateTimePicker1.BoldedDates) { dateTimePicker1.SelectionStart }.ToArray();
                    isControl = false;
                }
                else
                {
                    List<DateTime> selectedDays = new List<DateTime>(dateTimePicker1.BoldedDates);
                    if (dateTimePicker1.SelectionRange.Start == dateTimePicker1.SelectionRange.End)
                        selectedDays = new List<DateTime>();

                    for (DateTime day = dateTimePicker1.SelectionRange.Start; day <= dateTimePicker1.SelectionRange.End; day = day.AddDays(1))
                    {
                        selectedDays.Add(day);
                    }
                    dateTimePicker1.BoldedDates = selectedDays.ToArray();
                }
            }
        }
        private bool isControl;


        public int AgrupadoPor => chkTipoAgrupador.Checked ? 1 : 0;
        public List<int> TiposPolizasSeleccionados
        {
            get
            {
                var seleccionados = new List<int>();
                foreach (var item in chkListTipoPoliza.CheckedIndices)
                {
                    seleccionados.Add((int)item);
                }
                return seleccionados;
            }
        }
        private void dateTimePicker1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            isControl = e.Control;
        }

        private void chkListTipoPoliza_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // Habilita el botón si hay al menos un ítem seleccionado (considerando el cambio actual)
            btnGenerarPolizas.Enabled = chkListTipoPoliza.CheckedItems.Count > 1 || e.NewValue == CheckState.Checked;

            string currentItem = chkListTipoPoliza.Items[e.Index].ToString();
            bool isCurrentMonthly = currentItem.StartsWith("CM") || currentItem.StartsWith("IP");

            // Construir la lista de seleccionados antes del cambio propuesto
            var checkedItemsBefore = new List<string>();
            for (int i = 0; i < chkListTipoPoliza.Items.Count; i++)
            {
                if (chkListTipoPoliza.GetItemChecked(i))
                    checkedItemsBefore.Add(chkListTipoPoliza.Items[i].ToString());
            }

            // Simular la lista después del cambio propuesto
            var checkedItems = new List<string>(checkedItemsBefore);
            if (e.NewValue == CheckState.Checked && !checkedItems.Contains(currentItem))
                checkedItems.Add(currentItem);
            if (e.NewValue == CheckState.Unchecked && checkedItems.Contains(currentItem))
                checkedItems.Remove(currentItem);

            // Monthly logic: CM and IP can be selected together, but not with others
            bool anyMonthlySelected = checkedItemsBefore.Any(item => item.StartsWith("CM") || item.StartsWith("IP"));
            bool anyNonMonthlySelected = checkedItemsBefore.Any(item => !item.StartsWith("CM") && !item.StartsWith("IP"));

            // Si se intenta seleccionar un Monthly y ya hay un no-Monthly seleccionado, cancelar
            if (isCurrentMonthly && e.NewValue == CheckState.Checked && anyNonMonthlySelected)
            {
                MessageBox.Show("No puedes seleccionar una opción mensual junto con otras opciones.", "Selección inválida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.NewValue = CheckState.Unchecked;
                return;
            }
            // Si se intenta seleccionar un no-Monthly y ya hay un Monthly seleccionado, cancelar
            if (!isCurrentMonthly && e.NewValue == CheckState.Checked && anyMonthlySelected)
            {
                MessageBox.Show("No puedes seleccionar otras opciones si ya tienes una opción mensual seleccionada.", "Selección inválida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.NewValue = CheckState.Unchecked;
                return;
            }

            // Actualizar el estado del calendario según la selección actualizada
            UpdateTipoCalendarioByOpcion(checkedItems);
        }

        private void frmGeneradorPolizas_FormClosing(object sender, FormClosingEventArgs e)
        {
            var date = dateTimePicker1.BoldedDates.FirstOrDefault();
            var key = date.ToString("MMMyy").ToUpperInvariant();
            if (chkTipoCalendario.CheckState == CheckState.Checked)
                key += "_S" + GetWeekOfMonth(date);
            else if (chkTipoCalendario.CheckState == CheckState.Unchecked)
            {
                var last = dateTimePicker1.BoldedDates.LastOrDefault();
                key += date.ToString("dd") + "-" + last.ToString("dd");
            }

            Dates = new Dictionary<string, DateTime[]>();
            Dates.Add(key.Replace(".", ""), dateTimePicker1.BoldedDates.ToArray());
        }

        private void chkListTipoPoliza_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (chkListTipoPoliza.SelectedIndex >= 0)
            {
                int idx = chkListTipoPoliza.SelectedIndex;
                bool isChecked = chkListTipoPoliza.GetItemChecked(idx);
                chkListTipoPoliza.SetItemChecked(idx, !isChecked);
            }
        }

        private void UpdateTipoCalendarioByOpcion(List<string> checkedItems)
        {
            // IP y CM: mensual y readonly
            bool isIPorCMSelected = checkedItems.Contains("CM - Cierre Mensual") || checkedItems.Contains("IP- Ingresos Pequeños");
            if (isIPorCMSelected)
            {
                chkTipoCalendario.CheckState = CheckState.Indeterminate; // Indeterminate = "Por Mes"
                chkTipoCalendario.Enabled = false;
                chkTipoCalendario.Text = "Por Mes";
                chkTipoCalendario_CheckStateChanged(null, null); // Update UI
            }
            else
            {
                chkTipoCalendario.Enabled = true;
                // No change to CheckState, user can select
                chkTipoCalendario.CheckState = CheckState.Checked; 
                chkTipoCalendario.Text = "Por Semana";
                chkTipoCalendario_CheckStateChanged(null, null); // Update UI
            }
        }
    }
}
