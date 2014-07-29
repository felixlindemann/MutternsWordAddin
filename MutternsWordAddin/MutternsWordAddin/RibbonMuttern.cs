using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace MutternsWordAddin
{
    public partial class RibbonMuttern
    {
        private void RibbonMuttern_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void buttonAveryZweckform_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Process.Start("http://www.avery-zweckform.com/avery/de_de/Vorlagen-und-Software/Software/Avery-Zweckform-Assistent-fuer-Microsoft-Office.htm");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler beim Öffnen der Website des Avery Assistenten");
            }
        }

        private void btnStep1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                MessageBox.Show("Bitte nun mit Hilfe des Avery-Zweckform-Assistenten eine geeignete Blanko-Vorlage auswählen und ein leeres(!) neues(!) Dokument erstellen.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler beim Anzeigen der Fehlermeldung");
            }
        }

        private void buttonFillContent_Click(object sender, RibbonControlEventArgs e)
        {

            RegistryKey key = Registry.CurrentUser.OpenSubKey("Software", true);
            FormAssistent frm = new FormAssistent();
            string prefix = "";
            long offset = 1;
            long pages = 1;
            long step = 1;

            key = key.CreateSubKey("LIMETREE", RegistryKeyPermissionCheck.ReadWriteSubTree);
            key = key.CreateSubKey("MUTTERNSTOOL", RegistryKeyPermissionCheck.ReadWriteSubTree);
             
            try
            {
                frm.textBoxPrefix.Text = key.GetValue("prefix", prefix).ToString();
                frm.numericUpDownPages.Value = long.Parse(key.GetValue("pages", pages).ToString());
                frm.numericUpDownOffset.Value = long.Parse(key.GetValue("offset", offset).ToString());
                frm.numericUpDownStep.Value = long.Parse(key.GetValue("step", step).ToString());
            }
            catch (Exception ex)
            { 
                MessageBox.Show(ex.ToString(), "Fehler beim Lesen der Default-Werte.");
            }
            try
            {
                 

                if (frm.ShowDialog() == DialogResult.OK)
                {

                    offset = (long)frm.numericUpDownOffset.Value;
                    step = (long)frm.numericUpDownStep.Value;
                    pages = (long)frm.numericUpDownPages.Value;
                    prefix = frm.textBoxPrefix.Text;


                    frm = null;

                    Document doc = Globals.ThisAddIn.Application.ActiveDocument;

                    if (doc.Tables.Count != 1)
                    {
                        throw new Exception("Achtung es wurde keine gültige Vorlage gefunden. Bitte Prüfen Sie die Vorlage des aktuellen Dokuments!");
                    }
                    Table tbl = doc.Tables[1];
                    int i = 1;
                    int j = 1;
                    long r = (pages) * tbl.Rows.Count * tbl.Columns.Count;
                    while (i <= r)
                    {

                        tbl.Rows.Add();
                        if (getNumPages(doc) > pages)
                        {
                            tbl.Rows[tbl.Rows.Count].Delete();
                            break;
                        }
                        i = i + 1;
                    }
                    for (i = 1; i <= tbl.Rows.Count; i++)
                    {
                        for (j = 1; j <= tbl.Columns.Count; j++)
                        {
                            tbl.Cell(i, j).Range.Text = prefix + offset;
                            offset = offset + step;
                        }
                    }

                }

                key.SetValue("offset", offset);
                key.SetValue("step", step);
                key.SetValue("pages", pages);
                key.SetValue("prefix", prefix);

                key.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Fehler beim Füllen des Dokuments.");

            }
        }
        private long getNumPages(Document doc)
        {

            WdStatistic stat = WdStatistic.wdStatisticPages;
            return doc.ComputeStatistics(stat);
        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            AboutBox1 frm = new AboutBox1();
            frm.ShowDialog();

        }

    }
}
