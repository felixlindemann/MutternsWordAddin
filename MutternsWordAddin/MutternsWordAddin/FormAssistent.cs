using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MutternsWordAddin
{
    public partial class FormAssistent : Form
    {

         


        public FormAssistent()
        {
            InitializeComponent();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();

        }

        private void buttonDoit_Click(object sender, EventArgs e)
        {
            if (this.numericUpDownStep.Value != 1)
            {
                switch (MessageBox.Show("Der eingegebene Wert für die Schrittweite ist nicht 1. Ist dies so gewollt?", "Wert der Schrittweite weicht ab"))
                {
                    case System.Windows.Forms.DialogResult.Yes:
                        break;
                    default:
                        return;
                }
            }
           

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();

        }

        private void FormAssistent_Load(object sender, EventArgs e)
        {
           

        }
    }
}
