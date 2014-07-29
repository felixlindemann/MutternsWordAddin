namespace MutternsWordAddin
{
    partial class FormAssistent
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormAssistent));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.numericUpDownOffset = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownPages = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownStep = new System.Windows.Forms.NumericUpDown();
            this.textBoxPrefix = new System.Windows.Forms.TextBox();
            this.buttonDoit = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownOffset)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownPages)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStep)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150F));
            this.tableLayoutPanel1.Controls.Add(this.label5, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.label4, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.numericUpDownOffset, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.numericUpDownPages, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.numericUpDownStep, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.textBoxPrefix, 1, 3);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(12, 41);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(442, 119);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // label5
            // 
            this.label5.Dock = System.Windows.Forms.DockStyle.Top;
            this.label5.Location = new System.Drawing.Point(3, 87);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(286, 23);
            this.label5.TabIndex = 4;
            this.label5.Text = "Soll es ein Prefix geben?";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label4
            // 
            this.label4.Dock = System.Windows.Forms.DockStyle.Top;
            this.label4.Location = new System.Drawing.Point(3, 58);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(286, 23);
            this.label4.TabIndex = 3;
            this.label4.Text = "Welcher Abstand soll zwischen zwei Nummern sein?";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label2
            // 
            this.label2.Dock = System.Windows.Forms.DockStyle.Top;
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(286, 23);
            this.label2.TabIndex = 0;
            this.label2.Text = "Welches ist die erste Nummer die gedruckt werden soll?";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.Dock = System.Windows.Forms.DockStyle.Top;
            this.label3.Location = new System.Drawing.Point(3, 29);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(286, 23);
            this.label3.TabIndex = 1;
            this.label3.Text = "Wieviele Seiten sollen gedruckt werden?";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // numericUpDownOffset
            // 
            this.numericUpDownOffset.Dock = System.Windows.Forms.DockStyle.Fill;
            this.numericUpDownOffset.Location = new System.Drawing.Point(295, 3);
            this.numericUpDownOffset.Maximum = new decimal(new int[] {
            -1593835521,
            466537709,
            54210,
            0});
            this.numericUpDownOffset.Name = "numericUpDownOffset";
            this.numericUpDownOffset.Size = new System.Drawing.Size(144, 20);
            this.numericUpDownOffset.TabIndex = 2;
            this.toolTip1.SetToolTip(this.numericUpDownOffset, "Bitte hier einen Wert >0 eingeben");
            this.numericUpDownOffset.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left;
            // 
            // numericUpDownPages
            // 
            this.numericUpDownPages.Dock = System.Windows.Forms.DockStyle.Fill;
            this.numericUpDownPages.Location = new System.Drawing.Point(295, 32);
            this.numericUpDownPages.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDownPages.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownPages.Name = "numericUpDownPages";
            this.numericUpDownPages.Size = new System.Drawing.Size(144, 20);
            this.numericUpDownPages.TabIndex = 2;
            this.toolTip1.SetToolTip(this.numericUpDownPages, "Bitte hier einen Wert zwischen 1 und 10 eingeben");
            this.numericUpDownPages.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left;
            this.numericUpDownPages.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // numericUpDownStep
            // 
            this.numericUpDownStep.Dock = System.Windows.Forms.DockStyle.Fill;
            this.numericUpDownStep.Location = new System.Drawing.Point(295, 61);
            this.numericUpDownStep.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDownStep.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDownStep.Name = "numericUpDownStep";
            this.numericUpDownStep.Size = new System.Drawing.Size(144, 20);
            this.numericUpDownStep.TabIndex = 2;
            this.toolTip1.SetToolTip(this.numericUpDownStep, "Bitte hier einen Wert zwischen 1 und 10 eingeben.");
            this.numericUpDownStep.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left;
            this.numericUpDownStep.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // textBoxPrefix
            // 
            this.textBoxPrefix.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxPrefix.Location = new System.Drawing.Point(295, 90);
            this.textBoxPrefix.Name = "textBoxPrefix";
            this.textBoxPrefix.Size = new System.Drawing.Size(144, 20);
            this.textBoxPrefix.TabIndex = 5;
            this.toolTip1.SetToolTip(this.textBoxPrefix, "Default: Kein Prefix");
            // 
            // buttonDoit
            // 
            this.buttonDoit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonDoit.Location = new System.Drawing.Point(379, 166);
            this.buttonDoit.Name = "buttonDoit";
            this.buttonDoit.Size = new System.Drawing.Size(75, 23);
            this.buttonDoit.TabIndex = 0;
            this.buttonDoit.Text = "Doit";
            this.buttonDoit.UseVisualStyleBackColor = true;
            this.buttonDoit.Click += new System.EventHandler(this.buttonDoit_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.Location = new System.Drawing.Point(298, 166);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 1;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(296, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "Bitte hier den Etikettendruck einstellen";
            // 
            // FormAssistent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(466, 201);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonDoit);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormAssistent";
            this.Text = "Etikett-O-Mat";
            this.Load += new System.EventHandler(this.FormAssistent_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownOffset)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownPages)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownStep)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button buttonDoit;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Label label5;
        public System.Windows.Forms.NumericUpDown numericUpDownOffset;
        public System.Windows.Forms.NumericUpDown numericUpDownPages;
        public System.Windows.Forms.NumericUpDown numericUpDownStep;
        public System.Windows.Forms.TextBox textBoxPrefix;
    }
}