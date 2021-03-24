namespace ExcelWordCreation
{
    partial class Form1
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.labelCreatePath = new System.Windows.Forms.Label();
            this.buttonFile = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.comboBox4 = new System.Windows.Forms.ComboBox();
            this.comboBox5 = new System.Windows.Forms.ComboBox();
            this.comboBox6 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.textBoxAmountOfLines = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBox7 = new System.Windows.Forms.ComboBox();
            this.comboBox8 = new System.Windows.Forms.ComboBox();
            this.comboBox9 = new System.Windows.Forms.ComboBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.button1.Location = new System.Drawing.Point(2325, 470);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(321, 69);
            this.button1.TabIndex = 0;
            this.button1.Text = "Daten erstellen";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // labelCreatePath
            // 
            this.labelCreatePath.AutoSize = true;
            this.labelCreatePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.labelCreatePath.Location = new System.Drawing.Point(261, 505);
            this.labelCreatePath.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelCreatePath.Name = "labelCreatePath";
            this.labelCreatePath.Size = new System.Drawing.Size(31, 29);
            this.labelCreatePath.TabIndex = 1;
            this.labelCreatePath.Text = "...";
            // 
            // buttonFile
            // 
            this.buttonFile.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.buttonFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.buttonFile.Location = new System.Drawing.Point(2027, 470);
            this.buttonFile.Margin = new System.Windows.Forms.Padding(4);
            this.buttonFile.Name = "buttonFile";
            this.buttonFile.Size = new System.Drawing.Size(290, 69);
            this.buttonFile.TabIndex = 2;
            this.buttonFile.Text = "Ordner wählen...";
            this.buttonFile.UseVisualStyleBackColor = false;
            this.buttonFile.Click += new System.EventHandler(this.buttonFile_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.label1.Location = new System.Drawing.Point(45, 505);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(185, 29);
            this.label1.TabIndex = 3;
            this.label1.Text = "Ausgabeordner:";
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.textBox1.Location = new System.Drawing.Point(42, 37);
            this.textBox1.Margin = new System.Windows.Forms.Padding(4);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(285, 40);
            this.textBox1.TabIndex = 4;
            this.textBox1.Text = "1. Überschrift";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(39, 7);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(191, 25);
            this.label2.TabIndex = 10;
            this.label2.Text = "Spaltenüberschriften";
            // 
            // textBox2
            // 
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.textBox2.Location = new System.Drawing.Point(334, 37);
            this.textBox2.Margin = new System.Windows.Forms.Padding(4);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(285, 40);
            this.textBox2.TabIndex = 11;
            this.textBox2.Text = "2. Überschrift";
            // 
            // textBox3
            // 
            this.textBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.textBox3.Location = new System.Drawing.Point(623, 37);
            this.textBox3.Margin = new System.Windows.Forms.Padding(4);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(285, 40);
            this.textBox3.TabIndex = 12;
            this.textBox3.Text = "3. Überschrift";
            // 
            // textBox4
            // 
            this.textBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.textBox4.Location = new System.Drawing.Point(913, 37);
            this.textBox4.Margin = new System.Windows.Forms.Padding(4);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(285, 40);
            this.textBox4.TabIndex = 13;
            this.textBox4.Text = "4. Überschrift";
            // 
            // textBox5
            // 
            this.textBox5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.textBox5.Location = new System.Drawing.Point(1203, 37);
            this.textBox5.Margin = new System.Windows.Forms.Padding(4);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(285, 40);
            this.textBox5.TabIndex = 14;
            this.textBox5.Text = "5. Überschrift";
            // 
            // textBox6
            // 
            this.textBox6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.textBox6.Location = new System.Drawing.Point(1492, 37);
            this.textBox6.Margin = new System.Windows.Forms.Padding(4);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(285, 40);
            this.textBox6.TabIndex = 15;
            this.textBox6.Text = "6. Überschrift";
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Vorname",
            "Nachname",
            "Vorname + Nachname",
            "€ - Werte",
            "Datum",
            "Fortlaufende Zahl"});
            this.comboBox1.Location = new System.Drawing.Point(42, 118);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(285, 37);
            this.comboBox1.TabIndex = 16;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.textBox_Elements);
            // 
            // comboBox2
            // 
            this.comboBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "Vorname",
            "Nachname",
            "Vorname + Nachname",
            "€ - Werte",
            "Datum",
            "Fortlaufende Zahl"});
            this.comboBox2.Location = new System.Drawing.Point(334, 118);
            this.comboBox2.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(285, 37);
            this.comboBox2.TabIndex = 17;
            this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.textBox_Elements);
            // 
            // comboBox3
            // 
            this.comboBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Items.AddRange(new object[] {
            "Vorname",
            "Nachname",
            "Vorname + Nachname",
            "€ - Werte",
            "Datum",
            "Fortlaufende Zahl"});
            this.comboBox3.Location = new System.Drawing.Point(623, 118);
            this.comboBox3.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(285, 37);
            this.comboBox3.TabIndex = 18;
            this.comboBox3.SelectedIndexChanged += new System.EventHandler(this.textBox_Elements);
            // 
            // comboBox4
            // 
            this.comboBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.comboBox4.FormattingEnabled = true;
            this.comboBox4.Items.AddRange(new object[] {
            "Vorname",
            "Nachname",
            "Vorname + Nachname",
            "€ - Werte",
            "Datum",
            "Fortlaufende Zahl"});
            this.comboBox4.Location = new System.Drawing.Point(913, 118);
            this.comboBox4.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox4.Name = "comboBox4";
            this.comboBox4.Size = new System.Drawing.Size(285, 37);
            this.comboBox4.TabIndex = 19;
            this.comboBox4.SelectedIndexChanged += new System.EventHandler(this.textBox_Elements);
            // 
            // comboBox5
            // 
            this.comboBox5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.comboBox5.FormattingEnabled = true;
            this.comboBox5.Items.AddRange(new object[] {
            "Vorname",
            "Nachname",
            "Vorname + Nachname",
            "€ - Werte",
            "Datum",
            "Fortlaufende Zahl"});
            this.comboBox5.Location = new System.Drawing.Point(1203, 118);
            this.comboBox5.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox5.Name = "comboBox5";
            this.comboBox5.Size = new System.Drawing.Size(285, 37);
            this.comboBox5.TabIndex = 20;
            this.comboBox5.SelectedIndexChanged += new System.EventHandler(this.textBox_Elements);
            // 
            // comboBox6
            // 
            this.comboBox6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.comboBox6.FormattingEnabled = true;
            this.comboBox6.Items.AddRange(new object[] {
            "Vorname",
            "Nachname",
            "Vorname + Nachname",
            "€ - Werte",
            "Datum",
            "Fortlaufende Zahl"});
            this.comboBox6.Location = new System.Drawing.Point(1492, 118);
            this.comboBox6.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox6.Name = "comboBox6";
            this.comboBox6.Size = new System.Drawing.Size(285, 37);
            this.comboBox6.TabIndex = 21;
            this.comboBox6.SelectedIndexChanged += new System.EventHandler(this.textBox_Elements);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(39, 90);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(118, 25);
            this.label3.TabIndex = 22;
            this.label3.Text = "Zufallsdaten";
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(965, 441);
            this.richTextBox1.Margin = new System.Windows.Forms.Padding(6);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(585, 98);
            this.richTextBox1.TabIndex = 25;
            this.richTextBox1.Text = "";
            // 
            // textBoxAmountOfLines
            // 
            this.textBoxAmountOfLines.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.textBoxAmountOfLines.Location = new System.Drawing.Point(1863, 499);
            this.textBoxAmountOfLines.Margin = new System.Windows.Forms.Padding(4);
            this.textBoxAmountOfLines.Name = "textBoxAmountOfLines";
            this.textBoxAmountOfLines.Size = new System.Drawing.Size(134, 40);
            this.textBoxAmountOfLines.TabIndex = 26;
            this.textBoxAmountOfLines.Text = "10";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.14286F);
            this.label4.Location = new System.Drawing.Point(1869, 470);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(117, 25);
            this.label4.TabIndex = 27;
            this.label4.Text = "Datensätze:";
            // 
            // comboBox7
            // 
            this.comboBox7.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.comboBox7.FormattingEnabled = true;
            this.comboBox7.Items.AddRange(new object[] {
            "Vorname",
            "Nachname",
            "Vorname + Nachname",
            "€ - Werte",
            "Datum",
            "Fortlaufende Zahl"});
            this.comboBox7.Location = new System.Drawing.Point(1785, 118);
            this.comboBox7.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox7.Name = "comboBox7";
            this.comboBox7.Size = new System.Drawing.Size(285, 37);
            this.comboBox7.TabIndex = 28;
            this.comboBox7.SelectedIndexChanged += new System.EventHandler(this.textBox_Elements);
            // 
            // comboBox8
            // 
            this.comboBox8.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.comboBox8.FormattingEnabled = true;
            this.comboBox8.Items.AddRange(new object[] {
            "Vorname",
            "Nachname",
            "Vorname + Nachname",
            "€ - Werte",
            "Datum",
            "Fortlaufende Zahl"});
            this.comboBox8.Location = new System.Drawing.Point(2078, 118);
            this.comboBox8.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox8.Name = "comboBox8";
            this.comboBox8.Size = new System.Drawing.Size(285, 37);
            this.comboBox8.TabIndex = 29;
            this.comboBox8.SelectedIndexChanged += new System.EventHandler(this.textBox_Elements);
            // 
            // comboBox9
            // 
            this.comboBox9.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.14286F);
            this.comboBox9.FormattingEnabled = true;
            this.comboBox9.Items.AddRange(new object[] {
            "Vorname",
            "Nachname",
            "Vorname + Nachname",
            "€ - Werte",
            "Datum",
            "Fortlaufende Zahl"});
            this.comboBox9.Location = new System.Drawing.Point(2371, 118);
            this.comboBox9.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox9.Name = "comboBox9";
            this.comboBox9.Size = new System.Drawing.Size(285, 37);
            this.comboBox9.TabIndex = 30;
            this.comboBox9.SelectedIndexChanged += new System.EventHandler(this.textBox_Elements);
            // 
            // textBox7
            // 
            this.textBox7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.textBox7.Location = new System.Drawing.Point(1785, 37);
            this.textBox7.Margin = new System.Windows.Forms.Padding(4);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(285, 40);
            this.textBox7.TabIndex = 31;
            this.textBox7.Text = "7. Überschrift";
            // 
            // textBox8
            // 
            this.textBox8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.textBox8.Location = new System.Drawing.Point(2078, 37);
            this.textBox8.Margin = new System.Windows.Forms.Padding(4);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(285, 40);
            this.textBox8.TabIndex = 32;
            this.textBox8.Text = "8. Überschrift";
            // 
            // textBox9
            // 
            this.textBox9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.14286F);
            this.textBox9.Location = new System.Drawing.Point(2371, 37);
            this.textBox9.Margin = new System.Windows.Forms.Padding(4);
            this.textBox9.Name = "textBox9";
            this.textBox9.Size = new System.Drawing.Size(285, 40);
            this.textBox9.TabIndex = 33;
            this.textBox9.Text = "9. Überschrift";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(2680, 572);
            this.Controls.Add(this.textBox9);
            this.Controls.Add(this.textBox8);
            this.Controls.Add(this.textBox7);
            this.Controls.Add(this.comboBox9);
            this.Controls.Add(this.comboBox8);
            this.Controls.Add(this.comboBox7);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBoxAmountOfLines);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBox6);
            this.Controls.Add(this.comboBox5);
            this.Controls.Add(this.comboBox4);
            this.Controls.Add(this.comboBox3);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonFile);
            this.Controls.Add(this.labelCreatePath);
            this.Controls.Add(this.button1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Excel Dokumente erstellen";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label labelCreatePath;
        private System.Windows.Forms.Button buttonFile;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.ComboBox comboBox4;
        private System.Windows.Forms.ComboBox comboBox5;
        private System.Windows.Forms.ComboBox comboBox6;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.TextBox textBoxAmountOfLines;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboBox7;
        private System.Windows.Forms.ComboBox comboBox8;
        private System.Windows.Forms.ComboBox comboBox9;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.TextBox textBox9;
    }
}

