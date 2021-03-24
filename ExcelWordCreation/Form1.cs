using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelWordCreation
{
    public partial class Form1 : Form
    {
       public Form1()
        {
            InitializeComponent();
            
        }
        private
            List<int> usedColumns = new List<int> { };
  
        //RANDOM CREATIONS ###########################
        private string createRandomCompleteName()
        {
            string randomCustomer ="";
            string [] customerNames = { "Bernd Friedemann", 
                                        "Möller, Ralf",
                                        "Theo Landwig",
                                        "Mael Klaus",
                                        "Laus, Siegbert",
                                        "Jenzig, Mareen",
                                        "Theo Landwig",
                                        "Schmidt, Gunnar",
                                        "Liebmann, Falco",
                                        "Justus Hofmann",
                                        "Jens Melke",
                                        "Pomm Gerd",
                                        "Jungklaus Karsten",
                                        "Löber, Lars",};
            Random random = new Random(Guid.NewGuid().GetHashCode()); //Random Value Checken####################
            randomCustomer= customerNames[random.Next(0, customerNames.Length - 1)];

        

            


            return randomCustomer;

        }
        private string createRandomSurname()
        {
            string line;
            string randomSurname = "";
            List<string> surnames = new List<string> { };
            string workingDir = Directory.GetCurrentDirectory();

            System.IO.StreamReader file = new System.IO.StreamReader($"{workingDir}\\VornamenListe.txt");
            while ((line = file.ReadLine()) != null)
            {
                surnames.Add(line);
            }
            file.Close();
            
            Random random = new Random(Guid.NewGuid().GetHashCode());
            randomSurname= surnames[random.Next(0, surnames.Count - 1)];

            return randomSurname;
        }
        private string createRandomFamilyName()
        {
            string line;
            string randomName = "";
            List<string> names = new List<string> { };
            string workingDir = Directory.GetCurrentDirectory();

            System.IO.StreamReader file = new System.IO.StreamReader($"{workingDir}\\NachnamenListe.txt");
            while ((line = file.ReadLine()) != null)
            {
                names.Add(line);
            }
            file.Close();

            Random random = new Random(Guid.NewGuid().GetHashCode());
            randomName = names[random.Next(0, names.Count - 1)];

            return randomName;
        }
        private string createRandomGood()
        {
            string randomCompany = "";
            string[] companyNames = { "Gerste","Weizen","Dünger EzBN1", "Kartoffel - Rosaria","Hopfen Friess","Kornspäht","Saat 1F","Saat 1B","Kernsaat 4J"};
            Random random = new Random();
            randomCompany = companyNames[random.Next(0, companyNames.Length - 1)];
            return randomCompany;
        }
        private string createRandomFilename()
        {
            

            string filenamePrefix = "Liste Kunde-";
            string addedFilenameParameter = RandomString(4);



            return filenamePrefix+addedFilenameParameter;
        }
        public string RandomString(int size, bool lowerCase = false)
        {
            var builder = new StringBuilder(size);
            Random random = new Random();
            // Unicode/ASCII Letters are divided into two blocks
            // (Letters 65–90 / 97–122):
            // The first group containing the uppercase letters and
            // the second group containing the lowercase.  

            // char is a single Unicode character  
            char offset = lowerCase ? 'a' : 'A';
            const int lettersOffset = 26; // A...Z or a..z: length=26  

            for (var i = 0; i < size; i++)
            {
                var @char = (char)random.Next(offset, offset + lettersOffset);
                builder.Append(@char);
            }

            return lowerCase ? builder.ToString().ToLower() : builder.ToString();
        }
        private void createRandomExcel()
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            string fileName = folderBrowserDialog1.SelectedPath+"\\"+createRandomFilename();
            
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = textBox1.Text;
                oSheet.Cells[1, 2] = textBox2.Text;
                oSheet.Cells[1, 3] = textBox3.Text;
                oSheet.Cells[1, 4] = textBox4.Text;
                oSheet.Cells[1, 5] = textBox5.Text;
                oSheet.Cells[1, 6] = textBox6.Text;
                oSheet.Cells[1, 7] = textBox7.Text;
                oSheet.Cells[1, 8] = textBox8.Text;
                oSheet.Cells[1, 9] = textBox9.Text;

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "I1").Font.Bold = true;
                oSheet.get_Range("A1", "I1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                int parsedString = int.TryParse(textBoxAmountOfLines.Text, out parsedString) ? parsedString : 10;
                //string[,] saNames = new string[parsedString, 6];//in Foreach Schleife für Debugging
                string[,] saNames = new string[parsedString, 9];
                List<Appendix> appendedOptions = new List<Appendix> { };

                foreach (int element in usedColumns)
                {
                    string rebuildCaller = "comboBox" + element;

                    ComboBox comboBox = this.Controls.Find(rebuildCaller, true).FirstOrDefault() as ComboBox;
                    string chosenOption = comboBox.SelectedItem.ToString();

                    //Translate Cell Coordinates
                    string[] columnIndex = { "A", "B", "C", "D", "E", "F","G","H","I" };
                    string startCell = columnIndex[element - 1] + 2;
                    string endCell = columnIndex[element - 1] + textBoxAmountOfLines.Text;
                    
                   
                    switch (chosenOption)
                    {
                        case "Vorname":
                            
                            for (int i = 0; i < parsedString; i++)
                            {
                                saNames[i, element - 1] = createRandomSurname();
                            }

                            break;

                        case "Nachname":
                            for (int i = 0; i < parsedString; i++)
                            {
                                saNames[i, element - 1] = createRandomFamilyName();
                            }
                            
                            break;

                        case "Vorname + Nachname":
                            
                            for (int i = 0; i < parsedString; i++)
                            {
                                saNames[i, element - 1] = createRandomSurname()+" "+createRandomFamilyName();
                            }

                            
                            break;

                        case "Fortlaufende Zahl":
                            for (int i = 0; i < parsedString; i++)
                            {
                                saNames[i, element - 1] = (i+1).ToString();
                            }
                            break;

                        default:
                            endCell = columnIndex[element - 1] + (parsedString+1).ToString();
                            
                            Appendix appendix = new Appendix();
                            appendix.ChosenOption = chosenOption;
                            appendix.StartCell = startCell;
                            appendix.EndCell = endCell;
                            appendix.CalledBox = rebuildCaller;
                            appendedOptions.Add(appendix);
                            break;
                    }
                }
                

                oSheet.get_Range("A2", "I" + (parsedString + 1).ToString()).Value2 = saNames;

                //All Appended Options (Appendix-Object) that are Stored in the List are performt after all other options in order
                //not to get overwritten
                foreach (Appendix Object in appendedOptions)
                {
                    switch (Object.ChosenOption)
                    {
                        case "€ - Werte":
                            string start = "TextBoxStart"+Object.CalledBox;
                            string end = "TextBoxEnd" + Object.CalledBox;

                            string startValue= "0";
                            string endValue = "100";


                            foreach (Control item in this.Controls)
                            {
                                if(item.AccessibleName==start)
                                {
                                    startValue = item.Text;
                                }
                                if (item.AccessibleName == end)
                                {
                                    endValue = item.Text;
                                }
                            }

                            oRng = oSheet.get_Range(Object.StartCell, Object.EndCell);
                            oRng.Formula = $"=RAND()*({endValue}-{startValue})+{startValue}";
                         
                            oRng.NumberFormat = "$0.00";
                            oRng.Copy(Type.Missing);
                            oRng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                            //  = ZUFALLSZAHL() * (b - a) + a //Zufallsbereich
                            break;

                        case "Datum":
                            richTextBox1.AppendText("Datum");

                            break;

                        default:

                            break;



                    }

                }

                
                //Fill with an array of values the specified amount starting at A2
                //oSheet.get_Range("A2", "F"+ (parsedString + 1).ToString()).Value2 = saNames;

                //Wichtig für Später um Formel zu kopieren und einzufügen
                //oRng.Copy(Type.Missing);
                //oRng.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                //oSheet.get_Range("A2", "F" + (parsedString + 1).ToString()).Value2 = saNames;//debug



                // Fill F2:F6 with a relative formula (=E2*1,1). DEBUG
                /*oRng = oSheet.get_Range("F2", "F6");
                oRng.Formula = "=E2*1.1";
                oRng.NumberFormat = "$0.00";

                //Ware auffüllen
                oSheet.get_Range("D2", "D6").Value2 = createRandomGood();*/


                //AutoFit columns A:D.
                oRng = oSheet.get_Range("A1", "I1");
                oRng.EntireColumn.AutoFit();

                oXL.Visible = false;
                oXL.UserControl = false;
                oWB.SaveAs(fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                
                oWB.Close(0);
                oXL.Quit();

                Process[] excelProcs = Process.GetProcessesByName("EXCEL");
                foreach (Process proc in excelProcs)
                {
                    proc.Kill();
                }
            }
            catch
            {
             
            }
        }
      

        //BUTTONS #####################################
        private void buttonFile_Click(object sender, EventArgs e)
        {
            richTextBox1.AppendText(Directory.GetCurrentDirectory());
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    labelCreatePath.Text = folderBrowserDialog1.SelectedPath;
                    button1.Enabled = true;
                    
                }
                catch
                {
                    string message = "Ungültiges Dateiformat / nicht lesbare Datei";
                    string title = "Ladefehler";
                    MessageBox.Show(message, title);
                }
            }
        }
        private void button1_Click(object sender, EventArgs e) //Daten erstellen Knopf
        {
            createRandomExcel();
        }
        private void textBox_Elements(object sender, EventArgs e)
        {
            string chosenItem = (sender as ComboBox).SelectedItem.ToString();
            string calledComboBox = (sender as ComboBox).Name;

            int locationX = (sender as ComboBox).Location.X;
            int locationY = (sender as ComboBox).Location.Y;

            //add the ComboBox to the List for later use in Excel Columns
            int numberOfComboBox = int.TryParse(calledComboBox.Substring(calledComboBox.Length - 1), out numberOfComboBox) ? numberOfComboBox : 10;
            if (!usedColumns.Contains(numberOfComboBox))
            {
                usedColumns.Add(numberOfComboBox);
                richTextBox1.AppendText(numberOfComboBox.ToString() + Environment.NewLine);//DEBUG
                richTextBox1.AppendText(calledComboBox);//DEBUG
            }

            //delete all former Elements -> ist erstmal holzhammer wird dann aber noch spezifiziert und vielleicht zu methode
            foreach (Control item in this.Controls)
            {
                if (item.AccessibleName == "TextBoxStart" + calledComboBox || item.AccessibleName == "TextBoxEnd" + calledComboBox)
                {
                    this.Controls.Remove(item);
                    break; //important step
                }
            }
            foreach (Control item in this.Controls)
            {
                if (item.AccessibleName == "TextBoxEnd" + calledComboBox)
                {
                    this.Controls.Remove(item);
                    break; //important step
                }
            }

            //labelCalledBox.Text = calledComboBox;

            switch(chosenItem)
            {
                case "Vorname":
                    
                    break;
                case "Nachname":
                  

                    break;

                case "Vorname + Nachname":
                    
                    break;

                case "Fortlaufende Zahl":

                    break;

                case "€ - Werte":
                    TextBox startEuro = new TextBox();
                    TextBox endEuro = new TextBox();
                    startEuro.Text = "0";
                    endEuro.Text = "100";
                    startEuro.AccessibleName = "TextBoxStart" + calledComboBox;
                    endEuro.AccessibleName = "TextBoxEnd" + calledComboBox; ;
                    this.Controls.Add(startEuro);
                    this.Controls.Add(endEuro);
                    startEuro.Location = new Point(locationX, locationY+40); //43; 183
                    endEuro.Location = new Point(locationX, locationY+80);
                    startEuro.Height = 25;//157; 25
                    startEuro.Width = 157;
                    endEuro.Height = 25;
                    endEuro.Width = 157;
                    break;

                default:
                    

                    break;
            }




        }

        // Create a TextBox object  
       /* TextBox dynamicTextBox = newTextBox();


           // Set background and foreground  
           dynamicTextBox.BackColor = Color.Red;  
            dynamicTextBox.ForeColor = Color.Blue;  
    dynamicTextBox.Text = "I am Dynamic TextBox";  
    dynamicTextBox.Name = "DynamicTextBox";  
    dynamicTextBox.Font = newFont("Georgia", 16);



            Controls.Add(dynamicTextBox);*/


    }
    class Appendix
    {
        public string ChosenOption
        {
            get { return chosenOption; }
            set { chosenOption = value; }
        }
        public string StartCell
        {
            get { return startCell; }
            set { startCell = value; }
        }
        public string EndCell
        {
            get { return endCell; }
            set { endCell = value; }
        }
        public string CalledBox
        {
            get { return calledBox; }
            set { calledBox = value; }
        }
        string chosenOption;
        string startCell;
        string endCell;
        string calledBox;
        
    }
}
