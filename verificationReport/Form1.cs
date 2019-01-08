using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;


namespace verificationReport
{
    public partial class Form1 : Form
    {
        //WORKBOOK1
        Excel.Application app1 = new Excel.Application();

        public Excel.Workbook app1Workbook;
        public Excel._Worksheet app1Worksheet;
        public Excel.Range app1Range;

        //WORKBOOK2
        Excel.Application app2 = new Excel.Application();

        public Excel.Workbook app2Workbook;
        public Excel._Worksheet app2Worksheet;
        public Excel.Range app2Range;

        //WORKBOOK3
        Excel.Application app3 = new Excel.Application();

        public Excel.Workbook app3Workbook;
        public Excel._Worksheet app3Worksheet;
        public Excel.Range app3Range;

        //WORKBOOK4
        Excel.Application app4 = new Excel.Application();

        public Excel.Workbook app4Workbook;
        public Excel._Worksheet app4Worksheet;
        public Excel.Range app4Range;



        //misc
        int zero = 0;
        string errProneApps = "0";
        int appCounts = 0;
        public Form1()
        {
            InitializeComponent();
            textBox14.Text = zero.ToString();
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void howToUseToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            try
            {

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "C:\\";
                    openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //get path of file
                        filePath = openFileDialog.FileName;

                        //read contents into stream
                        var fileStream = openFileDialog.OpenFile();

                        using (StreamReader reader = new StreamReader(fileStream))
                        {
                            fileContent = reader.ReadToEnd();
                        }


                        textBox1.Text = filePath;
                        app1Workbook = app1.Workbooks.Open(filePath);
                        app1Worksheet = app1Workbook.Sheets[1];
                        app1Range = app1Worksheet.UsedRange;

                        //5.4 checkbox
                        errProneApps = app1Range.Cells[21, 6].Value2.ToString();

                        //4.1 get values for categorical + other source + mixed 
                        double numb1 = 0;
                        numb1 += Convert.ToInt32(app1Range.Cells[8, 8].Value2);
                        numb1 += Convert.ToInt32(app1Range.Cells[9, 8].Value2);
                        numb1 += Convert.ToInt32(app1Range.Cells[11, 8].Value2);
                        textBox8.Text = numb1.ToString();
                        //4.2
                        textBox9.Text = app1Range.Cells[14, 8].Value2.ToString();
                        //4.3
                        textBox10.Text = app1Range.Cells[15, 8].Value2.ToString();
                        //5.5
                        textBox15.Text = numb1.ToString();
                        //GC-ish
                        GC.WaitForPendingFinalizers();

                        app1Workbook.Close(0);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app1);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app1Workbook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app1Range);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app1Worksheet);

                        app1 = null;

                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("An Exception occured. Close the sheet before using this or restart the tool and try again.\n\n More Info: \n\n " + ex);
                if (app1Workbook != null && app1 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app1Workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app1);
                }
                if (app1Range != null && app1Worksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app1Range);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app1Worksheet);
                }

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            try
            {

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "C:\\";
                    openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //get path of file
                        filePath = openFileDialog.FileName;

                        //read contents into stream
                        var fileStream = openFileDialog.OpenFile();

                        using (StreamReader reader = new StreamReader(fileStream))
                        {
                            fileContent = reader.ReadToEnd();
                        }

                        textBox2.Text = filePath;
                        app2Workbook = app2.Workbooks.Open(filePath);
                        app2Worksheet = app2Workbook.Sheets[1];
                        app2Range = app2Worksheet.UsedRange;

                        //4.1b
                        textBox11.Text = app2Range.Cells[11, 38].Value2.ToString();
                        //4.2b
                        textBox12.Text = app2Range.Cells[15, 38].Value2.ToString();
                        //4.3b
                        textBox13.Text = app2Range.Cells[22, 38].Value2.ToString();


                        //GC-ish
                        GC.WaitForPendingFinalizers();

                        app2Workbook.Close(0);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app2Workbook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app2Worksheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app2Range);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app2);
                        app2 = null;


                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("An error occured. Close the sheet being used before using this or restart the tool and try again.\n\n More Info: \n\n " + ex);
                if (app2Workbook != null && app2 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app2Workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app2);
                }
                if (app2Range != null && app2Worksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app2Range);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app2Worksheet);
                }

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "C:\\";
                    openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //get path of file
                        filePath = openFileDialog.FileName;

                        //read contents into stream
                        var fileStream = openFileDialog.OpenFile();

                        using (StreamReader reader = new StreamReader(fileStream))
                        {
                            fileContent = reader.ReadToEnd();
                        }

                        textBox3.Text = filePath;
                        app3Workbook = app3.Workbooks.Open(filePath);
                        app3Worksheet = app3Workbook.Sheets[1];
                        app3Range = app3Worksheet.UsedRange;

                        //this writes to textbox NEED TO PARSE
                        
                        //1.1b
                        int numbOfSchools = 0;
                        for (int x = 12; x < 30; x++)
                        {
                            //string checkString = null;

                            if (app3Range.Cells[x, 2].Value2 != null && app3Range.Cells[x, 2].Value2.ToString() != "Total:")
                            {
                                numbOfSchools += 1;
                            }
                            else { break; };
                        }
                        textBox41.Text = numbOfSchools.ToString();

                        //1.1a
                        textBox5.Text = app3Range.Cells[12 + numbOfSchools, 3].Value2.ToString();

                        //3.2
                        textBox6.Text = app3Worksheet.Cells[17, 8].Value2.ToString();
                        //3.3 this adds three cells adhereing to 2dot rules hence verbosity 
                        Console.WriteLine(numbOfSchools);
                        double threeThree = app3Range.Cells[numbOfSchools + 12, 5].Value2;
                        threeThree += app3Range.Cells[numbOfSchools + 12, 6].Value2;
                        threeThree += app3Range.Cells[numbOfSchools + 12, 8].Value2;
                        textBox7.Text = threeThree.ToString();





                        //not working
                        //textBox6.Text = app3.ActiveSheet[2].app3Range.Cells[17,8].Value2.ToString();
                        //textBox6.Text = app3Range.Cells[17, 8].Value2.ToString();


                        //garbage

                        GC.WaitForPendingFinalizers();

                        app3Workbook.Close(0);

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app3);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app3Worksheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app3Range);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app3Workbook);
                        app3 = null;





                    }
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show("An Exception occured. Close the sheet before using this or restart the tool and try again.\n\n More Info: \n\n " + ex);
                if (app3Workbook != null && app3 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app3Workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app3);
                }
                if (app3Range != null && app3Worksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app3Range);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app3Worksheet);
                }

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            try
            {

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "C:\\";
                    openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //get path of file
                        filePath = openFileDialog.FileName;

                        //read contents into stream
                        var fileStream = openFileDialog.OpenFile();

                        using (StreamReader reader = new StreamReader(fileStream))
                        {
                            fileContent = reader.ReadToEnd();
                        }

                        textBox4.Text = filePath;
                        app4Workbook = app4.Workbooks.Open(filePath);
                        app4Worksheet = app4Workbook.Sheets[1];
                        app4Range = app4Worksheet.UsedRange;

                        //5.8   
                        //1.a,b --Free-Cata
                        textBox17.Text = app4Range.Cells[24, 23].Value2.ToString();
                        textBox18.Text = app4Range.Cells[25, 23].Value2.ToString();
                        //2.1,b
                        textBox20.Text = app4Range.Cells[26, 23].Value2.ToString();
                        textBox19.Text = app4Range.Cells[27, 23].Value2.ToString();
                        //3.1a,b
                        textBox22.Text = app4Range.Cells[28, 23].Value2.ToString();
                        textBox21.Text = app4Range.Cells[29, 23].Value2.ToString();
                        //4.1a.b
                        textBox24.Text = app4Range.Cells[30, 23].Value2.ToString();
                        textBox23.Text = app4Range.Cells[31, 23].Value2.ToString();
                        //2 - 1.a,b
                        textBox32.Text = app4Range.Cells[24, 40].Value2.ToString();
                        textBox31.Text = app4Range.Cells[25, 40].Value2.ToString();
                        //2 - 2.a,b
                        textBox30.Text = app4Range.Cells[26, 40].Value2.ToString();
                        textBox28.Text = app4Range.Cells[27, 40].Value2.ToString();
                        //2 - 3.a,b
                        textBox29.Text = app4Range.Cells[28, 40].Value2.ToString();
                        textBox27.Text = app4Range.Cells[29, 40].Value2.ToString();
                        //2 - 4.a,b
                        textBox26.Text = app4Range.Cells[30, 40].Value2.ToString();
                        textBox25.Text = app4Range.Cells[31, 40].Value2.ToString();
                        //3 - 1
                        textBox40.Text = app4Range.Cells[26, 54].Value2.ToString();
                        textBox39.Text = app4Range.Cells[27, 54].Value2.ToString();
                        //3 - 2
                        textBox38.Text = app4Range.Cells[24, 54].Value2.ToString();
                        textBox36.Text = app4Range.Cells[25, 54].Value2.ToString();
                        //3 - 3
                        textBox37.Text = app4Range.Cells[28, 54].Value2.ToString();
                        textBox35.Text = app4Range.Cells[29, 54].Value2.ToString();
                        //3 - 4
                        textBox34.Text = app4Range.Cells[30, 54].Value2.ToString();
                        textBox33.Text = app4Range.Cells[31, 54].Value2.ToString();



                        //GC-ish
                        GC.WaitForPendingFinalizers();

                        app4Workbook.Close(0);

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app4);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app4Worksheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app4Range);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(app4Workbook);
                        app4 = null;


                    }
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show("An Exception occured. Close the sheet before using this or restart the tool and try again.\n More Info: \n\n " + ex);
                if (app4Workbook != null && app4 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app4Workbook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app4);
                }
                if (app4Range != null && app4Worksheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app4Range);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app4Worksheet);
                }

            }
        }

        private void label6_Click(object sender, EventArgs e)
        {
           
        }

        private void button6_Click(object sender, EventArgs e)
        {
            
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                textBox14.Text = errProneApps.ToString();
            }
            else { textBox14.Text = "0"; }
        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                textBox16.Text = "0";
            }
            else { textBox16.Text = "n/a"; }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void label31_Click(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label54_Click(object sender, EventArgs e)
        {






        }

        private void howToUseToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            string message = "Get the reports from Payschools Admin and save them as Excel files. \n\nDo not open or use the reports generated for the tool and use the tool at the same time, until its fixed.\n\nReports:\n\nReports > Verification Reports > App Statistics at Verification.\nReports > Application > As of App Status (As of Oct 31)\nReports > Eligibility > Eligibility Breakout\nReports > Verfication Report > Verification Statistics";
            MessageBox.Show(message);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void label29_Click(object sender, EventArgs e)
        {

        }
    }//end class 
}
