using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OBCFix
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        private string outputDir;
        private string[] files;
        private List<string> reportSheets;
        private string pwd;

        public Form1()
        {
            InitializeComponent();

            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = false;
            xlApp.Visible = false;
            reportSheets = new List<string>();


        }

        private void button1_Click(object sender, EventArgs e)
        {
            //openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel Files|*.xls*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Multiselect = true;
            openFileDialog1.FileName = String.Empty;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {


                files = openFileDialog1.FileNames;
                PopulateSheetNames(files[0]);
                sheetsList.SelectedIndex = 0;

                fixButton.Text = $"Fix {files.Length} files";

                fixButton.Enabled = true;
            }
        }

        private void PopulateSheetNames(string fileName)
        {
            if (xlApp == null)
                xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileName);

            //sheetNamesCombo.Items.Clear();
            sheetsList.Items.Clear();

            foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
            {
                //sheetNamesCombo.Items.Add(sheet.Name);
                sheetsList.Items.Add(sheet.Name);
            }
            //xlWorkBook.Close();
            //xlApp.Quit();
        }



        private void PopulateColumns(string sheetName, int rowOffset)
        {
            lstCols.Items.Clear();

            if (xlApp == null)
                xlApp = new Excel.Application();
            if (xlWorkBook == null)
                xlWorkBook = xlApp.Workbooks.Open(openFileDialog1.FileNames[0]);


            xlWorkSheet = xlWorkBook.Worksheets[sheetName];

            int colCount = xlWorkSheet.UsedRange.Columns.Count;
            for (int i = 1; i <= colCount; i++)
            {
                try
                {
                    string tmp = (((xlWorkSheet.Cells[rowOffset, i] as Excel.Range).Value2).ToString());
                    //tmp = tmp.RemoveSpecialCharacters();
                    //Debug.Print(tmp);
                    lstCols.Items.Add(tmp);
                }
                catch (Exception e)
                {
                    lstCols.Items.Add("Null");
                }
            }


            //xlWorkBook.Close();
            //xlApp.Quit();
        }

        private void fixButton_Click(object sender, EventArgs e)
        {
            if (reportSheets.Count > 0)
                reportSheets.Clear();

            foreach (int i in sheetsList.SelectedIndices)
                reportSheets.Add(sheetsList.Items[i].ToString());


            FixFiles(outputDir, files);
        }
        public void FixFiles(string outputDir, params string[] srcFiles)
        {

            string[] dateCols = txtDateHeader.Text.Split(',');
            string[] ignoredCols = txtIgnoredCols.Text.Split(',');

            if (xlApp == null)
                xlApp = new Excel.Application();

            foreach (string srcFile in srcFiles)
            {
                try
                {

                    xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                    xlApp.DisplayAlerts = false;
                    xlApp.Visible = false;
                    xlWorkBook = xlApp.Workbooks.Open(srcFile); //Quick Repair MS Office installation if error Unable to cast COM object of type 'Microsoft.Office.Interop.Excel.ApplicationClass' to interface type 'Microsoft.Office.Interop.Excel._Application'. 
                    xlApp.ScreenUpdating = false;

                    foreach (string reportSheet in reportSheets)
                    {
                        xlWorkSheet = xlWorkBook.Worksheets[reportSheet]; //sheet

                        try
                        {
                            if (passwordCheckBox.Checked)
                                xlWorkSheet.Unprotect(pwd); //pwd
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("Wrong password used for " + xlWorkSheet.Name, "Error");
                        }

                        int colCount = xlWorkSheet.UsedRange.Columns.Count;
                        for (int i = 1; i <= colCount; i++)
                        {
                            try
                            {
                                string colName = (((xlWorkSheet.Cells[(int)numHeaderOffset.Value, i] as Excel.Range).Value2).ToString());


                                if (chkIgnoreCols.Checked)
                                {
                                    int pos = Array.IndexOf(ignoredCols, colName);
                                    if (pos > -1)
                                        continue;
                                }

                                xlWorkSheet.Columns[i].TextToColumns();

                                if (chkFixDates.Checked)
                                {
                                    int pos = Array.IndexOf(dateCols, colName);
                                    if (pos > -1)
                                        xlWorkSheet.Columns[i].NumberFormat = "dd mmmm yyyy";
                                    else
                                        xlWorkSheet.Columns[i].NumberFormat = "0";
                                }
                                else
                                {
                                    xlWorkSheet.Columns[i].NumberFormat = "0";
                                }
                                //Debug.Print($"{colName} ignored");

                                //fix name
                                if (chkCleanHeaderNames.Checked)
                                    (xlWorkSheet.Cells[(int)numHeaderOffset.Value, i] as Excel.Range).Value2 =
                                        colName.RemoveSpecialCharacters();

                            }
                            catch (Exception e) { continue; }
                        }
                    }
                    outputDir = Path.GetDirectoryName(openFileDialog1.FileNames[0]) + "\\Fixed\\";
                    Directory.CreateDirectory(outputDir);

                    string editedFileName = outputDir + Path.GetFileName(srcFile);
                    xlWorkBook.SaveAs(editedFileName);
                    xlWorkBook.Close();


                    //if (chkOpenAfterFix.Checked)Process.Start(editedFileName);
                }
                catch (Exception)
                {
                    MessageBox.Show("Failed, exitting...", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                    xlApp.Quit();//recent
                    Application.Exit();
                }
            }

            //Cleanup
            /*try
            {
                xlApp.Quit();


                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception e) { }*/

            MessageBox.Show(srcFiles.Length + " files were fixed successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

            if (chkOpenAfterFix.Checked)
                Process.Start(Path.GetDirectoryName(outputDir));
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Location = new System.Drawing.Point(this.Location.X, this.Location.Y - 200);
        }

        private void passwordCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            pwdBox.Enabled = passwordCheckBox.Checked;
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            pwdBox.UseSystemPasswordChar = false;
        }

        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            pwdBox.UseSystemPasswordChar = true;
        }

        private void numHeaderOffset_ValueChanged(object sender, EventArgs e)
        {
            PopulateColumns(sheetsList.SelectedItem.ToString(), (int)numHeaderOffset.Value);
        }

        private void sheetsList_SelectedIndexChanged(object sender, EventArgs e)
        {

            PopulateColumns(sheetsList.SelectedItem.ToString(), (int)numHeaderOffset.Value);
            lblColumns.Text = $"{sheetsList.SelectedItems[0]} Columns";

            fixButton.Enabled = sheetsList.SelectedItems.Count > 0;

            fixButton.Text = sheetsList.SelectedItems.Count > 1 ? $"Fix Multiple Sheets in {files.Length} files"
                : $"Fix {sheetsList.SelectedItems[0]} in {files.Length} files";
        }

        private void chkFixDates_CheckedChanged(object sender, EventArgs e)
        {
            txtDateHeader.Enabled = chkFixDates.Checked;
        }

        private void chkIgnoreCols_CheckedChanged(object sender, EventArgs e)
        {
            txtIgnoredCols.Enabled = chkIgnoreCols.Checked;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("https://panettonegames.com/");
        }

        private void lstCols_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedCol = ((ListBox)sender).SelectedItem.ToString();
            toolTip1.SetToolTip((ListBox)sender,
                "Double Click to Ignore or mark as a Date"
                );
        }

        private void lstCols_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            toolTip1.SetToolTip((ListBox)sender, null);
            string colName = lstCols.SelectedItem.ToString();
            if (colName.ToUpper().Contains("DATE"))
            {
                string[] dateCols = txtDateHeader.Text.Split(',');
                int pos = Array.IndexOf(dateCols, colName);
                if (pos == -1)
                {
                    chkFixDates.Checked = true;
                    txtDateHeader.Enabled = true;
                    if (txtDateHeader.Text.Length > 0)
                        txtDateHeader.Text += $",{colName}";
                    else
                        txtDateHeader.Text += $"{colName}";
                }
            }
            else //ignore
            {
                string[] ignoredCols = txtIgnoredCols.Text.Split(',');
                int pos = Array.IndexOf(ignoredCols, colName);
                if (pos == -1)
                {
                    chkIgnoreCols.Checked = true;
                    txtIgnoredCols.Enabled = true;
                    if (txtIgnoredCols.Text.Length > 0)
                        txtIgnoredCols.Text += $",{colName}";
                    else
                        txtIgnoredCols.Text += $"{colName}";

                }
            }

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            xlWorkBook.Close();
            xlApp.Quit();//recent

        }
    }


}

public static class StringUtils
{
    public static string RemoveSpecialCharacters(this string str)
    {
        return Regex.Replace(str, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
    }
}

