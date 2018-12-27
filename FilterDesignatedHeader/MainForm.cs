using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;

namespace FilterDesignatedHeader
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        Excel.Application _Excel = null;
        Excel.Workbook book = null;
        Excel.Worksheet sheet = null;
        Excel.Range range = null;

        List<SheetItems> sheetItems = new List<SheetItems>();

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                //顯示版本於Form title
                var Program_Title = this.Text;
                var Program_Version = FileVersionInfo.GetVersionInfo(this.GetType().Assembly.Location).ProductVersion;
                this.Text = string.Format(@"{0}  V{1}", Program_Title, Program_Version);

                checkTopmost();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //throw;
            }
        }

        private void button_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void checkBox_Topmost_CheckedChanged(object sender, EventArgs e)
        {
            checkTopmost();
        }

        private void button_SelectFile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog file = new OpenFileDialog
                {
                    Filter = "Excel files (*.xlsx, *.xls)|*.xlsx;*.xls|All files (*.*)|*.*"
                };
                file.ShowDialog();
                if (file.FileName == string.Empty || file == null)
                {
                    return;
                }
                this.label_File.Text = file.FileName;

                initailExcel();
                getSheetsName();

                //this._Excel.Quit();
                //this._Excel = null;
                ////確認已經沒有excel工作再回收
                //GC.Collect();

                comboBox_Sheet.SelectedIndex = 0;

                //MessageBox.Show("Read Complete!!\r\n \r\nPlease choice Excel sheet name.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //throw;
            }
        }

        private void comboBox_Sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                listBox_SelectItems.Items.Clear();
                foreach (var item in sheetItems)
                {
                    if (item.SheetName == comboBox_Sheet.Text)
                    {
                        sheet = (Excel.Worksheet)_Excel.Sheets[item.SheetName];
                        sheet.Activate();
                        listBox_SelectItems.Items.Add(item.HeaderItem);
                    }
                }
                if (_Excel.Visible == false)
                {
                    _Excel.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //throw;
            }
        }

        private void listBox_SelectItems_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                textBox_Input.Text = string.Empty;
                foreach (var item in listBox_SelectItems.SelectedItems)
                {
                    textBox_Input.Text = (textBox_Input.Text == string.Empty ? textBox_Input.Text : textBox_Input.Text + ", ") + item.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //throw;
            }
        }

        private void textBox_Input_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (listBox_SelectItems.SelectedItems.Count != 0)
                {
                    getOutputs();
                    this._Excel.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //throw;
            }
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //clipboard();
        }

        private void listBox_Ouput_MouseClick(object sender, MouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Right)
            //{
            //    contextMenuStrip_Output.Show(listBox_Ouput, e.X, e.Y);
            //}

        }

        private void textBox_Input_MouseClick(object sender, MouseEventArgs e)
        {
            
        }




        private void checkTopmost()
        {
            if (checkBox_Topmost.Checked)
            {
                this.TopMost = true;
            }
            else
            {
                this.TopMost = false;
            }
        }

        private void initailExcel()
        {
            //檢查PC有無Excel在執行
            bool flag = false;
            foreach (var item in Process.GetProcesses())
            {
                if (item.ProcessName == "EXCEL")
                {
                    flag = true;
                    break;
                }
            }

            if (!flag)
            {
                this._Excel = new Excel.Application();
            }
            else
            {
                object obj = Marshal.GetActiveObject("Excel.Application");//引用已在執行的Excel
                _Excel = obj as Excel.Application;
            }

            this._Excel.Visible = false;//設false效能會比較好
        }

        private void getSheetsName()
        {
            sheet = null;
            string path = label_File.Text;
            try
            {
                book = _Excel.Workbooks.Open(path);
                for (int i = 1; i <= book.Sheets.Count; i++)
                {
                    sheet = book.Worksheets[i];
                    comboBox_Sheet.Items.Add(sheet.Name);

                    int totalColumns = sheet.UsedRange.Columns.Count;

                    for (int col = 1; col <= totalColumns; col++)
                    {
                        range = (Excel.Range)sheet.Cells[1, col];
                        if (range.Value2 != null && range.Value2.ToString().Trim() != string.Empty)
                        {
                            sheetItems.Add(new SheetItems()
                            {
                                SheetName = sheet.Name,
                                HeaderItem = range.Value2.ToString().Trim()
                            });
                        }
                    }
                }
            }
            finally
            {
                //book.Close();
                //book = null;
            }
        }

        private void getOutputs()
        {
            List<string> results = new List<string>();

            sheet.Activate();
            sheet.Application.ActiveWindow.FreezePanes = false;
            sheet.Application.ActiveWindow.SplitRow = 1;
            sheet.Application.ActiveWindow.FreezePanes = true;

            string[] separator = { "," };
            string[] selectedItems = textBox_Input.Text.Split(separator, StringSplitOptions.RemoveEmptyEntries);

            if (sheet.AutoFilter != null)
            {
                sheet.AutoFilterMode = false;
            }

            //Excel.Range firstRow = sheet.Range["1:1"];
            int totalColumns = sheet.UsedRange.Columns.Count;
            int totalRows = sheet.UsedRange.Rows.Count;
            //篩選
            for (int i = 1; i <= totalColumns; i++)
            {
                range = (Excel.Range)sheet.Cells[1, i];
                foreach (var item in selectedItems)
                {
                    if (range.Value2 != null && range.Value2.ToString().Trim() != string.Empty)
                    {
                        if (range.Value2.ToString().Trim() == item.Trim())
                        {
                            sheet.UsedRange.AutoFilter(i, "<>", Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                        }
                    }
                }
            }
            textBox_Output.Text = string.Empty;
            //取得結果
            for (int col = 1; col <= totalColumns; col++)
            {
                range = (Excel.Range)sheet.Cells[1, col];
                if (range.Value2 != null && range.Value2.ToString().Trim() != string.Empty)
                {
                    if (range.Value2.ToString().Trim() == "結果" || range.Value2.ToString().Trim() == "Result")
                    {
                        for (int row = 2; row <= totalRows; row++)
                        {
                            //MessageBox.Show(sheet.Cells[row, col].Address(false, false));

                            if (sheet.Rows[row].Hidden == false && sheet.Cells[row, col].Value2 != null && sheet.Cells[row, col].Value2.ToString().Trim() != string.Empty)
                            {
                                if (textBox_Output.Text.Trim() == string.Empty)
                                {
                                    textBox_Output.Text = sheet.Cells[row, col].Value2.ToString().Trim();
                                }
                                else
                                {
                                    textBox_Output.Text += "\r\n" + sheet.Cells[row, col].Value2.ToString().Trim();
                                }
                            }
                        }
                    }
                }
            }
        }

        //private void clipboard()
        //{
        //    if (listBox_Ouput.SelectedItems.Count != 0)
        //    {
        //        Clipboard.SetText(string.Join(Environment.NewLine, listBox_Ouput.SelectedItems.OfType<string>().ToArray()));
        //    }
        //}

        
    }
}
