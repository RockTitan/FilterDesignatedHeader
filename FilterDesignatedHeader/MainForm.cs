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
        DataTable dt = new DataTable();
        private bool isFilter = false;


        #region Events
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

        private void checkBox_Filter_CheckedChanged(object sender, EventArgs e)
        {
            checkFilter();
        }

        private void button_SelectFile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog file = new OpenFileDialog
                {
                    Filter = "Excel files (*.xlsx, *.xlsm, *.xls)|*.xlsx;*.xlsm;*.xls|All files (*.*)|*.*"
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
                dt = null;
                listBox_SelectItems.Items.Clear();
                foreach (var item in sheetItems)
                {
                    if (item.SheetName == comboBox_Sheet.Text)
                    {
                        sheet = (Excel.Worksheet)_Excel.Sheets[item.SheetName];
                        sheet.Activate();
                        listBox_SelectItems.Items.Add(item.HeaderItem);

                        Excel.Range allRange = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                        int lastRow = allRange.Row;
                        int lastColumn = allRange.Column;
                        //將Excel存入object, 再存入DataTable
                        object[,] cellValues = (object[,])sheet.Range[(Excel.Range)sheet.Cells[1, 1], (Excel.Range)sheet.Cells[lastRow, lastColumn]].Value2;
                        dt = cellValues.ToDataTable();
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
                if (textBox_Input.Text.Trim() != string.Empty)
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

        #endregion


        #region Methods

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

        private void checkFilter()
        {
            if (checkBox_Filter.Checked)
            {
                isFilter = true;
            }
            else
            {
                isFilter = false;
            }
        }

        private void initailExcel()
        {
            ExcelUtility excelUtility = new ExcelUtility();
            excelUtility.ClearExcelProcess();

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

            //this._Excel.Visible = false;//設false效能會比較好
        }

        private void getSheetsName()
        {
            comboBox_Sheet.Items.Clear();
            sheet = null;
            string path = label_File.Text;
            try
            {
                sheetItems.Clear();
                book = _Excel.Workbooks.Open(path);
                for (int i = 1; i <= book.Sheets.Count; i++)
                {
                    sheet = book.Worksheets[i];
                    comboBox_Sheet.Items.Add(sheet.Name);

                    int totalColumns = sheet.UsedRange.Columns.Count;
                    int totalRows = sheet.UsedRange.Rows.Count;

                    if (radioButton_ResultTable.Checked)
                    {
                        int inputColIndex = 1;

                        //先取得輸入欄位Index
                        for (int col = 1; col <= totalColumns; col++)
                        {
                            range = (Excel.Range)sheet.Cells[1, col];
                            if (range.Value2.ToString().ToUpper().Contains(textBox_Comb.Text.Trim().ToUpper()))
                            {
                                inputColIndex = col;
                            }
                        }

                        //填入篩選項目
                        for (int row = 2; row <= totalRows; row++)
                        {
                            range = (Excel.Range)sheet.Cells[row, inputColIndex];
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
                    else if (radioButton_OriginalTable.Checked)
                    {
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
                sheetItems = sheetItems.Distinct(new SheetItemsComparer()).ToList();
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

            string[] separator = { };
            if (radioButton_ResultTable.Checked)
            {
                separator = new string[] { "," };
            }
            else if (radioButton_OriginalTable.Checked)
            {
                separator = new string[] { ",", " " };
            }
            string[] selectedItems = textBox_Input.Text.Split(separator, StringSplitOptions.RemoveEmptyEntries);

            textBox_Output.Text = string.Empty;

            if (isFilter)
            {
                sheet.Activate();
                sheet.Application.ActiveWindow.FreezePanes = false;
                sheet.Application.ActiveWindow.SplitRow = 1;
                sheet.Application.ActiveWindow.FreezePanes = true;

                if (sheet.AutoFilter != null) { sheet.AutoFilterMode = false; }

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

                //取得結果
                for (int col = 1; col <= totalColumns; col++)
                {
                    range = (Excel.Range)sheet.Cells[1, col];
                    if (range.Value2 != null && range.Value2.ToString().Trim() != string.Empty)
                    {
                        if (range.Value2.ToString().Trim() == "結果" || range.Value2.ToString().Trim().ToUpper() == "RESULT")
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
            else
            {
                int dataTableColCount = dt.Columns.Count;
                int dataTableRowCount = dt.Rows.Count;
                int resultIndex = 0;
                int inputColIndex = 0;

                //先取得結果欄位Index
                for (int col = 0; col < dataTableColCount; col++)
                {
                    if (dt.Columns[col].ColumnName.ToUpper().Contains(textBox_OutputHeader.Text.Trim().ToUpper()))
                    {
                        resultIndex = col;
                    }
                }

                //取得符合選擇項目的match list
                List<MatchItems> matchResults = new List<MatchItems>();
                if (radioButton_ResultTable.Checked)
                {
                    //先取得輸入欄位Index
                    for (int col = 0; col < dataTableColCount; col++)
                    {
                        if (dt.Columns[col].ColumnName.ToUpper().Contains(textBox_Comb.Text.Trim().ToUpper()))
                        {
                            inputColIndex = col;
                        }
                    }

                    for (int row = 0; row < dataTableRowCount; row++)
                    {
                        //MessageBox.Show($"[{row + 1}:{col + 1}] ({dt.Rows[row].Table.Columns[col].ColumnName}) : " + dt.Rows[row][col].ToString().Trim());

                        foreach (var item in selectedItems)
                        {
                            if (item == dt.Rows[row][inputColIndex].ToString().Trim() && dt.Rows[row][inputColIndex].ToString().Trim() != string.Empty)
                            {
                                if (textBox_Output.Text.Trim() == string.Empty)
                                {
                                    textBox_Output.Text = dt.Rows[row][resultIndex].ToString().Trim();
                                }
                                else
                                {
                                    textBox_Output.Text += dt.Rows[row][resultIndex].ToString().Trim() == string.Empty ? string.Empty : "\r\n" + dt.Rows[row][resultIndex].ToString().Trim();
                                }
                            }
                        }
                    }
                }
                else if (radioButton_OriginalTable.Checked)
                {
                    for (int col = 0; col < dataTableColCount; col++)
                    {
                        for (int row = 0; row < dataTableRowCount; row++)
                        {
                            //MessageBox.Show($"[{row + 1}:{col + 1}] ({dt.Rows[row].Table.Columns[col].ColumnName}) : " + dt.Rows[row][col].ToString().Trim());

                            foreach (var item in selectedItems)
                            {
                                if (item == dt.Rows[row].Table.Columns[col].ColumnName && dt.Rows[row][col].ToString().Trim() != string.Empty)
                                {
                                    matchResults.Add(new MatchItems()
                                    {
                                        SelectedHeader = dt.Rows[row].Table.Columns[col].ColumnName,
                                        MatchIndex = row,
                                        MatchItem = dt.Rows[row][col].ToString().Trim(),
                                        Result = dt.Rows[row][resultIndex].ToString().Trim()
                                    });
                                }
                            }
                        }
                        //matchResults = matchResults.Distinct().ToList();
                    }

                    //確認row中出現數目與選擇數目相同則輸出
                    for (int row = 0; row < dataTableRowCount; row++)
                    {
                        int i = 0;
                        foreach (var item in matchResults)
                        {
                            if (item.MatchIndex == row) { i++; }
                        }
                        if (i == selectedItems.Length)
                        {
                            if (textBox_Output.Text.Trim() == string.Empty)
                            {
                                textBox_Output.Text = dt.Rows[row][resultIndex].ToString().Trim();
                            }
                            else
                            {
                                textBox_Output.Text += dt.Rows[row][resultIndex].ToString().Trim() == string.Empty ? string.Empty : "\r\n" + dt.Rows[row][resultIndex].ToString().Trim();
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

        #endregion
        
    }
}
