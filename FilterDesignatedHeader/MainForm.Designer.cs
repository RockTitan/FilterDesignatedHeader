﻿namespace FilterDesignatedHeader
{
    partial class MainForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.button_Exit = new System.Windows.Forms.Button();
            this.listBox_SelectItems = new System.Windows.Forms.ListBox();
            this.comboBox_Sheet = new System.Windows.Forms.ComboBox();
            this.textBox_Input = new System.Windows.Forms.TextBox();
            this.checkBox_Topmost = new System.Windows.Forms.CheckBox();
            this.label_Sheet = new System.Windows.Forms.Label();
            this.label_Input = new System.Windows.Forms.Label();
            this.groupBox_SelectInput = new System.Windows.Forms.GroupBox();
            this.groupBox_Output = new System.Windows.Forms.GroupBox();
            this.textBox_Output = new System.Windows.Forms.TextBox();
            this.button_SelectFile = new System.Windows.Forms.Button();
            this.label_File = new System.Windows.Forms.Label();
            this.contextMenuStrip_Output = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.checkBox_Filter = new System.Windows.Forms.CheckBox();
            this.label_OutputHeader = new System.Windows.Forms.Label();
            this.textBox_OutputHeader = new System.Windows.Forms.TextBox();
            this.radioButton_ResultTable = new System.Windows.Forms.RadioButton();
            this.radioButton_OriginalTable = new System.Windows.Forms.RadioButton();
            this.label_Comb = new System.Windows.Forms.Label();
            this.textBox_Comb = new System.Windows.Forms.TextBox();
            this.groupBox_SelectInput.SuspendLayout();
            this.groupBox_Output.SuspendLayout();
            this.contextMenuStrip_Output.SuspendLayout();
            this.SuspendLayout();
            // 
            // button_Exit
            // 
            this.button_Exit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Exit.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Exit.Location = new System.Drawing.Point(472, 595);
            this.button_Exit.Name = "button_Exit";
            this.button_Exit.Size = new System.Drawing.Size(100, 35);
            this.button_Exit.TabIndex = 0;
            this.button_Exit.Text = "Exit";
            this.button_Exit.UseVisualStyleBackColor = true;
            this.button_Exit.Click += new System.EventHandler(this.button_Exit_Click);
            // 
            // listBox_SelectItems
            // 
            this.listBox_SelectItems.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBox_SelectItems.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox_SelectItems.FormattingEnabled = true;
            this.listBox_SelectItems.ItemHeight = 19;
            this.listBox_SelectItems.Location = new System.Drawing.Point(6, 26);
            this.listBox_SelectItems.Name = "listBox_SelectItems";
            this.listBox_SelectItems.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBox_SelectItems.Size = new System.Drawing.Size(204, 289);
            this.listBox_SelectItems.TabIndex = 1;
            this.listBox_SelectItems.SelectedValueChanged += new System.EventHandler(this.listBox_SelectItems_SelectedValueChanged);
            // 
            // comboBox_Sheet
            // 
            this.comboBox_Sheet.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox_Sheet.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox_Sheet.FormattingEnabled = true;
            this.comboBox_Sheet.Location = new System.Drawing.Point(140, 102);
            this.comboBox_Sheet.Name = "comboBox_Sheet";
            this.comboBox_Sheet.Size = new System.Drawing.Size(317, 27);
            this.comboBox_Sheet.TabIndex = 2;
            this.comboBox_Sheet.SelectedIndexChanged += new System.EventHandler(this.comboBox_Sheet_SelectedIndexChanged);
            // 
            // textBox_Input
            // 
            this.textBox_Input.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_Input.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_Input.Location = new System.Drawing.Point(140, 219);
            this.textBox_Input.Name = "textBox_Input";
            this.textBox_Input.Size = new System.Drawing.Size(317, 27);
            this.textBox_Input.TabIndex = 3;
            this.textBox_Input.MouseClick += new System.Windows.Forms.MouseEventHandler(this.textBox_Input_MouseClick);
            this.textBox_Input.TextChanged += new System.EventHandler(this.textBox_Input_TextChanged);
            // 
            // checkBox_Topmost
            // 
            this.checkBox_Topmost.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBox_Topmost.AutoSize = true;
            this.checkBox_Topmost.Checked = true;
            this.checkBox_Topmost.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_Topmost.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_Topmost.Location = new System.Drawing.Point(489, 18);
            this.checkBox_Topmost.Name = "checkBox_Topmost";
            this.checkBox_Topmost.Size = new System.Drawing.Size(83, 23);
            this.checkBox_Topmost.TabIndex = 5;
            this.checkBox_Topmost.Text = "Topmost";
            this.checkBox_Topmost.UseVisualStyleBackColor = true;
            this.checkBox_Topmost.CheckedChanged += new System.EventHandler(this.checkBox_Topmost_CheckedChanged);
            // 
            // label_Sheet
            // 
            this.label_Sheet.AutoSize = true;
            this.label_Sheet.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_Sheet.Location = new System.Drawing.Point(14, 105);
            this.label_Sheet.Name = "label_Sheet";
            this.label_Sheet.Size = new System.Drawing.Size(96, 19);
            this.label_Sheet.TabIndex = 6;
            this.label_Sheet.Text = "Select Sheet :";
            // 
            // label_Input
            // 
            this.label_Input.AutoSize = true;
            this.label_Input.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_Input.Location = new System.Drawing.Point(14, 222);
            this.label_Input.Name = "label_Input";
            this.label_Input.Size = new System.Drawing.Size(108, 19);
            this.label_Input.TabIndex = 7;
            this.label_Input.Text = "Input Headers :";
            // 
            // groupBox_SelectInput
            // 
            this.groupBox_SelectInput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox_SelectInput.Controls.Add(this.listBox_SelectItems);
            this.groupBox_SelectInput.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox_SelectInput.Location = new System.Drawing.Point(12, 258);
            this.groupBox_SelectInput.Name = "groupBox_SelectInput";
            this.groupBox_SelectInput.Size = new System.Drawing.Size(216, 326);
            this.groupBox_SelectInput.TabIndex = 8;
            this.groupBox_SelectInput.TabStop = false;
            this.groupBox_SelectInput.Text = "Select Input Headers";
            // 
            // groupBox_Output
            // 
            this.groupBox_Output.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox_Output.Controls.Add(this.textBox_Output);
            this.groupBox_Output.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox_Output.Location = new System.Drawing.Point(234, 258);
            this.groupBox_Output.Name = "groupBox_Output";
            this.groupBox_Output.Size = new System.Drawing.Size(338, 326);
            this.groupBox_Output.TabIndex = 9;
            this.groupBox_Output.TabStop = false;
            this.groupBox_Output.Text = "Outputs";
            // 
            // textBox_Output
            // 
            this.textBox_Output.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_Output.Location = new System.Drawing.Point(6, 26);
            this.textBox_Output.Multiline = true;
            this.textBox_Output.Name = "textBox_Output";
            this.textBox_Output.Size = new System.Drawing.Size(326, 289);
            this.textBox_Output.TabIndex = 0;
            // 
            // button_SelectFile
            // 
            this.button_SelectFile.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_SelectFile.Location = new System.Drawing.Point(12, 11);
            this.button_SelectFile.Name = "button_SelectFile";
            this.button_SelectFile.Size = new System.Drawing.Size(100, 35);
            this.button_SelectFile.TabIndex = 10;
            this.button_SelectFile.Text = "Select File";
            this.button_SelectFile.UseVisualStyleBackColor = true;
            this.button_SelectFile.Click += new System.EventHandler(this.button_SelectFile_Click);
            // 
            // label_File
            // 
            this.label_File.AutoSize = true;
            this.label_File.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_File.Location = new System.Drawing.Point(14, 60);
            this.label_File.Name = "label_File";
            this.label_File.Size = new System.Drawing.Size(0, 19);
            this.label_File.TabIndex = 11;
            // 
            // contextMenuStrip_Output
            // 
            this.contextMenuStrip_Output.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copyToolStripMenuItem});
            this.contextMenuStrip_Output.Name = "contextMenuStrip_MainForm";
            this.contextMenuStrip_Output.Size = new System.Drawing.Size(105, 26);
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(104, 22);
            this.copyToolStripMenuItem.Text = "Copy";
            this.copyToolStripMenuItem.Click += new System.EventHandler(this.copyToolStripMenuItem_Click);
            // 
            // checkBox_Filter
            // 
            this.checkBox_Filter.AutoSize = true;
            this.checkBox_Filter.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_Filter.Location = new System.Drawing.Point(123, 18);
            this.checkBox_Filter.Name = "checkBox_Filter";
            this.checkBox_Filter.Size = new System.Drawing.Size(99, 23);
            this.checkBox_Filter.TabIndex = 12;
            this.checkBox_Filter.Text = "Filter Excel";
            this.checkBox_Filter.UseVisualStyleBackColor = true;
            this.checkBox_Filter.CheckedChanged += new System.EventHandler(this.checkBox_Filter_CheckedChanged);
            // 
            // label_OutputHeader
            // 
            this.label_OutputHeader.AutoSize = true;
            this.label_OutputHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_OutputHeader.Location = new System.Drawing.Point(14, 182);
            this.label_OutputHeader.Name = "label_OutputHeader";
            this.label_OutputHeader.Size = new System.Drawing.Size(113, 19);
            this.label_OutputHeader.TabIndex = 13;
            this.label_OutputHeader.Text = "Output Header :";
            // 
            // textBox_OutputHeader
            // 
            this.textBox_OutputHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_OutputHeader.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_OutputHeader.Location = new System.Drawing.Point(140, 179);
            this.textBox_OutputHeader.Name = "textBox_OutputHeader";
            this.textBox_OutputHeader.Size = new System.Drawing.Size(317, 27);
            this.textBox_OutputHeader.TabIndex = 14;
            this.textBox_OutputHeader.Text = "結果";
            // 
            // radioButton_ResultTable
            // 
            this.radioButton_ResultTable.AutoSize = true;
            this.radioButton_ResultTable.Checked = true;
            this.radioButton_ResultTable.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButton_ResultTable.Location = new System.Drawing.Point(234, 17);
            this.radioButton_ResultTable.Name = "radioButton_ResultTable";
            this.radioButton_ResultTable.Size = new System.Drawing.Size(107, 23);
            this.radioButton_ResultTable.TabIndex = 15;
            this.radioButton_ResultTable.Text = "Result Table";
            this.radioButton_ResultTable.UseVisualStyleBackColor = true;
            // 
            // radioButton_OriginalTable
            // 
            this.radioButton_OriginalTable.AutoSize = true;
            this.radioButton_OriginalTable.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButton_OriginalTable.Location = new System.Drawing.Point(347, 17);
            this.radioButton_OriginalTable.Name = "radioButton_OriginalTable";
            this.radioButton_OriginalTable.Size = new System.Drawing.Size(118, 23);
            this.radioButton_OriginalTable.TabIndex = 16;
            this.radioButton_OriginalTable.Text = "Original Table";
            this.radioButton_OriginalTable.UseVisualStyleBackColor = true;
            // 
            // label_Comb
            // 
            this.label_Comb.AutoSize = true;
            this.label_Comb.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_Comb.Location = new System.Drawing.Point(14, 144);
            this.label_Comb.Name = "label_Comb";
            this.label_Comb.Size = new System.Drawing.Size(103, 19);
            this.label_Comb.TabIndex = 18;
            this.label_Comb.Text = "Input Column :";
            // 
            // textBox_Comb
            // 
            this.textBox_Comb.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_Comb.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_Comb.Location = new System.Drawing.Point(140, 141);
            this.textBox_Comb.Name = "textBox_Comb";
            this.textBox_Comb.Size = new System.Drawing.Size(317, 27);
            this.textBox_Comb.TabIndex = 17;
            this.textBox_Comb.Text = "組合";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 641);
            this.Controls.Add(this.label_Comb);
            this.Controls.Add(this.textBox_Comb);
            this.Controls.Add(this.radioButton_OriginalTable);
            this.Controls.Add(this.radioButton_ResultTable);
            this.Controls.Add(this.textBox_OutputHeader);
            this.Controls.Add(this.label_OutputHeader);
            this.Controls.Add(this.checkBox_Filter);
            this.Controls.Add(this.label_File);
            this.Controls.Add(this.button_SelectFile);
            this.Controls.Add(this.groupBox_Output);
            this.Controls.Add(this.groupBox_SelectInput);
            this.Controls.Add(this.label_Input);
            this.Controls.Add(this.label_Sheet);
            this.Controls.Add(this.checkBox_Topmost);
            this.Controls.Add(this.textBox_Input);
            this.Controls.Add(this.comboBox_Sheet);
            this.Controls.Add(this.button_Exit);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(600, 680);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Filter Tool";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox_SelectInput.ResumeLayout(false);
            this.groupBox_Output.ResumeLayout(false);
            this.groupBox_Output.PerformLayout();
            this.contextMenuStrip_Output.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_Exit;
        private System.Windows.Forms.ListBox listBox_SelectItems;
        private System.Windows.Forms.ComboBox comboBox_Sheet;
        private System.Windows.Forms.TextBox textBox_Input;
        private System.Windows.Forms.CheckBox checkBox_Topmost;
        private System.Windows.Forms.Label label_Sheet;
        private System.Windows.Forms.Label label_Input;
        private System.Windows.Forms.GroupBox groupBox_SelectInput;
        private System.Windows.Forms.GroupBox groupBox_Output;
        private System.Windows.Forms.Button button_SelectFile;
        private System.Windows.Forms.Label label_File;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip_Output;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.TextBox textBox_Output;
        private System.Windows.Forms.CheckBox checkBox_Filter;
        private System.Windows.Forms.Label label_OutputHeader;
        private System.Windows.Forms.TextBox textBox_OutputHeader;
        private System.Windows.Forms.RadioButton radioButton_ResultTable;
        private System.Windows.Forms.RadioButton radioButton_OriginalTable;
        private System.Windows.Forms.Label label_Comb;
        private System.Windows.Forms.TextBox textBox_Comb;
    }
}

