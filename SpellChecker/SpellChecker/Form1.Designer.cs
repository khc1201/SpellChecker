
namespace SpellChecker
{
    partial class MainForm
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.input_sheetName = new System.Windows.Forms.TextBox();
            this.input_ExcelPath = new System.Windows.Forms.TextBox();
            this.btn_BrowseExcel = new System.Windows.Forms.Button();
            this.input_ExceptText = new System.Windows.Forms.TextBox();
            this.input_RangeOfColumn = new System.Windows.Forms.TextBox();
            this.lb_Column = new System.Windows.Forms.Label();
            this.input_StartRow = new System.Windows.Forms.TextBox();
            this.lb_RowRange = new System.Windows.Forms.Label();
            this.input_EndRow = new System.Windows.Forms.TextBox();
            this.lb_ParsingRange = new System.Windows.Forms.Label();
            this.lb_ExcelPath = new System.Windows.Forms.Label();
            this.lb_sheetName = new System.Windows.Forms.Label();
            this.lb_ExceptText = new System.Windows.Forms.Label();
            this.btn_LoadExcel = new System.Windows.Forms.Button();
            this.lb_workComplete = new System.Windows.Forms.Label();
            this.input_resultBox = new System.Windows.Forms.TextBox();
            this.lb_resultPath = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // input_sheetName
            // 
            this.input_sheetName.Location = new System.Drawing.Point(97, 104);
            this.input_sheetName.Name = "input_sheetName";
            this.input_sheetName.Size = new System.Drawing.Size(100, 21);
            this.input_sheetName.TabIndex = 0;
            this.input_sheetName.TextChanged += new System.EventHandler(this.input_sheetName_TextChanged);
            // 
            // input_ExcelPath
            // 
            this.input_ExcelPath.Location = new System.Drawing.Point(97, 45);
            this.input_ExcelPath.Name = "input_ExcelPath";
            this.input_ExcelPath.Size = new System.Drawing.Size(438, 21);
            this.input_ExcelPath.TabIndex = 1;
            this.input_ExcelPath.TextChanged += new System.EventHandler(this.input_ExcelPath_TextChanged);
            // 
            // btn_BrowseExcel
            // 
            this.btn_BrowseExcel.Location = new System.Drawing.Point(584, 45);
            this.btn_BrowseExcel.Name = "btn_BrowseExcel";
            this.btn_BrowseExcel.Size = new System.Drawing.Size(75, 23);
            this.btn_BrowseExcel.TabIndex = 2;
            this.btn_BrowseExcel.Text = "찾기";
            this.btn_BrowseExcel.UseVisualStyleBackColor = true;
            this.btn_BrowseExcel.Click += new System.EventHandler(this.btn_BrowseExcel_Click);
            // 
            // input_ExceptText
            // 
            this.input_ExceptText.Location = new System.Drawing.Point(97, 158);
            this.input_ExceptText.Name = "input_ExceptText";
            this.input_ExceptText.Size = new System.Drawing.Size(438, 21);
            this.input_ExceptText.TabIndex = 3;
            this.input_ExceptText.TextChanged += new System.EventHandler(this.input_ExceptText_TextChanged);
            // 
            // input_RangeOfColumn
            // 
            this.input_RangeOfColumn.Location = new System.Drawing.Point(97, 223);
            this.input_RangeOfColumn.Name = "input_RangeOfColumn";
            this.input_RangeOfColumn.Size = new System.Drawing.Size(100, 21);
            this.input_RangeOfColumn.TabIndex = 4;
            // 
            // lb_Column
            // 
            this.lb_Column.AutoSize = true;
            this.lb_Column.Location = new System.Drawing.Point(97, 205);
            this.lb_Column.Name = "lb_Column";
            this.lb_Column.Size = new System.Drawing.Size(99, 12);
            this.lb_Column.TabIndex = 5;
            this.lb_Column.Text = "대상 열(Column)";
            // 
            // input_StartRow
            // 
            this.input_StartRow.Location = new System.Drawing.Point(259, 223);
            this.input_StartRow.Name = "input_StartRow";
            this.input_StartRow.Size = new System.Drawing.Size(100, 21);
            this.input_StartRow.TabIndex = 6;
            // 
            // lb_RowRange
            // 
            this.lb_RowRange.AutoSize = true;
            this.lb_RowRange.Location = new System.Drawing.Point(260, 205);
            this.lb_RowRange.Name = "lb_RowRange";
            this.lb_RowRange.Size = new System.Drawing.Size(80, 12);
            this.lb_RowRange.TabIndex = 7;
            this.lb_RowRange.Text = "대상 행(Row)";
            this.lb_RowRange.Click += new System.EventHandler(this.label1_Click);
            // 
            // input_EndRow
            // 
            this.input_EndRow.Location = new System.Drawing.Point(393, 223);
            this.input_EndRow.Name = "input_EndRow";
            this.input_EndRow.Size = new System.Drawing.Size(100, 21);
            this.input_EndRow.TabIndex = 8;
            // 
            // lb_ParsingRange
            // 
            this.lb_ParsingRange.AutoSize = true;
            this.lb_ParsingRange.Location = new System.Drawing.Point(373, 226);
            this.lb_ParsingRange.Name = "lb_ParsingRange";
            this.lb_ParsingRange.Size = new System.Drawing.Size(14, 12);
            this.lb_ParsingRange.TabIndex = 9;
            this.lb_ParsingRange.Text = "~";
            // 
            // lb_ExcelPath
            // 
            this.lb_ExcelPath.AutoSize = true;
            this.lb_ExcelPath.Location = new System.Drawing.Point(97, 30);
            this.lb_ExcelPath.Name = "lb_ExcelPath";
            this.lb_ExcelPath.Size = new System.Drawing.Size(57, 12);
            this.lb_ExcelPath.TabIndex = 10;
            this.lb_ExcelPath.Text = "엑셀 주소";
            // 
            // lb_sheetName
            // 
            this.lb_sheetName.AutoSize = true;
            this.lb_sheetName.Location = new System.Drawing.Point(99, 86);
            this.lb_sheetName.Name = "lb_sheetName";
            this.lb_sheetName.Size = new System.Drawing.Size(57, 12);
            this.lb_sheetName.TabIndex = 11;
            this.lb_sheetName.Text = "시트 이름";
            // 
            // lb_ExceptText
            // 
            this.lb_ExceptText.AutoSize = true;
            this.lb_ExceptText.Location = new System.Drawing.Point(97, 140);
            this.lb_ExceptText.Name = "lb_ExceptText";
            this.lb_ExceptText.Size = new System.Drawing.Size(57, 12);
            this.lb_ExceptText.TabIndex = 12;
            this.lb_ExceptText.Text = "제외 단어";
            // 
            // btn_LoadExcel
            // 
            this.btn_LoadExcel.Location = new System.Drawing.Point(97, 362);
            this.btn_LoadExcel.Name = "btn_LoadExcel";
            this.btn_LoadExcel.Size = new System.Drawing.Size(562, 41);
            this.btn_LoadExcel.TabIndex = 13;
            this.btn_LoadExcel.Text = "실행";
            this.btn_LoadExcel.UseVisualStyleBackColor = true;
            this.btn_LoadExcel.Click += new System.EventHandler(this.btn_LoadExcel_Click);
            // 
            // lb_workComplete
            // 
            this.lb_workComplete.AutoSize = true;
            this.lb_workComplete.Location = new System.Drawing.Point(622, 225);
            this.lb_workComplete.Name = "lb_workComplete";
            this.lb_workComplete.Size = new System.Drawing.Size(59, 12);
            this.lb_workComplete.TabIndex = 14;
            this.lb_workComplete.Text = "Complete";
            // 
            // input_resultBox
            // 
            this.input_resultBox.Location = new System.Drawing.Point(624, 265);
            this.input_resultBox.Name = "input_resultBox";
            this.input_resultBox.Size = new System.Drawing.Size(100, 21);
            this.input_resultBox.TabIndex = 15;
            // 
            // lb_resultPath
            // 
            this.lb_resultPath.AutoSize = true;
            this.lb_resultPath.Location = new System.Drawing.Point(624, 247);
            this.lb_resultPath.Name = "lb_resultPath";
            this.lb_resultPath.Size = new System.Drawing.Size(57, 12);
            this.lb_resultPath.TabIndex = 16;
            this.lb_resultPath.Text = "결과 위치";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.lb_resultPath);
            this.Controls.Add(this.input_resultBox);
            this.Controls.Add(this.lb_workComplete);
            this.Controls.Add(this.btn_LoadExcel);
            this.Controls.Add(this.lb_ExceptText);
            this.Controls.Add(this.lb_sheetName);
            this.Controls.Add(this.lb_ExcelPath);
            this.Controls.Add(this.lb_ParsingRange);
            this.Controls.Add(this.input_EndRow);
            this.Controls.Add(this.lb_RowRange);
            this.Controls.Add(this.input_StartRow);
            this.Controls.Add(this.lb_Column);
            this.Controls.Add(this.input_RangeOfColumn);
            this.Controls.Add(this.input_ExceptText);
            this.Controls.Add(this.btn_BrowseExcel);
            this.Controls.Add(this.input_ExcelPath);
            this.Controls.Add(this.input_sheetName);
            this.Name = "MainForm";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox input_sheetName;
        private System.Windows.Forms.TextBox input_ExcelPath;
        private System.Windows.Forms.Button btn_BrowseExcel;
        private System.Windows.Forms.TextBox input_ExceptText;
        private System.Windows.Forms.TextBox input_RangeOfColumn;
        private System.Windows.Forms.Label lb_Column;
        private System.Windows.Forms.TextBox input_StartRow;
        private System.Windows.Forms.Label lb_RowRange;
        private System.Windows.Forms.TextBox input_EndRow;
        private System.Windows.Forms.Label lb_ParsingRange;
        private System.Windows.Forms.Label lb_ExcelPath;
        private System.Windows.Forms.Label lb_sheetName;
        private System.Windows.Forms.Label lb_ExceptText;
        private System.Windows.Forms.Button btn_LoadExcel;
        private System.Windows.Forms.Label lb_workComplete;
        private System.Windows.Forms.TextBox input_resultBox;
        private System.Windows.Forms.Label lb_resultPath;
    }
}

