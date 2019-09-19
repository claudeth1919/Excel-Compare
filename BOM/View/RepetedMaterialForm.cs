using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BOM.Model;
using BOM.Tool;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace BOM.View
{
    public partial class RepetedMaterialForm : Form
    {
        private List<Material> _repetedMaterialList;
        private string _filePath;
        private Excel.Workbook workbook;
        public RepetedMaterialForm(List<Material> repetedMaterialList, string filePath)
        {
            InitializeComponent();
            _repetedMaterialList = repetedMaterialList;
            _filePath = filePath;
            this.Panel.HorizontalScroll.Maximum = 0;
            this.Panel.AutoScroll = false;
            this.Panel.VerticalScroll.Visible = false;
            this.Panel.AutoScroll = true;
            this.Panel.FlowDirection = FlowDirection.TopDown;
            this.Panel.Width = this.Width - 10;
            this.Panel.Height = this.Height;
            SetItemList();
            Panel.Focus();
        }

        private void SetItemList()
        {
            Font boldFont = new System.Drawing.Font("Arial Narrow", 12, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            Font standarFont = new System.Drawing.Font("Arial Narrow", 12, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            int heightSpace = 40;
            FlowLayoutPanel flowItemHeader = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Height = 40,
                Width = Panel.Width
            };

            int padingSpace = (int)Panel.Width / 4;
            Label originalCodeHeader = new Label
            {
                Text = $"Original Code",
                Width = padingSpace - 10,
                Height = heightSpace,
                Font = boldFont
            };
            
            
            Label amountHeader = new Label
            {
                Text = $"Amount",
                Width = padingSpace - 10,
                Height = heightSpace,
                Font = boldFont
            };

            Label rowHeader = new Label
            {
                Text = $"Num Row",
                Width = padingSpace - 10,
                Height = heightSpace,
                Font = boldFont
            };
            
            
            Label openExcelBtn = new Label
            {
                Text = $"Acción",
                Width = padingSpace - 10,
                Height = heightSpace,
                Font = boldFont
            };

            flowItemHeader.Controls.Add(originalCodeHeader);

            flowItemHeader.Controls.Add(amountHeader);

            flowItemHeader.Controls.Add(rowHeader);

            flowItemHeader.Controls.Add(openExcelBtn);

            Panel.Controls.Add(flowItemHeader);

            foreach (Material material in _repetedMaterialList)
            {
                FlowLayoutPanel flowItem = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.LeftToRight,
                    Height = heightSpace,
                    Width = Panel.Width
                };

                Label originalCodeItem = new Label
                {
                    Text = $"{material.OriginalCode}",
                    Width = padingSpace,
                    Height = heightSpace,
                    Font = standarFont
                };
                Label amountItem = new Label
                {
                    Text = $"{material.Amount}",
                    Width = padingSpace,
                    Height = heightSpace,
                    Font = standarFont
                };
                
                Label row = new Label
                {
                    Text = $"{material.RowNum} ({material.SheetName})",
                    Width = padingSpace,
                    Height = heightSpace,
                    Font = standarFont
                };
                
                flowItem.Controls.Add(originalCodeItem);

                flowItem.Controls.Add(amountItem);

                flowItem.Controls.Add(row);

                Button btn_1 = null;
                btn_1 = new Button()
                {
                    Text = "Open",
                    Name = $"{material.RowNum};{material.ColNum};{material.SheetName};{this._filePath}"
                };
                btn_1.Click += Click_OpenExcel;
                flowItem.Controls.Add(btn_1);
                Panel.Controls.Add(flowItem);
            }
        }

        private void Click_OpenExcel(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            string[] dataFromExcel = btn.Name.Split(';');
            if (dataFromExcel.Length != 4)
            {
                Util.ShowMessage(AlarmType.ERROR, "There was a problem in Click Event, call IT");
                return;
            }
            int rowNum = (int)Util.ConvertDynamicToDouble(dataFromExcel[0]);
            int colNum = (int)Util.ConvertDynamicToDouble(dataFromExcel[1]);
            string sheetName = dataFromExcel[2];
            string filePath = dataFromExcel[3];
            try
            {
                OpenExcel(filePath, rowNum, colNum, sheetName);
            }
            catch (Exception ex)
            {
                Util.ShowMessage(AlarmType.ERROR, $"There was a problem opening instance (Click_OpenExcel) {ex.Message}");
            }
        }

        private void OpenExcel(string filePath, int rowNum, int colNum, string sheetName)
        {
            Excel.Workbook workbook;
            if (this.workbook == null)
            {
                workbook = ExcelUtil.CreateWorkbook(filePath, true);
                this.workbook = workbook;
            }
            else
            {
                try
                {
                    workbook = this.workbook;
                    workbook.Worksheets[sheetName].Activate();
                    bool openedCell = workbook.Worksheets[sheetName].UsedRange.Cells[rowNum, colNum].Select;
                }
                catch
                {
                    workbook = ExcelUtil.CreateWorkbook(filePath, true);
                    this.workbook = workbook;
                    workbook.Worksheets[sheetName].Activate();
                    bool openedCell = workbook.Worksheets[sheetName].UsedRange.Cells[rowNum, colNum].Select;
                }
            }
            
        }
    }
}


