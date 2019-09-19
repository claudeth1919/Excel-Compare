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
    //Repeted Materials Message Button
    public class Btn : Button
    {
        List<Material> RepetedMaterials { get; set; }
        public Btn(List<Material> repetedMaterials)
        {
            RepetedMaterials = repetedMaterials;
        }
    }
    public partial class ComparationView : Form
    {
        List<Material> key_1;
        List<Material> key_2;
        List<Dictionary<List<Material>, Material>> materialList;
        private string filePath_1;
        private string filePath_2;
        private Dictionary<string, Excel.Workbook> dict;
        public ComparationView(List<Dictionary<List<Material>, Material>> materialList, List<Material> key_1, List<Material> key_2, string filePath_1, string filePath_2)
        {
            InitializeComponent();
            dict = new Dictionary<string, Excel.Workbook>();
            dict.Add(filePath_1, null);
            dict.Add(filePath_2, null);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            Panel.HorizontalScroll.Maximum = 0;
            Panel.AutoScroll = false;
            Panel.VerticalScroll.Visible = false;
            Panel.AutoScroll = true;
            Panel.WrapContents = false;
            Panel.Width = this.Width-70;
            FileNameLabel_1.Text = Path.GetFileName(filePath_1);
            FileNameLabel_2.Text = Path.GetFileName(filePath_2);
            this.filePath_1 = filePath_1;
            this.filePath_2 = filePath_2;
            this.materialList = materialList;
            this.key_1 = key_1;
            this.key_2 = key_2;
            SetItemList();
            Panel.Focus();
            this.Tittle.Text += $": {materialList.Count}";
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

            int padingSpace = (int) Panel.Width / 7;
            Label originalCodeHeader = new Label
            {
                Text = $"Original Code",
                Width = padingSpace-10,
                Height = heightSpace,
                Font = boldFont
            };

            Label codeHeader = new Label
            {
                Text = $"Code",
                Width = padingSpace-10,
                Height = heightSpace,
                Font = boldFont
            };

            Label amountHeader_1 = new Label
            {
                Text = $"Amount File 1",
                Width = padingSpace - 10,
                Height = heightSpace,
                Font = boldFont
            };

            Label amountHeader_2 = new Label
            {
                Text = $"Amount File 2",
                Width = padingSpace - 10,
                Height = heightSpace,
                Font = boldFont
            };

            Label rowHeader_1 = new Label
            {
                Text = $"Num Row File 1",
                Width = padingSpace - 10,
                Height = heightSpace,
                Font = boldFont
            };

            Label rowHeader_2 = new Label
            {
                Text = $"Num Row File 2",
                Width = padingSpace - 10,
                Height = heightSpace,
                Font = boldFont
            };

            Label openExcelBtn_1 = new Label
            {
                Text = $"Acción File 1",
                Width = padingSpace - 10,
                Height = heightSpace,
                Font = boldFont
            };
            Label openExcelBtn_2 = new Label
            {
                Text = $"Acción File 2",
                Width = padingSpace - 10,
                Height = heightSpace,
                Font = boldFont
            };

            flowItemHeader.Controls.Add(originalCodeHeader);
            //flowItemHeader.Controls.Add(codeHeader);

            flowItemHeader.Controls.Add(amountHeader_1);
            flowItemHeader.Controls.Add(amountHeader_2);
            
            flowItemHeader.Controls.Add(rowHeader_1);
            flowItemHeader.Controls.Add(rowHeader_2);

            flowItemHeader.Controls.Add(openExcelBtn_1);
            flowItemHeader.Controls.Add(openExcelBtn_2);

            Panel.Controls.Add(flowItemHeader);
            int index = 1;
            foreach(Dictionary<List<Material>, Material> comparedMaterial in materialList)
            {
                FlowLayoutPanel flowItem = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.LeftToRight,
                    Height = heightSpace,
                    Width = Panel.Width
                };

                string originalCodeText = comparedMaterial[key_1] == null ? comparedMaterial[key_2].OriginalCode : comparedMaterial[key_1].OriginalCode;
                Label originalCodeItem = new Label
                {
                    Text = $"{index}) {originalCodeText}",
                    Width = padingSpace,
                    Height = heightSpace,
                    Font = standarFont
                };
                //string codeText = comparedMaterial[key_1] == null ? comparedMaterial[key_2].Code : comparedMaterial[key_1].Code;
                //Label codeItem = new Label
                //{
                //    Text = $"{codeText}",
                //    Width = padingSpace,
                //    Height = heightSpace,
                //    Font = standarFont
                //};
                string amountText_1 = comparedMaterial[key_1] == null ? String.Empty: comparedMaterial[key_1].Amount+""; 
                Label amountItem_1 = new Label
                {
                    Text = $"{amountText_1}",
                    Width = padingSpace,
                    Height = heightSpace,
                    Font = standarFont
                };
                if (amountText_1 == String.Empty) amountItem_1.Text = "Ninguno";

                string amountText_2 = comparedMaterial[key_2] == null ? String.Empty : comparedMaterial[key_2].Amount + "";
                Label amountItem_2 = new Label
                {
                    Text = $"{amountText_2}",
                    Width = padingSpace,
                    Height = heightSpace,
                    Font = standarFont
                };
                if (amountText_2 == String.Empty) amountItem_2.Text = "Ninguno";

                string rowText_1 = comparedMaterial[key_1] == null ? String.Empty : comparedMaterial[key_1].RowNum + "";
                string sheetText_1 = comparedMaterial[key_1] == null ? String.Empty : comparedMaterial[key_1].SheetName + "";
                Label row_1 = new Label
                {
                    Text = $"{rowText_1} ({sheetText_1})",
                    Width = padingSpace,
                    Height = heightSpace,
                    Font = standarFont
                };
                if (sheetText_1 == String.Empty) row_1.Text = "Ninguno";
                if (comparedMaterial[key_1] != null)
                {
                    if (comparedMaterial[key_1].IsRepeted)
                    {
                        row_1.Text = "Repetido";
                    }
                }
                    
                string rowText_2 = comparedMaterial[key_2] == null ? String.Empty : comparedMaterial[key_2].RowNum + "";
                string sheetText_2 = comparedMaterial[key_2] == null ? String.Empty : comparedMaterial[key_2].SheetName + "";
                Label row_2 = new Label
                {
                    Text = $"{rowText_2} ({sheetText_2})",
                    Width = padingSpace,
                    Height = heightSpace,
                    Font = standarFont
                };
                if (sheetText_2 == String.Empty) row_2.Text = "Ninguno";
                if (comparedMaterial[key_2] != null)
                {
                    if (comparedMaterial[key_2].IsRepeted)
                    {
                        row_2.Text = "Repetido";
                    }
                }
                flowItem.Controls.Add(originalCodeItem);
                //flowItem.Controls.Add(codeItem);

                flowItem.Controls.Add(amountItem_1);
                flowItem.Controls.Add(amountItem_2);

                flowItem.Controls.Add(row_1);
                flowItem.Controls.Add(row_2);

                Label action_1 = null;
                if (comparedMaterial[key_1] == null)
                {
                    action_1 = new Label
                    {
                        Text = $""
                    };
                    flowItem.Controls.Add(action_1);
                }
                else if (!comparedMaterial[key_1].IsRepeted)
                {
                    Button btn_1 = null;
                    btn_1 = new Button()
                    {
                        Text = "Open",
                        Name = $"{comparedMaterial[key_1].RowNum};{comparedMaterial[key_1].ColNum};{comparedMaterial[key_1].SheetName};{this.filePath_1}"
                    };
                    btn_1.Click += Click_OpenExcel;

                    flowItem.Controls.Add(btn_1);
                }
                else
                {
                    Button btn = null;
                    btn = new Button()
                    {
                        Text = $"List",
                    };
                    btn.Click += new System.EventHandler((sender, e) => Open_RepetedMaterialsView(sender, e, comparedMaterial[key_1].RepetedItemList, filePath_1));
                    flowItem.Controls.Add(btn);
                }

                Label action_2 = null;
                if (comparedMaterial[key_2] == null)
                {
                    action_2 = new Label
                    {
                        Text = $""
                    };
                    flowItem.Controls.Add(action_2);
                }
                else if (!comparedMaterial[key_2].IsRepeted)
                {
                    Button btn_2 = null;
                    btn_2 = new Button()
                    {
                        Text = "Open",
                        Name = $"{comparedMaterial[key_2].RowNum};{comparedMaterial[key_2].ColNum};{comparedMaterial[key_2].SheetName};{this.filePath_2}"
                    };
                    btn_2.Click += Click_OpenExcel;
                    flowItem.Controls.Add(btn_2);
                }
                else
                {
                    Button btn = null;
                    btn = new Button()
                    {
                        Text = $"List",
                    };
                    btn.Click += new System.EventHandler((sender, e) => Open_RepetedMaterialsView(sender, e, comparedMaterial[key_2].RepetedItemList, filePath_2));
                    flowItem.Controls.Add(btn);
                }

                Panel.Controls.Add(flowItem);
                index++;
            }
        }
        private void Open_RepetedMaterialsView(object sender, EventArgs e, List<Material> repetedMaterialList, string filePath)
        {
            RepetedMaterialForm view = new RepetedMaterialForm(repetedMaterialList, filePath);
            view.Show();
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
            catch(Exception ex)
            {
                Util.ShowMessage(AlarmType.ERROR, $"There was a problem opening instance (Click_OpenExcel) {ex.Message}");
            }
        }

        private void OpenExcel(string filePath, int rowNum, int colNum, string sheetName)
        {
            Excel.Workbook workbook;
            if (dict[filePath] == null)
            {
                workbook = ExcelUtil.CreateWorkbook(filePath, true);
                dict[filePath] = workbook;
                workbook.Worksheets[sheetName].Activate();
                bool openedCell = workbook.Worksheets[sheetName].UsedRange.Cells[rowNum, colNum].Select;
            }
            else
            {
                try
                {
                    workbook = dict[filePath];
                    workbook.Worksheets[sheetName].Activate();
                    bool openedCell = workbook.Worksheets[sheetName].UsedRange.Cells[rowNum, colNum].Select;
                }
                catch
                {
                    workbook = ExcelUtil.CreateWorkbook(filePath, true);
                    dict[filePath] = workbook;
                    workbook.Worksheets[sheetName].Activate();
                    bool openedCell = workbook.Worksheets[sheetName].UsedRange.Cells[rowNum, colNum].Select;
                }
            }
            
           
        }

        private void Click_Cancel(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Click_AssignOrder(object sender, EventArgs e)
        {
           
        }
    }
}
