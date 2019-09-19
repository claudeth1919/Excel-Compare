using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using BOM.Model;
using BOM.Tool;
using BOM.View;
using System.Threading;

namespace BOM
{
    public partial class DropperView : Form
    {
        private string path_1;
        private string path_2;
        public DropperView()
        {
            InitializeComponent();
            this.LoadingImageFile_1.Hide();
            this.LoadingImageFile_2.Hide();
            this.LabelFileName_1.Text = String.Empty;
            this.LabelFileName_2.Text = String.Empty;
            this.Combo_1.Enabled = false;
            this.Combo_2.Enabled = false;
            this.Combo_3.Enabled = false;
            this.Combo_4.Enabled = false;
            this.ComboExtra_1.Enabled = false;
            this.ComboExtra_2.Enabled = false;
            this.BtnCompare.Enabled = false;
            this.path_1 = String.Empty;
            this.path_2 = String.Empty;
        }

        private void CheckFile(string filesPathItem, Source origin)
        {
            List<Column> columns = ExcelUtil.GetHeaderFromExcel(filesPathItem);
            
            if (columns.Count == 0)
            {
                Util.ShowMessage(AlarmType.ERROR, "Revise el formato por favor, hay algo raro con las columnas");
                return;
            }
            List<Column> columnsCopy = new List<Column>(columns);
            List<Column> columnsCopy_2 = new List<Column>(columns);
            columnsCopy_2.Add(new Column(Defs.NONE_COL, Defs.NONE_INDEX_COL));
            switch (origin)
            {
                case Source.FILE_1:

                    Combo_1.DataSource = columns;
                    Combo_2.DataSource = columnsCopy;
                    this.ComboExtra_1.DataSource = columnsCopy_2;
                    path_1 = filesPathItem;
                    Combo_1.Enabled = true;
                    Combo_2.Enabled = true;
                    this.ComboExtra_1.Enabled = true;
                    SetDefaultComboBoxItems(Combo_1, Combo_2, ComboExtra_1);
                    break;
                case Source.FILE_2:
                    Combo_3.DataSource = columns;
                    Combo_4.DataSource = columnsCopy;
                    this.ComboExtra_2.DataSource = columnsCopy_2;
                    path_2 = filesPathItem;
                    Combo_3.Enabled = true;
                    Combo_4.Enabled = true;
                    this.ComboExtra_2.Enabled = true;
                    SetDefaultComboBoxItems(Combo_3, Combo_4, ComboExtra_2);
                    break;
            }
            this.BtnCompare.Enabled = true;
        }

        private void SetDefaultComboBoxItems(ComboBox textCombo, ComboBox numCombo, ComboBox extraCombo)
        {
            List<Column> textColList = (List<Column>)textCombo.DataSource;
            List<Column> numColList = (List<Column>)numCombo.DataSource;
            List<Column> extraColList = (List<Column>)extraCombo.DataSource;
            int textColIndex = textColList.FindIndex(x => Util.FindPatternMatch(x.Name,Defs.COL_NUMBER_ARTICULE_POSSIBLE_LIST));
            int numColIndex = numColList.FindIndex(x => Util.FindPatternMatch(x.Name, Defs.COL_AMOUNT_ARTICULE_POSSIBLE_LIST));
            int extraColIndex = extraColList.FindIndex(x => x.Name == Defs.NONE_COL);
            textCombo.SelectedIndex = textColIndex;
            numCombo.SelectedIndex = numColIndex;
            extraCombo.SelectedIndex = extraColIndex;
        }

        #region Interfaces Interaction

        #region Btn UploadFile
        private void UploadFile(PictureBox imageBoxFile, PictureBox loadingImageFile, Button btnUploadFile, Source origin, Label fileNameLabel)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Excel | *.xlsx", // file types, that will be allowed to upload
                Multiselect = false // allow/deny user to upload more than one file at a time
            };
            if (dialog.ShowDialog() == DialogResult.OK) // if user clicked OK
            {
                imageBoxFile.Hide();
                loadingImageFile.Show();
                btnUploadFile.Enabled = false;
                String path = dialog.FileName; // get name of file
                //Thread thread = new Thread(() => UploadFileAction(path));
                //thread.Name = "UploadFileActionThread";
                //thread.Start();
                if((path_1 == path && Source.FILE_2 == origin )|| (path_2 == path && Source.FILE_1 == origin))
                {
                    Util.ShowMessage(AlarmType.WARNING, "Ya tiene ese archivo cargado");
                }
                else
                {
                    fileNameLabel.Text = Path.GetFileName(path);
                    CheckFile(path, origin);
                }
            }
            imageBoxFile.Show();
            loadingImageFile.Hide();
            btnUploadFile.Enabled = true;
        }

        private void UploadFileBtnFile_1(object sender, EventArgs e)
        {
            UploadFile(this.ImageBoxFile_1, this.LoadingImageFile_1, this.BtnUploadFile_1, Source.FILE_1, this.LabelFileName_1);
        }
        private void UploadFileBtnFile_2(object sender, EventArgs e)
        {
            UploadFile(this.ImageBoxFile_2, this.LoadingImageFile_2, this.BtnUploadFile_2, Source.FILE_2, this.LabelFileName_2);
        }
        #endregion

        #region DragDrop
        private void DragDropFile(DragEventArgs e, PictureBox imageBoxFile, PictureBox loadingImageFile, Button btnUploadFile, Source origin)
        {
            this.Invoke(new MethodInvoker(delegate () {
                imageBoxFile.Hide();
                loadingImageFile.Show();
                btnUploadFile.Enabled = false;
                string[] filePathArray = e.Data.GetData(DataFormats.FileDrop) as string[]; // get all files path droppeds  
                if (filePathArray != null && filePathArray.Any())
                {
                    foreach (string filesPathItem in filePathArray)
                    {
                        if (filesPathItem.ToUpper().IndexOf(".XLSX") != -1)
                        {
                            //Thread thread = new Thread(() => DropFileAction(filesPathItem));
                            //thread.Name = "DropFileActionThread";
                            //thread.Start();
                            CheckFile(filesPathItem, origin);
                        }
                        else
                        {
                            MessageBox.Show("El archivo no es del tipo solicitado");
                        }
                    }
                }
                imageBoxFile.Show();
                loadingImageFile.Hide();
                btnUploadFile.Enabled = true;
            }));
        }

        private void DragDropFile_2(object sender, DragEventArgs e)
        {
            DragDropFile(e, this.ImageBoxFile_1, this.LoadingImageFile_1, this.BtnUploadFile_1, Source.FILE_1);
        }
        private void DragDropFile_1(object sender, DragEventArgs e)
        {
            DragDropFile(e, this.ImageBoxFile_2, this.LoadingImageFile_2, this.BtnUploadFile_2, Source.FILE_2);
        }
        #endregion

        private void Click_Compare(object sender, EventArgs e)
        {
            this.BtnCompare.Enabled = false;
            this.BtnUploadFile_1.Enabled = false;
            this.BtnUploadFile_2.Enabled = false;
            MyMessageBox messa = Util.ShowMessage(AlarmType.LOADING,"Iniciando Proceso...\n");
            List<Column> headers_1 = (List<Column>)Combo_1.DataSource;
            List<Column> headers_2 = (List<Column>)Combo_3.DataSource;
            Column col_1 = (Column)Combo_1.SelectedItem;
            Column col_2 = (Column)Combo_2.SelectedItem;
            Column col_3 = (Column)Combo_3.SelectedItem;
            Column col_4 = (Column)Combo_4.SelectedItem;
            List<Material> materials_excel_1 = ExcelUtil.CompareExcelInformation(path_1, messa.MyRichTextBox, headers_1, col_1.Name, col_2.Name);
            messa.BringToFront();
            List<Material> materials_excel_2 = ExcelUtil.CompareExcelInformation(path_2, messa.MyRichTextBox, headers_2, col_3.Name, col_4.Name);
            List<Dictionary<List<Material>, Material>>  differentList = Util.CompareList(materials_excel_1, materials_excel_2);
            ComparationView comparationView = new ComparationView(differentList, materials_excel_1, materials_excel_2, path_1, path_2);
            comparationView.Show();
            messa.Close();
            CleanForm();
        }

        private void CleanForm()
        {
            Combo_1.DataSource = null;
            Combo_2.DataSource = null;
            Combo_3.DataSource = null;
            Combo_4.DataSource = null;
            ComboExtra_1.DataSource = null;
            ComboExtra_2.DataSource = null;

            Combo_1.Enabled = false;
            Combo_2.Enabled = false;
            Combo_3.Enabled = false;
            Combo_4.Enabled = false;
            ComboExtra_1.Enabled = false;
            ComboExtra_2.Enabled = false;

            this.BtnUploadFile_1.Enabled = true;
            this.BtnUploadFile_2.Enabled = true;
            this.LabelFileName_1.Text = String.Empty;
            this.LabelFileName_2.Text = String.Empty;

            this.path_1 = String.Empty;
            this.path_2 = String.Empty;
        }

        #endregion

    }
}
