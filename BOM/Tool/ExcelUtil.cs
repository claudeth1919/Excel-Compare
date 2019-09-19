using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;     
using BOM.Model;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using BOM.View;
using System.Windows.Forms;

namespace BOM.Tool
{
    public static class ExcelUtil
    {
        private static string CurrentExcelOpenPath = String.Empty;
        private const int MIN_COLUMNS_BOM_AMOUNT = 2;
        private const int MIN_COLUMNS_WH_AMOUNT = 5;
        private const int HEADER_COLUMN_TOLERANCE = 4;
        private const int EMPTINESS_ROW_TOLERANCE = 5; //Usado
        private const int NORMAL_COLUMN_AMOUNT = 22;

        private const int MIN_COLUMNS_GENERAL_AMOUNT = 4;

        public static Excel.Workbook CreateWorkbook(string filePath, bool visible = false)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = visible;
            Excel.Workbook workbook;
            try
            {
                workbook = app.Workbooks.Open(filePath, UpdateLinks: 2);
            }
            catch (Exception e)
            {
                Util.ShowMessage(AlarmType.ERROR, "No se pudo abrir el archivo: " + e.Message);
                return null;
            }
            return workbook;
        }
        public static List<Column> GetHeaderFromExcel(string filePath)
        {
            List<string> errorList = new List<string>();
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = CreateWorkbook(filePath);
            List<Column> headercolumns = new List<Column>();
            bool finish = false;
            foreach (Excel._Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Visible == Excel.XlSheetVisibility.xlSheetHidden) continue;
                Excel.Range range = sheet.UsedRange;
                int rowCount = range.Rows.Count;
                int colCount = range.Columns.Count;
                List<Column> tempHeadercolumns = new List<Column>(colCount);
                string sheetName = sheet.Name.ToUpper();
                if (sheetName == Defs.EXCLUDED_SHEET_COSTOS || sheetName == Defs.EXCLUDED_SHEET_LEYENDA)
                {
                    continue;
                }
                for (int rowIndex = 1; rowIndex <= rowCount && !finish; rowIndex++)
                {
                    if (tempHeadercolumns.Count < MIN_COLUMNS_BOM_AMOUNT) //Get Columns names
                    {
                        colCount = colCount > NORMAL_COLUMN_AMOUNT ? NORMAL_COLUMN_AMOUNT  : colCount;
                        for (int colIndex = 1; colIndex <= colCount && !finish; colIndex++)
                        {
                            try
                            {
                                if (range.Cells[rowIndex, colIndex] != null && range.Cells[rowIndex, colIndex].Value2 != null)
                                {
                                    string columnName = (string)range.Cells[rowIndex, colIndex].Value2.ToString();
                                    Column column = new Column(columnName, colIndex);
                                    tempHeadercolumns.Add(column);
                                }
                            }
                            catch (Exception e)
                            {
                                break;
                            }
                            
                        }
                        if (rowIndex == HEADER_COLUMN_TOLERANCE)
                        {
                            headercolumns = tempHeadercolumns;
                            break;
                        }else if (tempHeadercolumns.Count >= MIN_COLUMNS_GENERAL_AMOUNT)
                        {
                            headercolumns = tempHeadercolumns;
                            finish = true;
                            break;
                        }
                    }
                }
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(sheet);
                if (finish) break;
            }
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();


            //close and release
            workbook.Close(0);
            Marshal.ReleaseComObject(workbook);

            //quit and release
            app.Quit();
            Marshal.ReleaseComObject(app);

            return headercolumns;
        }

        public static List<Material> CompareExcelInformation(string filePath, RichTextBox processTextBox, List<Column> columns, string selectedTextColumn, string selectedNumberColumn)
        {
            List<Material> materialList = new List<Material>();
            List<string> errorList = new List<string>();
            ShowChangesRichTextBox(processTextBox, "\nAbriendo Libro de Excel...");
            
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = CreateWorkbook(filePath);

            foreach (Excel._Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Visible == Excel.XlSheetVisibility.xlSheetHidden) continue;
                Excel.Range range = sheet.UsedRange;
                int rowCount = range.Rows.Count;
                int colCount = range.Columns.Count;
                string sheetName = sheet.Name.ToUpper();
                if (sheetName == Defs.EXCLUDED_SHEET_COSTOS || sheetName == Defs.EXCLUDED_SHEET_LEYENDA)
                {
                    continue;
                }
                int rowToleranceIndex = 0;
                List<Column> headercolumns = GetCorrectColumns(sheet, columns, selectedTextColumn, selectedNumberColumn);
                for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    Material material = new Material();
                    string errorItem = String.Empty;
                    string materialMessageItem = String.Empty;
                    Column colMaterialCode = headercolumns.Find(col => Util.IsLike(col.Name, selectedTextColumn));
                    Column colAmount = headercolumns.Find(col => Util.IsLike(col.Name, selectedNumberColumn));
                    var dynamicCode = range.Cells[rowIndex, colMaterialCode.Index].Value;
                    var dynamicAmount = range.Cells[rowIndex, colAmount.Index].Value;

                    string materialCode = Util.ConvertDynamicToString(dynamicCode);

                    if (Util.IsEmptyString(materialCode) || materialCode == selectedTextColumn)
                    {
                        rowToleranceIndex++;
                        if (EMPTINESS_ROW_TOLERANCE < rowToleranceIndex) break;
                        continue;
                    }
                    if (colMaterialCode == null)
                    {
                        errorItem = $"\n*ERROR: Formato no reconocido en la hoja {sheet.Name.ToString()}, la columna del número de material, debe de decir: \"TEXTO BREVE Clave de material del proveedor\"";
                        errorList.Add(errorItem);
                        ShowChangesRichTextBox(processTextBox, errorItem);
                        continue;
                    }
                    if (colAmount == null)
                    {
                        errorItem = $"\n*ERROR: Formato no reconocido en la hoja {sheet.Name.ToString()}, la columna de la Cantidad del producto debe de decir: \"Cantidad\"";
                        errorList.Add(errorItem);
                        ShowChangesRichTextBox(processTextBox, errorItem);
                        continue;
                    }
                    material.Id = Guid.NewGuid();
                    material.Code = Util.NormalizeString(materialCode);
                    material.OriginalCode = materialCode;
                    material.Amount = Util.ConvertDynamicToDouble(dynamicAmount);
                    material.RowNum = rowIndex;
                   
                    //Add Valid material to list
                    if (!Util.IsEmptyString(material.Code) && material.Amount != 0)
                    {
                        if (Util.IsEmptyString(material.Unit))
                        {
                            material.Unit = "(Sin Unidad)";
                        }
                        material.SheetName = sheetName;
                        material.ColNum = colAmount.Index;
                        Material materialAlreadyInsideList= materialList.Find(x => x.Code == material.Code);
                        if (materialAlreadyInsideList!=null) //Weird validation to repeted items, I must change it cuz it doesn't have any sense
                        {
                            materialAlreadyInsideList.IsRepeted = true;
                            material.IsRepeted = true;
                            if (materialAlreadyInsideList.RepetedItemList.Count == 0)
                            {
                                materialAlreadyInsideList.RepetedItemList.Add(Material.Clone(materialAlreadyInsideList));
                            }
                            materialAlreadyInsideList.RepetedItemList.Add(material);
                            materialAlreadyInsideList.Amount += material.Amount;
                        }
                        else
                        {
                            materialList.Add(material);
                        }
                        materialMessageItem = $"\nPROCESANDO: Hoja: {sheet.Name} Num Parte: {material.Code} Cantidad: {material.Amount} Unidad {material.Unit}";
                        ShowChangesRichTextBox(processTextBox, materialMessageItem);
                        rowToleranceIndex = 0;
                    }
                    else if (!Util.IsEmptyString(material.Code))
                    {
                        errorItem = $"No hay información en la columna de {Defs.COL_AMOUNT} en la fila {rowIndex} de la hoja {sheet.Name.ToString()}";
                        errorList.Add(errorItem);
                        //materialList.Add(material);
                        rowToleranceIndex = 0;
                    }
                }
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(sheet);
            }
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();


            //close and release
            workbook.Close(0);
            Marshal.ReleaseComObject(workbook);

            //quit and release
            app.Quit();
            Marshal.ReleaseComObject(app);

            if (errorList.Count > 1)
            {
                Util.ShowMessage(AlarmType.ERROR, errorList);
            }

            return materialList;
        }


        private static List<Column> GetCorrectColumns(Excel._Worksheet sheet, List<Column> currentColumns, string selectedTextColumn, string selectedNumberColumn)
        {
            
            Excel.Range range = sheet.UsedRange;
            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;
            try
            {
                Column colMaterialCode = currentColumns.Find(col => Util.IsLike(col.Name, selectedTextColumn));
                Column colAmount = currentColumns.Find(col => Util.IsLike(col.Name, selectedNumberColumn));
                string nameColFromFirstIteration = range.Cells[1, colMaterialCode.Index].Value2.ToString();
                string amountColFromFirstIteration = range.Cells[1, colAmount.Index].Value2.ToString();
                if (Util.NormalizeString(nameColFromFirstIteration) == Util.NormalizeString(colMaterialCode.Name) && Util.NormalizeString(colAmount.Name) == Util.NormalizeString(amountColFromFirstIteration))
                {
                    return currentColumns;
                }
            }
            catch
            {
                //Think
            }
            List<Column> tempHeadercolumns = new List<Column>();

            for (int rowIndex = 1; rowIndex <= rowCount ; rowIndex++)
            {
                if (tempHeadercolumns.Count < MIN_COLUMNS_BOM_AMOUNT) //Get Columns names
                {
                    colCount = colCount > NORMAL_COLUMN_AMOUNT ? NORMAL_COLUMN_AMOUNT : colCount;
                    for (int colIndex = 1; colIndex <= colCount ; colIndex++)
                    {
                        try
                        {
                            if (range.Cells[rowIndex, colIndex] != null && range.Cells[rowIndex, colIndex].Value2 != null)
                            {
                                string columnName = (string)range.Cells[rowIndex, colIndex].Value2.ToString();
                                Column column = new Column(columnName, colIndex);
                                tempHeadercolumns.Add(column);
                            }
                        }
                        catch (Exception e)
                        {
                            break;
                        }

                    }
                    if (rowIndex == HEADER_COLUMN_TOLERANCE)
                    {
                        return currentColumns;
                    }
                    else if (tempHeadercolumns.Count >= MIN_COLUMNS_GENERAL_AMOUNT)
                    {
                        return tempHeadercolumns;
                    }
                }
            }
            return currentColumns;
        }


        private static void ShowChangesRichTextBox(RichTextBox processTextBox, string messageItem)
        {
            Font font = new Font("Arial Narrow", 13, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            processTextBox.Font = font;
            if(!Util.IsEmptyString(messageItem)) processTextBox.AppendText($"\n {messageItem}");
            processTextBox.SelectionStart = processTextBox.Text.Length;
            // scroll it automatically
            processTextBox.ScrollToCaret();
        }
        

    }
}
