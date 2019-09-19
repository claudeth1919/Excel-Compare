using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOM.Model
{
    public class Material
    {
        private Guid id;
        private string name;
        private double amount;
        private int rowNum;
        private string code;
        private string sheetName;
        private int colNum;
        private string originalCode;
        private string providerName;
        private string unit;
        private bool isRepeted;
        private List<Material> repetedItemList;

        public Material()
        {
            isRepeted = false;
            repetedItemList = new List<Material>();
        }

        public static Material Clone(Material material)
        {
            Material tempMaterial = new Material()
            {
                Id = Guid.NewGuid(),
                Amount = material.Amount,
                Code = material.Code,
                ColNum = material.ColNum,
                OriginalCode = material.OriginalCode,
                ProviderName = material.ProviderName,
                SheetName = material.SheetName,
                RowNum = material.RowNum
            };
            return tempMaterial;
        }

        public List<Material> RepetedItemList
        {
            get { return repetedItemList; }
            set { repetedItemList = value; }
        }

        public bool IsRepeted
        {
            get { return isRepeted; }
            set { isRepeted = value; }
        }
        public string OriginalCode
        {
            get { return originalCode; }
            set { originalCode = value; }
        }
        public string SheetName
        {
            get { return sheetName; }
            set { sheetName = value; }
        }
        public double Amount
        {
            get { return amount; }
            set { amount = value; }
        }
        public int RowNum
        {
            get { return rowNum; }
            set { rowNum = value; }
        }
        public string ProviderName
        {
            get { return providerName; }
            set { providerName = value; }
        }
        public string Unit
        {
            get { return unit; }
            set { unit = value; }
        }
        public Guid Id
        {
            get { return id; }
            set { id = value; }
        }
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        public string Code
        {
            get { return code; }
            set { code = value; }
        }
        public int ColNum
        {
            get { return colNum; }
            set { colNum = value; }
        }
    }
}
