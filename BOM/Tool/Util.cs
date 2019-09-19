using BOM.Model;
using BOM.View;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace BOM.Tool
{
    public static class Util
    {
        public static readonly string CURRENT_PATH = Path.GetDirectoryName(Application.ExecutablePath);
        public static List<MyMessageBox> MyMessages = new List<MyMessageBox>();
        public static MyMessageBox lastMessage;


        public static bool IsLike(string completeString, string conteinedString)
        {
            completeString = NormalizeString(completeString);
            conteinedString = NormalizeString(conteinedString);
            if (completeString.IndexOf(conteinedString)!=-1)
            {
                return true;
            }
            return false;
        }
        
        public static void DeleteFileIfExist(string path)
        {
            if (File.Exists(path))
            {
                try
                {
                    File.Delete(path);
                }
                catch (System.IO.IOException ex)
                {
                    return;
                }
            }
        }

        public static bool ExistFile(string path)
        {
            if (File.Exists(path))
            {
                return true;
            }
            return false;
        }

        public static double ConvertDynamicToDouble(dynamic dynamicNumber)
        {
           

            double amount;
            try
            {
                amount = (double)dynamicNumber;
            }
            catch (Exception e)
            {
                amount = 0;
            }
            if(amount == 0)
            {
                string numberString;
                try
                {
                    numberString = (string)dynamicNumber;
                    amount = double.Parse(numberString);
                }
                catch (Exception e)
                {
                    amount = 0;
                }
            }
            return amount;
        }

        public static string ConvertDynamicToString(dynamic dynamicString)
        {
            string newString;
            try
            {
                newString = (string)(dynamicString + "");
            }
            catch (Exception e)
            {
                newString = String.Empty;
            }
            
            return newString;
        }
        

        public static string NormalizeString(string chain)
        {
            string newString = RemoveDiacritics(chain);
            newString = RemoveSpecialCharacters(chain).ToUpper();
            return newString;
        }

        public static string NormalizeStringList(List<string> list)
        {
            string newString = String.Empty;
            foreach (string chain in list)
            {
                newString += NormalizeString(chain) + ' ';
            }
            return newString;
        }

        public static string GetStringListDummie(List<string> list)
        {
            string newString = String.Empty;
            foreach (string chain in list)
            {
                newString += chain + ' ';
            }
            return newString;
        }

        public static bool IsEmptyString(string chain)
        {
            if (chain == ""||chain==null)
            {
                return true;
            }
            return false;
        }

        public static bool IsEmptyGuid(Guid chain)
        {
            if (chain == null)
            {
                return true;
            }
            if (chain.ToString() == "00000000-0000-0000-0000-000000000000")
            {
                return true;
            }
            return false;
        }

        public static bool IsEmail(string chain)
        {
            if (IsEmptyString(chain))
            {
                return false;
            }
            if (chain.IndexOf("@")!=-1&&chain.IndexOf(".")!=-1)
            {
                return true;
            }
            return false;
        }

        public static MyMessageBox ShowMessage(AlarmType alarm, List<string> messages)
        {
            MyMessageBox message = new MyMessageBox(alarm, messages);
            MyMessages.Add(message);
            message.Show();
            message.FormClosed += Message_FormClosed;
            lastMessage = message;
            return message;
        }

        public static MyMessageBox ShowMessage(AlarmType alarm, string message)
        {
            string messageString = message;
            MyMessageBox myMessage = new MyMessageBox(alarm, new List<string>() { messageString });
            myMessage.Show();
            lastMessage = myMessage;
            return myMessage;
        }

        private static void Message_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        public static void CloseMessages()
        {
            foreach (MyMessageBox message in MyMessages)
            {
                MyMessages.Remove(message);
                message.Close();
            }
        }

        public static void CloseMessagesByType(AlarmType alarmType)
        {
            foreach (MyMessageBox message in MyMessages)
            {
                MyMessages.Remove(message);
                if (message.Type == alarmType) message.Close();
            }
        }

        public static bool StringToBool(string chain)
        {
            if (chain == "1" || chain.ToUpper() == "TRUE") return true;
            return false;
        }

        

        public static bool FindAndKillProcess(string name)
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.StartsWith(name))
                {
                    clsProcess.Kill();
                    return true;
                }
            }
            return false;
        }

        static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }
        

        public static string GetNetworkPath(string uncPath, string initialString, string rootString)
        {
            try
            {
                // remove the "\\" from the UNC path and split the path
                string path = String.Empty;
                uncPath = uncPath.Replace(@"\\", "");
                string[] uncParts = uncPath.Split(new char[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
                bool isRoot = false;
                foreach (string part in uncParts)
                {
                    if (isRoot|| part.ToUpper()==rootString.ToUpper())
                    {
                        path += $@"\{part}";
                        isRoot = true;
                    }
                }
                if (isRoot) path = initialString + path;
                return path;
            }
            catch (Exception ex)
            {
                return "[ERROR RESOLVING UNC PATH: " + uncPath + ": " + ex.Message + "]";
            }
        }
        
        public static string RemoveSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        public static List<Dictionary<List<Material>, Material>> CompareList(List<Material> list_1, List<Material> list_2)
        {
            List<Dictionary<List<Material>, Material>> differentItems = new List<Dictionary<List<Material>, Material>>();
            List<Material> removeFromList_1 = new List<Material>();
            List<Material> removeFromList_2 = new List<Material>();
            foreach (Material material_1 in list_1)
            {
                Dictionary<List<Material>, Material> dict = new Dictionary<List<Material>, Material>();
                Material material_2 = list_2.Find(x=> x.Code == material_1.Code);
                if (material_2 == null)
                {
                    dict.Add(list_1, material_1);
                    dict.Add(list_2, null);
                    removeFromList_1.Add(material_1);
                    differentItems.Add(dict);
                }
                else if ((material_1.Code == material_2.Code && (material_1.Amount - material_2.Amount) != 0) || material_1.IsRepeted)
                {
                    dict.Add(list_1, material_1);
                    dict.Add(list_2, material_2);
                    differentItems.Add(dict);
                    removeFromList_2.Add(material_2);
                    removeFromList_1.Add(material_1);
                }
            }

            foreach (Material material in removeFromList_1) list_1.RemoveAll(x => x.Id == material.Id);
            foreach (Material material in removeFromList_2) list_2.RemoveAll(x => x.Id == material.Id);

            foreach (Material material_2 in list_2)
            {
                Dictionary<List<Material>, Material> dict = new Dictionary<List<Material>, Material>();
                Material material_1 = list_1.Find(x => x.Code == material_2.Code);
                if (material_1 == null)
                {
                    dict.Add(list_1, null);
                    dict.Add(list_2, material_2);
                    differentItems.Add(dict);
                }
            }
            return differentItems;
        }
        public static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }
        public static string ColumnIndexToExcelColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
        public static bool FindPatternMatch(string orignalWord, List<string> possibleMatchList)
        {
            foreach (string word in possibleMatchList)
            {
                if (IsLike(orignalWord, word))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
