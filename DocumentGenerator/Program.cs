using Novacode;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = "\\Data\\" + DateTime.Now.ToString("MMMM", CultureInfo.InvariantCulture) + "_" + DateTime.Now.Year + "_UtilityBill.docx";

            var newUtilityBillPath = CreateFileFromTemplate("\\Data\\Test_Template.docx", fileName);

            var utilityBill = GetTemplate(newUtilityBillPath);

            utilityBill.ReplaceText("{{current_date}}", DateTime.Now.ToLongDateString());
            utilityBill.ReplaceText("{{user_name}}", "Austen");
            utilityBill.ReplaceText("{{boss_name}}", "Boss Austen");

            utilityBill.Save();
        }

        static DocX GetTemplate(string filePath)
        {
            try
            {
                DocX docx = DocX.Load(filePath);

                return docx;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        static string CreateFileFromTemplate(string templatePath, string fileName)
        {
            try
            {
                var sourceFilePath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + templatePath;
                var destinationFilePath = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName + fileName;

                File.Copy(sourceFilePath, destinationFilePath, true);

                return destinationFilePath;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
