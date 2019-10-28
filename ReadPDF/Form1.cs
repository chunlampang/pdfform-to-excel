using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Pdf;
using Microsoft.Office.Interop.Excel;
using ReadPDF.utils;

namespace ReadPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            string inputPath = SettingHelper.ReadSetting(SettingHelper.KEY_INPUT_PATH);
            string outputPath = SettingHelper.ReadSetting(SettingHelper.KEY_OUTPUT_PATH);
            string exportType = SettingHelper.ReadSetting(SettingHelper.KEY_EXPORT_TYPE);

            if (inputPath == null)
                inputPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (outputPath == null)
                outputPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
            if (exportType == null)
                exportType = "excel";

            txtFile.Text = inputPath;
            txtOut.Text = outputPath;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (txtFile.Text.Length == 0 || txtOut.Text.Length == 0)
                return;

            btnExport.Enabled = false;

            try
            {
                DirectoryInfo d = new DirectoryInfo(txtFile.Text);
                string filename = d.Name + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss");

                var files = d.GetFiles("*.pdf");
                if (files.Length == 0)
                    throw new Exception("No PDF file found");

                int success = WriteExcel(files, filename);
                int fail = files.Length - success;

                MessageBox.Show("Saved " + success + " record(s) to " + filename + (fail > 0 ? "\n" + fail + " PDF file(s) fail to read." : "."));
                Process.Start(txtOut.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                btnExport.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = txtFile.Text;
            DialogResult result = folderBrowserDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK)
            {
                txtFile.Text = folderBrowserDialog1.SelectedPath;
                SettingHelper.AddUpdateAppSettings(SettingHelper.KEY_INPUT_PATH, folderBrowserDialog1.SelectedPath);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = txtOut.Text;
            DialogResult result = folderBrowserDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK)
            {
                txtOut.Text = folderBrowserDialog1.SelectedPath;
                SettingHelper.AddUpdateAppSettings(SettingHelper.KEY_OUTPUT_PATH, folderBrowserDialog1.SelectedPath);
            }
        }

        private string[] GetHeaders(PdfAcroForm form)
        {
            
            IDictionary<string, PdfFormField> fields = form.GetFormFields();
            List<string> headers = new List<string>();
            foreach (string key in fields.Keys)
            {
                PdfFormField field = fields[key];

                if (field.GetFieldName() != null)
                    headers.Add(field.GetFieldName().ToString());
            }

            return headers.ToArray();
        }

        private void SetValueToCell(Range cell, PdfFormField field)
        {
            if (field == null)
                return;
            if (field is PdfButtonFormField)
            {
                string trueValue = "Y", falseValue = "N";
                string v;

                if (field.GetValue() == null || (v = field.GetValueAsString()) == "Off")
                {
                    cell.Value = falseValue;
                    cell.Font.Color = Color.Red;
                    return;
                }
                if (v == "Yes" || v.Length == 0)
                    cell.Value = trueValue;
                else
                    cell.Value = v;
            }
            else
            {
                cell.Value = field.GetValueAsString().Replace("\r", "\n");
            }
        }

        private static string ToLiteral(string input)
        {
            using (var writer = new StringWriter())
            {
                using (var provider = CodeDomProvider.CreateProvider("CSharp"))
                {
                    provider.GenerateCodeFromExpression(new CodePrimitiveExpression(input), writer, null);
                    return writer.ToString();
                }
            }
        }
        private int WriteExcel(FileInfo[] files, string filename)
        {
            return WriteExcel(files, filename, null);
        }
        private int WriteExcel(FileInfo[] files, string filename, string[] headers)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            Workbook workbook = excelApp.Workbooks.Add();
            Worksheet worksheet = workbook.Worksheets[1];

            int cell, row = 2, failCount = 0;
            for (int i = 0; i < files.Length; i++)
            {
                var file = files[i];
                PdfReader reader = new PdfReader(file.FullName);
                PdfDocument pdfDoc = new PdfDocument(reader);
                try
                {
                    PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, false);
                    IDictionary<string, PdfFormField> fields = form.GetFormFields();

                    if (headers == null)
                        headers = GetHeaders(form);

                    worksheet.Cells[row, 1] = file.Name;
                    cell = 1;
                    foreach (string header in headers)
                    {
                        PdfFormField field = form.GetField(header);
                        SetValueToCell(worksheet.Cells[row, cell + 1], field);
                        cell++;
                    }
                    row++;
                }
                catch (Exception e)
                {
                    failCount++;
                }
                finally
                {
                    pdfDoc.Close();
                }
            }
            //header
            worksheet.Cells[1, 1] = "File";
            cell = 1;
            foreach (string header in headers)
            {
                worksheet.Cells[1, cell + 1] = header;
                cell++;
            }
            //format
            Range allCells = worksheet.Cells;
            allCells.VerticalAlignment = XlVAlign.xlVAlignTop;
            worksheet.Columns.AutoFit();

            //save and quit
            workbook.Close(true, txtOut.Text + "\\" + filename + ".xlsx");
            excelApp.Quit();

            return files.Length - failCount;
        }
    }
}
