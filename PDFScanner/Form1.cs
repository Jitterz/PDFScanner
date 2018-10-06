using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace PDFScanner
{    
    public partial class Form1 : Form
    {
        // rectA
        private float coordinate1a;
        private float coordinate2a;
        private float coordinate3a;
        private float coordinate4a;
        // rectB
        private float coordinate1b;
        private float coordinate2b;
        private float coordinate3b;
        private float coordinate4b;

        // column row counts
        private int rowCount;

        private string textA;
        private string textB;

        public string pdfFileName;

        private Microsoft.Office.Interop.Excel.Application MyExcel;
        private Microsoft.Office.Interop.Excel.Workbook MyWorkbook;
        private Microsoft.Office.Interop.Excel.Worksheet MyWorksheet;
        private Microsoft.Office.Interop.Excel.Range MyCells;

        private PdfReader reader;

        public Form1()
        {
            InitializeComponent();

            
            rowCount = 1;
            InitializeCoordinates();
        }

        private void InitializeCoordinates()
        {
            coordinate1a = 390.9f;
            coordinate2a = 486.7f;
            coordinate3a = 427.6f;
            coordinate4a = 499.6f;

            coordinate1b = 71f;
            coordinate2b = 634f;
            coordinate3b = 21f;
            coordinate4b = 648f;
        }

        private void EndAndSaveExcel()
        {
            // --------------------------------------------------- END SAVE THE EXCEL FILE ----------------------------------------///
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xls";
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                MyWorkbook.SaveAs(saveFileDialog.FileName);
            }
            MessageBox.Show(saveFileDialog.InitialDirectory + Environment.NewLine + "Filename: " + saveFileDialog.FileName);

            MyExcel.Visible = true;
        }

        private void ReadText()
        {
            //List<string> linestringlist = new List<string>();
            //PdfReader reader = new PdfReader(pdfFileName);
            iTextSharp.text.Rectangle rectA = new iTextSharp.text.Rectangle(coordinate1a, coordinate2a, coordinate3a, coordinate4a);
            iTextSharp.text.Rectangle rectB = new iTextSharp.text.Rectangle(coordinate1b, coordinate2b, coordinate3b, coordinate4b);
            RenderFilter[] renderFilter = new RenderFilter[2];
            renderFilter[0] = new RegionTextRenderFilter(rectA);
            renderFilter[1] = new RegionTextRenderFilter(rectB);
            ITextExtractionStrategy textExtractionStrategyA = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), renderFilter[0]);
            ITextExtractionStrategy textExtractionStrategyB = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), renderFilter[1]);
            textA = PdfTextExtractor.GetTextFromPage(reader, 1, textExtractionStrategyA);
            textB = PdfTextExtractor.GetTextFromPage(reader, 1, textExtractionStrategyB);
        }

        private void AddTextToExcel()
        {              
            MyCells.Item[rowCount, "A"].Value = textA;
            MyCells.Item[rowCount, "B"].Value = textB;
            rowCount++;
        }

        private void buttonOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "PDF Files|*.pdf";
            openFileDialog1.Title = "Select a PDF File";
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // Initialize Excel
                MyExcel = new Microsoft.Office.Interop.Excel.Application();
                MyWorkbook = MyExcel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                MyExcel.StandardFont = "Arial";
                MyExcel.StandardFontSize = 12;
                MyWorksheet = MyExcel.Worksheets.Item[1];
                MyWorksheet.Name = "Exported Items";
                // access the cells
                MyCells = MyWorksheet.Cells;

                reader = new PdfReader(openFileDialog1.OpenFile());
                progressBar1.Visible = true;
                labelReadingPDF.Visible = true;
                progressBar1.Maximum = reader.NumberOfPages;
                progressBar1.Step = 1;
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    ReadText();
                    AddTextToExcel();
                    progressBar1.PerformStep();
                }
                EndAndSaveExcel();
                progressBar1.Visible = false;
                labelReadingPDF.Visible = false;
                rowCount = 1;
            }
            else
            {
                pdfFileName = null;
            }
        }
    }
}
