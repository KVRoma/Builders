using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Section = MigraDoc.DocumentObjectModel.Section;
using Orientation = MigraDoc.DocumentObjectModel.Orientation;
using Paragraph = MigraDoc.DocumentObjectModel.Paragraph;
using Table = MigraDoc.DocumentObjectModel.Tables.Table;
using Image = MigraDoc.DocumentObjectModel.Shapes.Image;
using Style = MigraDoc.DocumentObjectModel.Style;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using System.Xml.XPath;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System.Diagnostics;
using System.Windows.Documents;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using Microsoft.Win32;
using Builders.Models;
using System.IO;

namespace Builders.PDF
{
    public class PrintMaterialProfit
    {
        public PrintMaterialProfit()
        {
            Materials = new List<Material>();           
        }

        private readonly Color colorHeader = Colors.LightBlue;
        private readonly Color colorNotes = Colors.Blue;
        private readonly Color colorCost = Colors.Silver;
        private readonly string format = "0.00";
        private int countRows;
        private Document document;

        public Quotation Quota { get; set; }
        public Client Client { get; set; }
        public Invoice Invoice { get; set; }
        public List<Material> Materials { get; set; }
        public MaterialProfit MaterialProfitSelect { get; set; }


        public void Print()
        {
            CreateDocument();

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();

            string company = (Quota?.CompanyName == "CMO") ? "CMO" : "NL";
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (Directory.Exists(folder + "\\Builders File") == false)
            {
                Directory.CreateDirectory(folder + "\\Builders File");
            }
            string filename = folder + @"\\Builders File\" + company + "_Material Profits" + "-" + Invoice.NumberInvoice +
                                                                                             "-" + Client.PrimaryFullName +
                                                                                             "-" + MaterialProfitSelect.InvoiceDate.Day.ToString("00") +
                                                                                             "." + MaterialProfitSelect.InvoiceDate.Month.ToString("00") +
                                                                                             "." + MaterialProfitSelect.InvoiceDate.Year.ToString("0000") + ".pdf";

            pdfRenderer.PdfDocument.Save(filename);// сохраняем            

            Process.Start(filename);
        }

        private Document CreateDocument()
        {
            // Create a new MigraDoc document
            this.document = new Document();
            this.document.Info.Title = "Material Profits";
            this.document.Info.Subject = "Create an material profits PDF.";
            this.document.Info.Author = "Builders";

            DefineStyles();
            CreatePage();

            return this.document;
        }
        private void DefineStyles()
        {
            // Get the predefined style Normal.
            Style style = this.document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Verdana";


            style = this.document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);


            style = this.document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);

            // Create a new style called Table based on style Normal
            style = this.document.Styles.AddStyle("Table", "Normal");
            style.Font.Name = "Verdana";
            style.Font.Name = "Times New Roman";
            style.Font.Size = 9;



            // Create a new style called Reference based on style Normal
            style = this.document.Styles.AddStyle("Reference", "Normal");
            style.ParagraphFormat.SpaceBefore = "5mm";
            style.ParagraphFormat.SpaceAfter = "5mm";
            style.ParagraphFormat.TabStops.AddTabStop("16cm", TabAlignment.Right);


        }
        private void CreatePage()
        {
            // Each MigraDoc document needs at least one section.
            Section section = this.document.AddSection();
            section.PageSetup.Orientation = Orientation.Landscape;
            section.PageSetup.LeftMargin = "1.0cm";
            section.PageSetup.RightMargin = "1.0cm";
            section.PageSetup.TopMargin = "0.5cm";
            section.PageSetup.BottomMargin = "0.5cm";
            section.PageSetup.DifferentFirstPageHeaderFooter = false;



            #region Document Name    

            Paragraph paragraph = section.AddParagraph();
            paragraph.AddFormattedText("MATERIAL PROFITS", TextFormat.Bold);
            paragraph.Format.Font.Size = 20;
            paragraph.Format.Font.Color = Colors.Red;
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddLineBreak();


            #endregion
            //*******************************************************************************************


            // Додаємо таблицю з інформацією про вибраного клієнта ***************************************
            #region Client info
            // Створюємо таблицю з інформацією про клієнта
            Table table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;
            // Створюємо стовбчики таблиці 
            Column column = table.AddColumn("4.0cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("10cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("4.0cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("10cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            // Додаємо рядок
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Description:");
            row.Cells[1].AddParagraph((Quota.JobDescription?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[1].MergeRight = 2;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Invoice:");
            row.Cells[1].AddParagraph(Invoice?.NumberInvoice ?? "").Format.Font.Bold = false;
            row.Cells[2].Shading.Color = colorHeader;
            row.Cells[2].AddParagraph("Date:");
            row.Cells[3].AddParagraph((MaterialProfitSelect.InvoiceDate.ToShortDateString()?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Company Name:");
            row.Cells[1].AddParagraph((Client.CompanyName?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[1].MergeRight = 2;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("1st Contact:");
            row.Cells[1].AddParagraph((Client.PrimaryFullName?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].Shading.Color = colorHeader;
            row.Cells[2].AddParagraph("2st Contact:");
            row.Cells[3].AddParagraph((Client.SecondaryFirstName?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Phone:");
            row.Cells[1].AddParagraph((Client.PrimaryPhoneNumber?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].Shading.Color = colorHeader;
            row.Cells[2].AddParagraph("Phone:");
            row.Cells[3].AddParagraph((Client.SecondaryPhoneNumber?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("E-mail:");
            row.Cells[1].AddParagraph((Client.PrimaryEmail?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].Shading.Color = colorHeader;
            row.Cells[2].AddParagraph("E-mail:");
            row.Cells[3].AddParagraph((Client.SecondaryEmail?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Address:");
            row.Cells[1].AddParagraph((Client.AddressBillFull?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].Shading.Color = colorHeader;
            row.Cells[2].AddParagraph("Address:");
            row.Cells[3].AddParagraph((Client.AddressSiteFull?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Обводим лініями
            table.SetEdge(0, 0, 4, 7, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            #endregion
            //********************************************************************************************


            #region Material Profits page 1

            paragraph = section.AddParagraph();

            table = section.AddTable();
            table.Style = "Table";

            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Shading.Color = Colors.LightGreen;
            row.Cells[0].AddParagraph("CUSTOMER PAID").Format.Font.Size = 10;
            row.Cells[0].MergeRight = 4;
            row.Cells[5].Format.Font.Bold = true;
            row.Cells[5].Shading.Color = Colors.Yellow;
            row.Cells[5].AddParagraph("OUR COST").Format.Font.Size = 10;
            row.Cells[5].MergeRight = 5;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Item").Format.Font.Bold = true;
            row.Cells[1].AddParagraph("Description").Format.Font.Bold = true;
            row.Cells[2].AddParagraph("Quantity").Format.Font.Bold = true;
            row.Cells[3].AddParagraph("Rate").Format.Font.Bold = true;
            row.Cells[4].AddParagraph("Price").Format.Font.Bold = true;
            row.Cells[5].AddParagraph("Quantity").Format.Font.Bold = true;
            row.Cells[6].AddParagraph("Unit Price").Format.Font.Bold = true;
            row.Cells[7].AddParagraph("Subtotal").Format.Font.Bold = true;
            row.Cells[8].AddParagraph("EP Rate").Format.Font.Bold = true;
            row.Cells[9].AddParagraph("TAX").Format.Font.Bold = true;
            row.Cells[10].AddParagraph("Total").Format.Font.Bold = true;

            countRows = 2;
            foreach (var item in Materials)
            {
                if (item.Groupe == "FLOORING")
                {
                    countRows++;
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = false;
                    row.Cells[0].AddParagraph(item.Groupe + " - " + item.Item ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].AddParagraph(item?.Description ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[2].AddParagraph(item.Quantity.ToString(format));
                    row.Cells[3].AddParagraph(item.Rate.ToString(format));
                    row.Cells[4].AddParagraph(item.Price.ToString(format));
                    row.Cells[5].Shading.Color = colorCost;
                    row.Cells[5].AddParagraph(item.CostQuantity.ToString(format));
                    row.Cells[6].Shading.Color = colorCost;
                    row.Cells[6].AddParagraph(item.CostUnitPrice.ToString(format));
                    row.Cells[7].Shading.Color = colorCost;
                    row.Cells[7].AddParagraph(item.CostSubtotal.ToString(format));
                    row.Cells[8].Shading.Color = colorCost;
                    row.Cells[8].AddParagraph(item.CostEPRate.ToString(format));
                    row.Cells[9].Shading.Color = colorCost;
                    row.Cells[9].AddParagraph(item.CostTax.ToString(format));
                    row.Cells[10].Shading.Color = colorCost;
                    row.Cells[10].AddParagraph(item.CostTotal.ToString(format));
                }
                if (item.Groupe == "ACCESSORIES")
                {
                    countRows++;
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = false;
                    row.Cells[0].AddParagraph(item.Groupe + " - " + item.Item ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].AddParagraph(item?.Description ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[2].AddParagraph(item.Quantity.ToString(format));
                    row.Cells[3].AddParagraph(item.Rate.ToString(format));
                    row.Cells[4].AddParagraph(item.Price.ToString(format));
                    row.Cells[5].Shading.Color = colorCost;
                    row.Cells[5].AddParagraph(item.CostQuantity.ToString(format));
                    row.Cells[6].Shading.Color = colorCost;
                    row.Cells[6].AddParagraph(item.CostUnitPrice.ToString(format));
                    row.Cells[7].Shading.Color = colorCost;
                    row.Cells[7].AddParagraph(item.CostSubtotal.ToString(format));
                    row.Cells[8].Shading.Color = colorCost;
                    row.Cells[8].AddParagraph(item.CostEPRate.ToString(format));
                    row.Cells[9].Shading.Color = colorCost;
                    row.Cells[9].AddParagraph(item.CostTax.ToString(format));
                    row.Cells[10].Shading.Color = colorCost;
                    row.Cells[10].AddParagraph(item.CostTotal.ToString(format));
                }
            }

            countRows++;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("MATERIAL SUBTOTAL").Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 7;
            row.Cells[8].AddParagraph("$ " + MaterialProfitSelect.MaterialSubtotal.ToString(format));
            row.Cells[8].MergeRight = 2;

            countRows++;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("TAX¹").Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 7;
            row.Cells[8].AddParagraph("$ " + MaterialProfitSelect.MaterialTax.ToString(format));
            row.Cells[8].MergeRight = 2;

            countRows++;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("MATERIAL TOTAL").Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 7;
            row.Cells[8].AddParagraph("$ " + MaterialProfitSelect.MaterialTotal.ToString(format));
            row.Cells[8].MergeRight = 2;

            table.SetEdge(0, 0, 11, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            #endregion
            //********************************************************************************************

            #region Material Profits page 2

            section.AddPageBreak();           

            table = section.AddTable();
            table.Style = "Table";

            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            column = table.AddColumn("8cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("8cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;


            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].Shading.Color = colorCost;
            row.Cells[0].AddParagraph("OTHER INCURRED EXPENSES").Format.Font.Size = 10;
            row.Cells[0].MergeRight = 7;            
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Item").Format.Font.Bold = true;
            row.Cells[1].AddParagraph("Description").Format.Font.Bold = true;
            row.Cells[2].AddParagraph("Quantity").Format.Font.Bold = true;
            row.Cells[3].AddParagraph("Unit Price").Format.Font.Bold = true;
            row.Cells[4].AddParagraph("Subtotal").Format.Font.Bold = true;            
            row.Cells[5].AddParagraph("EP Rate").Format.Font.Bold = true;
            row.Cells[6].AddParagraph("TAX").Format.Font.Bold = true;
            row.Cells[7].AddParagraph("Total").Format.Font.Bold = true;

            countRows = 2;
            foreach (var item in Materials)
            {
                if (item.Groupe == "OTHER")
                {
                    countRows++;
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = false;
                    row.Cells[0].AddParagraph(item.Groupe + " - " + item.Item ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].AddParagraph(item?.Description ?? "").Format.Alignment = ParagraphAlignment.Left;                    
                    row.Cells[2].AddParagraph(item.CostQuantity.ToString(format));
                    row.Cells[3].AddParagraph(item.CostUnitPrice.ToString(format));
                    row.Cells[4].AddParagraph(item.CostSubtotal.ToString(format));
                    row.Cells[5].AddParagraph(item.CostEPRate.ToString(format));
                    row.Cells[6].AddParagraph(item.CostTax.ToString(format));
                    row.Cells[7].AddParagraph(item.CostTotal.ToString(format));
                }
                
            }

            table.SetEdge(0, 0, 8, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            paragraph = section.AddParagraph();
            table = section.AddTable();
            table.Style = "Table";

            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            column = table.AddColumn("8cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");

            countRows = 1;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("MATERIAL COST SUBTOTAL").Format.Alignment = ParagraphAlignment.Right;            
            row.Cells[1].AddParagraph("$ " + MaterialProfitSelect.CostMaterialSubtotal.ToString(format));
           

            countRows++;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("TAX¹").Format.Alignment = ParagraphAlignment.Right;            
            row.Cells[1].AddParagraph("$ " + MaterialProfitSelect.CostMaterialTax.ToString(format));
            

            countRows++;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("MATERIAL COST TOTAL").Format.Alignment = ParagraphAlignment.Right;            
            row.Cells[1].AddParagraph("$ " + MaterialProfitSelect.CostMaterialTotal.ToString(format));            

            table.SetEdge(0, 0, 2, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            paragraph = section.AddParagraph();

            table = section.AddTable();
            table.Style = "Table";

            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            column = table.AddColumn("8cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");

            // Додаємо рядок
            countRows = 1;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].Shading.Color = Colors.LightGreen;
            row.Cells[0].AddParagraph("MATERIAL PROFIT BEFORE TAX");
            row.Cells[1].AddParagraph("$ " + MaterialProfitSelect.ProfitBeforTax.ToString(format));
            // Додаємо рядок
            countRows++;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].Shading.Color = Colors.LightGreen;
            row.Cells[0].AddParagraph("MATERIAL TAX PROFIT");
            row.Cells[1].AddParagraph("$ " + MaterialProfitSelect.ProfitTax.ToString(format));
            // Додаємо рядок
            countRows++;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].Shading.Color = Colors.LightGreen;
            row.Cells[0].AddParagraph("MATERIAL PROFIT INCL. TAX");
            row.Cells[1].AddParagraph("$ " + MaterialProfitSelect.ProfitInclTax.ToString(format));
            // Додаємо рядок
            countRows++;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].Shading.Color = Colors.Silver;
            row.Cells[0].AddParagraph("DISCOUNT DEDUCTIONS");
            row.Cells[1].Shading.Color = Colors.Silver;
            row.Cells[1].AddParagraph("$ " + MaterialProfitSelect.ProfitDiscount.ToString(format));
            // Додаємо рядок
            countRows++;
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].Shading.Color = Colors.LightGreen;
            row.Cells[0].AddParagraph("TOTAL MATERIAL PROFITS");
            row.Cells[1].Shading.Color = Colors.Yellow;
            row.Cells[1].AddParagraph("$ " + MaterialProfitSelect.ProfitTotal.ToString(format));

            table.SetEdge(0, 0, 2, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);

            #endregion

        }
    }
}
