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
    public class PrintLabourProfit
    {
        public PrintLabourProfit()
        {
            Labours = new List<Labour>();
            Contractors = new List<LabourContractor>();

            if (Quota?.CompanyName == "CMO")
            {
                company = "CMO";
            }
            else if (Quota?.CompanyName == "Leveling")
            {
                company = "NL";
            }
            else
            {
                company = "";
            }
        }

        private readonly Color colorHeader = Colors.LightBlue;
        private readonly Color colorNotes = Colors.Blue;
        private readonly Color colorCost = Colors.Silver;
        private readonly string format = "0.00";
        private int countRows;
        private Document document;
        private string company;

        public Quotation Quota { get; set; }
        public Client Client { get; set; }
        public Invoice Invoice { get; set; }
        public List<Labour> Labours { get; set; }
        public List<LabourContractor> Contractors { get; set; }
        public LabourProfit LabourProfitSelect { get; set; }



        public void Print()
        {
            CreateDocument();

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();

            
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (Directory.Exists(folder + "\\Builders File") == false)
            {
                Directory.CreateDirectory(folder + "\\Builders File");
            }
            string filename = folder + @"\\Builders File\" + company + "_Labour Profits" + "-" + Invoice.NumberInvoice +
                                                                                           "-" + Client.PrimaryFullName +
                                                                                           "-" + LabourProfitSelect.InvoiceDate.Day.ToString("00") +
                                                                                           "." + LabourProfitSelect.InvoiceDate.Month.ToString("00") +
                                                                                           "." + LabourProfitSelect.InvoiceDate.Year.ToString("0000") + ".pdf";

            pdfRenderer.PdfDocument.Save(filename);// сохраняем            

            Process.Start(filename);
        }

        private Document CreateDocument()
        {
            // Create a new MigraDoc document
            this.document = new Document();
            this.document.Info.Title = "Labour Profits";
            this.document.Info.Subject = "Create an labour profits PDF.";
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
            paragraph.AddFormattedText("LABOUR PROFITS", TextFormat.Bold);
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
            row.Cells[3].AddParagraph((LabourProfitSelect.InvoiceDate.ToShortDateString()?.ToUpper()) ?? "").Format.Font.Bold = false;
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
            column = table.AddColumn("6cm");
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
                        
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Item").Format.Font.Bold = true;
            row.Cells[1].AddParagraph("Description").Format.Font.Bold = true;
            row.Cells[2].AddParagraph("Contractor").Format.Font.Bold = true;
            row.Cells[3].AddParagraph("Quantity").Format.Font.Bold = true;
            row.Cells[4].AddParagraph("Rate").Format.Font.Bold = true;
            row.Cells[5].AddParagraph("Price").Format.Font.Bold = true;
            row.Cells[6].Shading.Color = Colors.Yellow;
            row.Cells[6].AddParagraph("%").Format.Font.Bold = true;
            row.Cells[7].Shading.Color = Colors.Yellow;
            row.Cells[7].AddParagraph("Payout").Format.Font.Bold = true;
            row.Cells[8].Shading.Color = Colors.Yellow;
            row.Cells[8].AddParagraph("Profit").Format.Font.Bold = true;
            

            countRows = 1;
            foreach (var item in Labours)
            {
                if (item.Groupe == "INSTALLATION")
                {
                    countRows++;
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = false;
                    row.Cells[0].AddParagraph(item.Groupe + " - " + item.Item ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].AddParagraph(item?.Description ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[2].AddParagraph(item?.Contractor ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[3].AddParagraph(item.Quantity.ToString(format));
                    row.Cells[4].AddParagraph("$ " + item.Rate.ToString(format));
                    row.Cells[5].AddParagraph("$ " + item.Price.ToString(format));
                    row.Cells[6].Shading.Color = colorCost;
                    row.Cells[6].AddParagraph(item.Percent.ToString(format));
                    row.Cells[7].Shading.Color = colorCost;
                    row.Cells[7].AddParagraph("$ " + item.Payout.ToString(format));
                    row.Cells[8].Shading.Color = colorCost;
                    row.Cells[8].AddParagraph("$ " + item.Profit.ToString(format));                                        
                }
                if (item.Groupe == "DEMOLITION")
                {
                    countRows++;
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = false;
                    row.Cells[0].AddParagraph(item.Groupe + " - " + item.Item ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].AddParagraph(item?.Description ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[2].AddParagraph(item?.Contractor ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[3].AddParagraph(item.Quantity.ToString(format));
                    row.Cells[4].AddParagraph("$ " + item.Rate.ToString(format));
                    row.Cells[5].AddParagraph("$ " + item.Price.ToString(format));
                    row.Cells[6].Shading.Color = colorCost;
                    row.Cells[6].AddParagraph(item.Percent.ToString(format));
                    row.Cells[7].Shading.Color = colorCost;
                    row.Cells[7].AddParagraph("$ " + item.Payout.ToString(format));
                    row.Cells[8].Shading.Color = colorCost;
                    row.Cells[8].AddParagraph("$ " + item.Profit.ToString(format));
                }
                if (item.Groupe == "OPTIONAL SERVICES")
                {
                    countRows++;
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = false;
                    row.Cells[0].AddParagraph(item.Groupe + " - " + item.Item ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].AddParagraph(item?.Description ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[2].AddParagraph(item?.Contractor ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[3].AddParagraph(item.Quantity.ToString(format));
                    row.Cells[4].AddParagraph("$ " + item.Rate.ToString(format));
                    row.Cells[5].AddParagraph("$ " + item.Price.ToString(format));
                    row.Cells[6].Shading.Color = colorCost;
                    row.Cells[6].AddParagraph(item.Percent.ToString(format));
                    row.Cells[7].Shading.Color = colorCost;
                    row.Cells[7].AddParagraph("$ " + item.Payout.ToString(format));
                    row.Cells[8].Shading.Color = colorCost;
                    row.Cells[8].AddParagraph("$ " + item.Profit.ToString(format));
                }
                if (item.Groupe == "FLOORING DELIVERY")
                {
                    countRows++;
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = false;
                    row.Cells[0].AddParagraph(item.Groupe + " - " + item.Item ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].AddParagraph(item?.Description ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[2].AddParagraph(item?.Contractor ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[3].AddParagraph(item.Quantity.ToString(format));
                    row.Cells[4].AddParagraph("$ " + item.Rate.ToString(format));
                    row.Cells[5].AddParagraph("$ " + item.Price.ToString(format));
                    row.Cells[6].Shading.Color = colorCost;
                    row.Cells[6].AddParagraph(item.Percent.ToString(format));
                    row.Cells[7].Shading.Color = colorCost;
                    row.Cells[7].AddParagraph("$ " + item.Payout.ToString(format));
                    row.Cells[8].Shading.Color = colorCost;
                    row.Cells[8].AddParagraph("$ " + item.Profit.ToString(format));
                }
            }

            

            table.SetEdge(0, 0, 9, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


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
            
            column = table.AddColumn("7cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("7cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("7cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("7cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;            
            row.Cells[0].Shading.Color = Colors.Yellow;
            row.Cells[0].AddParagraph("LABOUR PROFIT SUMMARY").Format.Font.Size = 10;
            row.Cells[0].MergeRight = 3;            
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = Colors.LightGreen;
            row.Cells[0].AddParagraph("");
            row.Cells[1].AddParagraph("COLLECTED");
            row.Cells[2].AddParagraph("PAYOUT");
            row.Cells[3].AddParagraph("STORE PROFIT");
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;            
            row.Cells[0].AddParagraph("LABOUR TOTAL").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].AddParagraph("$ " + LabourProfitSelect.CollectedSubtotal.ToString(format));
            row.Cells[2].AddParagraph("$ " + LabourProfitSelect.PayoutSubtotal.ToString(format));
            row.Cells[3].AddParagraph("$ " + LabourProfitSelect.StoreSubtotal.ToString(format));
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("GST").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].AddParagraph("$ " + LabourProfitSelect.CollectedGST.ToString(format));
            row.Cells[2].AddParagraph("$ " + LabourProfitSelect.PayoutGST.ToString(format));
            row.Cells[3].AddParagraph("$ " + LabourProfitSelect.StoreGST.ToString(format));
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("LABOUR GRAND TOTAL").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[1].AddParagraph("$ " + LabourProfitSelect.CollectedTotal.ToString(format));
            row.Cells[2].AddParagraph("$ " + LabourProfitSelect.PayoutTotal.ToString(format));
            row.Cells[3].AddParagraph("$ " + LabourProfitSelect.StoreTotal.ToString(format));
            // Додаємо рядок
            row = table.AddRow();
            row.Cells[0].AddParagraph("");
            row.Cells[0].MergeRight = 3;
            row.Cells[0].Shading.Color = Colors.Silver;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("DISCOUNTS DEDUCTIONS").Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 2;
            row.Cells[3].AddParagraph("$ " + LabourProfitSelect.Discount.ToString(format));
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("TOTAL LABOUR PROFIT").Format.Alignment = ParagraphAlignment.Right;
            row.Cells[0].MergeRight = 2;
            row.Cells[3].Shading.Color = Colors.Yellow;
            row.Cells[3].AddParagraph("$ " + LabourProfitSelect.ProfitTotal.ToString(format));

            table.SetEdge(0, 0, 4, 8, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);




            paragraph = section.AddParagraph();
            table = section.AddTable();
            table.Style = "Table";

            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            column = table.AddColumn("10cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = Colors.LightGreen;
            row.Cells[0].AddParagraph("Contractor").Format.Font.Bold = true;
            row.Cells[1].AddParagraph("Payout").Format.Font.Bold = true;
            row.Cells[2].AddParagraph("Adjust").Format.Font.Bold = true;
            row.Cells[3].AddParagraph("Total").Format.Font.Bold = true;
            row.Cells[4].AddParagraph("TAX").Format.Font.Bold = true;
            row.Cells[5].AddParagraph("GST").Format.Font.Bold = true;
            row.Cells[6].Shading.Color = Colors.Yellow;
            row.Cells[6].AddParagraph("TOTAL").Format.Font.Bold = true;
            


            countRows = 1;
            foreach (var item in Contractors)
            {               
                countRows++;
                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Center;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = false;
                row.Cells[0].AddParagraph(item.Contractor ?? "").Format.Alignment = ParagraphAlignment.Left;
                row.Cells[1].AddParagraph("$ " + item.Payout.ToString(format));
                row.Cells[2].AddParagraph("$ " + item.Adjust.ToString(format));
                row.Cells[3].AddParagraph("$ " + item.Total.ToString(format));
                row.Cells[4].AddParagraph(item.TAX ?? "");
                row.Cells[5].AddParagraph("$ " + item.GST.ToString(format));
                row.Cells[6].AddParagraph("$ " + item.TotalContractor.ToString(format));
               
                

            }

            table.SetEdge(0, 0, 7, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            

            #endregion

        }
    }
}
