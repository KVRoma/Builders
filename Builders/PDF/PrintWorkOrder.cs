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
    public class PrintWorkOrder
    {
        public PrintWorkOrder()
        {
            works = new List<WorkOrder_Work>();
            accessories = new List<WorkOrder_Accessories>();
            installations = new List<WorkOrder_Installation>();
            contractors = new List<WorkOrder_Contractor>();
        }



        private readonly Color colorHeader = Colors.LightBlue;
        private readonly Color colorNotes = Colors.Blue;
        private readonly string format = "0.00";
        private int countRows;
        private Document document;

        public Quotation Quota { get; set; }        
        public Client Client { get; set; }
        public WorkOrder WorkOrder { get; set; }
        public List<WorkOrder_Work> works { get; set; }
        public List<WorkOrder_Accessories> accessories  { get; set; }
        public List<WorkOrder_Installation> installations { get; set; }
        public List<WorkOrder_Contractor> contractors { get; set; }


        public void Print()
        {
            CreateDocument();

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            DateTime date = new DateTime(2020, 12, 15);

            string company = (Quota?.CompanyName == "CMO") ? "CMO" : "NL";
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (Directory.Exists(folder + "\\Builders File") == false)
            {
                Directory.CreateDirectory(folder + "\\Builders File");
            }
            string filename = folder + @"\\Builders File\" + company + "_Work Order" + "-" + Quota.NumberQuota +
                                                                                       "-" + Client.PrimaryFullName + 
                                                                                       "-" + WorkOrder.DateServices.Day.ToString("00") +
                                                                                       "." + WorkOrder.DateServices.Month.ToString("00") +
                                                                                       "." + WorkOrder.DateServices.Year.ToString("0000") + 
                                                                                       "-" + contractors.FirstOrDefault().Contractor + ".pdf";
            
            pdfRenderer.PdfDocument.Save(filename);// сохраняем            

            Process.Start(filename);
        }

        private Document CreateDocument()
        {
            // Create a new MigraDoc document
            this.document = new Document();
            this.document.Info.Title = "Work Order";
            this.document.Info.Subject = "Create an work order PDF.";
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
            paragraph.AddFormattedText("WORK ORDER", TextFormat.Bold);
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
            row.Cells[0].AddParagraph("Estimate / Client:");
            row.Cells[1].AddParagraph(Quota.NumberQuota + " / " + Client.NumberClient).Format.Font.Bold = false;
            row.Cells[2].Shading.Color = colorHeader;
            row.Cells[2].AddParagraph("Parking:");
            row.Cells[3].AddParagraph((WorkOrder.Parking?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Service Date:");
            row.Cells[1].AddParagraph((WorkOrder.DateServices.ToShortDateString()?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].Shading.Color = colorHeader;
            row.Cells[2].AddParagraph("Completion Date:");
            row.Cells[3].AddParagraph((WorkOrder.DateCompletion.ToShortDateString()?.ToUpper()) ?? "").Format.Font.Bold = false;
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
            table.SetEdge(0, 0, 4, 8, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            // Створюєм таблицю з одним рядком NOTES
            paragraph = section.AddParagraph();

            table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            column = table.AddColumn("4.0cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("24.0cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("NOTES:");
            row.Cells[1].AddParagraph((WorkOrder.Notes) ?? "").Format.Font.Bold = false;

            table.SetEdge(0, 0, 2, 1, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            #endregion
            //********************************************************************************************


            #region Work Order room details

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
            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("11cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Room Details").Format.Font.Size = 13;            
            row.Cells[0].MergeRight = 4;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = colorHeader;            
            row.Cells[0].AddParagraph("Area").Format.Font.Bold = true;
            row.Cells[1].AddParagraph("Room").Format.Font.Bold = true;
            row.Cells[2].AddParagraph("Existing Floor").Format.Font.Bold = true;
            row.Cells[3].AddParagraph("New Floor").Format.Font.Bold = true;
            row.Cells[4].AddParagraph("Furniture Moving").Format.Font.Bold = true;
            
            countRows = 2;
            foreach (var item in works)
            {
                countRows++;
                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Center;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = false;                
                row.Cells[0].AddParagraph(item?.Area ?? "");
                row.Cells[1].AddParagraph(item?.Room ?? "");
                row.Cells[2].AddParagraph(item?.Existing ?? "");
                row.Cells[3].AddParagraph(item?.NewFloor ?? "");
                row.Cells[4].AddParagraph(item?.Furniture ?? "");

                countRows++;
                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Left;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = false;
                row.Cells[0].Shading.Color = colorHeader;
                row.Cells[0].AddParagraph("Caution / Notes: ");
                row.Cells[1].AddParagraph(item?.Misc ?? "").Format.Font.Bold = true;
                row.Cells[1].Format.Font.Color = colorNotes;
                row.Cells[1].Format.Font.Italic = true;
                row.Cells[1].MergeRight = 3;


            }

            table.SetEdge(0, 0, 5, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            #endregion
            //********************************************************************************************

            #region Work Order accessories details

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
            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("13cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            

            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Accessories Details").Format.Font.Size = 13;
            row.Cells[0].MergeRight = 3;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Area").Format.Font.Bold = true;
            row.Cells[1].AddParagraph("Room").Format.Font.Bold = true;
            row.Cells[2].AddParagraph("Old Accessories").Format.Font.Bold = true;
            row.Cells[3].AddParagraph("New Accessories").Format.Font.Bold = true;            

            countRows = 2;
            foreach (var item in accessories)
            {
                countRows++;
                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Center;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = false;
                row.Cells[0].AddParagraph(item?.Area ?? "");
                row.Cells[1].AddParagraph(item?.Room ?? "");
                row.Cells[2].AddParagraph(item?.OldAccessories ?? "");
                row.Cells[3].AddParagraph(item?.NewAccessories ?? "");               

                countRows++;
                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Left;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = false;
                row.Cells[0].Shading.Color = colorHeader;
                row.Cells[0].AddParagraph("Caution / Notes: ");
                row.Cells[1].AddParagraph(item?.Notes ?? "").Format.Font.Bold = true;
                row.Cells[1].Format.Font.Color = colorNotes;
                row.Cells[1].Format.Font.Italic = true;
                row.Cells[1].MergeRight = 2;


            }

            table.SetEdge(0, 0, 4, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            #endregion

            #region Work Order page 2
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
            column = table.AddColumn("9cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("4cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;


            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Earning Calculator").Format.Font.Size = 13;
            row.Cells[0].MergeRight = 5;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Installation").Format.Font.Bold = true;
            row.Cells[1].AddParagraph("Descriptions").Format.Font.Bold = true;
            row.Cells[2].AddParagraph("Contractor").Format.Font.Bold = true;
            row.Cells[3].AddParagraph("Quantity").Format.Font.Bold = true;
            row.Cells[4].AddParagraph("Rate or %").Format.Font.Bold = true;
            row.Cells[5].AddParagraph("Payout").Format.Font.Bold = true;

            countRows = 2;
            foreach (var item in installations)
            {
                if (item.Groupe == "DEMOLITION")
                {
                    countRows++;
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = false;
                    row.Cells[0].AddParagraph(item?.Groupe + 
                                             ((item?.RoomDetail == "  ") ? " - " : (" - " + item?.RoomDetail + " - ")) + 
                                              item?.Item).Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].AddParagraph(item?.Description ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[2].AddParagraph(item?.Contractor ?? "");
                    row.Cells[3].AddParagraph(item.Quantity.ToString(format));
                    row.Cells[4].AddParagraph(item.Procent.ToString(format));
                    row.Cells[5].AddParagraph("$ " + item.Payout.ToString(format));
                }
                if (item.Groupe == "INSTALLATION")
                {
                    countRows++;
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = false;
                    row.Cells[0].AddParagraph(item?.Groupe +
                                             ((item?.RoomDetail == "  ") ? " - " : (" - " + item?.RoomDetail + " - ")) +
                                              item?.Item).Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].AddParagraph(item?.Description ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[2].AddParagraph(item?.Contractor ?? "");
                    row.Cells[3].AddParagraph(item.Quantity.ToString(format));
                    row.Cells[4].AddParagraph(item.Procent.ToString(format));
                    row.Cells[5].AddParagraph("$ " + item.Payout.ToString(format));
                }
                if (item.Groupe == "OPTIONAL SERVICES")
                {
                    countRows++;
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Center;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = false;
                    row.Cells[0].AddParagraph(item?.Groupe +
                                             ((item?.RoomDetail == "  ") ? " - " : (" - " + item?.RoomDetail + " - ")) +
                                              item?.Item).Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[1].AddParagraph(item?.Description ?? "").Format.Alignment = ParagraphAlignment.Left;
                    row.Cells[2].AddParagraph(item?.Contractor ?? "");
                    row.Cells[3].AddParagraph(item.Quantity.ToString(format));
                    row.Cells[4].AddParagraph(item.Procent.ToString(format));
                    row.Cells[5].AddParagraph("$ " + item.Payout.ToString(format));
                }



            }

            table.SetEdge(0, 0, 6, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            #endregion

            #region Work Order page 2 Contractor
            section.AddParagraph();

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
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("4cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("4cm");
            column.Format.Alignment = ParagraphAlignment.Center;


            
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = colorHeader;
            row.Cells[0].AddParagraph("Contractor").Format.Font.Bold = true;
            row.Cells[1].AddParagraph("Payout").Format.Font.Bold = true;
            row.Cells[2].AddParagraph("Adjust").Format.Font.Bold = true;
            row.Cells[3].AddParagraph("Total").Format.Font.Bold = true;
            row.Cells[4].AddParagraph("TAX").Format.Font.Bold = true;
            row.Cells[5].AddParagraph("GST").Format.Font.Bold = true;
            row.Cells[6].AddParagraph("TOTAL").Format.Font.Bold = true;

            countRows = 1;
            foreach (var item in contractors)
            {                
                countRows++;
                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Center;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = false;
                row.Cells[0].AddParagraph(item?.Contractor ?? "").Format.Alignment = ParagraphAlignment.Left;
                row.Cells[1].AddParagraph("$ " + item.Payout.ToString(format));
                row.Cells[2].AddParagraph("$ " + item.Adjust.ToString(format));
                row.Cells[3].AddParagraph("$ " + item.Total.ToString(format));
                row.Cells[4].AddParagraph(item?.TAX ?? "");
                row.Cells[5].AddParagraph("$ " + item.GST.ToString(format));
                row.Cells[6].Shading.Color = Colors.Yellow;
                row.Cells[6].Format.Font.Bold = true;
                row.Cells[6].AddParagraph("$ " + item.TotalContractor.ToString(format)).Format.Font.Color = Colors.Red;
            }

            table.SetEdge(0, 0, 7, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            #endregion



        }



    }
}
