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

namespace Builders.PDF
{
    public class PrintDeliveryDriver
    {
        public PrintDeliveryDriver()
        {
            Deliveries = new List<Delivery>();
            DeliveryMaterials = new List<DeliveryMaterial>();
            Suppliers = new List<DIC_Supplier>();
        }

        private Quotation quota;
        private Client client;
        private List<Delivery> deliveries;
        private List<DeliveryMaterial> deliveryMaterials;
        private List<DIC_Supplier> suppliers;
        private readonly string format = "0.00";       
        private Document document;
        
        public Quotation Quota
        {
            get { return quota; }
            set
            {
                quota = value;
            }
        }
        public Client Client
        {
            get { return client; }
            set
            {
                client = value;
            }
        }
        public List<Delivery> Deliveries
        {
            get { return deliveries; }
            set
            {
                deliveries = value;
            }
        }
        public List<DeliveryMaterial> DeliveryMaterials
        {
            get { return deliveryMaterials; }
            set
            {
                deliveryMaterials = value;
            }
        }
        public List<DIC_Supplier> Suppliers
        {
            get { return suppliers; }
            set
            {
                suppliers = value;
            }
        }

        public void Print()
        {
            CreateDocument();

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            DateTime date = new DateTime(2020, 12, 15);
            string filename = @"Output Files\Delivery" + "-" + Quota.NumberQuota +
                                                         "-" + DateTime.Today.Day.ToString("00") +
                                                         "." + DateTime.Today.Month.ToString("00") +
                                                         "." + DateTime.Today.Year.ToString("0000") + ".pdf";
            
            pdfRenderer.PdfDocument.Save(filename);// сохраняем 

            
            Process.Start(filename);
        }

        private Document CreateDocument()
        {
            // Create a new MigraDoc document
            this.document = new Document();
            this.document.Info.Title = "Delivery";
            this.document.Info.Subject = "Create an delivery PDF.";
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
            section.PageSetup.LeftMargin = "1cm";
            section.PageSetup.RightMargin = "1cm";
            section.PageSetup.TopMargin = "3.5cm";
            section.PageSetup.DifferentFirstPageHeaderFooter = true;


            // Додаємо логотип в заголовок + інформацію про власника та розміщуємо під логотипом *******
            #region Heder logo
            TextFrame frame = section.Headers.FirstPage.AddTextFrame();
            frame.Width = "7.0cm";
            frame.Left = ShapePosition.Left;
            frame.RelativeHorizontal = RelativeHorizontal.Margin;
            frame.Top = "1.0cm";
            frame.RelativeVertical = RelativeVertical.Page;

            Image image = frame.AddImage(@"User\Logo.jpg");
            image.Height = "2.5cm";
            image.LockAspectRatio = true;
            image.RelativeVertical = RelativeVertical.Line;
            image.RelativeHorizontal = RelativeHorizontal.Margin;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Center;
            image.WrapFormat.Style = WrapStyle.TopBottom;

            Paragraph paragraph = frame.AddParagraph("Delivery time:   3pm - 5pm");
            //paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("*expected delivery time to the client's house", TextFormat.Italic).Font.Size = 7;
            //paragraph.Format.Font.Size = 7;
            //paragraph.Format.SpaceAfter = 3;
            #endregion
            //*******************************************************************************************

            // Додаємо назву документу та заповнюємо дату, номер та номер квоти *************************
            #region Heder data

            //frame = section.AddTextFrame();
            //frame.Left = ShapePosition.Right;            
            //frame.Top = "1.0cm";

            frame = section.Headers.FirstPage.AddTextFrame();
            frame.Left = ShapePosition.Right;
            frame.Width = "6.0cm";

            paragraph = frame.AddParagraph();
            paragraph.AddFormattedText("Instructions for Drivers", TextFormat.Bold);
            paragraph.Format.Font.Size = 20;
            paragraph.Format.Font.Color = Colors.Red;
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();

            paragraph = frame.AddParagraph();
            paragraph.Format.Font.Size = 10;
            paragraph.AddText("Date: ");
            paragraph.AddFormattedText(DateTime.Today.ToShortDateString(), TextFormat.Bold);
            paragraph.AddLineBreak();            
            paragraph.AddText("Estimate: ");
            paragraph.AddFormattedText(Quota.NumberQuota, TextFormat.Bold);
            #endregion
            //*******************************************************************************************
            // Додаємо логотип в заголовок + інформацію про власника та розміщуємо під логотипом *******
            #region Heder logo page 2            
            frame = section.Headers.Primary.AddTextFrame();
            frame.Width = "7.0cm";
            frame.Left = ShapePosition.Left;
            frame.RelativeHorizontal = RelativeHorizontal.Margin;
            frame.Top = "1.0cm";
            frame.RelativeVertical = RelativeVertical.Page;

            image = frame.AddImage(@"User\Logo.jpg");
            image.Height = "2.5cm";
            image.LockAspectRatio = true;
            image.RelativeVertical = RelativeVertical.Line;
            image.RelativeHorizontal = RelativeHorizontal.Margin;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Center;
            image.WrapFormat.Style = WrapStyle.TopBottom;


            #endregion
            //*******************************************************************************************
            
            // Додаємо таблицю з інформацією про вибраного клієнта ***************************************
            #region Client info
            // Тут вставляю пустий параграф для контролю відступу від заголовку
            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "2cm";
            //paragraph.Style = "Quota";


            // Створюємо таблицю з інформацією про клієнта
            Table table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;
            // Створюємо стовбчики таблиці 
            Column column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("6.5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("6.5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            // Додаємо рядок
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].AddParagraph("Client Name:");
            row.Cells[1].AddParagraph((Client.PrimaryFullName?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[1].MergeRight = 2;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
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
            row.Cells[0].AddParagraph("Billing Address:");
            row.Cells[1].AddParagraph((Client.AddressBillFull?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].AddParagraph("Site Address:");
            row.Cells[3].AddParagraph((Client.AddressSiteFull?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].AddParagraph("Phone:");
            row.Cells[1].AddParagraph((Client.PrimaryPhoneNumber?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].AddParagraph("E-mail:");
            row.Cells[3].AddParagraph((Client.PrimaryEmail?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Обводим лініями
            table.SetEdge(0, 0, 4, 4, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
                        
            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "1cm";

            paragraph.AddFormattedText("1. Rent Van or Cargo using Modo (depending on the size and weight of the delivery choose appropiate size of the vehicle)" + Environment.NewLine +
                                        "2.Make sure that you booked the car in advance and have enough time to return it by the end of the booking" + Environment.NewLine +
                                        "3.Modo is a car - sharing service, please always keep the vehicle clean for the next member, empty the vehicle" + 
                                        " of any garbage and personal items" + Environment.NewLine +
                                        "4.Your first destination is always the closest warehouse ti your current location(location where you getting a Modo car)." + 
                                        "You can change your route depending on where you start your trip, it doesn't matter what material you pick up first,"  +
                                        " but important to arrive to the warehouse during their Open Hours." + Environment.NewLine +
                                        "5.You can have a few locations to pick up the materials(for example surrey for the laminate, Richmond for the underlay).", TextFormat.Italic);
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("Important!!!", TextFormat.Bold).Font.Size = 15;
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("The client's adress is always the final destination. We expect you to deliever all the materials to their homes.", TextFormat.Italic);
            #endregion
            //********************************************************************************************

            // Створюєм таблицю інформацією по чеку*******************************************************
            #region Delivery for Drivers info
            paragraph = section.AddParagraph();

            table = section.AddTable();
            table.Style = "Table";

            //frame = section.AddTextFrame();
            //frame.Left = ShapePosition.Left;            
            //frame.Top = "0.5cm";

            //table = frame.AddTable();
            //table.Style = "Table";
            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            column = table.AddColumn("6.5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("6.5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
                       

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("Supplier").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[1].AddParagraph("Address").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].AddParagraph("Hours").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[3].AddParagraph("Quantity").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[4].AddParagraph("Order number").Format.Alignment = ParagraphAlignment.Center;

            int countRows = 1;
            foreach (var item in Deliveries)
            {
                var sup = Suppliers.FirstOrDefault(s => s.Id == item.SupplierId);

                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Left;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = true;
                row.Cells[0].AddParagraph((sup.Supplier) ?? "");
                row.Cells[1].AddParagraph((sup.Address) ?? "");
                row.Cells[2].AddParagraph((sup.Hours) ?? "").Format.Alignment = ParagraphAlignment.Center;
                row.Cells[3].AddParagraph("-").Format.Alignment = ParagraphAlignment.Center;
                row.Cells[4].AddParagraph((item.OrderNumber) ?? "").Format.Alignment = ParagraphAlignment.Center;

                var materialAll = DeliveryMaterials.Where(d => d.DeliveryId == item.Id);
                var materialName = materialAll.Select(m => m.Description).Distinct();
                int countMaterial = 1;
                countRows++;
                foreach (var temp in materialName)
                {
                    var material = materialAll.FirstOrDefault(m => m.Description == temp);
                    var materialQuantity = materialAll.Where(m => m.Description == temp).Select(s => s.Quantity).Sum();
                    row = table.AddRow();
                    row.HeadingFormat = false;
                    row.Format.Alignment = ParagraphAlignment.Left;
                    row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                    row.Format.Font.Bold = true;
                    row.Cells[0].AddParagraph(countMaterial.ToString() + ". " + (material.Description) ?? "");
                    row.Cells[0].MergeRight = 2;
                    row.Cells[3].AddParagraph((materialQuantity.ToString(format)) ?? "");                    

                    countMaterial++;
                    countRows++;
                }
            }
            

            table.SetEdge(0, 0, 5, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Colors.Black);



            #endregion
            //********************************************************************************************

           
           

        }

    }
}
