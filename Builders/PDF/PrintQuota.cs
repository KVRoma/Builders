using Builders.Models;
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

namespace Builders.PDF
{
    public class PrintQuota
    {
        public PrintQuota(User _user)
        {
            user = _user;
            MaterialQuotations = new List<MaterialQuotation>();
        }

        private Quotation quota;
        private Client client;
        private List<MaterialQuotation> materialQuotations;
        private IEnumerable<Payment> payments;
        private string nameQuota;
        private User user;

        private readonly string format = "0.00";                
        private string[] file = new string[4];
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
        public List<MaterialQuotation> MaterialQuotations
        {
            get { return materialQuotations; }
            set
            {
                materialQuotations = value;
            }
        }
        public IEnumerable<Payment> Payments
        {
            get { return payments; }
            set
            {
                payments = value;
            }
        }
        public string NameQuota
        {
            get { return nameQuota; }
            set
            {
                nameQuota = value;
            }
        }

        public void Print()
        {
            CreateDocument();

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);            
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            DateTime date = new DateTime(2020, 12, 15);
            string filename = @"Output Files\" + NameQuota + "-" + Client.NumberClient + 
                                                         "-" + Quota.NumberQuota + 
                                                         "-" + Quota.QuotaDate.Day.ToString("00") + 
                                                         "." + Quota.QuotaDate.Month.ToString("00") + 
                                                         "." + Quota.QuotaDate.Year.ToString("0000") + ".pdf";
            string filenameAgreement = @"User\Agreement" + user?.Id.ToString() ?? "" + ".pdf";
            pdfRenderer.PdfDocument.Save(filename);// сохраняем 

            PrintSignaturePage signature = new PrintSignaturePage();
            signature.Client = Client;
            signature.Quotation = Quota;
            signature.Date = Quota.QuotaDate;
            signature.User = user;
            signature.Print();

            PrintSignaturePageFirst first = new PrintSignaturePageFirst();
            first.Client = Client;
            first.Quotation = Quota;
            first.Date = Quota.QuotaDate;
            first.User = user;
            first.Print();

            file[0] = filename;
            file[1] = first.filename;
            file[2] = filenameAgreement;
            file[3] = signature.filename;
            MergeMultiplePDFIntoSinglePDF(filename, file);
            Process.Start(filename);
        }

        private Document CreateDocument()
        {
            // Create a new MigraDoc document
            this.document = new Document();
            this.document.Info.Title = "Estimate";
            this.document.Info.Subject = "Create an estimate PDF.";
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

            Image image = frame.AddImage(@"User\Logo" + (user?.Id.ToString() ?? "") + ".jpg");
            image.Height = "2.5cm";
            image.LockAspectRatio = true;
            image.RelativeVertical = RelativeVertical.Line;
            image.RelativeHorizontal = RelativeHorizontal.Margin;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Center;
            image.WrapFormat.Style = WrapStyle.TopBottom;

            
            Paragraph paragraph = frame.AddParagraph(user?.Name + ", " + Environment.NewLine +
                                          user?.Post + Environment.NewLine +
                                          user?.Address + Environment.NewLine +
                                          user?.Additional + Environment.NewLine +
                                          user?.Phone + Environment.NewLine + 
                                          user?.Mail + Environment.NewLine + 
                                          user?.WebSite);
            //paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.Format.Font.Size = 7;
            paragraph.Format.SpaceAfter = 3;
            #endregion
            //*******************************************************************************************

            // Додаємо назву документу та заповнюємо дату, номер та номер квоти *************************
            #region Heder data

            //frame = section.AddTextFrame();
            //frame.Left = ShapePosition.Right;            
            //frame.Top = "1.0cm";

            frame = section.Headers.FirstPage.AddTextFrame();
            frame.Left = ShapePosition.Right;
            frame.Width = "6.5cm";

            paragraph = frame.AddParagraph();
            paragraph.AddFormattedText(NameQuota, TextFormat.Bold);
            paragraph.Format.Font.Size = 15;
            paragraph.Format.Font.Color = Colors.Red;
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddLineBreak();            

            paragraph = frame.AddParagraph();
            paragraph.Format.Font.Size = 10;
            paragraph.AddText("Date: ");
            paragraph.AddFormattedText(Quota.QuotaDate.ToShortDateString(), TextFormat.Bold);
            paragraph.AddLineBreak();
            paragraph.AddText("Client Number: ");
            paragraph.AddFormattedText(Client.NumberClient, TextFormat.Bold);
            paragraph.AddLineBreak();
            paragraph.AddText("Estimate Number: ");
            paragraph.AddFormattedText(Quota.NumberQuota, TextFormat.Bold);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();

            paragraph = frame.AddParagraph();
            paragraph.Format.Font.Size = 7;
            if (!string.IsNullOrWhiteSpace(user?.GST))
            {
                paragraph.AddText("GST # ");
                paragraph.AddFormattedText(user?.GST + "", TextFormat.Bold);
                paragraph.AddLineBreak();                
            }
            if (!string.IsNullOrWhiteSpace(user?.PST))
            {
                paragraph.AddText("PST # ");
                paragraph.AddFormattedText(user?.PST + "", TextFormat.Bold);
                paragraph.AddLineBreak();                
            }
            if (!string.IsNullOrWhiteSpace(user?.WCB))
            {
                paragraph.AddText("WCB # ");
                paragraph.AddFormattedText(user?.WCB + "", TextFormat.Bold);
                paragraph.AddLineBreak();                
            }
            if (!string.IsNullOrWhiteSpace(user?.Liability))
            {
                paragraph.AddText("GENERAL LIABILITY # ");
                paragraph.AddFormattedText(user?.Liability + "", TextFormat.Bold);
                paragraph.AddLineBreak();                
            }
            if (!string.IsNullOrWhiteSpace(user?.Licence))
            {
                paragraph.AddText("BUSINESS LICENSE # ");
                paragraph.AddFormattedText(user?.Licence + "", TextFormat.Bold);
                paragraph.AddLineBreak();                
            }
            if (!string.IsNullOrWhiteSpace(user?.Incorporation))
            {
                paragraph.AddText("INCORPORATION # ");
                paragraph.AddFormattedText(user?.Incorporation + "", TextFormat.Bold);
                paragraph.AddLineBreak();                
            }
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

            image = frame.AddImage(@"User\Logo" + (user?.Id.ToString() ?? "") + ".jpg");
            image.Height = "2.5cm";
            image.LockAspectRatio = true;
            image.RelativeVertical = RelativeVertical.Line;
            image.RelativeHorizontal = RelativeHorizontal.Margin;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Center;
            image.WrapFormat.Style = WrapStyle.TopBottom;


            #endregion
            //*******************************************************************************************


            //// Create footer
            //paragraph = section.Footers.Primary.AddParagraph();
            //paragraph.AddText("PowerBooks Inc · Sample Street 42 · 56789 Cologne · Germany");
            //paragraph.Format.Font.Size = 9;
            //paragraph.Format.Alignment = ParagraphAlignment.Center;



            // Додаємо таблицю з інформацією про вибраного клієнта ***************************************
            #region Client info
            // Тут вставляю пустий параграф для контролю відступу від заголовку
            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "2.5cm";
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
            row.Cells[0].AddParagraph("1st Contact:");
            row.Cells[1].AddParagraph((Client.PrimaryFullName?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].AddParagraph("2st Contact:");
            row.Cells[3].AddParagraph((Client.SecondaryFirstName?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].AddParagraph("Phone:");
            row.Cells[1].AddParagraph((Client.PrimaryPhoneNumber?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].AddParagraph("Phone:");
            row.Cells[3].AddParagraph((Client.SecondaryPhoneNumber?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].AddParagraph("E-mail:");
            row.Cells[1].AddParagraph((Client.PrimaryEmail?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].AddParagraph("E-mail:");
            row.Cells[3].AddParagraph((Client.SecondaryEmail?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].AddParagraph("Address:");
            row.Cells[1].AddParagraph((Client.AddressBillFull?.ToUpper()) ?? "").Format.Font.Bold = false;
            row.Cells[2].AddParagraph("Address:");
            row.Cells[3].AddParagraph((Client.AddressSiteFull?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Обводим лініями
            table.SetEdge(0, 0, 4, 6, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            // Створюєм таблицю з одним рядком NOTES
            paragraph = section.AddParagraph();

            table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("17cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("NOTES:");
            row.Cells[1].AddParagraph((Quota.JobNote) ?? "").Format.Font.Bold = false;

            table.SetEdge(0, 0, 2, 1, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            #endregion
            //********************************************************************************************

            // Створюєм таблицю інформацією по чеку*******************************************************
            #region Quota Material info
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

            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("10cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].Format.Font.Size = 8;
            row.Cells[0].AddParagraph("Please carefully read the details, expectations and liabilitie").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[0].MergeRight = 1;
            row.Cells[2].Format.Font.Color = Colors.Red;
            row.Cells[2].AddParagraph("MATERIAL").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].MergeRight = 2;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("GROUP").Format.Alignment = ParagraphAlignment.Center; 
            row.Cells[1].AddParagraph("DESCRIPTION").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].AddParagraph("QUANTITY").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[3].AddParagraph("RATE").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[4].AddParagraph("PRICE").Format.Alignment = ParagraphAlignment.Center;

            var material = MaterialQuotations.Where(m => m.Groupe == "FLOORING" || m.Groupe == "ACCESSORIES");
            int countMaterialRow = 2;
            foreach (var item in material)
            {                
                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Right;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = true;
                row.Cells[0].AddParagraph((item.Groupe) ?? "").Format.Alignment = ParagraphAlignment.Left;
                if (item.Groupe == "FLOORING" && user?.Id == 2)
                {
                    if (!string.IsNullOrWhiteSpace(item.Depth))   // NL Quota - work user
                    {
                        row.Cells[1].AddParagraph(item.Description + " - Size: " + item.QuantityNL + " - Depth: " + item.Depth ).Format.Alignment = ParagraphAlignment.Left;
                        row.Cells[2].AddParagraph(item.Mapei?.ToString(format));
                    }
                    else
                    {
                        string[] words = item.Description.Split(' ');
                        int count = 0;
                        int size = 0;
                        string depth = "";
                        decimal qty = 0m;
                        string decript = "";
                        foreach (string word in words)
                        {
                            if (word == "Size:")
                            {
                                size = int.TryParse(words[count + 1], out int res) ? res : 0;
                            }
                            if (word == "Depth:")
                            {
                                depth = words[count + 1];
                            }
                            if (word == "QTY:")
                            {
                                qty = decimal.TryParse(words[count + 1], out decimal res1) ? res1 : 0m;
                                break;
                            }

                            decript += word + " ";
                            count++;
                        }
                        
                        row.Cells[1].AddParagraph(decript ?? "").Format.Alignment = ParagraphAlignment.Left;   // NL Quota - work Generator
                        row.Cells[2].AddParagraph(qty.ToString(format) ?? "");
                    }
                }
                else
                {
                    row.Cells[1].AddParagraph((item.Description) ?? "").Format.Alignment = ParagraphAlignment.Left;    // CMO Quota
                    row.Cells[2].AddParagraph(item.Quantity.ToString(format));
                }
                row.Cells[3].AddParagraph("$ " + item.Rate.ToString(format));
                row.Cells[4].AddParagraph("$ " + item.Price.ToString(format));
                countMaterialRow++;
            }

            row = table.AddRow();
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row[0].MergeRight = 4;
            countMaterialRow++;
            //for (int i = 60; i > material.Count() + 1; i--)
            //{
            //    row = table.AddRow();
            //    countMaterialRow++;
            //}

            table.SetEdge(0, 0, 5, countMaterialRow, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            row = table.AddRow();           
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[2].AddParagraph("Material Subtotal").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].MergeRight = 1;
            row.Cells[4].AddParagraph("$ " + Quota.MaterialSubtotal.ToString(format));

            row = table.AddRow();            
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;            
            row.Cells[2].AddParagraph("TAX¹").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].MergeRight = 1;            
            row.Cells[4].AddParagraph("$ " + Quota.MaterialTax.ToString(format));

            row = table.AddRow();            
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[2].MergeRight = 1;
            if (Quota.MaterialDiscountAmount != 0m)
            {
                row.Cells[2].AddParagraph("Discount").Format.Alignment = ParagraphAlignment.Left;
                
                row.Cells[4].AddParagraph("$ " + Quota.MaterialDiscountAmount.ToString(format));
            }

            row = table.AddRow();            
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[2].AddParagraph("Material TOTAL").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[2].MergeRight = 1;
            row.Cells[4].AddParagraph("$ " + Quota.MaterialTotal.ToString(format));

            table.SetEdge(2, countMaterialRow, 3, 4, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Colors.Black);



            #endregion
            //********************************************************************************************
            
            #region Quota Labour info


            //frame = section.AddTextFrame();
            //frame.Left = ShapePosition.Left;
            //frame.Top = "0.5cm";


            //table = frame.AddTable();
            section.AddPageBreak();
            table = section.AddTable();
            
            table.Style = "Table";
            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;            
            table.Rows.LeftIndent = 0;          

            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
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

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[3].Format.Font.Color = Colors.Red;
            row.Cells[3].AddParagraph("LABOUR").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[3].MergeRight = 2;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("GROUP").Format.Alignment = ParagraphAlignment.Center; 
            row.Cells[1].AddParagraph("ITEMS").Format.Alignment = ParagraphAlignment.Center; 
            row.Cells[2].AddParagraph("DESCRIPTION").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[3].AddParagraph("QUANTITY").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[4].AddParagraph("RATE").Format.Alignment = ParagraphAlignment.Center; 
            row.Cells[5].AddParagraph("PRICE").Format.Alignment = ParagraphAlignment.Center;

            var labour = MaterialQuotations.Where(m => m.Groupe != "FLOORING" && m.Groupe != "ACCESSORIES");
            int countLabourRow = 2;

            foreach (var item in labour)
            {
                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Right;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = true;
                row.Cells[0].AddParagraph((item.Groupe) ?? "").Format.Alignment = ParagraphAlignment.Left;                
                row.Cells[1].AddParagraph((item.Item) ?? "").Format.Alignment = ParagraphAlignment.Left;
                row.Cells[2].AddParagraph((item.Description) ?? "").Format.Alignment = ParagraphAlignment.Left; ;
                row.Cells[3].AddParagraph(item.Quantity.ToString(format));
                row.Cells[4].AddParagraph("$ " + item.Rate.ToString(format));
                row.Cells[5].AddParagraph("$ " + item.Price.ToString(format));
                countLabourRow++;
            }

            row = table.AddRow();
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row[0].MergeRight = 5;
            countLabourRow++;

            //for (int i = 40; i > labour.Count() + 1; i--)
            //{
            //    row = table.AddRow();
            //    countLabourRow++;
            //}

            table.SetEdge(0, 0, 6, countLabourRow, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);

            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[3].AddParagraph("Labour Subtotal").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].MergeRight = 1;
            row.Cells[5].AddParagraph("$ " + Quota.LabourSubtotal.ToString(format));

            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[3].AddParagraph("GST").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].MergeRight = 1;
            row.Cells[5].AddParagraph("$ " + Quota.LabourTax.ToString(format));

            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[3].MergeRight = 1;
            if (Quota.MaterialDiscountAmount != 0m)
            {
                row.Cells[3].AddParagraph("Discount").Format.Alignment = ParagraphAlignment.Left;

                row.Cells[5].AddParagraph("$ " + Quota.LabourDiscountAmount.ToString(format));
            }

            row = table.AddRow();
            row.HeadingFormat = false;
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[3].AddParagraph("Labour TOTAL").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].MergeRight = 1;
            row.Cells[5].AddParagraph("$ " + Quota.LabourTotal.ToString(format));

            table.SetEdge(3, countLabourRow, 3, 4, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Colors.Black);


            #endregion

            #region Quota Total info
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            //paragraph.Format.SpaceBefore = "2cm";

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
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            

            row = table.AddRow();            
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;                      
            row.Cells[0].AddParagraph("Do you want financing?").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[1].AddParagraph((Quota.FinancingYesNo) ? ("Yes") : ("No")).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].AddParagraph("PROJECT TOTAL").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].AddParagraph("$ " + Quota.ProjectTotal.ToString(format));

            row = table.AddRow();
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Financing amount").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[1].AddParagraph("$ " + Quota.FinancingAmount.ToString(format)).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].AddParagraph("FINANCING FEE").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].AddParagraph("$ " + Quota.FinancingFee.ToString(format));

            row = table.AddRow();
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Amount Paid by Credit Card").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[1].AddParagraph("$ " + Quota.AmountPaidByCreditCard.ToString(format)).Format.Alignment = ParagraphAlignment.Center;
            row.Cells[2].AddParagraph("PROCESSING FEE").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].AddParagraph("$ " + Quota.ProcessingFee.ToString(format));

            row = table.AddRow();
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;            
            row.Cells[2].AddParagraph("GRAND TOTAL").Format.Alignment = ParagraphAlignment.Left;
            row.Cells[3].AddParagraph("$ " + Quota.InvoiceGrandTotal.ToString(format));

            row = table.AddRow();
            row.Format.Alignment = ParagraphAlignment.Right;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[2].AddParagraph("REQUIRED DEPOSIT").Format.Alignment = ParagraphAlignment.Left;
            decimal deposit;
            if (Quota.FinancingYesNo)
            {
                deposit = decimal.Round((Quota.FinancingAmount * 0.08m) - Quota.ProjectTotal, 2);
            }
            else
            {
                deposit = decimal.Round((Quota.LabourTotal * 0.25m) + Quota.MaterialTotal, 2);
            }
            row.Cells[3].AddParagraph("$ " + deposit.ToString(format));

            table.SetEdge(0, 0, 4, 5, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            #endregion

            // Виводить таблицю з проведеними оплатами ***************************************************
            #region Payment view                                

            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "2cm";
            //frame = section.AddTextFrame();
            //frame.Left = ShapePosition.Left;
            ////frame.Top = "0.5cm";
            //table = frame.AddTable();
            table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;
            // Створюємо стовбчики таблиці 
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("5cm");
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
            row.Cells[0].AddParagraph("Payment");
            row.Cells[1].AddParagraph("Date paid");
            row.Cells[2].AddParagraph("Amount Paid");            
            row.Cells[3].AddParagraph("Payment Method");
            row.Cells[4].AddParagraph("Balance");
            // Додаємо рядок
            int countPayment = 1;
            foreach (var item in Payments)
            {                
                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Right;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = true;
                row.Cells[0].AddParagraph(countPayment.ToString()).Format.Alignment = ParagraphAlignment.Center; 
                row.Cells[1].AddParagraph(item.PaymentDatePaid.ToShortDateString()).Format.Alignment = ParagraphAlignment.Center; 
                row.Cells[2].AddParagraph("$ " + item.PaymentAmountPaid.ToString(format));
                row.Cells[3].AddParagraph(((string.IsNullOrEmpty(item.NumberPayment)) ? (item.PaymentMethod) : (item.PaymentMethod + Environment.NewLine + "TID:" + item.NumberPayment))).Format.Alignment = ParagraphAlignment.Center;
                //row.Cells[4].AddParagraph("$ " + item.Balance.ToString(format));
                if (item.Balance > 0)
                {
                    row.Cells[4].AddParagraph("$ " + item.Balance.ToString(format));
                }
                else
                {
                    row.Cells[4].AddParagraph("Paid in full").Format.Font.Color = Colors.Red; ;
                }
                countPayment++;
            }

            table.SetEdge(0, 0, 5, Payments.Count() + 1, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            #endregion
            //********************************************************************************************

        }


        private static void MergeMultiplePDFIntoSinglePDF(string outputFilePath, string[] pdfFiles)
        {            
            PdfDocument outputPDFDocument = new PdfDocument();
            foreach (string pdfFile in pdfFiles)
            {
                PdfDocument inputPDFDocument;
                try
                {
                    inputPDFDocument = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import);
                    outputPDFDocument.Version = inputPDFDocument.Version;
                    foreach (PdfPage page in inputPDFDocument.Pages)
                    {
                        outputPDFDocument.AddPage(page);
                    }
                }
                catch 
                {                   
                }             
                
            }
            
            outputPDFDocument.Save(outputFilePath);            
        }
    }
}
