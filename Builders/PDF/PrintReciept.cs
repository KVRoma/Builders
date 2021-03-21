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

namespace Builders.PDF
{
    public class PrintReciept
    {
        public PrintReciept(User _user)
        {
            user = _user;
        }

        private Document document;        

        private Reciept reciept;
        private Payment paymentSelect;
        private IEnumerable<Payment> payments;
        private Quotation quotation;
        private Client client;
        private string format = "0.00";
        private User user;

        public Reciept Reciept
        {
            get { return reciept; }
            set
            {
                reciept = value;
            }
        }
        public Payment PaymentSelect
        {
            get { return paymentSelect; }
            set
            {
                paymentSelect = value;
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
        public Quotation Quotation
        {
            get { return quotation; }
            set
            {
                quotation = value;
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
        
        public void Print()
        {
            CreateDocument();

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            string filename = @"Output Files\Reciept" + "-" + Reciept.Number +                                                         
                                                         "-" + PaymentSelect.PaymentDatePaid.Day.ToString("00") +
                                                         "." + PaymentSelect.PaymentDatePaid.Month.ToString("00") +
                                                         "." + PaymentSelect.PaymentDatePaid.Year.ToString("0000") + ".pdf";            
            pdfRenderer.PdfDocument.Save(filename);// сохраняем            
            Process.Start(filename);
        }

        private Document CreateDocument()
        {
            // Create a new MigraDoc document
            this.document = new Document();
            this.document.Info.Title = "Reciept";
            this.document.Info.Subject = "Create an reciept PDF.";
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

            // Додаємо логотип в заголовок + інформацію про власника та розміщуємо під логотипом *******
            #region Heder logo
            TextFrame frame = section.Headers.Primary.AddTextFrame();
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

            
            Paragraph paragraph = frame.AddParagraph(user.Name + ", " + Environment.NewLine +
                                          user.Post + Environment.NewLine +
                                          user.Address + Environment.NewLine +
                                          user.Additional + Environment.NewLine +
                                          user.Phone + Environment.NewLine + 
                                          user.Mail + Environment.NewLine + 
                                          user.WebSite);
            //paragraph.Format.Font.Name = "Times New Roman";
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.Format.Font.Size = 7;
            paragraph.Format.SpaceAfter = 3;
            #endregion
            //*******************************************************************************************

            // Додаємо назву документу та заповнюємо дату, номер та номер квоти *************************
            #region Heder data
            frame = section.Headers.Primary.AddTextFrame();
            frame.Left = ShapePosition.Right;
            frame.Width = "6.0cm";

            paragraph = frame.AddParagraph();
            paragraph.AddFormattedText("RECIEPT", TextFormat.Bold);
            paragraph.Format.Font.Size = 22;
            paragraph.Format.Font.Color = Colors.Red;
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();

            paragraph = frame.AddParagraph();
            paragraph.Format.Font.Size = 10;
            paragraph.AddText("Date: ");
            paragraph.AddFormattedText(PaymentSelect.PaymentDatePaid.ToShortDateString(), TextFormat.Bold);
            paragraph.AddLineBreak();
            paragraph.AddText("Reciept Number: ");
            paragraph.AddFormattedText(Reciept?.Number, TextFormat.Bold);
            paragraph.AddLineBreak();
            paragraph.AddText("Estimate Number: ");
            paragraph.AddFormattedText(Quotation.NumberQuota, TextFormat.Bold);
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
            paragraph.Format.SpaceBefore = "4.5cm";


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
            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            // Додаємо рядок
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            //row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.White;
            row.Cells[0].AddParagraph("Description:");
            row.Cells[1].AddParagraph((Quotation.JobDescription?.ToUpper()) ?? "").Format.Font.Bold = false;
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
            column = table.AddColumn("14cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("NOTES:");
            row.Cells[1].AddParagraph((Quotation.JobNote) ?? "").Format.Font.Bold = false;

            table.SetEdge(0, 0, 2, 1, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            #endregion
            //********************************************************************************************

            // Створюєм таблицю інформацією по чеку*******************************************************
            #region Reciept info
            //paragraph = section.AddParagraph();

            frame = section.AddTextFrame();
            frame.Left = ShapePosition.Right;
            frame.Width = "7.9cm";
            frame.Top = "0.5cm";

            table = frame.AddTable();
            table.Style = "Table";
            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Format.Alignment = ParagraphAlignment.Right;
            //table.Rows.LeftIndent = 0;          

            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Right;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("MATERIAL");
            row.Cells[0].MergeRight = 1;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Subtotal:");
            row.Cells[1].AddParagraph("$ " + Quotation.MaterialSubtotal.ToString(format)).Format.Font.Bold = false;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("TAX:");
            row.Cells[1].AddParagraph("$ " + Quotation.MaterialTax.ToString(format)).Format.Font.Bold = false;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;

            if (Quotation.MaterialDiscountAmount != 0m)
            {                
                row.Cells[0].AddParagraph("Discount");
                row.Cells[1].AddParagraph("$ " + Quotation.MaterialDiscountAmount.ToString(format)).Format.Font.Bold = false;
            }

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Total:");
            row.Cells[1].AddParagraph("$ " + Quotation.MaterialTotal.ToString(format));

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("LABOUR");
            row.Cells[0].MergeRight = 1;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Subtotal:");
            row.Cells[1].AddParagraph("$ " + Quotation.LabourSubtotal.ToString(format)).Format.Font.Bold = false;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("TAX:");
            row.Cells[1].AddParagraph("$ " + Quotation.LabourTax.ToString(format)).Format.Font.Bold = false;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            if (Quotation.LabourDiscountAmount != 0m)
            {                
                row.Cells[0].AddParagraph("Discount");
                row.Cells[1].AddParagraph("$ " + Quotation.LabourDiscountAmount.ToString(format)).Format.Font.Bold = false;
            }

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Total:");
            row.Cells[1].AddParagraph("$ " + Quotation.LabourTotal.ToString(format));

            row = table.AddRow();
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Total Material and Labour:");
            row.Cells[1].AddParagraph("$ " + decimal.Round(Quotation.MaterialTotal + Quotation.LabourTotal, 2).ToString(format));

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Financing Cost:");
            row.Cells[1].AddParagraph("$ " + decimal.Round(Quotation.ProcessingFee + Quotation.FinancingFee, 2).ToString(format)).Format.Font.Bold = false;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            if (Quotation.MaterialDiscountAmount != 0m || Quotation.LabourDiscountAmount != 0m)
            {                
                row.Cells[0].AddParagraph("Discount");
                row.Cells[1].AddParagraph("$ " + decimal.Round(Quotation.MaterialDiscountAmount + Quotation.LabourDiscountAmount, 2).ToString(format)).Format.Font.Bold = false;
            }

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Total Invoice Amount:");
            row.Cells[1].AddParagraph("$ " + decimal.Round(Quotation.MaterialTotal + Quotation.LabourTotal + Quotation.ProcessingFee + Quotation.FinancingFee, 2).ToString(format));

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Total Paid to Date:");
            row.Cells[1].AddParagraph("$ " + decimal.Round(Payments.Select(p => p.PaymentPrincipalPaid).Sum(), 2).ToString(format)).Format.Font.Bold = false;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Total Remaining Balance:");
            row.Cells[1].AddParagraph("$ " + decimal.Round((Quotation.MaterialTotal + Quotation.LabourTotal + Quotation.ProcessingFee + Quotation.FinancingFee) - (Payments.Select(p => p.PaymentPrincipalPaid).Sum()), 2).ToString(format)).Format.Font.Bold = false;

            table.SetEdge(0, 0, 2, 17, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);


            paragraph = frame.AddParagraph();

            table = frame.AddTable();
            table.Style = "Table";
            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Format.Alignment = ParagraphAlignment.Right;
            //table.Rows.LeftIndent = 0;          

            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Right;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Right;

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("PAYMENT AMOUNT RECEIVED:");
            row.Cells[1].AddParagraph("$ " + PaymentSelect.PaymentAmountPaid.ToString(format));

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("PAYMENT METHOD:");
            row.Cells[1].AddParagraph(PaymentSelect.PaymentMethod);

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("Processing Fee:");
            row.Cells[1].AddParagraph("$ " + PaymentSelect.ProcessingFee.ToString(format));

            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("Amount Applied to Principal:");
            row.Cells[1].AddParagraph("$ " + PaymentSelect.PaymentPrincipalPaid.ToString(format));

            table.SetEdge(0, 0, 2, 4, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);

            // Виводить номера рахунків
            frame = section.AddTextFrame();
            frame.Left = ShapePosition.Left;
            frame.Width = "5.0cm";
            frame.Top = "0.5cm";

            paragraph = frame.AddParagraph();            
            paragraph.Format.Font.Size = 7;
            if (!string.IsNullOrWhiteSpace(user.GST))
            {
                paragraph.AddText("GST # ");
                paragraph.AddFormattedText(user.GST, TextFormat.Bold);
                paragraph.AddLineBreak();
                paragraph.AddLineBreak();
            }
            if (!string.IsNullOrWhiteSpace(user.PST))
            {
                paragraph.AddText("PST # ");
                paragraph.AddFormattedText(user.PST, TextFormat.Bold);
                paragraph.AddLineBreak();
                paragraph.AddLineBreak();
            }
            if (!string.IsNullOrWhiteSpace(user.WCB))
            {
                paragraph.AddText("WCB # ");
                paragraph.AddFormattedText(user.WCB, TextFormat.Bold);
                paragraph.AddLineBreak();
                paragraph.AddLineBreak();
            }
            if (!string.IsNullOrWhiteSpace(user.Liability))
            {
                paragraph.AddText("GENERAL LIABILITY # ");
                paragraph.AddFormattedText(user.Liability, TextFormat.Bold);
                paragraph.AddLineBreak();
                paragraph.AddLineBreak();
            }
            if (!string.IsNullOrWhiteSpace(user.Licence))
            {
                paragraph.AddText("BUSINESS LICENSE # ");
                paragraph.AddFormattedText(user.Licence, TextFormat.Bold);
                paragraph.AddLineBreak();
                paragraph.AddLineBreak();
            }
            if (!string.IsNullOrWhiteSpace(user.Incorporation))
            {
                paragraph.AddText("INCORPORATION # ");
                paragraph.AddFormattedText(user.Incorporation, TextFormat.Bold);
                paragraph.AddLineBreak();
                paragraph.AddLineBreak();
            }
            #endregion
            //********************************************************************************************

            // Виводить таблицю з проведеними оплатами ***************************************************
            #region Payment view                                
                        
            table = section.Footers.Primary.AddTable();            
            table.Style = "Table";
            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;
            // Створюємо стовбчики таблиці 
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2.5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("Date paid");
            row.Cells[1].AddParagraph("Amount Paid");
            row.Cells[2].AddParagraph("Principal Paid");
            row.Cells[3].AddParagraph("Administration Fee");
            row.Cells[4].AddParagraph("Payment Method");
            row.Cells[5].AddParagraph("Remaining Balance");
            // Додаємо рядок
            foreach (var item in Payments)
            {
                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Right;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = true;
                row.Cells[0].AddParagraph(item.PaymentDatePaid.ToShortDateString());
                row.Cells[1].AddParagraph("$ " + item.PaymentAmountPaid.ToString(format));
                row.Cells[2].AddParagraph("$ " + item.PaymentPrincipalPaid.ToString(format));
                row.Cells[3].AddParagraph("$ " + item.ProcessingFee.ToString(format));
                row.Cells[4].AddParagraph("$ " + ((string.IsNullOrEmpty(item.NumberPayment.ToString())) ? (item.PaymentMethod) : (item.PaymentMethod + Environment.NewLine + "TID:" + item.NumberPayment))).Format.Alignment = ParagraphAlignment.Center;
                if (item.Balance > 0)
                {
                    row.Cells[5].AddParagraph("$ " + item.Balance.ToString(format));
                }
                else
                {
                    row.Cells[5].AddParagraph("Paid in full").Format.Font.Color = Colors.Red; ;                   
                }                
            }

            table.SetEdge(0, 0, 6, Payments.Count() + 1, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            #endregion
            //********************************************************************************************
        }
    }
}
