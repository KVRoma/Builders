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
    public class PrintDelivery
    {
        public PrintDelivery(User _user)
        {
            user = _user;
            Material = new List<string>();
            DeliveryMaterials = new List<DeliveryMaterial>();

            if (user?.Id == 1)
            {
                company = "CMO";
            }
            else if (user?.Id == 2)
            {
                company = "NL";
            }
            else 
            {
                company = "";
            }
            
        }
        private User user;
        private Quotation quota;
        private Client client;
        private Delivery deliverySelect;
        private List<string> material;
        private DIC_Supplier supplier;
        private List<DeliveryMaterial> deliveryMaterials;

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
        public Delivery DeliverySelect
        {
            get { return deliverySelect; }
            set
            {
                deliverySelect = value;
            }
        }
        public List<string> Material
        {
            get { return material; }
            set
            {
                material = value;
            }
        }
        public DIC_Supplier Supplier
        {
            get { return supplier; }
            set
            {
                supplier = value;
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
        private string company; 

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
            string filename = folder + @"\\Builders File\" + company + "_MaterialPO" + "-" + Quota.NumberQuota +
                                                                                       "-" + DeliverySelect.DateCreating.Day.ToString("00") +
                                                                                       "." + DeliverySelect.DateCreating.Month.ToString("00") +
                                                                                       "." + DeliverySelect.DateCreating.Year.ToString("0000") + ".pdf";

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

            Image image = frame.AddImage(@"User\Logo" + company + ".jpg");
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

            //frame = section.AddTextFrame();
            //frame.Left = ShapePosition.Right;            
            //frame.Top = "1.0cm";

            frame = section.Headers.FirstPage.AddTextFrame();
            frame.Left = ShapePosition.Right;
            frame.Width = "6.5cm";

            paragraph = frame.AddParagraph();
            paragraph.AddFormattedText("Material Purchase Order", TextFormat.Bold);
            paragraph.Format.Font.Size = 15;
            paragraph.Format.Font.Color = Colors.Red;
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();

            paragraph = frame.AddParagraph();
            paragraph.Format.Font.Size = 10;
            //paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddText("Date: ");
            paragraph.AddFormattedText(DeliverySelect.DateCreating.ToShortDateString(), TextFormat.Bold);
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();

            paragraph = frame.AddParagraph();
            paragraph.Format.Font.Size = 7;
            if (!string.IsNullOrWhiteSpace(user.GST))
            {
                paragraph.AddText("GST # ");
                paragraph.AddFormattedText(user.GST, TextFormat.Bold);
                paragraph.AddLineBreak();
            }
            if (!string.IsNullOrWhiteSpace(user.PST))
            {
                paragraph.AddText("PST # ");
                paragraph.AddFormattedText(user.PST, TextFormat.Bold);
                paragraph.AddLineBreak();
            }
            if (!string.IsNullOrWhiteSpace(user.WCB))
            {
                paragraph.AddText("WCB # ");
                paragraph.AddFormattedText(user.WCB, TextFormat.Bold);
                paragraph.AddLineBreak();
            }
            if (!string.IsNullOrWhiteSpace(user.Liability))
            {
                paragraph.AddText("GENERAL LIABILITY # ");
                paragraph.AddFormattedText(user.Liability, TextFormat.Bold);
                paragraph.AddLineBreak();
            }
            if (!string.IsNullOrWhiteSpace(user.Licence))
            {
                paragraph.AddText("BUSINESS LICENSE # ");
                paragraph.AddFormattedText(user.Licence, TextFormat.Bold);
                paragraph.AddLineBreak();
            }
            if (!string.IsNullOrWhiteSpace(user.Incorporation))
            {
                paragraph.AddText("INCORPORATION # ");
                paragraph.AddFormattedText(user.Incorporation, TextFormat.Bold);
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

            image = frame.AddImage(@"User\Logo" + company + ".jpg");
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
            Column column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("17cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            string abreviatura = "";
            if (user.Id == 1)
            {
                abreviatura = "CMO";
            }
            if (user.Id == 2)
            {
                abreviatura = "NL";
            }

            // Додаємо рядок
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph(abreviatura + " PO # ");
            row.Cells[1].AddParagraph(("\"" + Client?.NumberClient + " - " + Quota?.NumberQuota + "\"") ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph(abreviatura + " TAG  ");
            row.Cells[1].AddParagraph(("\"" + abreviatura + " - " + Client?.PrimaryFullName?.ToUpper() + "\"") ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Supplier");
            row.Cells[1].AddParagraph((Supplier?.Supplier?.ToUpper()) ?? "").Format.Font.Bold = false;
            // Додаємо рядок
            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Cells[0].AddParagraph("Address");
            row.Cells[1].AddParagraph((Supplier?.Address?.ToUpper()) ?? "").Format.Font.Bold = false;

            table.SetEdge(0, 0, 2, 4, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            #endregion
            //********************************************************************************************

            // Створюєм таблицю інформацією по чеку*******************************************************
            #region Quota Material info
            paragraph = section.AddParagraph();

            table = section.AddTable();
            table.Style = "Table";


            table.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Rows.LeftIndent = 0;

            column = table.AddColumn("17cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;


            row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = MigraDoc.DocumentObjectModel.Colors.Silver;
            row.Cells[0].AddParagraph("Name").Format.Alignment = ParagraphAlignment.Center;
            row.Cells[1].AddParagraph("QUANTITY").Format.Alignment = ParagraphAlignment.Center;

            int countRows = 1;
            foreach (var item in Material)
            {
                decimal quantity = DeliveryMaterials.Where(m => m.Description == item)?.Select(m => m.Quantity)?.Sum() ?? 0m;

                row = table.AddRow();
                row.HeadingFormat = false;
                row.Format.Alignment = ParagraphAlignment.Right;
                row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center;
                row.Format.Font.Bold = true;
                row.Cells[0].AddParagraph((item) ?? "").Format.Alignment = ParagraphAlignment.Left;
                row.Cells[1].AddParagraph((quantity.ToString(format)) ?? "").Format.Alignment = ParagraphAlignment.Center;

                countRows++;
            }



            table.SetEdge(0, 0, 2, countRows, Edge.Box, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);






            #endregion
            //********************************************************************************************
        }
    }
}
