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
    public class PrintSignaturePageFirst
    {
        public PrintSignaturePageFirst()
        {
        }

        private Document document;

        public Quotation Quotation { get; set; }
        public Client Client { get; set; }
        public DateTime Date { get; set; }
        public string filename { get; } = @"Blanks\SignatureFirst.pdf";
        public User User { get; set; }

        public void Print()
        {
            CreateDocument();

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            //string filename = @"Blanks\Signature.pdf";

            pdfRenderer.PdfDocument.Save(filename);// сохраняем 

        }

        private Document CreateDocument()
        {
            // Create a new MigraDoc document
            this.document = new Document();
            this.document.Info.Title = "Signature";
            this.document.Info.Subject = "Create an 3 page PDF.";
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
            style.Font.Size = 13;



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
            section.PageSetup.LeftMargin = "2cm";
            section.PageSetup.RightMargin = "2cm";
            section.PageSetup.TopMargin = "2cm";
            section.PageSetup.DifferentFirstPageHeaderFooter = true;

            Paragraph paragraph;



            // Додаємо таблицю з інформацією про вибраного клієнта ***************************************
            #region Client info
            // Тут вставляю пустий параграф для контролю відступу від заголовку

            paragraph = section.AddParagraph();
            paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
            paragraph.Format.LineSpacing = "1.0cm";
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddFormattedText("LIABILITY WAIVER", TextFormat.Bold).Font.Size = 15;
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("“Condition of Subfloor”", TextFormat.NotBold).Font.Size = 11;

            paragraph = section.AddParagraph();
            paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
            paragraph.Format.LineSpacing = "0.5cm";
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.AddFormattedText("This liability waiver is an agreement is between the Client and " + User.CompanyName);
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("I, the client, acknowledge that the circumstances of my sub-floor " +
                                       "have been discussed with me, and it is not congruent to the flooring " +
                                       "installation guidelines, required by the manufacturer.");
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("A sub-floor is the concrete and/or plywood that is below the floor which is being installed.");
            paragraph.AddFormattedText("Deficiencies in flatness will create excessive movement / hollow spots which " +
                                       "can attribute to premature deterioration of the integrity of the floor.");
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("The condition of the sub-floor must be documented in writing and with supporting photos. " +
                                       "Both parties (Client and " + User.CompanyName + ") must compile this with your sales receipt " +
                                       "in order to claim your installation and/or material warranty.");
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("By proceeding with the flooring installation, without the required leveling to acquire the proper " +
                                       "surface tolerances, I understand I forfeit my installation and material warranty.");
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("Thank you for using our services, this agreement shall be binding upon both undersigned.");
            paragraph.AddLineBreak();

            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
            paragraph.Format.LineSpacing = "1.0cm";
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.AddFormattedText("SIGNED AS AN AGREEMENT", TextFormat.Bold).Font.Size = 12;
            paragraph.AddLineBreak();

            paragraph = section.AddParagraph();
            paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
            paragraph.Format.LineSpacing = "0.5cm";
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.AddFormattedText("IN THE WITNESS WHEREOF, the parties hereto have affixed their signatures " +
                                        "as on the date specified above.", TextFormat.Bold).Font.Size = 11;
            paragraph.AddLineBreak();

            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
            paragraph.Format.LineSpacing = "1.0cm";
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.AddFormattedText(" □   ", TextFormat.Bold).Font.Size = 20;
            paragraph.AddFormattedText("I choose to proceed ", TextFormat.NotBold).Font.Size = 10;
            paragraph.AddFormattedText("with ", TextFormat.Bold).Font.Size = 10;
            paragraph.AddFormattedText("the necessary recommended amount of subfloor leveling.", TextFormat.NotBold).Font.Size = 10;
            paragraph.AddLineBreak();

            paragraph.AddFormattedText(" □   ", TextFormat.Bold).Font.Size = 20;
            paragraph.AddFormattedText("I choose to proceed ", TextFormat.NotBold).Font.Size = 10;
            paragraph.AddFormattedText("without ", TextFormat.Bold).Font.Size = 10;
            paragraph.AddFormattedText("the necessary recommended amount of subfloor leveling.", TextFormat.NotBold).Font.Size = 10;
            paragraph.AddLineBreak();

            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
            paragraph.Format.LineSpacing = "0.5cm";
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.AddFormattedText("The Customer:", TextFormat.Bold).Font.Size = 11;

            // Створюємо таблицю з інформацією про клієнта
            Table table = section.AddTable();
            table.Style = "Table";
            table.Format.Alignment = ParagraphAlignment.Center;
            table.Rows.LeftIndent = 0;
            // Створюємо стовбчики таблиці 
            Column column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("10cm");
            column.Format.Alignment = ParagraphAlignment.Center;

            Row row;
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Client Name:");
            row.Cells[1].AddParagraph((Client.PrimaryFullName) ?? "");
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Phone number:");
            row.Cells[1].AddParagraph((Client.PrimaryPhoneNumber) ?? "");
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("E-mail:");
            row.Cells[1].AddParagraph((Client.PrimaryEmail) ?? "");
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Date:");
            row.Cells[1].AddParagraph((Date.ToShortDateString()) ?? "");
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Signature:");

            table.SetEdge(1, 0, 1, 1, Edge.Bottom, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            table.SetEdge(1, 0, 1, 2, Edge.Bottom, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            table.SetEdge(1, 0, 1, 3, Edge.Bottom, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            table.SetEdge(1, 0, 1, 4, Edge.Bottom, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            table.SetEdge(1, 0, 1, 5, Edge.Bottom, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);







            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
            paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
            paragraph.Format.LineSpacing = "0.5cm";
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.AddFormattedText(User.CompanyName + "Representative:", TextFormat.Bold).Font.Size = 11;

            // Створюємо таблицю з інформацією про клієнта
            table = section.AddTable();
            table.Style = "Table";
            table.Format.Alignment = ParagraphAlignment.Center;
            table.Rows.LeftIndent = 0;
            // Створюємо стовбчики таблиці 
            column = table.AddColumn("5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn("10cm");
            column.Format.Alignment = ParagraphAlignment.Center;


            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Name:");
            row.Cells[1].AddParagraph((User.Name) ?? "");
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Designation:");
            row.Cells[1].AddParagraph((User.Post) ?? "");
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Phone number:");
            row.Cells[1].AddParagraph((User.Phone) ?? "");
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("E-mail:");
            row.Cells[1].AddParagraph((User.Mail) ?? "");

            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;
            row.Cells[0].AddParagraph("Signature:");

            table.SetEdge(1, 0, 1, 1, Edge.Bottom, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            table.SetEdge(1, 0, 1, 2, Edge.Bottom, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            table.SetEdge(1, 0, 1, 3, Edge.Bottom, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            table.SetEdge(1, 0, 1, 4, Edge.Bottom, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);
            table.SetEdge(1, 0, 1, 5, Edge.Bottom, BorderStyle.Single, 0.75, MigraDoc.DocumentObjectModel.Color.Empty);

            #endregion
            //********************************************************************************************
        }

    }
}
