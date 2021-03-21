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
    public class PrintSignaturePage
    {
        public PrintSignaturePage()
        {
        }
              

        
        private Document document;

        public Quotation Quotation { get; set; }
        public Client Client { get; set; }
        public DateTime Date { get; set; }
        public string filename { get; } = @"Blanks\Signature.pdf";
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
            paragraph.Format.LineSpacing = "0.5cm";
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.AddFormattedText("I hereby have read, accept and fully understand this Invoice and the terms " + 
                                       "and conditions of this Agreement, and hereby state that I am the Client or am " + 
                                       "authorized by the Client to approve this Invoice and this Agreement.", TextFormat.Bold).Font.Size = 11;
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("");
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("I acknowledge that I am over the age of 18, and I am either the homeowner or a " + 
                                       "person with authority to authorize renovation at the above address. I am the person " + 
                                       "who can approve any additional work if required, including levelling or other " + 
                                       "unforeseen circumstances.", TextFormat.Bold).Font.Size = 11;
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("");
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("By signing below, I authorize " + User?.CompanyName + " to process my credit card for the " + 
                                       "above transactions.  I understand that when job is complete all sensitive credit card " + 
                                       "information will be deleted.", TextFormat.Bold).Font.Size = 11;
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("");
            paragraph.AddLineBreak();
            paragraph.AddFormattedText("I ACCEPT THE TERMS AND SPECIFICATIONS STATED HEREIN AND AUTHORIZE " + User?.CompanyName + 
                                       " TO PROCEED WITH THE PROJECT.", TextFormat.Bold).Font.Size = 11;
            paragraph.AddLineBreak();


            paragraph = section.AddParagraph();
            paragraph = section.AddParagraph();
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
            paragraph.AddFormattedText(User?.CompanyName + "Representative:", TextFormat.Bold).Font.Size = 11;

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
            row.Cells[1].AddParagraph((User?.Name + "") ?? "");
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;            
            row.Cells[0].AddParagraph("Designation:");
            row.Cells[1].AddParagraph((User?.Post + "") ?? "");
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;            
            row.Cells[0].AddParagraph("Phone number:");
            row.Cells[1].AddParagraph((User?.Phone + "") ?? "");
            // Додаємо рядок
            row = table.AddRow();
            row.Height = "0.8cm";
            row.Format.Alignment = ParagraphAlignment.Left;
            row.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom;            
            row.Cells[0].AddParagraph("E-mail:");
            row.Cells[1].AddParagraph((User?.Mail + "") ?? "");
            
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
