using System;
using System.Linq;
using GemBox.Document;
using GemBox.Pdf;
using GemBox.Pdf.Content;
using System.Drawing;

namespace StampFromDocToDocAndPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // лицензирование GemBox
            GemBox.Document.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            GemBox.Pdf.ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            SavePicDoc();
            SaveDoc();
            SavePdf();

            Console.WriteLine("Done saving doc and pdf");
            Console.ReadLine();
        }

        static void SaveDoc()
        {
            PdfConformanceLevel conformanceLevel = PdfConformanceLevel.PdfA3u;

            string destFileName = @"C:\Users\Выймова Елена\Documents\Input.docx";

            DocumentModel document = DocumentModel.Load(destFileName);

            // загрузка изображения с определенными размерами 
            // page.Size.Width = 595, уменьшаем, потому что в doc есть еще отступ слева
            // image.Height = 676, уменьшаем docx по высоте больше, чем pdf, потому что меньше и ширина
            Picture picture = new Picture(document, @"C:\Users\Выймова Елена\Documents\InputPic_Cropped.png", 595 - 90, 676 / 3 - 90);

            // последняя страница документа
            Section section = document.Sections.Last();
            // элемент "параграф" со штампом
            var paragraphPic = new Paragraph(document, picture);
            // добавление пустой строки
            section.Blocks.Add(new Paragraph(document));
            // добавление строки со штампом
            section.Blocks.Add(paragraphPic);

            var options = new GemBox.Document.PdfSaveOptions()
            {
                ConformanceLevel = conformanceLevel
            };

            // сохранение изменений
            document.Save(@"C:\Users\Выймова Елена\Documents\Output from docx 2.pdf", options);
        }

        static void SavePdf()
        {
            string destFileNamePdf = @"C:\Users\Выймова Елена\Documents\Input.pdf";

            PdfDocument documentPdf = PdfDocument.Load(destFileNamePdf);

            PdfImage image = PdfImage.Load(@"C:\Users\Выймова Елена\Documents\InputPic_Cropped.png");

            // последняя страница документа
            var page = documentPdf.Pages.Last();

            // вставка изображения с определенными размерами
            // page.Size.Width = 595
            // image.Height = 676
            page.Content.DrawImage(image, new PdfPoint(0, 0), new PdfSize(page.Size.Width, image.Height / 3 - 50));

            // сохранение изменений
            documentPdf.Save(@"C:\Users\Выймова Елена\Documents\Output from pdf 2.pdf");

        }

        static void SavePicDoc()
        {
            string destFileName = @"C:\Users\Выймова Елена\Documents\InputPicDoc.docx";

            DocumentModel document = DocumentModel.Load(destFileName);

            var imageOptions = new GemBox.Document.ImageSaveOptions(GemBox.Document.ImageSaveFormat.Png)
            {
                PageNumber = 0
            };

            string outputFile = @"C:\Users\Выймова Елена\Documents\InputPic.png";
            document.Save(outputFile, imageOptions);

            Bitmap croppedImage = new Bitmap(outputFile);

            // отступ сверху 100
            // по высоте обрезаем, чтобы остался на листе только штамп
            Rectangle rectangle = new Rectangle(0, 100, croppedImage.Width, croppedImage.Height / 4 - 200);
            croppedImage = CropAtRect(croppedImage, rectangle);

            croppedImage.Save(@"C:\Users\Выймова Елена\Documents\InputPic_Cropped.png");
            Console.WriteLine("Done converting image");
        }

        static Bitmap CropAtRect(Bitmap croppedImage, Rectangle rectangle)
        {
            Bitmap bmp = new Bitmap(rectangle.Width, rectangle.Height);

            using (Graphics gfx = Graphics.FromImage(bmp))
            {
                gfx.DrawImage(croppedImage, new Rectangle(-rectangle.X, -rectangle.Y, rectangle.Width, rectangle.Height), rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height, GraphicsUnit.Pixel);
                return bmp;
            }
        }
    }
}
