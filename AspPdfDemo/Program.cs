using ASPPDFLib;
using System;
using System.IO;
using System.Reflection;

namespace AspPdfDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var currentDirectoryPath = Directory.GetCurrentDirectory();

            var imagePath = currentDirectoryPath + "\\image.png";
            var pdfFolder = currentDirectoryPath + "\\PDF\\";
            var imageFolder = currentDirectoryPath + "\\Image\\";

            if (!File.Exists(imagePath))
            {
                return;
            }

            if (!Directory.Exists(pdfFolder))
            {
                Directory.CreateDirectory(pdfFolder);
            }

            if (!Directory.Exists(imageFolder))
            {
                Directory.CreateDirectory(imageFolder);
            }

            for (var i = 0; i <= 100; i++)
            {
                var fileName = i.ToString();
                var pdfPath = pdfFolder + "\\" + fileName + ".pdf";
                var imageFilePath = imageFolder + string.Format("{0}_{1}.png", fileName, "{0}");
                CreatePdf(imagePath, imagePath, pdfPath);
                CreateImageFromPdf(pdfPath, imageFilePath);
            }

        }

        private static string CreatePdf(
            string frontTemplateImagePath,
            string backTemplateImagePath,
            string pdfPath)
        {

            var pdfManagerObj = new PdfManager();
            var pdfDocumentObj = pdfManagerObj.CreateDocument();
            pdfDocumentObj.CreateColorSpace("DeviceCMYK");
            CreatePdfPage(pdfManagerObj, pdfDocumentObj, frontTemplateImagePath);
            CreatePdfPage(pdfManagerObj, pdfDocumentObj, backTemplateImagePath);
            return pdfDocumentObj.Save(pdfPath, true);
        }

        private static void CreateImageFromPdf(string sourcePath, string destinationPath)
        {
            try
            {
                var pdfManagerObj = new PdfManager();
                var objPdfDoc = pdfManagerObj.OpenDocument(sourcePath, Missing.Value);
                var objParam = pdfManagerObj.CreateParam(Missing.Value);
                objParam["ResolutionX"].Value = 36;
                objParam["ResolutionY"].Value = 36;
                objParam["ScaleX"].Value = 0.1f;
                objParam["ScaleY"].Value = 0.1f;
                var objFrontPage = objPdfDoc.Pages[1];
                var objFrontPreview = objFrontPage.ToImage(objParam);
                objFrontPreview.Save(String.Format(destinationPath, "front"), true);
                var objBackPage = objPdfDoc.Pages[2];
                var objBackPreview = objBackPage.ToImage(objParam);
                objBackPreview.Save(String.Format(destinationPath, "back"), true);
                objPdfDoc.Close();
            }
            catch (Exception exception)
            {
            }
        }
        private static void CreatePdfPage(IPdfManager pdfManagerObj,
         IPdfDocument pdfDocumentObj,
         string imagePath)
        {
            try
            {
                var objFont = pdfDocumentObj.Fonts["Times-BoldItalic", Missing.Value];
                var objPage = pdfDocumentObj.Pages.Add(9.25 * 72, 10.875 * 72, Missing.Value);
                var templateImage = pdfDocumentObj.OpenImage(imagePath, Missing.Value);
                objPage.Width = templateImage.Width;
                objPage.Height = templateImage.Height;
                var objParam = GenericCreateParam(pdfManagerObj, 0, 0,
                  objPage.Width, objPage.Height, 1, 1);
                objPage.Background.DrawImage(templateImage, objParam);
            }
            catch (Exception exception)
            {
                throw;
            }
        }

        private static IPdfParam GenericCreateParam(IPdfManager objPdf,
           float xValue, float yValue, float widthValue,
           float heightValue, float scaleXValue, float scaleYValue)
        {
            var objParam = objPdf.CreateParam(Missing.Value);
            objParam["x"].Value = xValue;
            objParam["y"].Value = yValue;
            objParam["width"].Value = widthValue;
            objParam["height"].Value = heightValue;
            objParam["ScaleX"].Value = scaleXValue;
            objParam["ScaleY"].Value = scaleYValue;
            return objParam;
        }
    }
}
