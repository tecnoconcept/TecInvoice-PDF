using Config.Net;
using iText.Kernel.Pdf;
using iText.Kernel.XMP;
using iText.Kernel.XMP.Options;
using NLog;
using System.Runtime.CompilerServices;
using TecInvoice_PDF.Config;
using its = iText.Kernel;

namespace TecInvoice_PDF
{
    internal class Program
    {
        private static Logger Logger { get; } = LogManager.GetCurrentClassLogger();

        private static void Main(string[] args)
        {
            var config = new ConfigurationBuilder<IConfig>()
                .UseCommandLineArgs()
                .Build();

            if (string.IsNullOrWhiteSpace(config.xmlFilePath) || string.IsNullOrWhiteSpace(config.pdfFilePath))
            {
                Logger.Error("Please provide the xml and pdf file paths.");
                return;
            }
            Logger.Info($"Adding {config.xmlFilePath} to {config.pdfFilePath}");
            AddXmlToPdf(config.xmlFilePath, config.pdfFilePath, config.Profile);
            Logger.Info("Successfully added XML to PDF.");
        }
        private static void AddXmlToPdf(string xmlFile, string pdfFile, string profile)
        {
            var tempFile = Path.GetTempFileName();

            var reader = new its.Pdf.PdfReader(pdfFile);
            var writer = new its.Pdf.PdfWriter(new FileStream(tempFile, FileMode.Create));
            var pdfDoc = new its.Pdf.PdfDocument(reader, writer);

            Logger.Info("Adding XMP Metadata...");
            byte[] xmpFile = File.ReadAllBytes($"TecFerd_{profile.ToUpper()}.xmp");
            XMPMeta xmp = XMPMetaFactory.ParseFromBuffer(xmpFile);
            SerializeOptions options = new SerializeOptions();
            options.SetUseCanonicalFormat(true);
            pdfDoc.SetXmpMetadata(xmp, options);

            Logger.Info("Attaching XML to the PDF...");
            PdfDictionary parameters = new PdfDictionary();
            parameters.Put(PdfName.ModDate, new PdfDate().GetPdfObject());
            var fs = its.Pdf.Filespec.PdfFileSpec.CreateEmbeddedFileSpec(pdfDoc, File.ReadAllBytes(xmlFile), "factur-x.xml", "factur-x.xml", new PdfName("text/xml"), parameters, PdfName.Alternative);
            pdfDoc.AddFileAttachment("factur-x.xml", fs);
            PdfArray array = new PdfArray();
            array.Add(fs.GetPdfObject().GetIndirectReference());
            pdfDoc.GetCatalog().Put(PdfName.AF, array);

            Logger.Info("Closing PDF...");
            pdfDoc.Close();
            reader.Close();

            File.Delete(pdfFile);
            File.Move(tempFile, pdfFile);
        }
    }
}
