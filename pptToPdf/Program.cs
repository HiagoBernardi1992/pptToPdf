using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationToPdfConverter;
using System.IO;

namespace pptToPdf
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Carrega a apresentação do power point via stream
            using(FileStream fileStreamInput = new FileStream(@"C:/Users/Integer/OneDrive - Integer Consulting/Ambiente de Trabalho/Documentos/Teste.pptx", FileMode.Open, FileAccess.ReadWrite))
            {
                //carrega o stream do ppt no Presentation
                using (IPresentation pptxDoc = Presentation.Open(fileStreamInput))
                {
                    foreach (ISlide slide in pptxDoc.Slides)
                    {
                        //Intera sobre as formas do powerpoint
                        foreach (IShape shape in slide.Shapes)
                        {
                            if (shape != null)
                            {
                                switch (shape.TextBody.Text)
                                {
                                    case "Teste 1":
                                        shape.TextBody.Text = "Título";
                                        break;
                                    case "Teste 2":
                                        shape.TextBody.Text = "Texto";
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                    }

                    //Cria o PDF para fazer a transferencia
                    using (MemoryStream pdfStream = new MemoryStream())
                    {
                        //Converte o power point para o pdf
                        using (PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc))
                        {
                            //Salva o PDF convertido em MemoryStream.
                            pdfDocument.Save(pdfStream);
                            pdfStream.Position = 0;
                        }

                        //Cria saida para o PDF
                        using (FileStream fileStreamOutput = File.Create("C:/Users/Integer/OneDrive - Integer Consulting/Ambiente de Trabalho/Documentos/Output.pdf"))
                        {
                            //copia o pdf convertido para a saída
                            pdfStream.CopyTo(fileStreamOutput);
                        }
                    }
                }
            }
        }
    }
}
