using CreatingPDF.Models;
using iTextSharp.text; //Instalar a iTextSharp
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace CreatingPDF.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult PDF()
        {
            using (MemoryStream ms = new MemoryStream()) //Inicializa uma nova instância da classe MemoryStream com uma capacidade expansível inicializada em zero
            {
                //Criando classe generica de um documento, e colocando suas caracteristicas como tipo de papel e etc
                Document document = new Document(PageSize.A4, 25, 25, 30, 30); 

                //Criando um PdfWriter e pegar o documento e o ms da MemoryStream
                PdfWriter writer = PdfWriter.GetInstance(document, ms);
                document.Open();

                //Criando variavel imagem e pegando sua localização e também colocando seu alinhamento(Para ficar com a marca da empresa no topo do relatório)
                var image = iTextSharp.text.Image.GetInstance("C:\\Users\\Lucas\\Desktop\\Generate PDF Report with Image and Table in ASP.NET\\CreatingPDF\\CreatingPDF\\wwwroot\\image\\conectiva img.png");
                image.Alignment = Element.ALIGN_CENTER;
                document.Add(image); //Adicionando a imagem ao documento

                //Criando um novo parágrafo e colocando suas caracteristicas(Texto, Fonte, Tamanho, Alinhamento, Tamanho do espaço entre o último elemento)
                Paragraph para1 = new Paragraph("Esse é o primeiro parágrafo", new Font(Font.FontFamily.HELVETICA, 20));
                para1.Alignment = Element.ALIGN_CENTER;
                para1.SpacingBefore = 7;
                document.Add(para1);

                Paragraph para2 = new Paragraph("Esse é o segundo parágrafo", new Font(Font.FontFamily.HELVETICA, 15));
                para2.Alignment = Element.ALIGN_CENTER;
                document.Add(para2);

                Paragraph para3 = new Paragraph("Esse é o terceiro parágrafo", new Font(Font.FontFamily.HELVETICA, 10));
                para3.Alignment = Element.ALIGN_CENTER;
                para3.SpacingAfter = 10;
                document.Add(para3);
             
                //Criando uma tabela em pdf e colocando a quantidade de colunas
                PdfPTable table = new PdfPTable(4);

                //Criando as células e colocando o nome de cada coluna do cabeçalho
                PdfPCell cell1 = new PdfPCell(new Phrase("Data", new Font(Font.FontFamily.HELVETICA, 10)));
                //Adicionando as caracteristicas como cor, tamanho etc
                cell1.BackgroundColor = BaseColor.LIGHT_GRAY;
                cell1.Border = Rectangle.BOTTOM_BORDER | Rectangle.TOP_BORDER | Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER;
                cell1.BorderWidthBottom = 1;
                cell1.BorderWidthTop = 1;
                cell1.BorderWidthLeft = 1;
                cell1.BorderWidthRight = 1;
                cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                cell1.VerticalAlignment = Element.ALIGN_CENTER;
                table.AddCell(cell1);

                PdfPCell cell2 = new PdfPCell(new Phrase("Nome", new Font(Font.FontFamily.HELVETICA, 10)));

                cell2.BackgroundColor = BaseColor.LIGHT_GRAY;
                cell2.Border = Rectangle.BOTTOM_BORDER | Rectangle.TOP_BORDER | Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER;
                cell2.BorderWidthBottom = 1;
                cell2.BorderWidthTop = 1;
                cell2.BorderWidthLeft = 1;
                cell2.BorderWidthRight = 1;
                cell2.HorizontalAlignment = Element.ALIGN_CENTER;
                cell2.VerticalAlignment = Element.ALIGN_CENTER;
                table.AddCell(cell2);

                PdfPCell cell3 = new PdfPCell(new Phrase("Sobrenome", new Font(Font.FontFamily.HELVETICA, 10)));

                cell3.BackgroundColor = BaseColor.LIGHT_GRAY;
                cell3.Border = Rectangle.BOTTOM_BORDER | Rectangle.TOP_BORDER | Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER;
                cell3.BorderWidthBottom = 1;
                cell3.BorderWidthTop = 1;
                cell3.BorderWidthLeft = 1;
                cell3.BorderWidthRight = 1;
                cell3.HorizontalAlignment = Element.ALIGN_CENTER;
                cell3.VerticalAlignment = Element.ALIGN_CENTER;
                table.AddCell(cell3);

                PdfPCell cell4 = new PdfPCell(new Phrase("Endereço", new Font(Font.FontFamily.HELVETICA, 10)));

                cell4.BackgroundColor = BaseColor.LIGHT_GRAY;
                cell4.Border = Rectangle.BOTTOM_BORDER | Rectangle.TOP_BORDER | Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER;
                cell4.BorderWidthBottom = 1;
                cell4.BorderWidthTop = 1;
                cell4.BorderWidthLeft = 1;
                cell4.BorderWidthRight = 1;
                cell4.HorizontalAlignment = Element.ALIGN_CENTER;
                cell4.VerticalAlignment = Element.ALIGN_CENTER;
                table.AddCell(cell4);

                //Fazendo um loop de 100 linhas para exemplo e adicionando todas nas linhas de cada coluna
                for (int i = 0; i < 100; i++)
                {
                    PdfPCell cell_1 = new PdfPCell(new Phrase(i.ToString()));
                    PdfPCell cell_2 = new PdfPCell(new Phrase(i.ToString()));
                    PdfPCell cell_3 = new PdfPCell(new Phrase(i.ToString()));
                    PdfPCell cell_4 = new PdfPCell(new Phrase(i.ToString()));

                    cell_1.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell_2.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell_3.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell_4.HorizontalAlignment = Element.ALIGN_CENTER;

                    table.AddCell(cell_1);
                    table.AddCell(cell_2);
                    table.AddCell(cell_3);
                    table.AddCell(cell_4);
                }
                document.Add(table);
                document.Close();
                writer.Close();
                var constant = ms.ToArray();
                return File(constant, "application/vnd", "PrimeiroPDF.pdf"); //Coloando como vai ser o nome do arquivo PDF
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}