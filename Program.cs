using ClosedXML.Excel;
using Polimorfico;
using System;
using System.IO;

namespace Polimorfico
{


    class ManipularExcel
    {

        private readonly string filePath;

        public ManipularExcel(string filePath)
        {
            this.filePath = filePath;
        }
        public void Exibir()
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();
                foreach (var row in rows)
                {
                    Console.WriteLine(row.Cell(1).Value + " " + row.Cell(2).Value + " " + row.Cell(3).Value + " " + row.Cell(4).Value + " " + row.Cell(5).Value);
                }
            }
        }
        public void Criar(Descricao novolivro)
        {
            try{
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                worksheet.Cell("A1").Value = novolivro.Codigo;
                worksheet.Cell("B1").Value = novolivro.Editora;
                worksheet.Cell("C1").Value = novolivro.Genero;
                worksheet.Cell("D1").Value = novolivro.Livro;
                worksheet.Cell("E1").Value = novolivro.Autor;
                workbook.Save();
            }} catch(Exception ex){

                Console.WriteLine($"Ocorreu um erro ao criar o livro: {ex.Message}");
            }

        }


        public void Adcionar(Descricao novolivro)
        {

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var lastRow = worksheet.LastRowUsed().RowNumber();
                    worksheet.Cell(lastRow + 1, 1).Value = novolivro.Codigo;
                    worksheet.Cell(lastRow + 1, 2).Value = novolivro.Editora;
                    worksheet.Cell(lastRow + 1, 3).Value = novolivro.Genero;
                    worksheet.Cell(lastRow + 1, 4).Value = novolivro.Livro;
                    worksheet.Cell(lastRow + 1, 5).Value = novolivro.Autor;
                    workbook.Save();
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine($"Ocorreu um erro ao criar o livro: {ex.Message}");
            }

        }
        // Exibe todos os valores da planilha

    }


}






class Program
{
    static void Main(string[] args)
    {

        string filePath = Path.Combine("arquivo", "titulo.xlsx");
        Descricao livro = new Descricao("codigo", "editora", "genero", "titulo", "autor");
        ManipularExcel bd = new ManipularExcel(filePath);
        bd.Criar(livro);

        Descricao livroN = new Descricao("132", "ABRIL", "FANTASIA", "HARRY POTER", "J.K");

        bd.Adcionar(livroN);

        bd.Exibir();


    }
}

