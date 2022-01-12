using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp_GerarScriptInsert_Excel
{

    public class ConsoleApp_GerarScriptInsert_Excel
    {
        static void Main(string[] args)
        {
           var path = System.IO.Path.GetDirectoryName("C:\\Users\\neto_\\source\\repos\\Excel\\"); //Diretório do arquivo de entrada e saída

            var products = CreateProductsScript(path);
            System.IO.File.WriteAllText(System.IO.Path.Combine(path, "sugestoes.sql"), products); //faz o Spool do arquivo .sql, passando o caminho de saída, o nome do arquivo + extensão, e o dados.
            Console.WriteLine("Arquivo gerado com sucesso!!!");
            Console.ReadKey();
            //

        }

            public static string CreateProductsScript(string path)
            {

                var workbook = new XLWorkbook(path + "\\recomendation.xlsx"); //Concatena o caminho + o nome do seu arquivo Excel de entrada (não esquece da extensão do arquivo!)
                var ws1 = workbook.Worksheet("recomendation"); //Passa o nome da Sheet do seu arquivo que você vai querer manipular

                var rows = ws1.RowsUsed();

                var sb = new StringBuilder();

                int counter = 0;

                Console.WriteLine("Lendo o Excel...");

                foreach (var row in rows)
                {
                    if (row.RowNumber() > 1)
                    {
                        counter++;

                        var id = Convert.ToString(row.Cell(1).Value); //le a 1 coluna
                        var name = Convert.ToString(row.Cell(2).Value); //le a 2 coluna

                        // executing

                        sb.AppendLine("INSERT INTO suggestionbox (id,name) VALUES(" + id + ",'" + name + "');"); //concatena as variaveis com as informações pra montar o Insert.

                    }
                }

                return sb.ToString();
            }

            /*
            protected string Trim(string text)
            {
                if (string.IsNullOrEmpty(text))
                    return string.Empty;

                return text.Trim();
            }
            */   
    }
}