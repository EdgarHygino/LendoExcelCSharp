using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LendoExcelC_
{
    public class LidandoComExcecoes
    {
        public void LidandoComExcecoes1()
        {
            var path = "C:\\Users\\edgar\\OneDrive\\Documentos\\ProjetosYutube\\LendoExcelC#\\ListaDeCompra.xlsx";

            // Cria um objeto XLWorkbook
            var xls = new XLWorkbook(path);

            // Verifica se o arquivo existe
            if (!File.Exists(path))
            {
                // Se o arquivo não existir, imprime uma mensagem de erro
                Console.WriteLine("Erro: arquivo não encontrado.");
                return;
            }

            var planilha = xls.Worksheets.First(w => w.Name == "Planilha1");
            var totalLinhas = planilha.RowsUsed().Count();

            // Cria uma lista do objeto Produto
            List<Produto> listaDeProdutos = new List<Produto>();


            //Aqui criamos um loop para percorrer a planilha e extrairmos os seus valores.
            for (int l = 2; l <= totalLinhas; l++)
            {
                // Ler dados da planilha
                var produto = planilha.Cell($"A{l}").Value.ToString();
                var preco = Decimal.Parse(planilha.Cell($"B{l}").Value.ToString());

                // Criar objeto Produto e adicionar à lista de Produtos
                try
                {
                    Produto novoProduto = new Produto
                    {
                        Nome = produto,
                        Preco = preco
                    };

                    listaDeProdutos.Add(novoProduto);
                }
                catch (IndexOutOfRangeException e)
                {
                    Console.WriteLine("Erro: linha não encontrada.");
                }

                Console.WriteLine($"{produto} - {preco}");
            }

        }

        public void LidandoComExcecoes2()
        {
            // Caminho do arquivo
            var path = "C:\\Users\\edgar\\OneDrive\\Documentos\\ProjetosYutube\\LendoExcelC#\\ListaDeCompra.xlsx";

            // Verifica se o arquivo existe
            if (!File.Exists(path))
            {
                Console.WriteLine("Erro: Arquivo não encontrado.");
                return;
            }

            try
            {
                // Leitura do arquivo e criação do objeto XLWorkbook
                var xls = LerArquivoExcel(path);

                var planilha = xls.Worksheets.First(w => w.Name == "Planilha1");
                var totalLinhas = planilha.RowsUsed().Count();

                // Lista de produtos
                List<Produto> listaDeProdutos = new List<Produto>();

                // Loop para percorrer a planilha e extrair valores
                for (int l = 2; l <= totalLinhas; l++)
                {
                    // Ler dados da planilha
                    string produto = planilha.Cell($"A{l}").Value.ToString();
                    decimal preco = Decimal.Parse(planilha.Cell($"B{l}").Value.ToString());

                    // Criar objeto Produto e adicionar à lista de Produtos
                    Produto novoProduto = new Produto
                    {
                        Nome = produto,
                        Preco = preco
                    };

                    listaDeProdutos.Add(novoProduto);

                    Console.WriteLine($"{produto} - {preco}");
                }
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("Erro: Arquivo não encontrado.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro inesperado: {ex.Message}");
            }
        }

        private XLWorkbook LerArquivoExcel(string path)
        {
            try
            {
                return new XLWorkbook(path);
            }
            catch (Exception ex)
            {
                // Tratar exceções específicas se necessário
                Console.WriteLine($"Erro ao ler o arquivo Excel: {ex.Message}");
                throw;
            }
        }
    }
}
