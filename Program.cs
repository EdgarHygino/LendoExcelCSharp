using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        //passa o path do arquivo
        var path = "C:\\Users\\edgar\\OneDrive\\Documentos\\ProjetosYutube\\LendoExcelC#\\ListaDeCompra.xlsx";

        var xls = new XLWorkbook(path);
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
            Produto novoProduto = new Produto
            {
                Nome = produto,
                Preco = preco
            };

            listaDeProdutos.Add(novoProduto);

            Console.WriteLine($"{produto} - {preco}");
        }

    }
}

public class Produto
{
    public string Nome { get; set; }
    public decimal Preco { get; set; }
}
