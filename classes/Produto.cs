using System;
using System.IO;
using NetOffice.ExcelApi;

namespace Produtos.classes

{
    public class Produto
    {
        public string nome;
        public string descricao;
        public decimal preco;
        public int quantidade;
        public Categoria categoria;
        public Fornecedor fornecedor;

        public string cadastro(){
           Application ex = new Application();
           FileInfo teste = new FileInfo(@"c:\Danilo\produtos.xlsx");     
           if(teste.Exists){

               ex.Visible = true;
               ex.Workbooks.Open(@"c:\Danilo\produtos.xlsx");

               for (int p =2; p <=50;p++)
               {
                   if(ex.Range("a"+p).Value == null)
                   {
                      ex.Range("A"+p).Value = nome;
                      ex.Range("B"+p).Value = descricao;
                      ex.Range("C"+p).Value = preco;
                      ex.Range("d"+p).Value = quantidade;
                      ex.Range("e"+p).Value = categoria.nome;
                      ex.Range("F"+p).Value = categoria.descricao;
                      ex.Range("G"+p).Value = fornecedor.razaosocial;
                      ex.Range("H"+p).Value = fornecedor.nomefantasia;
                      ex.Range("I"+p).Value = fornecedor.CNPJ;

                      break; 
                   }
               }

               ex.ActiveWorkbook.Save();
               ex.Quit();
           } 
           else{
               ex.Visible = true;
               ex.Workbooks.Add();

                ex.Range("A1").Value = "Nome do produto"; 
                ex.Range("B1").Value = "Descrição do produto"; 
                ex.Range("C1").Value = "Preço"; 
                ex.Range("D1").Value = "Quantidade"; 
                ex.Range("E1").Value = "Categoria"; 
                ex.Range("F1").Value = "Descriçao cat";
                ex.Range("G1").Value = "Razão Social"; 
                ex.Range("H1").Value = "Nome Fantasia";   
                ex.Range("I1").Value = "CNPJ";   

                ex.Range("A1:I1").Font.Name = "Tahoma";
                ex.Range("A1:I1").Font.Bold = true;
                ex.Range("A1:I1").Font.Size = 13;

                 ex.Range("A2").Value = nome;
                 ex.Range("B2").Value = descricao;
                 ex.Range("C2").Value = preco;
                 ex.Range("d2").Value = quantidade;
                 ex.Range("e2").Value = categoria.nome;
                 ex.Range("F2").Value = categoria.descricao;
                 ex.Range("G2").Value = fornecedor.razaosocial;
                 ex.Range("H2").Value = fornecedor.nomefantasia;
                 ex.Range("I2").Value = fornecedor.CNPJ;

                ex.ActiveWorkbook.SaveAs(@"c:\Danilo\produtos.xlsx");
                ex.Quit(); 
           }
            return "Produto cadastrado!";

        }     
        public string[,] listar()
         {
        //Vamos construir uma matriz de string para guardar os dados dos clientes
            string[,] dados = new string[10,10];

            Application excel = new Application();
            excel.Visible = true;
            excel.Workbooks.Open(@"c:\Danilo\produtos.xlsx");      
            for(int lin = 1; lin <= 10; lin++){
              for(int col = 1; col <=10; col++){
                  dados[lin-1,col-1] = excel.Cells[lin,col].Text.ToString();

              }
            } 
                excel.Quit();
            return dados;
        }
    }
}