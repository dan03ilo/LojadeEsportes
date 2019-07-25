using System;
using NetOffice.ExcelApi;
using Produtos.classes;

namespace Produtos
{
    class Program
    {
        static void Main(string[] args)
        {
            Produto pro = new Produto();

            Fornecedor fr = new Fornecedor();
            fr.razaosocial = "DD New Sports";
            fr.nomefantasia = "DNS";
            fr.CNPJ = "61.258.468/8451-96";

            Categoria cat = new Categoria();
            cat.nome = "Bola de basketbaal";
            cat.descricao = "Bola Spalding";

            pro.nome="Bola Spalding";
            pro.descricao="Bola de basquete";
            pro.preco=39.99M;
            pro.quantidade=1;
            pro.categoria=cat;
            pro.fornecedor=fr;
            Console.WriteLine(pro.cadastro());



        //      string[,] info = pro.listar();
        //         for(int i = 0;i < 10;i++){
        //             for(int x = 0;x < 10;x++){
        //                 Console.Write(info[i,x]+"\t");
        //             }
        //             Console.WriteLine();
        //  }  
         }
    } 
} 
