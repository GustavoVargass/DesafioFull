using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ListaEnderecos.Model;

namespace ListaEnderecos
{
    class Program
    {

        static void Main(string[] args)
        {
            List<Endereco> enderecos = new List<Endereco>();
            
            Console.WriteLine(" --- Iniciando processamento de arquivo com a lista de CEP's ---\n");

            var listaCeps = new XLWorkbook("../../../Planilha/ListaCEPS.xlsx");
            var planilha = listaCeps.Worksheet(1);

            var linha = 2;
            while (true)
            {
                var faixaCeps = planilha.Cell("A" + linha.ToString()).Value.ToString();
                var cepInicial = planilha.Cell("B" + linha.ToString()).Value.ToString();
                var cepFinal = planilha.Cell("C" + linha.ToString()).Value.ToString();

                if (string.IsNullOrEmpty(faixaCeps)) break;

                for (int i = Convert.ToInt32(cepInicial); i <= Convert.ToInt32(cepFinal); i++)
                {
                    Endereco enderecoEncontrado = Metodos.BuscarCEP(i.ToString());
                    if (enderecoEncontrado != null)
                        enderecos.Add(enderecoEncontrado);
                }
                linha++;
            }
            Metodos.GeraPlanilhaResultado(enderecos);
        }
    }
}
