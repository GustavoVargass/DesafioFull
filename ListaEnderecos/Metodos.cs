using ClosedXML.Excel;
using ListaEnderecos.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ListaEnderecos
{
    class Metodos
    {
        public static Endereco BuscarCEP(string cep)
        {
            Endereco endereco = new Endereco();

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://viacep.com.br/ws/" + cep + "/json/");

            request.AllowAutoRedirect = false;
            HttpWebResponse servidor = (HttpWebResponse)request.GetResponse();

            if (servidor.StatusCode != HttpStatusCode.OK)
            {
                Console.WriteLine("Servidor indisponível");
                return null;
            }

            using (Stream webStream = servidor.GetResponseStream())
            {
                if (webStream != null)
                {
                    using (StreamReader responseReader = new StreamReader(webStream))
                    {
                        string response = responseReader.ReadToEnd();
                        response = Regex.Replace(response, "[{},]", string.Empty);
                        response = response.Replace("\"", "");

                        String[] substrings = response.Split('\n');

                        int cont = 0;
                        foreach (var substring in substrings)
                        {
                            if (cont == 1)
                            {
                                string[] valor = substring.Split(":".ToCharArray());
                                endereco.CEP = valor[1];
                                if (valor[0] == "  erro")
                                {
                                    //Console.WriteLine(cep + " - CEP não encontrado");
                                    return null;
                                }
                            }

                            //Logradouro
                            if (cont == 2)
                            {
                                string[] valor = substring.Split(":".ToCharArray());
                                endereco.Logradouro = valor[1];
                            }

                            //Bairro
                            if (cont == 4)
                            {
                                string[] valor = substring.Split(":".ToCharArray());
                                endereco.Bairro = valor[1];
                            }

                            //Cidade
                            if (cont == 5)
                            {
                                string[] valor = substring.Split(":".ToCharArray());
                                endereco.Cidade = valor[1];
                            }

                            //Estado
                            if (cont == 6)
                            {
                                string[] valor = substring.Split(":".ToCharArray());
                                endereco.UF = valor[1];
                            }

                            cont++;
                        }
                    }
                }
                return endereco;
            }
        }

        public static void GeraPlanilhaResultado(List<Endereco> enderecos)
        {
            Console.WriteLine("Gerando arquivo...");

            var arquivo = new XLWorkbook();
            var planilha = arquivo.Worksheets.Add("Resultado");

            //Cabeçalho
            planilha.Cell("A1").Value = "CEP";
            planilha.Cell("B1").Value = "Logradouro";
            planilha.Cell("C1").Value = "Bairro/Distrito";
            planilha.Cell("D1").Value = "Localidade/UF";
            planilha.Cell("E1").Value = "Data/Hora processamento";

            var linha = 2;

            foreach (Endereco end in enderecos)
            {
                planilha.Cell("A" + linha.ToString()).Value = end.CEP;
                planilha.Cell("B" + linha.ToString()).Value = end.Logradouro;
                planilha.Cell("C" + linha.ToString()).Value = end.Bairro;
                planilha.Cell("D" + linha.ToString()).Value = end.Cidade + "/" + end.UF;
                planilha.Cell("E" + linha.ToString()).Value = DateTime.Now;
                linha++;
            }

            //Ajustar tamanho da coluna, conforme conteúdo da célula
            planilha.Columns("1-5").AdjustToContents();

            //Salvar arquivo
            arquivo.SaveAs("../../../Planilha/resultado.xlsx");

            arquivo.Dispose();

            Console.WriteLine("Criado arquivo com o resultado!\n");
            Console.WriteLine("Aperte qualquer tecla para encerrar!");
            Console.ReadKey();
        }
    }
}
