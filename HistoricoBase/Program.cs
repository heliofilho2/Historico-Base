using System;
using System.Configuration;
using System.Linq;

namespace HistoricoBases
{
    internal static class Program
    {
        /// <summary>
        /// Ponto de entrada principal para o aplicativo.
        /// </summary>
        static void Main()
        {
            // Executa a query na base genereic e armazena os resultados
            histBases.QueryView();

            // Itera sobre as bases configuradas e processa cada uma delas
            foreach (var @base in ConfigurationManager.AppSettings.Get("Bases").Split(';').ToList())
            {
                try
                {
                    // Cria uma instância da classe Magic para a base atual
                    histBases _magic = new histBases(@base);

                    // Processa a base atual (conecta, atualiza e desconecta)
                    _magic.ProcessBases();
                }
                catch (Exception ex)
                {
                    // Captura exceções e exibe a mensagem de erro
                    Console.WriteLine($"Erro no processamento da base {@base}: {ex.Message}");

                }
            }
        }
    }
}