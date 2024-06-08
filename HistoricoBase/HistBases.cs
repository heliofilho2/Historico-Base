using System;
using SAPbobsCOM;
using System.Configuration;
using System.Collections.Generic;
using System.IO;

namespace HistoricoBases
{
    // Classe principal responsável por gerenciar a conexão ao SAP e processar os lançamentos
    public class histBases
    {
        // Objeto que representa a empresa no SAP
        private SAPbobsCOM.Company cp = new Company();

        // Variável para armazenar a última mensagem de erro
        private string lastError = "";

        // Variável para indicar se está conectado ao SAP
        public bool conectado = false;

        // Nome da empresa para conexão
        public string _empresa;

        // Lista para armazenar os lançamentos retornados pela query
        private static List<Tuple<string, int, string, Dictionary<int, string>>> lancamentos = new List<Tuple<string, int, string, Dictionary<int, string>>>();

        // Construtor da classe, inicializa com o nome da empresa
        public histBases(string empresa)
        {
            _empresa = empresa;
        }

        // Método para conectar ao SAP
        private bool ConnectSap(string empresa)
        {
            try
            {
                // Configurações de conexão
                cp.DbServerType = BoDataServerTypes.dst_HANADB;
                cp.Server = ConfigurationManager.AppSettings.Get(empresa.EndsWith("2") ? "Server02" : "Server");
                cp.CompanyDB = ConfigurationManager.AppSettings.Get("DbName" + empresa);
                cp.DbUserName = ConfigurationManager.AppSettings.Get("DBUser" + empresa);
                cp.DbPassword = ConfigurationManager.AppSettings.Get("DBPwd" + empresa);
                cp.UserName = ConfigurationManager.AppSettings.Get("User" + empresa);
                cp.Password = ConfigurationManager.AppSettings.Get("Pwd" + empresa);
                cp.SLDServer = "LINKsld:PORT";
                cp.UseTrusted = true;

                // Tentativa de conexão
                Console.WriteLine("Conectando à empresa: " + empresa);
                var i = cp.Connect();

                if (i != 0)
                {
                    // Caso ocorra um erro na conexão, exibe o erro e retorna false
                    lastError = cp.GetLastErrorDescription();
                    Console.WriteLine("Erro - Conectar SAP: " + empresa + " " + lastError);
                    conectado = false;
                    return false;
                }
                else
                {
                    // Conexão bem-sucedida
                    Console.WriteLine("Conectado à empresa: " + empresa);
                    conectado = true;
                    return true;
                }
            }
            catch (Exception ex)
            {
                // Captura exceções e exibe a mensagem de erro
                Console.WriteLine("Erro: " + ex.Message);
                var auxEx = "Exception CntSAP --- " + ex.Message + "<br>Base: " + empresa + "<br> Stack Trace CntSAP --- " + ex.StackTrace + "<br> Source CntSAP ---" + ex.Source;
                Console.WriteLine(auxEx);
                conectado = false;
                return false;
            }
        }

        // Método para desconectar do SAP
        public bool DisconnectSap()
        {
            Console.WriteLine("Desconectando da empresa: " + _empresa);
            if (conectado)
            {
                cp.Disconnect();
                cp = new Company();
                conectado = false;
            }
            return true;
        }

        // Método para executar a query na base onde está a view e armazenar os resultados
        public static void QueryView()
        {
            // Cria um objeto temporário para a empresa SAP
            SAPbobsCOM.Company tempCompany = new Company();
            try
            {
                // Configurações de conexão para a base 
                tempCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                tempCompany.Server = ConfigurationManager.AppSettings.Get("Server");
                tempCompany.CompanyDB = ConfigurationManager.AppSettings.Get("DbName");
                tempCompany.DbUserName = ConfigurationManager.AppSettings.Get("DBUser");
                tempCompany.DbPassword = ConfigurationManager.AppSettings.Get("DBPwd");
                tempCompany.UserName = ConfigurationManager.AppSettings.Get("User");
                tempCompany.Password = ConfigurationManager.AppSettings.Get("Pwd");
                tempCompany.SLDServer = "linkSLD:PORT";
                tempCompany.UseTrusted = true;

                // Tentativa de conexão à base Generic
                Console.WriteLine("Conectando à base Generic para executar a query");
                var i = tempCompany.Connect();

                if (i != 0)
                {
                    // Caso ocorra um erro na conexão, lança uma exceção
                    throw new Exception("Erro ao conectar à base Generic: " + tempCompany.GetLastErrorDescription());
                }

                // Executa a query para obter os lançamentos
                Recordset recordset = (Recordset)tempCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                recordset.DoQuery("SELECT \"Base\", \"ParentKey\", \"Line_ID\", \"Memo\", \"LineMemo\" FROM \"Generic".\"Obs_Unificada\"");

                // Itera sobre os resultados da query e armazena na lista de lançamentos
                while (!recordset.EoF)
                {
                    string baseName = recordset.Fields.Item("Base").Value.ToString();
                    int parentKey = int.Parse(recordset.Fields.Item("ParentKey").Value.ToString());
                    string memo = recordset.Fields.Item("Memo").Value.ToString();
                    int lineId = int.Parse(recordset.Fields.Item("Line_ID").Value.ToString());
                    string lineMemo = recordset.Fields.Item("LineMemo").Value.ToString();

                    // Verifica se o lançamento já existe na lista
                    var lancamento = lancamentos.Find(x => x.Item1 == baseName && x.Item2 == parentKey);
                    if (lancamento == null)
                    {
                        // Adiciona um novo lançamento à lista
                        lancamento = new Tuple<string, int, string, Dictionary<int, string>>(baseName, parentKey, memo, new Dictionary<int, string>());
                        lancamentos.Add(lancamento);
                    }
                    lancamento.Item4[lineId] = lineMemo;

                    recordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
                // Captura exceções e exibe a mensagem de erro
                Console.WriteLine($"Erro ao executar QueryView: {ex.Message}");
            }
            finally
            {
                // Desconecta da base Generic e libera recursos
                if (tempCompany.Connected)
                {
                    tempCompany.Disconnect();
                }
                tempCompany = null;
            }
        }

        // Método para processar as bases específicas
        public void ProcessBases()
        {
            // Conecta-se à base específica para atualizar os lançamentos
            if (ConnectSap(_empresa))
            {
                // Chama o método para atualizar os lançamentos
                UpdateLc();
                // Desconecta da base
                DisconnectSap();
            }
        }

        // Método para atualizar os lançamentos na base específica
        private void UpdateLc()
        {
            // Lista para armazenar os IDs dos lançamentos atualizados
            List<(int, string)> updatedLcs = new List<(int, string)>();

            // Lista para armazenar os erros detalhados
            List<string> errors = new List<string>();

            if (conectado)
            {
                try
                {
                    // Objeto para gerenciar os lançamentos no SAP
                    JournalEntries lc = (JournalEntries)cp.GetBusinessObject(BoObjectTypes.oJournalEntries);

                    // Itera sobre os lançamentos armazenados
                    foreach (var lancamento in lancamentos)
                    {
                        // Verifica se o lançamento pertence à base atual
                        if (lancamento.Item1 == _empresa)
                        {
                            // Verifica se o lançamento existe no SAP
                            if (lc.GetByKey(lancamento.Item2))
                            {
                                // Atualiza o memo do cabeçalho do lançamento
                                lc.Memo = lancamento.Item3;

                                // Itera sobre as linhas do lançamento
                                for (int i = 0; i < lc.Lines.Count; i++)
                                {
                                    lc.Lines.SetCurrentLine(i);
                                    int lineId = lc.Lines.Line_ID;

                                    // Atualiza o memo da linha, se existir
                                    if (lancamento.Item4.ContainsKey(lineId))
                                    {
                                        lc.Lines.LineMemo = lancamento.Item4[lineId];
                                    }
                                }

                                // Tenta atualizar o lançamento no SAP
                                if (lc.Update() != 0)
                                {
                                    // Se houver erro, adiciona à lista de erros
                                    errors.Add($"Erro ao atualizar lançamento {lancamento.Item2}: {cp.GetLastErrorDescription()}");
                                }
                                else
                                {
                                    // Adiciona o ID do lançamento atualizado à lista
                                    updatedLcs.Add((lancamento.Item2, _empresa));
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    // Captura exceções e exibe a mensagem de erro
                    Console.WriteLine("Erro ao atualizar LC: " + e.Message);
                    errors.Add($"Erro ao atualizar LC: {e.Message}");
                }
            }

            // Escreve os resultados em arquivos de log
            WriteLogs(updatedLcs, errors);
        }


        // Método para escrever os resultados em arquivos de log
        private void WriteLogs(List<(int, string)> updatedLcs, List<string> errors)
        {
            // Define o nome dos arquivos de log
            string successLogFile = "Logs/success_log.txt";
            string errorLogFile = "Logs/error_log.txt";

            try
            {
                // Abre ou cria os arquivos de log e define o modo de escrita para sobrescrever o conteúdo existente
                using (StreamWriter successWriter = new StreamWriter(successLogFile, true))
                using (StreamWriter errorWriter = new StreamWriter(errorLogFile, true))
                {
                    // Escreve os lançamentos atualizados com sucesso no arquivo de log de sucesso
                    foreach (var (id, baseLC) in updatedLcs)
                    {
                        successWriter.WriteLine($"Lançamento atualizado  com sucesso: {id}, Base: {baseLC}");
                    }

                    // Escreve os erros no arquivo de log de erros
                    foreach (string error in errors)
                    {
                        errorWriter.WriteLine($"Erro durante a atualização dos LC's: {error}");
                    }
                }

                // Exibe mensagem de sucesso
                Console.WriteLine("Registros de atualização e erros foram escritos nos arquivos de log.");
            }
            catch (Exception ex)
            {
                // Em caso de erro ao escrever nos arquivos de log, exibe a mensagem de erro
                Console.WriteLine($"Erro ao escrever nos arquivos de log: {ex.Message}");
            }
        }

    }
}   