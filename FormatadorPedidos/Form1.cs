using ClosedXML.Excel;

namespace FormatadorPedidos
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Arquivos excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*",
                Title = "Abrir arquivo do Excel"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = openFileDialog.FileName;
            }
        }

        private void btnSaveFile_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtFilePath.Text))
            {
                try
                {
                    ProcessAndSaveFile(txtFilePath.Text);
                    MessageBox.Show("Arquivo formatado e salvo com sucesso!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ocorreu um erro: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Por favor selecione um arquivo primeiro");
            }
        }

        private void ProcessAndSaveFile(string filePath)
        {
            using var workbook = new XLWorkbook(filePath);
            var inputWorksheet = workbook.Worksheet(1);

            var cabecalhoPedido = GetHeadersDataOfPedido(inputWorksheet);

            var formattedWorksheet = workbook.AddWorksheet("Formatado");

            WriteHeader(cabecalhoPedido, formattedWorksheet);

            MovePositionOfProducts(inputWorksheet, formattedWorksheet);

            MoveProductsGrid(inputWorksheet, formattedWorksheet);

            AddDataFaturamentoColumm(formattedWorksheet);

            AddStyle(formattedWorksheet);

            string formattedFilePath = filePath.Replace(".xlsx", "_formatado.xlsx");

            try
            {
                workbook.SaveAs(formattedFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha ao salvar o arquivo. Erro: " + ex.Message);
            }
        }

        private static void AddDataFaturamentoColumm(IXLWorksheet formattedWorksheet)
        {
            formattedWorksheet.Cell(Constantes.A5).Value = Constantes.DataFaturamentoCabecalho;
        }

        private void AddStyle(IXLWorksheet formattedWorksheet)
        {
            EstilizarTabelaProdutos(formattedWorksheet);

            EstilizarGradesProdutos(formattedWorksheet);

        }

        private void EstilizarGradesProdutos(IXLWorksheet formattedWorksheet)
        {
            var lastUsedColumn = formattedWorksheet.Row(1).LastCellUsed();
            var range = formattedWorksheet.Range($"J1:{lastUsedColumn}");
            range.Style.Fill.SetPatternColor(XLColor.FromTheme(XLThemeColor.Accent4, 0.2));

            lastUsedColumn = formattedWorksheet.Row(2).LastCellUsed();
            range = formattedWorksheet.Range($"J2:{lastUsedColumn}");
            range.Style.Fill.SetPatternColor(XLColor.FromTheme(XLThemeColor.Accent3, 0.2));

            lastUsedColumn = formattedWorksheet.Row(3).LastCellUsed();
            range = formattedWorksheet.Range($"J3:{lastUsedColumn}");
            range.Style.Fill.SetPatternColor(XLColor.FromTheme(XLThemeColor.Accent2, 0.2));

            lastUsedColumn = formattedWorksheet.Row(4).LastCellUsed();
            range = formattedWorksheet.Range($"J4:{lastUsedColumn}");
            range.Style.Fill.SetPatternColor(XLColor.FromTheme(XLThemeColor.Accent1, 0.2));
        }

        private void EstilizarTabelaProdutos(IXLWorksheet formattedWorksheet)
        {
            var lastCellUsed = formattedWorksheet.LastColumnUsed()?.LastCellUsed();

            if (lastCellUsed != null)
            {
                var firstRowColumn = "A5";
                var lastRowColumn = lastCellUsed.Address.ToString();

                var rangeOfCells = $"{firstRowColumn}:{lastRowColumn}";
                var range = formattedWorksheet.Range(rangeOfCells);

                if (!range.IsEmpty() && range.CellsUsed().Count() > 0)
                {
                    try
                    {
                        ConvertHeadersToString(range);

                        var table = range.CreateTable("ListaProdutos");
                        table.Theme = XLTableTheme.TableStyleMedium9;
                        table.ShowHeaderRow = true;
                        table.ShowRowStripes = true;
                        table.ShowAutoFilter = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Falha ao criar a tabela. Erro: " + ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("A seleção de células utilizadas inclui células vazias ou estão em formato incorreto.");
                }
            }
            else
            {
                MessageBox.Show("Nenhum dado encontrado para a criação da tabela.");
            }
        }

        private void ConvertHeadersToString(IXLRange range)
        {
            foreach (var cell in range.Cells())
            {
                if(cell.Address.RowNumber == 5)
                {
                    if(cell.DataType != XLDataType.Text)
                    {
                        cell.Value = cell.Value.ToString();
                    }
                }
            }

        }

        private static void MoveProductsGrid(IXLWorksheet inputWorksheet, IXLWorksheet formattedWorksheet)
        {
            var rangeOfGrade = inputWorksheet.Range(Constantes.GradeProdutosIntervaloCelulas);
            var targetOfCopy = formattedWorksheet.Cell(Constantes.CelulaInicialGradeProdutos);
            rangeOfGrade.CopyTo(targetOfCopy);
        }

        private static void WriteHeader(CabecalhoPedido cabecalhoPedido, IXLWorksheet formattedWorksheet)
        {
            formattedWorksheet.Cell(Constantes.A1).Value = cabecalhoPedido.Nome.Keys.First();
            formattedWorksheet.Cell(Constantes.A2).Value = cabecalhoPedido.Nome.Values.First();
            formattedWorksheet.Cell(Constantes.D1).Value = cabecalhoPedido.Razao.Keys.First();
            formattedWorksheet.Cell(Constantes.D2).Value = cabecalhoPedido.Razao.Values.First();
            formattedWorksheet.Cell(Constantes.G1).Value = cabecalhoPedido.CNPJ.Keys.First();
            formattedWorksheet.Cell(Constantes.G2).Value = cabecalhoPedido.CNPJ.Values.First();
            formattedWorksheet.Cell(Constantes.A3).Value = cabecalhoPedido.Po.Keys.First();
            formattedWorksheet.Cell(Constantes.A4).Value = cabecalhoPedido.Po.Values.First();
            formattedWorksheet.Cell(Constantes.B3).Value = cabecalhoPedido.Customer.Keys.First();
            formattedWorksheet.Cell(Constantes.B4).Value = cabecalhoPedido.Customer.Values.First();
            formattedWorksheet.Cell(Constantes.C3).Value = cabecalhoPedido.Store.Keys.First();
            formattedWorksheet.Cell(Constantes.C4).Value = cabecalhoPedido.Store.Values.First();
            formattedWorksheet.Cell(Constantes.D3).Value = cabecalhoPedido.CondicaoPagamento.Keys.First();
            formattedWorksheet.Cell(Constantes.D4).Value = cabecalhoPedido.CondicaoPagamento.Values.First();
            formattedWorksheet.Cell(Constantes.F3).Value = cabecalhoPedido.QuantidadePecas.Keys.First();
            formattedWorksheet.Cell(Constantes.F4).Value = cabecalhoPedido.QuantidadePecas.Values.First();
            formattedWorksheet.Cell(Constantes.G3).Value = cabecalhoPedido.Total.Keys.First();
            formattedWorksheet.Cell(Constantes.G4).Value = cabecalhoPedido.Total.Values.First();
        }

        private static void MovePositionOfProducts(IXLWorksheet inputWorksheet, IXLWorksheet formattedWorksheet)
        {
            var lastCellUsed = inputWorksheet.Column("AN").LastCellUsed();

            if (lastCellUsed != null)
            {
                var lastRow = lastCellUsed.Address.RowNumber;

                var range = inputWorksheet.Range(string.Format(Constantes.ListaProdutosIntervaloCelulas, lastRow));
                var targetOfCopy = formattedWorksheet.Cell(Constantes.AlvoListaProdutos);
                range.CopyTo(targetOfCopy);
            }
        }

        private CabecalhoPedido GetHeadersDataOfPedido(IXLWorksheet inputWorksheet)
        {
            return new CabecalhoPedido
            {
                Po = GetLabelAndValueFromStringCells(inputWorksheet, Constantes.Po, Constantes.PoValor),
                Razao = GetLabelAndValueFromStringCells(inputWorksheet, Constantes.RazaoSocial, Constantes.RazaoSocialValor),
                Nome = GetLabelAndValueFromStringCells(inputWorksheet, Constantes.Nome, Constantes.NomeValor),
                Customer = GetLabelAndValueFromIntCells(inputWorksheet, Constantes.Customer, Constantes.CustomerValor),
                Store = GetLabelAndValueFromStringCells(inputWorksheet, Constantes.Store, Constantes.StoreValor),
                CNPJ = GetLabelAndValueFromStringCells(inputWorksheet, Constantes.Cnpj, Constantes.CnpjValor),
                CondicaoPagamento = GetLabelAndValueFromStringCells(inputWorksheet, Constantes.CondicaoPagamento, Constantes.CondicaoPagamentoValor),
                QuantidadePecas = GetLabelAndValueFromIntCells(inputWorksheet, Constantes.QuantidadePecas, Constantes.QuantidadePecasValor),
                Total = GetLabelAndValueFromDoubleCells(inputWorksheet, Constantes.Total, Constantes.TotalValor)
            };
        }


        private Dictionary<string, string> GetLabelAndValueFromStringCells(IXLWorksheet inputWorksheet, string label, string value)
        {
            var dic = new Dictionary<string, string>();
            var labelString = GetStringFromCell(inputWorksheet, label);
            var valueString = GetStringFromCell(inputWorksheet, value);
            dic.Add(labelString, valueString);
            return dic;
        }

        private Dictionary<string, double> GetLabelAndValueFromDoubleCells(IXLWorksheet inputWorksheet, string label, string value)
        {
            var dic = new Dictionary<string, double>();
            var labelString = GetStringFromCell(inputWorksheet, label);
            var valueDouble = GetDoubleFromCell(inputWorksheet, value);
            dic.Add(labelString, valueDouble);
            return dic;
        }

        private Dictionary<string, int> GetLabelAndValueFromIntCells(IXLWorksheet inputWorksheet, string label, string value)
        {
            var dic = new Dictionary<string, int>();
            var labelString = GetStringFromCell(inputWorksheet, label);
            var valueInt = GetIntFromCell(inputWorksheet, value);
            dic.Add(labelString, valueInt);
            return dic;
        }

        private string GetStringFromCell(IXLWorksheet inputWorksheet, string cell)
        {
            string saida;
            inputWorksheet.Cell(cell).Value.TryGetText(out saida);
            return saida;
        }

        private double GetDoubleFromCell(IXLWorksheet inputWorksheet, string cellPosition)
        {
            var cell = inputWorksheet.Cell(cellPosition);
            if (!cell.IsEmpty())
            {
                double value;
                cell.TryGetValue(out value);
                return value;
            }
            return 0;
        }

        private int GetIntFromCell(IXLWorksheet inputWorksheet, string cell)
        {
            return inputWorksheet.Cell(cell).GetValue<int>();
        }

    }
}
