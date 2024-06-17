namespace FormatadorPedidos
{
    public class CabecalhoPedido
    {
        public Dictionary<string, string> Po { get; set; }
        public Dictionary<string, string> Razao { get; set; }
        public Dictionary<string, string> Nome { get; set; }
        public Dictionary<string, int> Customer { get; set; }
        public Dictionary<string, string> Store { get; set; }
        public Dictionary<string, string> CNPJ { get; set; }
        public Dictionary<string, double> DescontoNacional { get; set; }
        public Dictionary<string, double> DescontoImportado { get; set; }
        public Dictionary<string, string> Representante { get; set; }
        public Dictionary<string, string> CondicaoPagamento { get; set; }
        public Dictionary<string, int> QuantidadePecas { get; set; }
        public Dictionary<string, double> Total { get; set; }

    }
}
