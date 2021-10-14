using System;
using System.Collections.Generic;

#nullable disable

namespace DOC.Models
{
    public partial class DocumentoAvalidar
    {
        public int ProtocoloId { get; set; }
        public int ClienteId { get; set; }
        public int DocumentoId { get; set; }
        public int AnalistaId { get; set; }
        public string Fornecedor { get; set; }
        public string ClienteUnidade { get; set; }
        public DateTime MesDeposito { get; set; }
        public string Empregado { get; set; }
        public DateTime DataInicio { get; set; }
        public DateTime DataInclusao { get; set; }
        public int TempoEmAnalise { get; set; }
        public bool Divida { get; set; }
        public DateTime InicioInadimplencia { get; set; }
        public DateTime FimInadimplencia { get; set; }
        public int Qlp { get; set; }

        public virtual Analista Analista { get; set; }
        public virtual Cliente Cliente { get; set; }
        public virtual Documento Documento { get; set; }
    }
}
