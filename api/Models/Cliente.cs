using System;
using System.Collections.Generic;

#nullable disable

namespace DOC.Models
{
    public partial class Cliente
    {
        public Cliente()
        {
            DocumentoAvalidars = new HashSet<DocumentoAvalidar>();
            Slas = new HashSet<Sla>();
        }

        public int ClienteId { get; set; }
        public int CelulaId { get; set; }
        public string Nome { get; set; }
        public byte Tipo { get; set; }

        public virtual Celula Celula { get; set; }
        public virtual ICollection<DocumentoAvalidar> DocumentoAvalidars { get; set; }
        public virtual ICollection<Sla> Slas { get; set; }
    }
}
