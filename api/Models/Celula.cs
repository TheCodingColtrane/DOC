using System;
using System.Collections.Generic;

#nullable disable

namespace DOC.Models
{
    public partial class Celula
    {
        public Celula()
        {
            Analista = new HashSet<Analista>();
            Clientes = new HashSet<Cliente>();
        }

        public int CelulaId { get; set; }
        public string Nome { get; set; }
        public byte Tipo { get; set; }

        public virtual ICollection<Analista> Analista { get; set; }
        public virtual ICollection<Cliente> Clientes { get; set; }
    }
}
