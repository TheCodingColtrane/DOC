using System;
using System.Collections.Generic;

#nullable disable

namespace DOC.Models
{
    public partial class Sla
    {
        public Sla()
        {
            Documentos = new HashSet<Documento>();
        }

        public int Slaid { get; set; }
        public int ClienteId { get; set; }

        public virtual Cliente Cliente { get; set; }
        public virtual ICollection<Documento> Documentos { get; set; }
    }
}
