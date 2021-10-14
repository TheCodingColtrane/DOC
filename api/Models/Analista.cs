using System;
using System.Collections.Generic;

#nullable disable

namespace DOC.Models
{
    public partial class Analista
    {
        public Analista()
        {
            DocumentoAvalidars = new HashSet<DocumentoAvalidar>();
        }

        public int AnalistaId { get; set; }
        public int CelulaId { get; set; }
        public string Nome { get; set; }
        public short Cargo { get; set; }
        public bool Eliderenca { get; set; }
        public string Lideranca { get; set; }
        public string Email { get; set; }
        public short CargoComplexidade { get; set; }
        public bool Elocal { get; set; } 

        public virtual Celula Celula { get; set; }
        public virtual ICollection<DocumentoAvalidar> DocumentoAvalidars { get; set; }
    }
}

