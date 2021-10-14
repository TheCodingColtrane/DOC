using System;
using System.Collections.Generic;

#nullable disable

namespace DOC.Models
{
    public partial class Documento
    {
        public Documento()
        {
            DocumentoAvalidars = new HashSet<DocumentoAvalidar>();
        }

        public int DocumentoId { get; set; }
        public string Nome { get; set; }
        public byte? PrazoMaximoAnalise { get; set; }
        public byte Complexidade { get; set; }
        public byte Tipo { get; set; }
        public int Slaid { get; set; }

        [System.Text.Json.Serialization.JsonIgnore]
        public TimeSpan TempoMedioAnalise { get; set; }
       
        public virtual Sla Sla { get; set; }
        public virtual ICollection<DocumentoAvalidar> DocumentoAvalidars { get; set; }

        [System.ComponentModel.DataAnnotations.Schema.NotMapped]
        public DateTime TempoMedioAnaliseBruto { get; set; }

        [System.ComponentModel.DataAnnotations.Schema.NotMapped]
        public string Cliente { get; set; }

    }
}
