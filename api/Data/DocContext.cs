using System;
using Microsoft.EntityFrameworkCore;
using DOC.Models;
using Microsoft.EntityFrameworkCore.Metadata;

#nullable disable

namespace DOC.Data
{
    public partial class DocContext : DbContext
    {
        public DocContext()
        {
        }

        public DocContext(DbContextOptions<DocContext> options)
            : base(options)
        {
        }

        public virtual DbSet<Analista> Analista { get; set; }
        public virtual DbSet<Celula> Celulas { get; set; }
        public virtual DbSet<Cliente> Clientes { get; set; }
        public virtual DbSet<Documento> Documentos { get; set; }
        public virtual DbSet<DocumentoAvalidar> DocumentoAvalidars { get; set; }
        public virtual DbSet<Sla> Slas { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.HasAnnotation("Relational:Collation", "Latin1_General_CI_AS");

            modelBuilder.Entity<Analista>(entity =>
            {
                entity.HasKey(e => e.AnalistaId)
                    .HasName("PK__Analista__128665FA8CC77854");

                entity.Property(e => e.AnalistaId).HasColumnName("AnalistaID");

                entity.Property(e => e.CelulaId).HasColumnName("CelulaID");

                entity.Property(e => e.Eliderenca).HasColumnName("ELiderenca");

                entity.Property(e => e.Elocal).HasColumnName("ELocal");

                entity.Property(e => e.Email)
                    .IsRequired()
                    .HasMaxLength(500)
                    .IsUnicode(false);

                entity.Property(e => e.Lideranca)
                    .IsRequired()
                    .HasMaxLength(500)
                    .IsUnicode(false);

                entity.Property(e => e.Nome)
                    .IsRequired()
                    .HasMaxLength(500)
                    .IsUnicode(false);

                entity.HasOne(d => d.Celula)
                    .WithMany(p => p.Analista)
                    .HasForeignKey(d => d.CelulaId)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK__Analista__Celula__71D1E811");
            });

            modelBuilder.Entity<Celula>(entity =>
            {
                entity.ToTable("Celula");

                entity.Property(e => e.CelulaId).HasColumnName("CelulaID");

                entity.Property(e => e.Nome)
                    .IsRequired()
                    .HasMaxLength(500)
                    .IsUnicode(false);
            });

            modelBuilder.Entity<Cliente>(entity =>
            {
                entity.ToTable("Cliente");

                entity.Property(e => e.ClienteId).HasColumnName("ClienteID");

                entity.Property(e => e.CelulaId).HasColumnName("CelulaID");

                entity.Property(e => e.Nome)
                    .IsRequired()
                    .HasMaxLength(1000)
                    .IsUnicode(false);

                entity.HasOne(d => d.Celula)
                    .WithMany(p => p.Clientes)
                    .HasForeignKey(d => d.CelulaId)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK__Cliente__CelulaI__72C60C4A");
            });

            modelBuilder.Entity<Documento>(entity =>
            {
                entity.ToTable("Documento");

                entity.Property(e => e.DocumentoId).HasColumnName("DocumentoID");

                entity.Property(e => e.Nome)
                    .IsRequired()
                    .HasMaxLength(1000)
                    .IsUnicode(false);

                entity.Property(e => e.Slaid).HasColumnName("SLAID");

                entity.HasOne(d => d.Sla)
                    .WithMany(p => p.Documentos)
                    .HasForeignKey(d => d.Slaid)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK__Documento__SLAID__73BA3083");
            });

            modelBuilder.Entity<DocumentoAvalidar>(entity =>
            {
                entity.HasKey(e => e.ProtocoloId)
                    .HasName("PK__Document__C002F5A23007CAB8");

                entity.ToTable("DocumentoAValidar");

                entity.Property(e => e.ProtocoloId)
                    .ValueGeneratedNever()
                    .HasColumnName("ProtocoloID");

                entity.Property(e => e.AnalistaId).HasColumnName("AnalistaID");

                entity.Property(e => e.ClienteId).HasColumnName("ClienteID");

                entity.Property(e => e.ClienteUnidade)
                    .IsRequired()
                    .HasMaxLength(250)
                    .IsUnicode(false);

                entity.Property(e => e.DataInclusao).HasColumnType("datetime");

                entity.Property(e => e.DataInicio).HasColumnType("date");

                entity.Property(e => e.DocumentoId).HasColumnName("DocumentoID");

                entity.Property(e => e.Empregado)
                    .IsRequired()
                    .HasMaxLength(250)
                    .IsUnicode(false);

                entity.Property(e => e.FimInadimplencia).HasColumnType("datetime");

                entity.Property(e => e.Fornecedor)
                    .IsRequired()
                    .HasMaxLength(250)
                    .IsUnicode(false);

                entity.Property(e => e.InicioInadimplencia).HasColumnType("datetime");

                entity.Property(e => e.MesDeposito).HasColumnType("date");

                entity.Property(e => e.Qlp).HasColumnName("QLP");

                entity.HasOne(d => d.Analista)
                    .WithMany(p => p.DocumentoAvalidars)
                    .HasForeignKey(d => d.AnalistaId)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK__Documento__Anali__76969D2E");

                entity.HasOne(d => d.Cliente)
                    .WithMany(p => p.DocumentoAvalidars)
                    .HasForeignKey(d => d.ClienteId)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK__Documento__Clien__74AE54BC");

                entity.HasOne(d => d.Documento)
                    .WithMany(p => p.DocumentoAvalidars)
                    .HasForeignKey(d => d.DocumentoId)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK__Documento__Docum__75A278F5");
            });

            modelBuilder.Entity<Sla>(entity =>
            {
                entity.ToTable("SLA");

                entity.Property(e => e.Slaid).HasColumnName("SLAID");

                entity.Property(e => e.ClienteId).HasColumnName("ClienteID");

                entity.HasOne(d => d.Cliente)
                    .WithMany(p => p.Slas)
                    .HasForeignKey(d => d.ClienteId)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK__SLA__ClienteID__778AC167");
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}

