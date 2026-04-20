using Microsoft.EntityFrameworkCore;

namespace Group4337
{
    public class ClientDbContext : DbContext
    {
        public DbSet<Client> Clients { get; set; } = null!;

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(
                "Server=localhost;Database=ClientsDB;Trusted_Connection=True;TrustServerCertificate=True;");
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Client>(entity =>
            {
                entity.ToTable("Clients");
                entity.HasKey(e => e.Id);
                entity.Property(e => e.ClientCode).HasMaxLength(20).IsRequired();
                entity.Property(e => e.FullName).HasMaxLength(150).IsRequired();
                entity.Property(e => e.BirthDate).IsRequired();
                entity.Property(e => e.PostalCode).HasMaxLength(10);
                entity.Property(e => e.City).HasMaxLength(100);
                entity.Property(e => e.Street).HasMaxLength(100);
                entity.Property(e => e.House).HasMaxLength(10);
                entity.Property(e => e.Apartment).HasMaxLength(10);
                entity.Property(e => e.Email).HasMaxLength(200);

                // Age и AgeCategory — вычисляемые свойства, не маппятся в БД
                entity.Ignore(e => e.Age);
                entity.Ignore(e => e.AgeCategory);
            });
        }
    }
}
