using Microsoft.EntityFrameworkCore;
using GeradorRotas.Models;

namespace GeradorRotas.Data
{
    public class GeradorRotasContext : DbContext
    {
        public GeradorRotasContext (DbContextOptions<GeradorRotasContext> options)
            : base(options)
        {
        }

        public DbSet<Cidade> Cidade { get; set; }
    }
}
