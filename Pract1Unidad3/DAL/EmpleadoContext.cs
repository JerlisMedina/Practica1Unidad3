using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Pract1Unidad3.Models;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;

namespace Pract1Unidad3.DAL
{
    public class EmpleadoContext : DbContext
    {
        public EmpleadoContext() : base("EmpleadoContext") {
        }
        public DbSet<Empleado> Empleados { get; set; }
        public DbSet<Registro> Registros { get; set; }
        public DbSet<Departamento> Departamentos { get; set; }
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
        }

    }
}
