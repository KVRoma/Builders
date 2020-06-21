using SQLite.CodeFirst;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class BuilderContext : DbContext
    {
        public BuilderContext() : base("ConStr") { }

        public DbSet<Client> Clients { get; set; }
        public DbSet<Quotation> Quotations { get; set; }
        public DbSet<MaterialQuotation> MaterialQuotations { get; set; }
        public DbSet<Payment> Payments { get; set; }
        public DbSet<Invoice> Invoices { get; set; }
        public DbSet<WorkOrder> WorkOrders { get; set; }
        public DbSet<WorkOrder_Work> WorkOrder_Works { get; set; }
        public DbSet<WorkOrder_Installation> WorkOrder_Installations { get; set; }
        public DbSet<WorkOrder_Accessories> WorkOrder_Accessories { get; set; }
        public DbSet<WorkOrder_Contractor> WorkOrder_Contractors { get; set; }
        public DbSet<Reciept> Reciepts { get; set; }
        public DbSet<MaterialProfit> MaterialProfits { get; set; }
        public DbSet<Material> Materials { get; set; }
        public DbSet<LabourProfit> LabourProfits { get; set; }
        public DbSet<Labour> Labours { get; set; }
        public DbSet<LabourContractor> LabourContractors { get; set; }
        public DbSet<Expenses> Expenses { get; set; }
        public DbSet<Debts> Debts { get; set; }
        public DbSet<Delivery> Deliveries { get; set; }
        public DbSet<DeliveryMaterial> DeliveryMaterials { get; set; }

        public DbSet<Generated> Generateds { get; set; }
        public DbSet<GeneratedAccessories> GeneratedAccessories { get; set; }
        public DbSet<GeneratedFlood> GeneratedFloods { get; set; }
        public DbSet<GeneratedMaterial> GeneratedMaterials { get; set; }
        public DbSet<GeneratedMolding> GeneratedMoldings { get; set; }
        public DbSet<GeneratedStairs> GeneratedStairs { get; set; }
        public DbSet<GeneratedSuplementary> GeneratedSuplementaries { get; set; }
        public DbSet<GeneratedList> GeneratedLists { get; set; }

        public DbSet<DIC_TypeOfClient> DIC_TypeOfClients { get; set; }
        public DbSet<DIC_HearAboutsUs> DIC_HearAboutsUse { get; set; }
        public DbSet<DIC_GroupeQuotation> DIC_GroupeQuotations { get; set; }
        public DbSet<DIC_ItemQuotation> DIC_ItemQuotations { get; set; }
        public DbSet<DIC_DescriptionQuotation> DIC_DescriptionQuotations { get; set; }
        public DbSet<DIC_PaymentMethod> DIC_PaymentMethods { get; set; }
        public DbSet<DIC_Area> DIC_Areas { get; set; }
        public DbSet<DIC_Room> DIC_Rooms { get; set; }
        public DbSet<DIC_ExistingFloor> DIC_ExistingFloors { get; set; }
        public DbSet<DIC_Contractor> DIC_Contractors { get; set; }
        public DbSet<DIC_Supplier> DIC_Suppliers { get; set; }
        public DbSet<DIC_DepthQuotation> DIC_DepthQuotations { get; set; }

        public DbSet<DIC_G_AccessoriesFloor> DIC_G_AccessoriesFloors { get; set; }
        public DbSet<DIC_G_Additional> DIC_G_Additionals { get; set; }
        public DbSet<DIC_G_Floor> DIC_G_Floors { get; set; }
        public DbSet<DIC_G_GradeLevel> DIC_G_GradeLevels { get; set; }
        public DbSet<DIC_G_Modeling> DIC_G_Modelings { get; set; }
        public DbSet<DIC_G_Painting> DIC_G_Paintings { get; set; }
        public DbSet<DIC_G_Partition> DIC_G_Partitions { get; set; }
        public DbSet<DIC_G_TypeAccessoriesFloor> DIC_G_TypeAccessoriesFloors { get; set; }
        public DbSet<DIC_G_TypeLeveling> DIC_G_TypeLevelings { get; set; }
        public DbSet<DIC_G_TypeModeling> DIC_G_TypeModelings { get; set; }
        public DbSet<DIC_G_TypeStairs> DIC_G_TypeStairs { get; set; }
        public DbSet<DIC_G_TypeTrim> DIC_G_TypeTrims { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            var sqliteConnectionInitializer = new SqliteCreateDatabaseIfNotExists<BuilderContext>(modelBuilder);
            Database.SetInitializer(sqliteConnectionInitializer);
        }

    }
}
