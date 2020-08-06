using System;
using System.Data.Entity;


namespace ConsoleApp2SQL
{
    public class MyDbContext : DbContext
    {
        public MyDbContext() : base("DbConnectionString")
        { }
        public DbSet<Group> Groups { get; set; }

    }
}
