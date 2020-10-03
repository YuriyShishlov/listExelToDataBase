using System.Data.Entity;

namespace listExelToDataBase.Data.DataModel
{

    public class LpuContext : DbContext
    {
        public LpuContext()
            : base("name=LpuContext")
        {
        }

        public DbSet<LpuData> LpuDatas { get; set; }
    }
}