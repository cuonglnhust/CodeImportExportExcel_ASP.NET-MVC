using ExcelTest.Models.EF;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace ExcelTest.Models.Dao
{
    public class ProductDao
    {
        EntityModel db = null;
        public ProductDao()
        {
            db = new EntityModel();
        }
        public List<Product> ListALL()
        {
            return db.Products.Where(x => x.ID != 0).OrderByDescending(x => x.Price).ToList();
        }

        
    }
}