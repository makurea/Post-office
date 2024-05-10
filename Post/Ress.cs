using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Post
{
    public static class Ress
    {
        public static readonly SqlConnection connect = new SqlConnection(@"Server=.\SQLEXPRESS;Initial Catalog=post;Integrated Security=True;");
    }
}
