using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace ExlsCoursesToBD
{
    class ToDB
    {
        private const string myConnectionString = @"Server=MORIA; Database=WebDev_GLEKUZ ;Trusted_Connection=Yes";
        public string errors = string.Empty;




        public void ExecUpdatePersonIdInOtherTables(string firstName, string secondName, string middleName, string birthdat)
        {
            SqlConnection connection = new SqlConnection(myConnectionString);
            try
            {
                connection.Open();
                using (SqlCommand comm = new SqlCommand("UpdatePersonIdInOtherTables", connection))
                {

                    comm.CommandType = CommandType.StoredProcedure;
                    //comm.Parameters.AddWithValue("newID", newID);
                    //comm.Parameters.AddWithValue("oldID", oldID);
                    comm.ExecuteNonQuery();
                }
            }
            catch (Exception exception)
            {
                errors += exception.Message;
            }
            finally
            {
                connection.Close();
            }
        }
    }
}
