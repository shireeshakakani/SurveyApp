using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace DataAccess
{
    public class SurveyDataAccess
    {
        public static string conString;

        public SurveyDataAccess()
        {

            if (ConfigurationManager.ConnectionStrings["ApplicationServices"].ConnectionString != null)
            {
                conString = ConfigurationManager.ConnectionStrings["ApplicationServices"].ConnectionString;
            }
            else
            {
                conString = null;
            }
        }

        private SqlConnection ReturnConnection()
        {
            SqlConnection conn = new SqlConnection();
            if (conString != null)
            {
                conn.ConnectionString = conString;
                return conn;
            }

            return null;
        }


        public DataSet GetQuestion()
        {

          
            SqlConnection conn = new SqlConnection();
           
            SqlCommand cmd ;
            SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet ds = new DataSet();
            conn=ReturnConnection();
            if (conn != null)
            {
              
                conn.Open();
                cmd =conn.CreateCommand();
               
                cmd.CommandText = "Select * from SurveyQuestion q inner join Section s on q.SectionId=s.Id ";
                cmd.CommandType = System.Data.CommandType.Text;
                adapter.SelectCommand = cmd;
                adapter.Fill(ds);
                return ds;
            }
            return null;
        }


    }
}
