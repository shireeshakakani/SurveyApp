using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DataAccess;
using System.Data;
using System.Linq;


namespace BusinessObjects
{  
    
    public class SurveyBusiness
    {

        public SurveyDataAccess sd = null;

        public SurveyBusiness()
        {
            sd = new SurveyDataAccess();
        }


        public List<DataTable> GetQuestionData()
        {

            List<DataTable> ld ;
            if (sd != null)
            {


               ld= GetCollection(sd);
                

            }

            else
            {
                sd = new SurveyDataAccess();
                ld=GetCollection(sd);
            }

            return ld;


        }

        public List<DataTable> GetCollection(SurveyDataAccess sd)
        {
            List<DataTable> ldata = new List<DataTable>();
            DataSet ds = sd.GetQuestion();
            if (ds != null)
            {
                var grouped = from table in ds.Tables[0].AsEnumerable()

                              group table by new { placeCol = table["SectionId"] } into groupby

                              select new

                              {

                                  // Value = groupby.Key,

                                  ColumnValues = groupby

                              };

                foreach (var key in grouped)
                {
                   
                    DataTable dt = new DataTable();
                    dt = key.ColumnValues.CopyToDataTable();
                    ldata.Add(dt);
                }
            }

            return ldata;


        }

    }


}
