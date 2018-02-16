using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using BusinessObjects;
using PracticePerformanceAssessmentDataAccess;

using System.Data;
using System.IO;
using System.Net;
using System.Globalization;
using System.Configuration;
using System.Data.SqlClient;

namespace SurveyApp
{
    public partial class Admin : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
           PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess ObjDataAccess = new PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess("", 1);
           // IEnumerable<GetAdminInfo_Result> ld = ObjDataAccess.GetAdminInfo();
            IEnumerable<GetSurveyTranscation_Result>  tr=  ObjDataAccess.GetSurveyTranscation();

            List<GetSurveyTranscation_Result> lstquestion = new List<GetSurveyTranscation_Result>();

            DataTable dt = new DataTable();
            dt.Columns.Add("User");
            dt.Columns.Add("Year");
            dt.Columns.Add("PracticeId");
            dt.Columns.Add("CSV");
            dt.Columns.Add("Detailed");
            dt.Columns.Add("Infographic");
            dt.Columns.Add("Executive");
            dt.Columns.Add("Date");

            foreach (GetSurveyTranscation_Result data in tr)
            {
                DataRow dr = dt.NewRow();
                dr["User"]=data.UserName;
                DateTime dd = (DateTime)data.CreationDate;
                dr["Year"] =dd.Year;    //data.ActiveYear+" - "+data.PreviousActiveYear;
                       dr["PracticeId"]=data.PracticeId;
                         dr["CSV"]=data.CSVPath;
                           dr["Detailed"]=data.DetailedPath;
                             dr["Infographic"]=data.InfographicPath;
                             dr["Executive"] = data.ExecutivePath;
                             DateTime dd1 = (DateTime)data.Entrydate;
                             string day = dd1.Day.ToString();
                             string month = dd1.ToString("MMMM", CultureInfo.InvariantCulture);
                             string year = dd1.Year.ToString();
                        dr["Date"] = day + " " + month + " " + year;
                             dt.Rows.Add(dr);

                //lstquestion.Add(data);
           
            }
           

          //  DirectoryInfo directory = new DirectoryInfo(Server.MapPath("~/pics"));
           

           

            //foreach (FileInfo file in directory.GetFiles())
            //{
            //    DataRow dr = dt.NewRow();

            //    dr["CSV"] = file.Name; // Server.MapPath("~/pics" + "/" + file.Name);
               
            //     DirectoryInfo directory1 = new DirectoryInfo(Server.MapPath("~/finalpdf"));
            //     foreach (FileInfo file1 in directory1.GetFiles())
            //     {
                   
            //         string name1 = Path.GetFileNameWithoutExtension(file.Name).Split('-')[0].Substring(4);
            //         string name2 = Path.GetFileNameWithoutExtension(file1.Name).Split('-')[0];
            //         if (name1 == name2)
            //         {

            //             dr["Detailed"] = file1.Name;// Server.MapPath("~/finalpdf" + "/" + file1.Name);
            //             dr["User"] = file1.Name.Split('_')[0];
            //             dr["PracticeId"] = file1.Name.Split('_')[1];
            //         }

                    
            //          name2 = Path.GetFileNameWithoutExtension(file1.Name).Split('-')[0];
                    
            //          name2 = name2.Substring(12);
            //          if (file1.Name.StartsWith("Infographic") && name1==name2)
            //          {

            //              dr["Infographic"] = file1.Name;// Server.MapPath("~/finalpdf" + "/" + file1.Name);
            //         }


            //          name2 = Path.GetFileNameWithoutExtension(file1.Name).Split('-')[0];

            //          name2 = name2.Substring(9);
            //          if (file1.Name.StartsWith("Executive") && name1==name2)
            //         {

            //             dr["Executive"] = file1.Name;// Server.MapPath("~/finalpdf" + "/" + file1.Name);
            //         }

                    

                    
            //     }
                 
            //     dt.Rows.Add(dr);

            //}

          
                int counter = 0;
                rptrsubdata.DataSource = dt;
                rptrsubdata.DataBind();




                string csv = Server.MapPath("~/pics");
               
            //"C:\\pics\\";

                //StreamReader rdr = new StreamReader(csv + "/MasterCSVTemplate.csv");
                //string master = rdr.ReadToEnd();
                //rdr.Close();

                //bool flag = true;

                //DirectoryInfo diTarget = new DirectoryInfo(csv);
                //StreamWriter wtr = new StreamWriter(csv + "/MasterCSV.csv");
                //wtr.Flush();
                //foreach (FileInfo fi in diTarget.GetFiles())
                //{

                //    if (fi.Name != "MasterCSV.csv" && fi.Name != "MasterCSV.csv")
                //    {
                //        rdr = new StreamReader(fi.FullName);
                //        string newdata = rdr.ReadToEnd();
                //        rdr.Close();
                //        rdr.Dispose();

                //        //The New Data .csv file will have headers. Need to remove those.
                //        newdata = newdata.Substring(newdata.IndexOf('\n') + 1);
                //        if (flag)
                //        {

                //            wtr.Write(master+newdata);

                //            flag = false;
                //        }

                //        else
                //        {
                //            wtr.Write(newdata);

                //        }
                       
                //    }

                //}

                //wtr.Close();
                //wtr.Dispose();

           
                lnkmastercsv.Text = "Download";  // dr["Detailed"].ToString();
                if (dt.Rows.Count > 0)
                {
                    lnkmastercsv.Enabled = true;

                }
                else
                {
                    lnkmastercsv.Enabled = false;

                }
                lnkmastercsv.Click += new EventHandler(lnkmastercsv_Click);
               // lnkmastercsv.CssClass = csv + "/MasterCSV.csv";  // Server.MapPath("~/finalpdf" + "/" + dr["Detailed"]);




               
        }

        void lnkmastercsv_Click(object sender, EventArgs e)
        {

            string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
            SqlConnection con = new SqlConnection(connStr);
            string path = Server.MapPath("/") + @"bin\Debug";
            using (SqlCommand cmd = new SqlCommand("ExportSurveytoCSV", con))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                
                cmd.Parameters.Add("@Path", SqlDbType.VarChar).Value = path;
                

                con.Open();
                cmd.ExecuteNonQuery();
            }
            LinkButton lb = sender as LinkButton;
            //string strURL = lb.CssClass;  // Server.MapPath("~/finalpdf" + "/" + lb.Text);
            string strURL = path+"\\MasterCSV.csv";
            WebClient req = new WebClient();
            HttpResponse response = HttpContext.Current.Response;
            response.Clear();
            response.ClearContent();
            response.ClearHeaders();
            response.Buffer = true;
            response.AddHeader("Content-Disposition", "attachment;filename=\"" + "MasterCSV.csv" + "\"");
            byte[] data = req.DownloadData(strURL);
            response.BinaryWrite(data);
            response.End();
        }

        protected void rptrsubdata_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {

          
            if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
            {
                DataRowView dr = (DataRowView)e.Item.DataItem;
                LinkButton hlTabLink = e.Item.FindControl("lnkcsv") as LinkButton;
                LinkButton hlpdfLink = e.Item.FindControl("lnkpdf") as LinkButton;
                LinkButton hlinfographicLink = e.Item.FindControl("lnkinfographicpdf") as LinkButton;
                LinkButton hlexecutiveLink = e.Item.FindControl("lnkexecutivecpdf") as LinkButton;
                Label lbluser = e.Item.FindControl("lbluser") as Label;
                Label lblpractice = e.Item.FindControl("lblpracticeid") as Label;
                Label lblyear = e.Item.FindControl("lblyear") as Label;
                Label lbldate = e.Item.FindControl("lbldate") as Label;

                DateTime dt = Convert.ToDateTime(dr["Date"].ToString());

                lbldate.Text = dr["Date"].ToString();   // dt.ToString("yyyy-MM-dd");
                lblyear.Text=dr["Year"].ToString();
                hlTabLink.Text = "Download";
                hlTabLink.Visible = true;
                hlTabLink.Click += new EventHandler(hlTabLink_Click);
                hlTabLink.CssClass = dr["CSV"] + ".csv";     //Server.MapPath("~/pics" + "/" + dr["CSV"]);
                if (dr["Detailed"] != null)
                {
                    hlpdfLink.Text = "Download";  // dr["Detailed"].ToString();
                    hlpdfLink.Click += new EventHandler(hlpdfLink_Click);
                    hlpdfLink.CssClass = dr["Detailed"] + ".pdf";  // Server.MapPath("~/finalpdf" + "/" + dr["Detailed"]);
                }

                if (dr["Infographic"] != null)
                {
                    hlinfographicLink.Text = "Download";  // dr["Infographic"].ToString();
                    hlinfographicLink.Click += new EventHandler(hlinfographicLink_Click);
                    hlinfographicLink.CssClass = dr["Infographic"] + ".pdf";  // Server.MapPath("~/finalpdf" + "/" + dr["Infographic"]);
                }

                if (dr["Executive"] != null)
                {
                    hlexecutiveLink.Text = "Download";  // dr["Executive"].ToString();
                    hlexecutiveLink.Click += new EventHandler(hlexecutiveLink_Click);
                    hlexecutiveLink.CssClass = dr["Executive"] + ".pdf";  // Server.MapPath("~/finalpdf" + "/" + dr["Executive"]);
                }
               // Session["filename"] = Server.MapPath("~/pics" + "/" + dr["CSV"]);
                lbluser.Text = dr["User"].ToString();
                lblpractice.Text = dr["PracticeId"].ToString();
               
            }

        }

        void hlexecutiveLink_Click(object sender, EventArgs e)
        {
            LinkButton lb = sender as LinkButton;
            string strURL = lb.CssClass;  // Server.MapPath("~/finalpdf" + "/" + lb.Text);
            WebClient req = new WebClient();
            HttpResponse response = HttpContext.Current.Response;
            response.Clear();
            response.ClearContent();
            response.ClearHeaders();
            response.Buffer = true;
            response.AddHeader("Content-Disposition", "attachment;filename=\"" + strURL + "\"");
            byte[] data = req.DownloadData(strURL);
            response.BinaryWrite(data);
            response.End();
        }

        void hlinfographicLink_Click(object sender, EventArgs e)
        {
            LinkButton lb = sender as LinkButton;
            string strURL = lb.CssClass; // Server.MapPath("~/finalpdf" + "/" + lb.Text);
            WebClient req = new WebClient();
            HttpResponse response = HttpContext.Current.Response;
            response.Clear();
            response.ClearContent();
            response.ClearHeaders();
            response.Buffer = true;
            response.AddHeader("Content-Disposition", "attachment;filename=\"" + strURL + "\"");
            byte[] data = req.DownloadData(strURL);
            response.BinaryWrite(data);
            response.End();
        }

        void hlpdfLink_Click(object sender, EventArgs e)
        {
            LinkButton lb = sender as LinkButton;
            string strURL = lb.CssClass;  // Server.MapPath("~/finalpdf" + "/" + lb.Text);
            WebClient req = new WebClient();
            HttpResponse response = HttpContext.Current.Response;
            response.Clear();
            response.ClearContent();
            response.ClearHeaders();
            response.Buffer = true;
            response.AddHeader("Content-Disposition", "attachment;filename=\"" + strURL + "\"");
            byte[] data = req.DownloadData(strURL);
            response.BinaryWrite(data);
            response.End();
        }

        void hlTabLink_Click(object sender, EventArgs e)
        {
           LinkButton lb =sender as LinkButton;
           string strURL = lb.CssClass;  // Server.MapPath(lb.CssClass); // Server.MapPath("~/pics" + "/" + lb.Text);
            WebClient req = new WebClient();
            HttpResponse response = HttpContext.Current.Response;
            response.Clear();
            response.ClearContent();
            response.ClearHeaders();
            response.Buffer = true;
            response.AddHeader("Content-Disposition", "attachment;filename=\"" + strURL+ "\"");
            byte[] data = req.DownloadData(strURL);
            response.BinaryWrite(data);
            response.End();
        }

        protected void rptpdf_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
            {
                FileInfo dr = (FileInfo)e.Item.DataItem;
                LinkButton hlTabLink1 = e.Item.FindControl("lnkpdf") as LinkButton;
                hlTabLink1.Text = dr.Name;
                hlTabLink1.Click += new EventHandler(hlTabLink1_Click);
                Session["filename"] = Server.MapPath("~/finalpdf" + "/" + dr.Name);
                hlTabLink1.CssClass = Server.MapPath("~/finalpdf" + "/" + dr.Name);
            }

        }

        void hlTabLink1_Click(object sender, EventArgs e)
        {
            LinkButton lb = sender as LinkButton;
            string strURL = Server.MapPath("~/finalpdf" + "/" + lb.Text);
            WebClient req = new WebClient();
            HttpResponse response = HttpContext.Current.Response;
            response.Clear();
            response.ClearContent();
            response.ClearHeaders();
            response.Buffer = true;
            response.AddHeader("Content-Disposition", "attachment;filename=\"" + strURL + "\"");
            byte[] data = req.DownloadData(strURL);
            response.BinaryWrite(data);
            response.End();
        }

       
    }
}