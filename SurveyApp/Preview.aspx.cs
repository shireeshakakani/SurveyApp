using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using PracticePerformanceAssessmentDataAccess;
using System.Web.UI.HtmlControls;
using System.Web.Services;
using System.Web.Script.Serialization;
using System.IO;
using System.Text;
using BusinessObjects;
using System.Reflection;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Web.UI.DataVisualization.Charting;
using System.Drawing;
using System.Globalization;
using System.Net;
using System.Threading;


namespace SurveyApp
{
    
    public partial class Preview : System.Web.UI.Page
    {
        public List<DataTable> ld;
        public bool fpage =true;
        public DataTable currenttable;

        public int currentpageno=1;
        private int TotalPages;
        private bool compare;
        private bool display3d;
        private bool showValues;
        private string charType = "Column";
        private string bsub=string.Empty;

        private string startdate = string.Empty;
        private string starttime = string.Empty;
        public int TotalPages1
        {
            get { return TotalPages; }
            set { TotalPages = value; }
        }
        private Dictionary<string, string> columnname=new Dictionary<string,string>();
        public Dictionary<string,string> ColumnName
        {
            get { return columnname; }
            set { columnname = value; }
        }

        public int practiceid
        {
            get { return Practiceid; }
            set { Practiceid = value; }
        }
        public int Practiceid;

        public int tabindex = 1;
    
        protected void Page_Load(object sender, EventArgs e)
        {
            
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");
                     
   
            string name = Environment.UserName;
            startdate = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);
            starttime = DateTime.Now.ToString("HH:mm:ss tt");

            if (!IsPostBack)
            {
                secondpage.Visible = true;

                RenderQuestion();
            }
        }

     
       
        public void RenderQuestion()
        {
                
                PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess ObjDataAccess = new PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess("", 1);


              

                Dictionary<string, string> dcsave = new Dictionary<string, string>();//Session["columnlist"] as Dictionary<string, string>;

                IEnumerable<GetQuestionData_Result> ld = ObjDataAccess.GetQuestionData(Convert.ToInt32(Session["practiceid"]), Convert.ToInt32(System.Web.HttpContext.Current.Session["YearValue"]));
                
                    List<GetQuestionData_Result> lstquestion = new List<GetQuestionData_Result>();
                    foreach (GetQuestionData_Result data in ld)
                    {
                        lstquestion.Add(data);

                        if (data.DbColumnName == "Q7")
                        {

                            string q17 = data.QuestionId.ToString();

                            Session["q17"] = q17;
                        }

                        if (data.DbColumnName == "Q18")
                        {

                            string q18 = data.QuestionId.ToString();

                            Session["q18"] = q18;
                        }


                        if (data.DbColumnName == "Q40")
                        {

                            string q26a = data.QuestionId.ToString();

                            Session["q26a"] = q26a;
                        }

                        if (data.DbColumnName == "Q79")
                        {

                            string q26b = data.QuestionId.ToString();

                            Session["q26b"] = q26b;
                        }
                    }
                    Session["table"] = lstquestion;
                    

                    //var lengthgroups = lstquestion.Where(t => (t.PageNo.ToString() == btnNext.ToolTip)).GroupBy(a => a.SectionId).ToList();

                    var lengthgroups = lstquestion.GroupBy(a => a.SectionId).ToList();

                    List<GetQuestionData_Result> lst = new List<GetQuestionData_Result>();

                    foreach (IGrouping<int?, GetQuestionData_Result> section in lengthgroups)
                    {

                        lst.Add(section.ElementAt(1));
                    }

                        rptrdata.DataSource = lst;

             rptrdata.DataBind();
            //lblyear1.Text=Session["YearValue"].ToString();
            //lblname1.Text = Session["namesave"].ToString();
            //lblpracticename1.Text = Session["practicenamesave"].ToString();
            //lblyeartext1.Text = Session["YearName"].ToString();
                       
             


          
        }


        protected void Backpagebutton_Click(object sender, EventArgs e)
        {
            //Session["preview"] = "1";
            //btnback.PostBackUrl = "/HomePage.aspx?id=6";
           
           
Response.Redirect("/HomePage.aspx?id=6",false);
        }

        private void NextButtonQuestionRendering(string btnname)
        {
            tabindex = 1;
            bool flag = true;
            Dictionary<string, string> dc = Session["columnlist"] as Dictionary<string, string>;
            Dictionary<string, string> dcfinal = new Dictionary<string, string>();
            foreach (RepeaterItem item in rptrdata.Items)
            {
                if (item.ItemType == ListItemType.Item || item.ItemType == ListItemType.AlternatingItem)
                {
                    Repeater rpt = (Repeater)item.FindControl("rptrchild");
                    foreach (RepeaterItem item1 in rpt.Items)
                    {
                        if (item1.ItemType == ListItemType.Item || item1.ItemType == ListItemType.AlternatingItem)
                        {
                            Label HC = (Label)item1.FindControl("Q44");

                            RadioButtonList radgender = (RadioButtonList)item1.FindControl("radgender");
                            RadioButtonList radOfficeManager = (RadioButtonList)item1.FindControl("radOfficeManager");


                            Repeater rptsub = (Repeater)item1.FindControl("rptrsubdata");
                            foreach (string k in dc.Keys)
                            {

                                if (k == HC.Attributes["class"] && k != "Q80" && k != "Q97")
                                {

                                    if (HC != null)
                                    {


                                        dcfinal.Add(k, HC.Text);

                                    }
                                }
                                if (radgender != null)
                                {
                                    if (k == radgender.Attributes["class"])
                                    {




                                        dcfinal.Add(k, radgender.SelectedValue);

                                    }
                                }
                                if (radOfficeManager != null)
                                {
                                    if (k == radOfficeManager.Attributes["class"])
                                    {




                                        dcfinal.Add(k, radOfficeManager.SelectedValue);

                                    }
                                }
                                foreach (RepeaterItem item2 in rptsub.Items)
                                {
                                    Label HC2 = (Label)item2.FindControl("txtsubquestion");
                                    RadioButtonList RL = (RadioButtonList)item2.FindControl("radsub");
                                    if (k == HC2.Attributes["class"])
                                    {

                                        if (HC2 != null)
                                        {


                                            dcfinal.Add(k, HC2.Text);

                                        }
                                    }


                                    if (RL != null)
                                    {
                                        if (k == RL.Attributes["class"])
                                        {


                                            dcfinal.Add(k, RL.SelectedValue);

                                        }
                                    }
                                }


                            }
                        }


                    }

                }

            }


            Dictionary<string, string> DCF = new Dictionary<string, string>();
            foreach (string key in dc.Keys)
            {
                if (!dcfinal.Keys.Contains(key))
                {
                    DCF.Add(key, dc[key]);
                }

                else
                {
                    DCF.Add(key, dcfinal[key]);
                }
            }

            Session["columnlist"] = DCF;
            //Dictionary<string,string> dc=Session["columnlist"] as Dictionary<string, string>;
            List<GetQuestionData_Result> lst = Session["table"] as List<GetQuestionData_Result>;



            if (flag)
            {
                //btnback.ToolTip = (Convert.ToInt32(btnback.ToolTip) + 1).ToString();
                //btnNext.ToolTip = (Convert.ToInt32(btnNext.ToolTip) + 1).ToString();

                //RenderPaging(Convert.ToInt32(btnNext.ToolTip) + 1);
                if (btnname == "Next")
                {
                    btnback.ToolTip = (Convert.ToInt32(btnback.ToolTip) + 1).ToString();


                    RenderPaging(Convert.ToInt32(btnback.ToolTip) + 1);
                    QuestionRendering(btnback);

                    int tab = btnback.TabIndex + 1;
                    // btnprint.TabIndex = Convert.ToInt16(tab);
                    // int tabbext = btnbackbottom.TabIndex + 1;
                    


                }

                else if (btnname == "Back")
                {
                    //btnback.ToolTip = (Convert.ToInt32(btnback.ToolTip) - 1).ToString();
                    //btnNext.ToolTip = (Convert.ToInt32(btnNext.ToolTip) - 1).ToString();
                    //RenderPaging(Convert.ToInt32(btnNext.ToolTip) + 1);
                    QuestionRendering(btnback);

                    int tab = btnback.TabIndex + 1;
                   // int tabbext = btnbackbottom.TabIndex + 1;
                   
                }

                else if (btnname == "Save")
                {

                }

            }

            if (Session["columnlist"] != null && Session["columnlist"] as Dictionary<string, string> != null)
            {

                Dictionary<string, string> dc1 = Session["columnlist"] as Dictionary<string, string>;

            }

         


        }

        public void RenderPaging(int currentpage)
        {

            if (TotalPages == 0)
            {
                practiceid = Convert.ToInt32(Session["practiceid"]);
                PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess ObjDataAccess = new PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess("", practiceid);
                TotalPages1 = ObjDataAccess.GetPageSection(Convert.ToInt32(Session["practiceid"]), Convert.ToInt16(Session["YearValue"]))+ 1;
                
            }


            HtmlGenericControl divprogress = Page.FindControl("divprogress") as HtmlGenericControl;

            if (divprogress != null)
            {
                divprogress.InnerText = "Page " + currentpage + " of " + TotalPages1;
            }

            if (currentpage == TotalPages1)
            {
                // btnSave.Visible = true;
                btnSubmit.Visible = true;
                
                btnback.Visible = true;
               // btnbackbottom.Visible = true;

            }

            else
            {
                btnSubmit.Visible = false;
                
                btnback.Visible = true;
                //btnbackbottom.Visible = true;
            }
            HtmlTable control = Page.FindControl("tblpaging") as HtmlTable;
            if (control != null)
            {
                HtmlTableRow tr = new HtmlTableRow();
                control.Rows.Add(tr);

                for (int i = 1; i <= TotalPages; i++)
                {
                    HtmlTableCell td = new HtmlTableCell();
                    td.Attributes.Add("style", "width:10%;");

                    td.Attributes.Add("class", "uncompleted-cell");

                    td.InnerHtml = "&nbsp;";
                    tr.Controls.Add(td);

                }

                for (int i = 1; i <= TotalPages; i++)
                {
                    HtmlTableCell tc = (HtmlTableCell)control.Rows[0].Cells[i - 1];
                    if (i <= currentpage)
                    {

                        tc.Attributes.Remove("class");
                        tc.Attributes.Add("class", "completed-cell");
                    }

                    else
                    {
                        tc.Attributes.Remove("class");
                        tc.Attributes.Add("class", "uncompleted-cell");
                    }
                }


            }



        }


        protected void Submitpagebutton_Click(object sender, EventArgs e)
        {
            lblsaveerror.Text = "";
            parenterrorlbl.Text = "";
            try
            {
                DataTable dtsave = null;
                save();
            }
            catch (Exception ex)
            {
                string filePath = null;
                if (ConfigurationManager.AppSettings["ErrorFilePath"] != null)
                {
                    filePath = ConfigurationManager.AppSettings["ErrorFilePath"].ToString();// @"C:\Error.txt";
                }
                if (filePath != null)
                {
                    using (StreamWriter writer = new StreamWriter(filePath, true))
                    {
                        writer.WriteLine("Messagesubmitpagebutton :" + ex.Message + "Stacktrace:" + ex.StackTrace + ex.InnerException);
                        writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                    }

                    //Response.Redirect("/Result.aspx?Result=E/P");
                    Response.Redirect("/Result.aspx?Result=S",false);
                }

                else
                {

                    Response.Redirect("/Result.aspx?Result=S",false);
                }

            }

            int NoOfRowsLimit = 1;
            PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess ObjDataAccess = new PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess("", NoOfRowsLimit);

            string strSourcePath = Session["strSourcePath"].ToString();
            string finalpath = Session["finalpath"].ToString();
            Report ObjFinalReport = Session["ObjFinalReport"] as Report;
            string strSourceInfographicPath = Session["strSourceInfographicPath"].ToString();
            string csvPath = Session["csvPath"].ToString();
            string strSourceexecutivePath = Session["strSourceexecutivePath"].ToString();

            //Dictionary<string,DataSet> x = Session["x"] as Dictionary<string,DataSet>;
            Dictionary<Dictionary<string, DataTable>, DataSet> x = Session["x"] as Dictionary<Dictionary<string, DataTable>, DataSet>;
            practiceid = Convert.ToInt32(Session["practiceid"]);
            string strStatus = ObjDataAccess.GenerateWordReport(Session["YearName"].ToString(), strSourcePath, finalpath, ObjFinalReport, practiceid.ToString(), x.Values.FirstOrDefault(), x.Keys.FirstOrDefault().Values.FirstOrDefault());
            // string strStatus = ObjDataAccess.GenerateWordReport(strSourcePath, strDOCFileTargetPath, ObjFinalReport, x.Values.FirstOrDefault(), x.Keys.FirstOrDefault().Values.FirstOrDefault());
            string csv = Server.MapPath("~/pics");
            string pdf = Server.MapPath("~/finalpdf");
            DirectoryInfo diSource = new DirectoryInfo(csvPath);
            DirectoryInfo diTarget = new DirectoryInfo(csv);
            DirectoryInfo diSourcepdf = new DirectoryInfo(finalpath);
            DirectoryInfo diTargetpdf = new DirectoryInfo(pdf);

            DirectoryInfo diSourceinfographicpdf = new DirectoryInfo(finalpath);
            DirectoryInfo diTargetinfographicpdf = new DirectoryInfo(pdf);
            DirectoryInfo diSourceexecutivepdf = new DirectoryInfo(finalpath);
            DirectoryInfo diTargetexecutivepdf = new DirectoryInfo(pdf);
            if (strStatus == "success")
            {
                string filenamestr = System.Web.HttpContext.Current.Session["varName"].ToString();
                string strinfographicStatus = ObjDataAccess.GenerateInfographicWordReport(Server.MapPath("~/ImageChart"), strSourceInfographicPath, finalpath, ObjFinalReport, practiceid.ToString(), x.Keys.FirstOrDefault().Keys.FirstOrDefault(), x.Keys.FirstOrDefault().Values.FirstOrDefault());
                if (strinfographicStatus == "success")
                {
                    string filenameinfographicstr = System.Web.HttpContext.Current.Session["varNameinfographic"].ToString();
                    string strexecutivetatus = ObjDataAccess.GenerateExecutiveWordReport(Server.MapPath("~/ImageChart"), strSourceexecutivePath, finalpath, ObjFinalReport, practiceid.ToString(), x.Keys.FirstOrDefault().Values.FirstOrDefault());

                    if (strexecutivetatus == "success")
                    {

                        if (!Directory.Exists(csv))
                        {
                            Directory.CreateDirectory(csv);
                        }

                        if (!Directory.Exists(pdf))
                        {
                            Directory.CreateDirectory(csv);
                        }





                        CopyAll(diSource, diTarget, Session["csvname"].ToString(), ".csv");

                        filenamestr = System.Web.HttpContext.Current.Session["varName"].ToString();


                        CopyAll(diSourcepdf, diTargetpdf, filenamestr, ".pdf");

                        filenameinfographicstr = System.Web.HttpContext.Current.Session["varNameinfographic"].ToString();
                        string filenameexecutivestr = System.Web.HttpContext.Current.Session["varNameexecutive"].ToString();


                        CopyAll(diSourceinfographicpdf, diTargetinfographicpdf, filenameinfographicstr, ".pdf");


                        CopyAll(diSourceexecutivepdf, diTargetexecutivepdf, filenameexecutivestr, ".pdf");

                        if (Session["table"] != null && Session["table"] as List<GetQuestionData_Result> != null)
                        {
                            string username = "";
                            if (Session["Username"] != null)
                            {
                                username = Session["Username"].ToString();
                            }

                            List<GetQuestionData_Result> lst = Session["table"] as List<GetQuestionData_Result>;
                            string status = ObjDataAccess.InsertSurveyTranscation(lst.Select(t => t.SurveyId).FirstOrDefault(),Session["YearValue"].ToString(), practiceid, username, DateTime.Now,
                                 pdf + "\\" + filenamestr, pdf + "\\" + filenameinfographicstr, pdf + "\\" + filenameexecutivestr, csv + "\\" + Session["csvname"].ToString());
                        }

                        Response.Redirect("/Result.aspx?Result=Y",false);
                    }

                    else
                    {
                        CopyAll(diSource, diTarget, Session["csvname"].ToString(), ".csv");
                        CopyAll(diSourcepdf, diTargetpdf, filenamestr, ".pdf");
                        CopyAll(diSourceinfographicpdf, diTargetinfographicpdf, filenameinfographicstr, ".pdf");
                        if (Session["table"] != null && Session["table"] as List<GetQuestionData_Result> != null)
                        {
                            filenameinfographicstr = System.Web.HttpContext.Current.Session["varNameinfographic"].ToString();
                            string username = "";
                            if (Session["Username"] != null)
                            {
                                username = Session["Username"].ToString();
                            }

                            List<GetQuestionData_Result> lst = Session["table"] as List<GetQuestionData_Result>;
                            string status = ObjDataAccess.InsertSurveyTranscation(lst.Select(t => t.SurveyId).FirstOrDefault(), Session["YearValue"].ToString(), practiceid, username, DateTime.Now,
                                 pdf + "\\" + filenamestr, pdf + "\\" + filenameinfographicstr, null, csv + "\\" + Session["csvname"].ToString());
                        }
                        Response.Redirect("/Result.aspx?Result=N/E");

                    }
                }
                else
                {
                    CopyAll(diSource, diTarget, Session["csvname"].ToString(), ".csv");
                    CopyAll(diSourcepdf, diTargetpdf, filenamestr, ".pdf");
                    if (Session["table"] != null && Session["table"] as List<GetQuestionData_Result> != null)
                    {
                        filenamestr = System.Web.HttpContext.Current.Session["varName"].ToString();
                        string username = "";
                        if (Session["Username"] != null)
                        {
                            username = Session["Username"].ToString();
                        }

                        List<GetQuestionData_Result> lst = Session["table"] as List<GetQuestionData_Result>;
                        string status = ObjDataAccess.InsertSurveyTranscation(lst.Select(t => t.SurveyId).FirstOrDefault(),Session["YearValue"].ToString(), practiceid, username, DateTime.Now,
                             pdf + "\\" + filenamestr, null, null, csv + "\\" + Session["csvname"].ToString());
                    }
                    Response.Redirect("/Result.aspx?Result=N/I",false);
                }
            }

            else
            {
                CopyAll(diSource, diTarget, Session["csvname"].ToString(), ".csv");
                if (Session["table"] != null && Session["table"] as List<GetQuestionData_Result> != null)
                {
                    string username = "";
                    if (Session["Username"] != null)
                    {
                        username = Session["Username"].ToString();
                    }

                    List<GetQuestionData_Result> lst = Session["table"] as List<GetQuestionData_Result>;
                    string status = ObjDataAccess.InsertSurveyTranscation(lst.Select(t => t.SurveyId).FirstOrDefault(),Session["YearValue"].ToString(), practiceid, username, DateTime.Now,
                         null, null, null, csv + "\\" + Session["csvname"].ToString());
                }

                Response.Redirect("/Result.aspx?Result=N/D",false);
            }



        }

      


        public void QuestionRendering(Button btn)
        {
            if (Session["table"] != null)
            {
                List<GetQuestionData_Result> lst = Session["table"] as List<GetQuestionData_Result>;
                //var lengthgroups = lst.Where(t =>(t.QuestionOrderNo > mincounter && t.QuestionOrderNo<=maxcounter)).GroupBy(a => a.SectionId).ToList();
                var lengthgroups = lst.Where(t => (t.PageNo.ToString() == btn.ToolTip)).OrderBy(x=>x.QuestionOrderNo).GroupBy(a => a.SectionId).ToList();

                List<GetQuestionData_Result> lst1 = new List<GetQuestionData_Result>();
                foreach (IGrouping<int?, GetQuestionData_Result> section in lengthgroups)
                {

                    lst1.Add(section.ElementAt(1));
                }

                //Session["CurrentPageData"] = lst.Where(t => (t.PageNo.ToString() == btn.ToolTip)).ToList();
                
                rptrdata.DataSource = null;
                rptrdata.DataBind();
                rptrdata.DataSource = lst1;
                rptrdata.DataBind();


            }
        }


      


        protected void Nextpagebutton_Click(object sender, EventArgs e)
        {
            lblsaveerror.Text = "";
            parenterrorlbl.Text = "";
            lblsaveerror.Style.Add("display", "none");
            if (Session["showtotalgroup"] != null)
            {

            }

            else
            {
                Dictionary<string, string> showtotalgroup = new Dictionary<string, string>();
                Session["showtotalgroup"] = showtotalgroup;
            }
            

           

        }

     


     

        protected void rptrdata_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
            {

                GetQuestionData_Result dr = (GetQuestionData_Result)e.Item.DataItem;

                Label hlTabLink = e.Item.FindControl("lblsection") as Label;
                Repeater rpt = e.Item.FindControl("rptrchild") as Repeater;
               
                if (dr.HavingSubQuestion != null)
                {
                   
                }
                hlTabLink.Text = dr.SectionText.ToString();
               
                if (Session["table"] != null)
                {
                    List<GetQuestionData_Result> lst = Session["table"] as List<GetQuestionData_Result>;



                    var groups = from p in lst
                                 where p.SectionId == dr.SectionId
                                 group p by p.QuestionOrderNo into g
                                 select new { GroupName = g.Key, Members = g };
                    List<GetQuestionData_Result> groupedObjects = new List<GetQuestionData_Result>();
                    foreach (var g in groups)
                    {
                        groupedObjects.Add(g.Members.FirstOrDefault());

                    }

                    rpt.DataSource = null;
                    rpt.DataBind();
                    rpt.DataSource = groupedObjects.OrderBy(t=>t.QuestionOrderNo);
                    rpt.DataBind();
                    
                    //btnNext.TabIndex = Convert.ToInt16(tabindex);
                    //btnback.TabIndex = Convert.ToInt16(tabindex + 1);

                    //btnNextbottom.TabIndex = Convert.ToInt16(tabindex+2);
                    //btnbackbottom.TabIndex = Convert.ToInt16(tabindex + 3);
                    //btnSave.TabIndex = Convert.ToInt16(tabindex + 4);
                    //btnprint.TabIndex = Convert.ToInt16(tabindex + 5);
                    
                }
            }
        }

                
               
               
               

            
        

        protected void rptrchild_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
            {
                GetQuestionData_Result dr = (GetQuestionData_Result)e.Item.DataItem;
                Label hlTabLink = e.Item.FindControl("lblqtext") as Label;
                Label lblno = e.Item.FindControl("lblqorder") as Label;
                Label lblerror = e.Item.FindControl("lblerrormsg") as Label;
                HiddenField hd = e.Item.FindControl("hidvalue") as HiddenField;
                HtmlImage image = e.Item.FindControl("imghelptext") as HtmlImage;
              // TextBox txt = e.Item.FindControl("Q44") as TextBox;
               // HtmlGenericControl hg = e.Item.FindControl("litcontrol") as HtmlGenericControl;
                Label txt = e.Item.FindControl("Q44") as Label;
                Literal currencytype = e.Item.FindControl("lblquescurrency") as Literal;
                Repeater rpt = e.Item.FindControl("rptrsubdata") as Repeater;
                RadioButtonList radgender  = e.Item.FindControl("radgender") as RadioButtonList;
                RadioButtonList radOfficeManager = e.Item.FindControl("radOfficeManager") as RadioButtonList;
                
                if (dr.RuleName == "percentagechk")
                {

                    hd.ID = "hdq" + "/" + dr.DbColumnName;
                    hd.Value = dr.RuleName;
                }


                if (dr.HelpText != null)
                {
                    hlTabLink.Text = dr.QuestionText.ToString() + " " + dr.HelpText;
                }

                else
                {

                    hlTabLink.Text = dr.QuestionText.ToString();
                }
                hlTabLink.ID = dr.DbColumnName;
                lblno.Text = dr.QuestionOrderNo.ToString()+". ";
                lblno.ID = "lblorderno_" + dr.DbColumnName;
                //txt.ID = dr.DbColumnName;
                
                lblerror.ID = "lblerror" + dr.DbColumnName;
                
                //TextBox tx = new TextBox();
                //tx.ClientIDMode = ClientIDMode.Static;
                //tx.ID = dr.DbColumnName;
                //hg.Controls.Add(tx);

                if (dr.QuestionDataType == "Male/Female" && dr.QuestionControl == "CheckBox")
                {
                    radgender.Visible = true;
                    radgender.Attributes.Add("class", dr.DbColumnName);
                    txt.Visible = false;
                    radgender.TabIndex = Convert.ToInt16(tabindex);
                    tabindex++;
                }
                else if(dr.QuestionId == 491 && dr.QuestionControl == "CheckBox")
                {
                    radgender.Visible = false;
                    radOfficeManager.Visible = true;
                    radOfficeManager.Attributes.Add("class", dr.DbColumnName);
                    txt.Visible = false;
                    radOfficeManager.TabIndex = Convert.ToInt16(tabindex);
                    tabindex++;
                }
                else
                {
                    radgender.Visible = false;
                    txt.Attributes.Add("class", dr.DbColumnName);
                    txt.Visible = true;
                    if (dr.DbColumnName == "Q38")
                    {

                        txt.Text = Session["namesave"].ToString();
                    }

                    if (dr.DbColumnName == "Q47")
                    {
                        txt.Text = Session["practicenamesave"].ToString();
                    }
                   
                }
                if (image != null)
                {

                    if (dr.HelpText != null)
                    {
                        image.Alt = dr.HelpText;
                       // image.Visible = true;
                        image.Attributes.Add("title",dr.HelpText);

                    }

                }
                if (dr.HavingSubQuestion == true)
                {
                    if (Session["table"] != null)
                    {
                        if (txt!=null)
                       txt.Visible = false;
                        List<GetQuestionData_Result> lst = Session["table"] as List<GetQuestionData_Result>;
                      
                        //var groups = from p in lst
                        //             where p.QuestionId == dr.QuestionId
                        //             group p by p.SubQuestionId into g
                        //             select new { GroupName = g.Key, Members = g };

                        var groups = from p in lst
                                     where p.QuestionOrderNo == dr.QuestionOrderNo
                                     group p by p.SubQuestionId into g
                                     select new { GroupName = g.Key, Members = g };
                        List<GetQuestionData_Result> groupedObjects = new List<GetQuestionData_Result>();

                        //var groups1 = from p in groups
                        //             where p == dr.QuestionOrderNo
                        //             group p by p.SubAdditionalText into g
                        //             select new { GroupName = g.Key, Members = g };
                        foreach (var g in groups)
                        {
                            groupedObjects.Add(g.Members.FirstOrDefault());

                        }

                        if (dr.RuleId != null)
                        {

                            if (dr.RuleName == "ChkQuesRef1" ||dr.RuleName=="ChQuesRef2")
                            {
                                if (dr.QuestionRef != null)
                                {
                                    string t = lst.Where(t1 => t1.QuestionOrderNo == dr.QuestionRef).Select(t1 => t1.DbColumnName).First().ToString();
                                   var group=from p in lst
                                             where p.QuestionOrderNo==dr.QuestionOrderNo
                                             group p by p.DbColumnName into g
                                             select new { GroupName = g.Key, Members = g };

                                   List<GetQuestionData_Result> groupedObjects1 = new List<GetQuestionData_Result>();
                                   foreach (var g in group)
                                   {
                                       groupedObjects1.Add(g.Members.FirstOrDefault());

                                   }

                                   if (groupedObjects1.Count > 1)
                                   {
                                       hd.ID = "hdq/" + t + "/" + groupedObjects1[0].DbColumnName + "/" + groupedObjects.Where(m => m.DbColumnName == groupedObjects1[0].DbColumnName).Count() + "$" + groupedObjects1[1].DbColumnName + "/" + groupedObjects.Where(m => m.DbColumnName == groupedObjects1[1].DbColumnName).Count();
                                   }
                                   else
                                   {
                                       hd.ID = "hdq/" + t + "/" + dr.DbColumnName + "/" + groupedObjects.Count;
                                   }
                                    hd.Value = dr.RuleName;
                                }
                            }

                            if (dr.RuleName == "ChkTotal")
                            {

                                hd.ID = "hdq" + "/" + dr.DbColumnName+"/" + dr.Total + "/" + groupedObjects.Count;
                                hd.Value = dr.RuleName;
                            }
                        }

                        else
                        {

                        }


                        rpt.DataSource = groupedObjects;
                        rpt.DataBind();
                        currencytype.Visible = false;
                    }
                }

               
                else
                {
                      txt.Attributes.Add("TabIndex", Convert.ToInt16(tabindex).ToString());
                       tabindex++;                   
                    if (currencytype != null)
                    {
                        if (dr.QuestionDataType != "int" && dr.QuestionDataType != "string" && dr.QuestionControl!="CheckBox")
                        {
                            currencytype.Text = dr.QuestionDataType;
                            currencytype.Visible = true;
                        }

                        if(dr.QuestionDataType=="$")
                        {
                            txt.Attributes.Add("maxlength", "9");

                        }

                        else if(dr.QuestionDataType=="%")
                        {

                            txt.Attributes.Add("maxlength", "3");
                            //if (dr.DbColumnName == "Q41" || dr.DbColumnName == "Q43")
                            //{
                            //    txt.Attributes.Add("maxlength", "3");

                            //}
                            //else
                            //{
                            //    txt.Attributes.Add("maxlength", "2");
                            //}

                           // txt.Attributes.Add("maxlength", "2");
                        }
                        else
                        {
                            txt.Attributes.Add("maxlength", "6");

                        }

                        
                    }


                    if (Session["columnlist"] != null && Session["columnlist"] as Dictionary<string, string> != null)
                    {
                        Dictionary<string, string> dc = Session["columnlist"] as Dictionary<string, string>;
                        if (txt.Attributes["class"] != null)
                        {
                            if (dc.Keys.Contains(txt.Attributes["class"]))
                            {
                                //dc[txt.Attributes["class"]] = txt.Value;
                            }

                            else
                            {
                                dc.Add(dr.DbColumnName.ToString(), txt.Text);
                            }
                            if (dc[txt.Attributes["class"]] != null)
                            {
                                txt.Text = dc[txt.Attributes["class"]];
                                if(string.IsNullOrEmpty(txt.Text))
                                {

                                    currencytype.Text = "";
                                }
                            }
                        }

                        if (radgender.Attributes["class"] != null)
                        {

                            if (dc.Keys.Contains(radgender.Attributes["class"]))
                            {
                                //dc[txt.Attributes["class"]] = txt.Value;
                            }

                            else
                            {
                                dc.Add(dr.DbColumnName, (radgender.SelectedItem == null ? "2" : radgender.SelectedItem.Value));
                              
                            }
                            if (dc[radgender.Attributes["class"]] != null)
                            {
                                radgender.SelectedIndex = Convert.ToInt32(dc[radgender.Attributes["class"]]) - 1;
                            }

                        }
                        if (radOfficeManager.Attributes["class"] != null)
                        {

                            if (dc.Keys.Contains(radOfficeManager.Attributes["class"]))
                            {
                                //dc[txt.Attributes["class"]] = txt.Value;
                            }

                            else
                            {
                                dc.Add(dr.DbColumnName, (radOfficeManager.SelectedItem == null ? "2" : radOfficeManager.SelectedItem.Value));

                            }
                            if (dc[radOfficeManager.Attributes["class"]] != null)
                            {
                                radOfficeManager.SelectedIndex = Convert.ToInt32(dc[radOfficeManager.Attributes["class"]]) - 1;
                            }

                        }

                        Session["columnlist"] = dc;
                        
                    }

                    else
                    {
                        Dictionary<string, string> dc = new Dictionary<string,string>();

                        dc.Add(dr.DbColumnName.ToString(), txt.Text);
                       
                       // txt.Value = dc[txt.ID];
                        Session["columnlist"] = dc;
                    }

                  
                  
                  // txt.ID=dr.DbColumnName;

                    txt.Attributes.Add("class", dr.DbColumnName);
                    rpt.DataSource = null;
                    rpt.DataBind();

                }




            }
            
        }


        protected void rptrsubdata_ItemDataBound(object sender, RepeaterItemEventArgs e)
        {
            if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
            {
                GetQuestionData_Result dr = (GetQuestionData_Result)e.Item.DataItem;
                Label hlTabLink = e.Item.FindControl("lblsubquest") as Label;
                Label lblsub = e.Item.FindControl("lblsub") as Label;
                Label lblno = e.Item.FindControl("txtsubquestion") as Label;
                RadioButtonList radsub = e.Item.FindControl("radsub") as RadioButtonList;
                Literal currencytype = e.Item.FindControl("lblsubquescurrency") as Literal;
                if (Session["showtotalgroup"] != null && Session["showtotalgroup"] as Dictionary<string, string> != null)
                {
                    Dictionary<string, string> dshowtotal = Session["showtotalgroup"] as Dictionary<string, string>;

                    if (dr.ShowTotal == true && !dshowtotal.Keys.Contains(dr.DbColumnName + "_" + dr.SubOrder))
                    {
                        dshowtotal.Add(dr.DbColumnName + "_" + dr.SubOrder,dr.QuestionId.ToString());
                        Session["showtotalgroup"] = dshowtotal;
                    }
                }
                int i = 0;
                if (dr.SubQuestionControl == "CheckBox" && radsub!=null)
                {    
                    radsub.Visible = true;
                    lblno.Visible = false;
                    if (dr.SubAdditionalText != null && bsub != dr.SubAdditionalText)
                    {
                        lblsub.Text = dr.SubAdditionalText;
                        bsub = dr.SubAdditionalText;

                    }

                    else
                    {

                        lblsub.Text = "";
                    }
                    radsub.Attributes.Add("class", dr.DbColumnName + "_" + dr.SubOrder);
                    
                    radsub.Attributes.Add("TabIndex", Convert.ToInt16(tabindex).ToString());

                    for (i = 0; i < radsub.Items.Count;i++ )
                    {
                        radsub.Items[i].Attributes.Add("TabIndex", Convert.ToInt16(tabindex).ToString());
                    }
                    tabindex++;
                    i++;
                }
                else
                {
                   // lblno.ID = dr.DbColumnName + "_" + dr.SubOrder;
                    if (dr.QuestionDataType != null && dr.QuestionDataType == "$")
                    {

                       // lblno.MaxLength = 9;

                    }

                    else if (dr.QuestionDataType != null && dr.QuestionDataType == "%")
                    {

                       // lblno.MaxLength = 2;

                       // lblno.MaxLength = 3;

                    }
                    else
                    {

                        //lblno.MaxLength = 6;
                    }
                    
                    
                    lblno.Attributes.Add("class", dr.DbColumnName + "_" + dr.SubOrder);
                    lblno.Visible = true;
                    radsub.Visible = false;
                    lblno.TabIndex = Convert.ToInt16(tabindex);
                    tabindex++;
                }
                hlTabLink.Text = dr.SubQuestionText.ToString();
                hlTabLink.ID = dr.SubQuestionId.ToString();
                KeyValuePair<string,string> k=new KeyValuePair<string,string>();
               // k.Key = dr.DbColumnName.ToString() + "_" + dr.SubOrder.ToString();
                //k.Value = "";
               // if (!columnname.Contains(new KeyValuePair<string, string>(dr.DbColumnName.ToString() + "_" + dr.SubOrder.ToString(), "")))
               // {
                    columnname.Add(dr.DbColumnName.ToString() + "_" + dr.SubOrder.ToString(), "");
              //  }
               

                if (currencytype != null)
                {
                    if (dr.QuestionDataType != "int" && dr.QuestionDataType != "string" && dr.QuestionControl != "CheckBox")
                    {
                        currencytype.Text = dr.QuestionDataType;
                        currencytype.Visible = true;
                    }

                    else
                    {
                        currencytype.Visible = false;
                    }
                }

                if (Session["columnlist"] != null && Session["columnlist"] as Dictionary<string, string> != null)
                {
                    Dictionary<string, string> dc = Session["columnlist"] as Dictionary<string, string>;

                   
                        if ( lblno.Attributes["class"] !=null && dc.Keys.Contains(lblno.Attributes["class"]))
                        {
                            //dc[lblno.Attributes["class"]] = lblno.Text;
                        }

                        else
                        {

                            if (dr.QuestionControl != "CheckBox")
                            {
                              
                                if (dc.Keys.Contains(lblno.Attributes["class"]) && dc[lblno.Attributes["class"]] != null)
                                {
                                    lblno.Text = dc[lblno.Attributes["class"]];

                                }

                                else
                                {
                                    dc.Add(dr.DbColumnName + "_" + dr.SubOrder, lblno.Text);

                                   
                                }
                            }

                            else
                            {
                                
                                if (dc.Keys.Contains(radsub.Attributes["class"]) && dc[radsub.Attributes["class"]] != null)
                                {
                                    radsub.SelectedIndex = Convert.ToInt32(dc[radsub.Attributes["class"]]) - 1;
                                }

                                else
                                {
                                    dc.Add(dr.DbColumnName + "_" + dr.SubOrder, (radsub.SelectedItem == null ? "2" : radsub.SelectedItem.Value));
                                }
                            }
                        }
                    

                    Session["columnlist"] = dc;
                    if (dr.QuestionControl != "CheckBox")
                    {
                        if (dc.Keys.Contains(lblno.Attributes["class"]) && dc[lblno.Attributes["class"]] != null)
                        {
                            lblno.Text = dc[lblno.Attributes["class"]];
                            if (string.IsNullOrEmpty(lblno.Text))
                            {

                                currencytype.Text = "";
                            }
                        }
                    }

                    else
                    {

                        if (dc.Keys.Contains(radsub.Attributes["class"]) && dc[radsub.Attributes["class"]] != null)
                        {
                            radsub.SelectedIndex = Convert.ToInt32(dc[radsub.Attributes["class"]])-1;
                        }
                    }
                   
                }

                else
                {
                    Dictionary<string, string> dc = new Dictionary<string, string>();
                    if (dr.QuestionControl != "CheckBox")
                    {
                        dc.Add(dr.DbColumnName.ToString(), lblno.Text);
                    }

                    else
                    {
                        dc.Add(dr.DbColumnName.ToString(), "2");
                    }

                    // txt.Value = dc[txt.ID];
                    Session["columnlist"] = dc;
                }

            }
        }
        protected void Savepagebutton_Click(object sender, EventArgs e)
        {
            string strSourcePath = Server.MapPath("/") + @"bin\Debug" + @"\MasterCSVStruct.csv";
            StreamReader sr = new StreamReader(strSourcePath);
            string[] headers = sr.ReadLine().Split(',');
            DataTable dt = new DataTable();
            foreach (string header in headers)
            {
                dt.Columns.Add(header);
            }
            while (!sr.EndOfStream)
            {
                string[] rows = sr.ReadLine().Split(',');
                DataRow dr = dt.NewRow();
                //for (int i = 0; i < headers.Length; i++)
                //{
                //    dr[i] = rows[i];
                //}
                dt.Rows.Add(dr);
            }
            try
            {
               
            }


            catch(Exception ex)
            {
                //lblsaveerror.Attributes.Add("style","display:block");
                lblsaveerror.Style.Add("display","block");

                lblsaveerror.Text = "Error Occured While saving Record";
            }


            //int tooltip = Convert.ToInt32(btnNext.ToolTip);
            //for (int i = tooltip; i < 7; i++)
            //{

            //    NextButtonQuestionRenderingForSave("Next");
            //}

           
        }


      

        private void save()
        {
           
            Dictionary<string, string> dshowtotal = Session["showtotalgroup"] as Dictionary<string, string>;

            DataTable dt = new DataTable();
            dt.Columns.Add("ID.format");
            dt.Columns.Add("ID.endDate");
            dt.Columns.Add("ID.end");
            dt.Columns.Add("ID.start");
            dt.Columns.Add("ID.date");
            dt.Columns.Add("ID.name");
            if (Session["columnlist"] != null && Session["columnlist"] as Dictionary<string, string> != null)
            {

                Dictionary<string, string> finaldc = Session["columnlist"] as Dictionary<string, string>;
                // int sum = 0;
                int i = 0;
                foreach (string key in finaldc.Keys)
                {
                    if (key == "Q48_1" || key == "Q48_2" || key == "Q48_3" || key == "Q48_4")
                    {
                        if (key == "Q48_4")
                        {

                            dt.Columns.Add("Q48");
                        }
                    }
                    else
                    {
                        dt.Columns.Add(key);
                    }

                    if (key == "Q34_A_11")
                    {
                        dt.Columns.Add("x");
                    }

                    if (key == "Q48_4")
                    {
                        dt.Columns.Add("x_0");
                        dt.Columns.Add("x_1");

                    }

                    if (dshowtotal.Keys.Contains(key))
                    {

                        string questionid = dshowtotal[key];

                        if (dshowtotal[key] == questionid)
                        {
                            var sum = dshowtotal.Where(t => (t.Value.ToString() == questionid)).Count();
                            // var sum1 = dshowtotal.Where(t => (t.Value.ToString() == questionid));

                            if (sum > 0)
                            {
                                i++;

                                if (i.ToString() == sum.ToString())
                                {

                                    if (Session["q17"].ToString() == questionid)
                                    {
                                        dt.Columns.Add("x_17");

                                    }

                                    else
                                    {

                                        dt.Columns.Add("x_" + questionid);
                                    }

                                    // Session["qid"]="x_" + questionid;
                                    if (questionid == "29" || questionid == Session["q26a"].ToString())
                                    {
                                        // dt.Columns.Add("x-" + questionid);
                                        dt.Columns.Add("x-29");
                                    }

                                    if (questionid == "33" || questionid == Session["q26b"].ToString())
                                    {
                                        //dt.Columns.Add("x-" + questionid);
                                        dt.Columns.Add("x-33");
                                    }
                                    i = 0;
                                }


                            }


                        }



                    }

                }


                dt.Columns.Add("Started");
                dt.Columns.Add("Completed");
                dt.Columns.Add("Branched Out");
                dt.Columns.Add("Over Quota");
                dt.Columns.Add("Last Modified");
                dt.Columns.Add("Culture");
                dt.Columns.Add("Last Page");
                dt.Columns.Add("Response Source");
                dt.Columns.Add("Referring URL");
                dt.Columns.Add("Web Browser's User");
                dt.Columns.Add("Respondent's IP Address");
                dt.Columns.Add("Respondent's Hostname");
                dt.Columns.Add("ParticipantURL");
                Dictionary<string, string> dshowtotal1 = Session["showtotalgroup"] as Dictionary<string, string>;
                DataRow dr = dt.NewRow();
                dr["ID.format"] = "Y";
                dr["ID.endDate"] = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);   //"7/2/2015  2:19:00 PM";
                dr["ID.end"] = DateTime.Now.ToString("HH:mm:ss tt");    //"2:31:54 AM";
                dr["ID.start"] = starttime;  // "11:50:12 PM";
                dr["ID.date"] = startdate; // "7/2/2015  2:19:00 PM";
                practiceid = Convert.ToInt32(Session["practiceid"]);
                dr["ID.name"] = practiceid;
                int i1 = 0;
                long total = 0;
                string address = string.Empty;
                foreach (string key in finaldc.Keys)
                {

                    if (key == "Q48_1" || key == "Q48_2" || key == "Q48_3" || key == "Q48_4")
                    {

                        if (finaldc[key] != null)
                        {
                            address = address + finaldc[key] + "-";
                        }

                        if (key == "Q48_4")
                        {
                            if (address.EndsWith("-"))
                            {
                                address = address.Substring(0, address.Length - 1);

                            }
                            dr["Q48"] = address;
                        }
                    }
                    else
                    {
                        dr[key] = finaldc[key];
                    }

                    if (dshowtotal1.Keys.Contains(key))
                    {

                        string questionid = dshowtotal1[key];

                        if (dshowtotal1[key] == questionid)
                        {
                            var sum = dshowtotal1.Where(t => (t.Value.ToString() == questionid)).Count();


                            // var sum1 = finaldc.Where(t => (t.Value.ToString() == questionid)).Sum(a => a.Value.);

                            if (sum > 0)
                            {

                                if (!string.IsNullOrEmpty(finaldc[key]))
                                {
                                    total = total + Convert.ToInt64(Convert.ToDecimal(finaldc[key]));
                                }

                                else
                                {

                                    total = total + 0;
                                }

                                i1++;

                                if (i1.ToString() == sum.ToString())
                                {
                                    if (Session["q17"].ToString() == questionid)
                                    {
                                        dr["x_17"] = total;

                                    }

                                    else
                                    {

                                        dr["x_" + questionid] = total;
                                    }
                                   // dr["x_" + questionid] = total;
                                    // dt.Columns.Add("x_" + questionid);
                                    total = 0;
                                    i1 = 0;
                                }


                            }


                        }



                    }

                }
                string url = HttpContext.Current.Request.Url.AbsoluteUri;
                string host = HttpContext.Current.Request.Url.Host;
                string ua = Request.UserAgent;
                dr["Started"] = startdate;  // "7/2/2015  2:19:00 PM";
                dr["Completed"] = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);    //"7/2/2015  5:00:00 PM";
                dr["Branched Out"] = "";
                dr["Over Quota"] = "";
                dr["Last Modified"] = startdate;  // "7/2/2015  5:00:00 PM";
                dr["Culture"] = "en-US";
                dr["Last Page"] = url;  //"/Community/surveys/1069620006/45dd2ab5007.htm";
                dr["Response Source"] = "0";
                dr["Referring URL"] = "";
                dr["Web Browser's User"] = ua;  // "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1985.125 Safari/537.36";
                dr["Respondent's IP Address"] = "10.26.230.147";
                dr["Respondent's Hostname"] = host;  // "10.26.230.147";
                dr["ParticipantURL"] = "https://deloittesurvey.deloitte.com/Community/se.ashx?s=3FC11B2645DD2AB508D1812AF48355B970";
                dt.Rows.Add(dr);

            }
            DataRow customerRow = dt.Rows[0];


            int Q6VAL = 0;

            if (customerRow["Q6"] != null && !string.IsNullOrEmpty(customerRow["Q6"].ToString()))
            {
                Q6VAL = Convert.ToInt32(customerRow["Q6"]);
            }


            decimal value1 = 0;
            decimal value2 = 0;
            decimal value3 = 0;
            if (!string.IsNullOrEmpty(customerRow["Q7_1"].ToString()))
            {
                value1 = Convert.ToDecimal(Convert.ToDecimal(customerRow["Q7_1"]) / 100);
            }

            if (!string.IsNullOrEmpty(customerRow["Q7_2"].ToString()))
            {

                value2 = Convert.ToDecimal(Convert.ToDecimal(customerRow["Q7_2"]) / 100);
            }

            if (!string.IsNullOrEmpty(customerRow["Q7_3"].ToString()))
            {

                value3 = Convert.ToDecimal(Convert.ToDecimal(customerRow["Q7_3"]) / 100);
            }



            string value11 = Math.Round(Q6VAL * value1).ToString();
            string value12 = Math.Round(Q6VAL * value2).ToString();
            string value13 = Math.Round(Q6VAL * value3).ToString();
            string value4 = (Convert.ToInt32(value11) + Convert.ToInt32(value12) + Convert.ToInt32(value13)).ToString();
            customerRow["Q7_1"] = value11;
            customerRow["Q7_2"] = value12;
            customerRow["Q7_3"] = value13;

            customerRow["x_17"] = value4;
           
            

            dt.AcceptChanges();

            StringBuilder sb = new StringBuilder();

            string[] columnNames = dt.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName).
                                              ToArray();
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dt.Rows)
            {
                string[] fields = row.ItemArray.Select(field => field.ToString()).
                                                ToArray();
                sb.AppendLine(string.Join(",", fields));
            }

            // string filename = string.Format("{0}_{1}", "Survey"+practiceid, DateTime.Now.Ticks);
            practiceid = Convert.ToInt32(Session["practiceid"]);
            DateTime CurrentDateTime = DateTime.Now;
            string filename = "CSV_" + dt.Rows[0]["Q38"] + "_" + practiceid + "_" + CurrentDateTime.ToString("MMddyyyy-hhmmss");
            Session["csvname"] = filename;
            Session["Username"] = dt.Rows[0]["Q38"];
            string csvPath = ConfigurationManager.AppSettings["LocalCSVFilePath"].ToString();    //"C:\\pics\\";
            string finalpath = ConfigurationManager.AppSettings["LocalPDFFilePath"].ToString();  //"C:\\finalpdf\\";
            if (!Directory.Exists(csvPath))
            {
                Directory.CreateDirectory(csvPath);

            }

            if (!Directory.Exists(finalpath))
            {
                Directory.CreateDirectory(finalpath);

            }

            File.WriteAllText(csvPath + filename + ".csv", sb.ToString());
            //  System.Threading.Thread.Sleep(1000);

            int NoOfRowsLimit = 1;
            PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess ObjDataAccess = new PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess("", NoOfRowsLimit);

            //string strSourcePath = Server.MapPath("/") + @"bin\Debug" + @"\MBA_PPA_BlankReport_Template.doc";

            string strSourcePath = Server.MapPath("/") + @"bin\Debug" + @"\MBA_PPA_BlankReport_Template_v3.doc";
            string strSourceInfographicPath = Server.MapPath("/") + @"bin\Debug" + @"\infographic_correctionnew.doc";

            string strSourceexecutivePath = Server.MapPath("/") + @"bin\Debug" + @"\executivesummary_Correction.doc";
            Session["strSourceInfographicPath"] = strSourceInfographicPath;

            Session["strSourceexecutivePath"] = strSourceexecutivePath;
            string IsRecordInserted = ObjDataAccess.ReadCSVAndInsertToSQL(csvPath, filename + ".csv", strSourcePath);

            string filePath = null;
            if (ConfigurationManager.AppSettings["ErrorFilePath"] != null)
            {
                filePath = ConfigurationManager.AppSettings["ErrorFilePath"].ToString();// @"C:\Error.txt";
            }
            if (filePath != null)
            {
                using (StreamWriter writer = new StreamWriter(filePath, true))
                {
                    writer.WriteLine("Message :" + IsRecordInserted);
                    writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                }
            }

            if (IsRecordInserted == "success")
            {
                practiceid = Convert.ToInt32(Session["practiceid"]);
                Report ObjFinalReport = ObjDataAccess.GetOutputData(practiceid.ToString());

                if (ObjFinalReport.lstInput.Count > 0)
                {

                    string msg = CheckVerification(dt);
                    string csv = Server.MapPath("~/pics");
                    string pdf = Server.MapPath("~/finalpdf");
                    //  Dictionary<string, DataSet> dc = new Dictionary<string, DataSet>();
                    // dc = CreateChart(ObjFinalReport);

                    Dictionary<Dictionary<string, DataTable>, DataSet> dc = new Dictionary<Dictionary<string, DataTable>, DataSet>();
                    dc = CreateChart1(ObjFinalReport);

                    if (filePath != null)
                    {
                        using (StreamWriter writer = new StreamWriter(filePath, true))
                        {
                            writer.WriteLine("Message :" + msg);
                            writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                        }
                    }


                    Session["x"] = dc;
                    if (msg == "success")
                    {
                       
                       
                        Session["ObjDataAccess"] = ObjDataAccess;
                        Session["strSourcePath"] = strSourcePath;
                        Session["finalpath"] = finalpath;
                        Session["ObjFinalReport"] = ObjFinalReport;
                        Session["csvPath"] = csvPath;

                    }

                    else
                    {
                        // Response.Redirect("Result.aspx?Result=Verification Falied"); 
                    }
                }
            }
            else
            {
                if (IsRecordInserted == "The Microsoft Jet database engine cannot open the file ''.  It is already opened exclusively by another user, or you need permission to view its data.")
                {
                    string error = "The Microsoft Jet database engine cannot open the file ''.  It is already opened exclusively by another user, or you need permission to view its data.";
                    // Response.Redirect("/Result.aspx?Result="+error);
                }

            }

        }

       

        private string CheckVerification(DataTable dt)
        {
            PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess ObjDataAccess = new PracticePerformanceAssessmentDataAccess.PracticePerformanceAssessmentDataAccess("", 1);
            practiceid = Convert.ToInt32(Session["practiceid"]);
            IEnumerable<GetSurveyData_Result> ld = ObjDataAccess.GetSurveyData(practiceid.ToString());
            List<GetSurveyData_Result> lstquestion = new List<GetSurveyData_Result>();

            dt.Columns.Remove("Completed");
             dt.Columns.Remove("Branched Out");
             dt.Columns.Remove("Over Quota");
             dt.Columns.Remove("Last Modified");
             dt.Columns.Remove("Culture");
             dt.Columns.Remove("Last Page");
             dt.Columns.Remove("Response Source");
             dt.Columns.Remove("Referring URL");
             dt.Columns.Remove("Web Browser's User");
             dt.Columns.Remove("Respondent's IP Address");
             dt.Columns.Remove("Respondent's Hostname");
             dt.Columns.Remove("ParticipantURL");
           
         
             foreach (GetSurveyData_Result data in ld)
            {
                if (data != null )
                {
                    Type t = data.GetType();
                    DataTable dt1 = new DataTable(t.Name);
                    foreach (PropertyInfo pi in t.GetProperties())
                    {
                        dt1.Columns.Add(new DataColumn(pi.Name));
                    }
                   
                        DataRow dr = dt1.NewRow();
                        foreach (DataColumn dc in dt1.Columns)
                        {
                            dr[dc.ColumnName] = data.GetType().GetProperty(dc.ColumnName).GetValue(data, null);
                        }
                        dt1.Rows.Add(dr);

                        dt1.Columns.Remove("RowId");

                        if (dt1.Rows.Count.ToString() == dt.Rows.Count.ToString())
                        {
                            int i = 0;
                            foreach (DataColumn a in dt.Columns)
                            {

                                if (i > 4 && i < dt1.Rows.Count-2)
                                {
                                    string source = dt.Rows[0][a] == null ? "" : dt.Rows[0][a].ToString();
                                    string target = dt1.Rows[0][i] == null ? "" : dt1.Rows[0][i].ToString();

                                    if(target.Contains('.'))
                                    {
                                        source=source+".00";
                                    }
                                    if (source.Trim() == target.Trim())
                                    {
                                        i++;
                                    }

                                    else
                                    {
                                        return "failed";
                                    }
                                }
                              

                            }
                        }
                   
                }

               
                 

                
            }

            return "success";
        }


        public DataTable GetDataTableFromObjects(object[] objects)
        {
            if (objects != null && objects.Length > 0)
            {
                Type t = objects[0].GetType();
                DataTable dt = new DataTable(t.Name);
                foreach (PropertyInfo pi in t.GetProperties())
                {
                    dt.Columns.Add(new DataColumn(pi.Name));
                }
                foreach (var o in objects)
                {
                    DataRow dr = dt.NewRow();
                    foreach (DataColumn dc in dt.Columns)
                    {
                        dr[dc.ColumnName] = o.GetType().GetProperty(dc.ColumnName).GetValue(o, null);
                    }
                    dt.Rows.Add(dr);
                }
                return dt;
            }
            return null;
        }

        public static void CopyAll(DirectoryInfo source, DirectoryInfo target, string fileName,string ext)
        {
            // Check if the target directory exists, if not, create it.
            if (Directory.Exists(target.FullName) == false)
            {
                Directory.CreateDirectory(target.FullName);
            }
            // Copy each file into it's new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                if (fi.Name.ToLower() == fileName.ToLower() + ext.ToLower())
                {
                    //Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                    fi.CopyTo(Path.Combine(target.ToString(), fi.Name), true);
                }
            }
            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir =
                target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir, fileName,ext.ToLower());
            }
        }

       

        public decimal Returntotal(List<string> arr)
        {
            decimal total = 0;
            foreach (string x in arr)
            {
                string x1 = "";
                if (x.Contains('-'))
                {
                    x1 = x.Split('-')[0];
                }
                else
                {

                    x1 = x;
                }

                if (x1.Contains('s'))
                {
                    total = total + Convert.ToDecimal(x1.Split('s')[0]);
                }
                else if (x1.Contains('t'))
                {
                    total = total + Convert.ToDecimal(x1.Split('t')[0]);
                }

                else if (x1.Contains('n'))
                {

                    total = total + Convert.ToDecimal(x1.Split('n')[0]);
                }

                else if (x1.Contains('r'))
                {

                    total = total + Convert.ToDecimal(x1.Split('r')[0]);
                }
                else if (x1 == null)
                {

                    total = total;
                }

                else
                {
                    total = total + Convert.ToDecimal(x1);
                }

            }

            return total;
        }
        public long ReturnValue(string arr)
        {
            //decimal total = 0;
            if (string.IsNullOrEmpty(arr))
            {
                arr = "0";
            }

            else
            {
                if (arr.Contains('-'))
                {

                    arr = arr.Split('-')[0].ToString();
                }


                if (arr.Contains('s'))
                {
                    arr = arr.Split('s')[0];
                }
                else if (arr.Contains('t'))
                {

                    arr = arr.Split('t')[0];

                }

                else if (arr.Contains('n'))
                {

                    arr = arr.Split('n')[0];
                }

                else if (arr.Contains('r'))
                {

                    arr = arr.Split('r')[0];
                }
                else if (arr == null)
                {

                    arr = "0";
                }

                else
                {
                    arr = arr;
                }


            }


            return Convert.ToInt64(arr);
        }


   

        public string Returnaverage(string arr)
        {

            decimal total = 0;
            string x1 = "";
            if (arr.Contains('-'))
            {
                x1 = arr.Split('-')[0];
            }
            else
            {

                x1 = arr;
            }

            if (x1.Contains('s'))
            {
                return x1;
            }
            else if (x1.Contains('t'))
            {
                return x1;
            }

            else if (x1.Contains('n'))
            {

                return x1;
            }

            else if (x1.Contains('r'))
            {

                return x1;
            }
            else if (x1 == null)
            {

                return string.Empty;
            }

            else if (x1.StartsWith("1") && x1.Length <= 1)
            {
                return x1 + "st";
            }
            else if (x1.StartsWith("1") && x1.Length > 1)
            {
                return x1 + "th";
            }
            else if (x1.StartsWith("2"))
            {
                return x1 + "nd";
            }
            else if (x1.StartsWith("3"))
            {
                return x1 + "rd";
            }
            else
            {
                return x1 + "th";
            }




            return x1;
        }

        private Dictionary<Dictionary<string, DataTable>, DataSet> CreateChart1(Report objReport)
        {
            DataSet ds = new DataSet();
            DataTable dt1 = new DataTable();
            dt1.Columns.Add("SortedData", typeof(long));
            dt1.Columns.Add("DisplayText");
            dt1.Columns.Add("DisplayValue");
            dt1.Columns.Add("Pageno");
            dt1.Columns.Add("75percentile");
            dt1.Columns.Add("75percentilesorted", typeof(long));

            DataTable dt2 = new DataTable();
            dt2.Columns.Add("SortedData", typeof(long));
            dt2.Columns.Add("DisplayText");
            dt2.Columns.Add("DisplayValue");
            dt2.Columns.Add("Pageno");
            dt2.Columns.Add("75percentile");
            dt2.Columns.Add("75percentilesorted", typeof(long));



            try
            {
                long x = 0;

                DataTable dt = new DataTable();




                dt = new DataTable();
                dt.Columns.Add("Y1", typeof(long));
                dt.Columns.Add("X");
                decimal total1 = ReturnValue(objReport.lstOutput[0].col9d) + ReturnValue(objReport.lstOutput[0].col7b)
                    + ReturnValue(objReport.lstOutput[0].col8c)
                    + ReturnValue(objReport.lstOutput[0].col4b) + ReturnValue(objReport.lstOutput[0].col9b) + ReturnValue(objReport.lstOutput[0].col6b) +
                    ReturnValue(objReport.lstOutput[0].col5b) + ReturnValue(objReport.lstOutput[0].col3b);
                decimal total11 = Math.Round(total1 / 8, 0);

                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col3b), "Gross Revenue per Complete Exam");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col5b), "Annual Gross Revenue per Active Patient");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col6b), "Annual Complete Exams per 100 Active Patients");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col9b), "Gross Revenue per Non-OD Staff Hour");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col4b), "Complete Exams per OD Hour");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col8c), "Annual Gross Revenue per FTE OD");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col7b), "Gross Revenue per OD Hour");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col9d), "Gross Revenue per Square Foot of Office Space");

                dt.Rows.Add(Convert.ToInt64(total11), "Practice Productivity Metrics Average Percentile Ranking");
                DataView _dv = dt.DefaultView;
                _dv.Sort = "Y1 ASC";

                dt = _dv.ToTable();






                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col3b), "Gross Revenue per Complete Exam", objReport.lstOutput[0].col3b, "6", objReport.lstOutput[0].col3d == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col3d).ToString("#,0.##"), objReport.lstOutput[0].col3d == null ? 0 : objReport.lstOutput[0].col3d);   //«M_3d»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col5b), "Annual Gross Revenue per Active Patient", objReport.lstOutput[0].col5b, "8", objReport.lstOutput[0].col5d == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col5d).ToString("#,0.##"), objReport.lstOutput[0].col5d == null ? 0 : objReport.lstOutput[0].col5d);   //«M_5d»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col6b), "Annual Complete Exams per 100 Active Patients", objReport.lstOutput[0].col6b, "9", objReport.lstOutput[0].col6e == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col6e).ToString("#,0.##"), objReport.lstOutput[0].col6e == null ? 0 : objReport.lstOutput[0].col6e); //«M_6e»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col9b), "Gross Revenue per Non-OD Staff Hour", objReport.lstOutput[0].col9b, "12", "0");   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col4b), "Complete Exams per OD Hour", objReport.lstOutput[0].col4b, "7", objReport.lstOutput[0].col4e == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col4e).ToString("#,0.##"), objReport.lstOutput[0].col4e == null ? 0 : objReport.lstOutput[0].col4e);   //«M_4e»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col8c), "Annual Gross Revenue per FTE OD", objReport.lstOutput[0].col8c, "11", objReport.lstOutput[0].col8e == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col8e).ToString("#,0.##"), objReport.lstOutput[0].col8e == null ? 0 : objReport.lstOutput[0].col8e);  //«M_8e»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col7b), "Gross Revenue per OD Hour", objReport.lstOutput[0].col7b, "10", objReport.lstOutput[0].col7d == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col7d).ToString("#,0.##"), objReport.lstOutput[0].col7d == null ? 0 : objReport.lstOutput[0].col7d);   //«M_7d»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col9d), "Gross Revenue per Square Foot of Office Space", objReport.lstOutput[0].col9d, "12", "0");   //0
                dt2 = dt1.Copy();

                dt1.Rows.Add(Convert.ToInt64(total11), "Practice Productivity Metrics Average Percentile Ranking", Returnaverage(total11.ToString()));
                DataView _dtv = dt1.DefaultView;
                _dtv.Sort = "SortedData DESC";

                dt1 = _dtv.ToTable();



                ds.Tables.Add(dt1);
                GenerateChartExecutive(dt, 740, 300, "chart1");


                dt = new DataTable();
                dt.Columns.Add("Y1", typeof(long));
                dt.Columns.Add("X");

                decimal total2 = ReturnValue(objReport.lstOutput[0].col15d) + ReturnValue(objReport.lstOutput[0].col16d)
                    + ReturnValue(objReport.lstOutput[0].col12b)
                    + ReturnValue(objReport.lstOutput[0].col13c) + ReturnValue(objReport.lstOutput[0].col14b)
                    + ReturnValue(objReport.lstOutput[0].col17b) +
                    ReturnValue(objReport.lstOutput[0].col20a) + ReturnValue(objReport.lstOutput[0].col19b) +
                    ReturnValue(objReport.lstOutput[0].col18b) + ReturnValue(objReport.lstOutput[0].col21a);
                decimal total21 = Math.Round(total2 / 10, 0);

                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col18b), "High Index Lens % of Eyewear Rxes");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col19b), "Photochromic Lens % of Eyewear Rxes");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col20a), "Eyewear Multiple Pair Sales % Eyewear Buyers");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col17b), "No-Glare (anti-reflective) Lens % of eyewear Rxes");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col14b), "Eyewear Gross Revenue per Eyewear Rx");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col13c), "Eyewear Rxes per 100 Complete Exams");

                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col12b), "Eyewear Sales % of Gross Revenue");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col16d), "Progressive Lens % of Presbyopic Rxes");

                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col15d), "Eyewear Gross Profit Margin %");

                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col21a), "% of Contact Lens Patients Purchasing Eyewear");

                dt.Rows.Add(Convert.ToInt64(total21), "Eyewear Metrics Average Percentile Ranking");
                DataView _dv1 = dt.DefaultView;
                _dv1.Sort = "Y1 ASC";

                dt = _dv1.ToTable();





                dt1 = new DataTable();
                dt1.Columns.Add("SortedData", typeof(long));
                dt1.Columns.Add("DisplayText");
                dt1.Columns.Add("DisplayValue");
                dt1.Columns.Add("Pageno");
                dt1.Columns.Add("75percentile");
                dt1.Columns.Add("75percentilesorted", typeof(long));

                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col18b), "High Index Lens % of Eyewear Rxes", objReport.lstOutput[0].col18b, "21", objReport.lstOutput[0].col18d == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col18d).ToString("#,0.##"), objReport.lstOutput[0].col18d == null ? 0 : objReport.lstOutput[0].col18d);   //«M_18d»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col19b), "Photochromic Lens % of Eyewear Rxes", objReport.lstOutput[0].col19b, "22", objReport.lstOutput[0].col19d == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col19d).ToString("#,0.##"), objReport.lstOutput[0].col19d == null ? 0 : objReport.lstOutput[0].col19d);    //«M_19d»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col20a), "Eyewear Multiple Pair Sales % Eyewear Buyers", objReport.lstOutput[0].col20a, "23", objReport.lstOutput[0].col20f == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col20f).ToString("#,0.##"), objReport.lstOutput[0].col20f == null ? 0 : objReport.lstOutput[0].col20f);   //«M_20f»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col17b), "No-Glare (anti-reflective) Lens % of eyewear Rxes", objReport.lstOutput[0].col17b, "20", objReport.lstOutput[0].col17d == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col17d).ToString("#,0.##"), objReport.lstOutput[0].col17d == null ? 0 : objReport.lstOutput[0].col17d);   //«M_17d»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col14b), "Eyewear Gross Revenue per Eyewear Rx", objReport.lstOutput[0].col14b, "17", objReport.lstOutput[0].col14d == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col14d).ToString("#,0.##"), objReport.lstOutput[0].col14d == null ? 0 : objReport.lstOutput[0].col14d);   //«M_14d»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col13c), "Eyewear Rxes per 100 Complete Exams", objReport.lstOutput[0].col13c, "16", objReport.lstOutput[0].col13g == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col13g).ToString("#,0.##"), objReport.lstOutput[0].col13g == null ? 0 : objReport.lstOutput[0].col13g);   //«M_13g»

                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col12b), "Eyewear Sales % of Gross Revenue", objReport.lstOutput[0].col12b, "15", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col16d), "Progressive Lens % of Presbyopic Rxes", objReport.lstOutput[0].col16d, "19", objReport.lstOutput[0].col16f == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col16f).ToString("#,0.##"), objReport.lstOutput[0].col16f == null ? 0 : objReport.lstOutput[0].col16f);  //«M_16f»

                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col15d), "Eyewear Gross Profit Margin %", objReport.lstOutput[0].col15d, "18", objReport.lstOutput[0].col15f == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col15f).ToString("#,0.##"), objReport.lstOutput[0].col15f == null ? 0 : objReport.lstOutput[0].col15f);   //«M_15f»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col21a), "% of Contact Lens Patients Purchasing Eyewear", objReport.lstOutput[0].col21a, "24", objReport.lstOutput[0].col21f == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col21f).ToString("#,0.##"), objReport.lstOutput[0].col21f == null ? 0 : objReport.lstOutput[0].col21f);   //«M_21f»
                dt1.AsEnumerable().CopyToDataTable(dt2, LoadOption.PreserveChanges);
                dt1.Rows.Add(Convert.ToInt64(total21), "Eyewear Metrics Average Percentile Ranking", Returnaverage(total21.ToString()));
                DataView _dtv1 = dt1.DefaultView;
                _dtv1.Sort = "SortedData DESC";

                dt1 = _dtv1.ToTable();

                total2 = total2 + ReturnValue(objReport.lstOutput[0].col21a);

                ds.Tables.Add(dt1);


                GenerateChartExecutive(dt, 740, 374, "chart2");

                dt = new DataTable();
                dt.Columns.Add("Y1", typeof(long));
                dt.Columns.Add("X");


                decimal total3 = ReturnValue(objReport.lstOutput[0].col25c) + ReturnValue(objReport.lstOutput[0].col24b)
                    + ReturnValue(objReport.lstOutput[0].col29a)
                    + ReturnValue(objReport.lstOutput[0].col30b) + ReturnValue(objReport.lstOutput[0].col27b)
                    + ReturnValue(objReport.lstOutput[0].col26c) +
                    ReturnValue(objReport.lstOutput[0].col26a) + ReturnValue(objReport.lstOutput[0].col29b) +
                    ReturnValue(objReport.lstOutput[0].col29c) + ReturnValue(objReport.lstOutput[0].col28b) + ReturnValue(objReport.lstOutput[0].col28c)
                    + ReturnValue(objReport.lstOutput[0].col30b);
                decimal total31 = Math.Round(total3 / 12, 0);


                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col29c), "Monthly Lens Wearer % of Soft Lens Wearers");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col29b), "Daily Disposable Wearer % of Soft Lens Wearers");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col26a), "Contact Lens Wearer % of Active Patients");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col26c), "Contact Lens Exams % of Total Complete Eye Exams");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col27b), "Annual Contact Lens Sales per Contact Lens Eye Exam");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col30b), "Soft Toric Lens Wearer % of Soft Lens Wearers");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col29a), "Silicone Hydrogel Wearer % of Soft Lens Wearers");

                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col24b), "Contact Lens Sales % of Gross Revenue");

                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col25c), "Contact Lens Gross Profit Margin %");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col28b), "Contact Lens New Fits per 100 Contact Lens Exams");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col28c), "Contact Lens Refits % of Contact Lens Exams");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col30b), "Soft Multi-focal Lens Wearer % of Soft Lens Wearers");

                dt.Rows.Add(Convert.ToInt64(total31), "Contact Lens Metrics Average Percentile");
                DataView _dv2 = dt.DefaultView;
                _dv2.Sort = "Y1 ASC";

                dt = _dv2.ToTable();



                dt1 = new DataTable();
                dt1.Columns.Add("SortedData", typeof(long));
                dt1.Columns.Add("DisplayText");
                dt1.Columns.Add("DisplayValue");
                dt1.Columns.Add("Pageno");
                dt1.Columns.Add("75percentile");
                dt1.Columns.Add("75percentilesorted", typeof(long));
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col29c), "Monthly Lens Wearer % of Soft Lens Wearers", objReport.lstOutput[0].col29c, "32", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col29b), "Daily Disposable Wearer % of Soft Lens Wearers", objReport.lstOutput[0].col29b, "32", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col26a), "Contact Lens Wearer % of Active Patients", objReport.lstOutput[0].col26a, "29", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col26c), "Contact Lens Exams % of Total Complete Eye Exams", objReport.lstOutput[0].col26c, "29", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col27b), "Annual Contact Lens Sales per Contact Lens Eye Exam", objReport.lstOutput[0].col27b, "30", objReport.lstOutput[0].col27d == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col27d).ToString("#,0.##"), objReport.lstOutput[0].col27d == null ? 0 : objReport.lstOutput[0].col27d);   //«M_27d»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col30b), "Soft Toric Lens Wearer % of Soft Lens Wearers", objReport.lstOutput[0].col30b, "33", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col29a), "Silicone Hydrogel Wearer % of Soft Lens Wearers", objReport.lstOutput[0].col29a, "32", "0", 0);  //0

                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col24b), "Contact Lens Sales % of Gross Revenue", objReport.lstOutput[0].col24b, "27", "0", 0);  //0

                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col25c), "Contact Lens Gross Profit Margin %", objReport.lstOutput[0].col25c, "28", objReport.lstOutput[0].col25e == null ? "" : "$" + Convert.ToDecimal(objReport.lstOutput[0].col25e).ToString("#,0.##"), objReport.lstOutput[0].col25e == null ? 0 : objReport.lstOutput[0].col25e);   //«M_25e»
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col28b), "Contact Lens New Fits per 100 Contact Lens Exams", objReport.lstOutput[0].col28b, "31", "0", 0); //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col28c), "Contact Lens Refits % of Contact Lens Exams", objReport.lstOutput[0].col28c, "31", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col30b), "Soft Multi-focal Lens Wearer % of Soft Lens Wearers", objReport.lstOutput[0].col30b, "33", "0", 0);  //0
                dt1.AsEnumerable().CopyToDataTable(dt2, LoadOption.PreserveChanges);
                dt1.Rows.Add(Convert.ToInt64(total31), "Contact Lens Metrics Average Percentile", Returnaverage(total31.ToString()));
                DataView _dtv2 = dt1.DefaultView;
                _dtv2.Sort = "SortedData DESC";

                dt1 = _dtv2.ToTable();
                total3 = total3 + ReturnValue(objReport.lstOutput[0].col25c) + ReturnValue(objReport.lstOutput[0].col28b) + ReturnValue(objReport.lstOutput[0].col28c);


                ds.Tables.Add(dt1);


                GenerateChartExecutive(dt, 740, 442, "chart3");


                dt = new DataTable();
                dt.Columns.Add("Y1", typeof(long));
                dt.Columns.Add("X");

                decimal total4 = ReturnValue(objReport.lstOutput[0].col33c) + ReturnValue(objReport.lstOutput[0].col34b)
               + ReturnValue(objReport.lstOutput[0].col33g)
               + ReturnValue(objReport.lstOutput[0].col34e);
                decimal total41 = Math.Round(total4 / 4, 0);

                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col34e), "Annual Pharmaceutical Rxes per 1,000 Active Patients");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col33g), "Medical Eye Care Visits % of Total Patient Visits");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col34b), "Annual Medical Eye Care Visits per 1,000 Active Patients");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col33c), "Non-refractive Fee Revenue % of Total Gross Revenue");


                dt.Rows.Add(Convert.ToInt64(total41), "Medical Eye Care Metrics Average Percentile Ranking");
                DataView _dv3 = dt.DefaultView;
                _dv3.Sort = "Y1 ASC";

                dt = _dv3.ToTable();



                dt1 = new DataTable();
                dt1.Columns.Add("SortedData", typeof(long));
                dt1.Columns.Add("DisplayText");
                dt1.Columns.Add("DisplayValue");
                dt1.Columns.Add("Pageno");
                dt1.Columns.Add("75percentile");
                dt1.Columns.Add("75percentilesorted", typeof(long));
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col34e), "Annual Pharmaceutical Rxes per 1,000 Active Patients", objReport.lstOutput[0].col34e, "37", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col33g), "Medical Eye Care Visits % of Total Patient Visits", objReport.lstOutput[0].col33g, "36", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col34b), "Annual Medical Eye Care Visits per 1,000 Active Patients", objReport.lstOutput[0].col34b, "36", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col33c), "Non-refractive Fee Revenue % of Total Gross Revenue", objReport.lstOutput[0].col33c, "36", "0", 0);   //0
                dt1.AsEnumerable().CopyToDataTable(dt2, LoadOption.PreserveChanges);
                dt1.Rows.Add(Convert.ToInt64(total41), "Medical Eye Care Metrics Average Percentile Ranking", Returnaverage(total41.ToString()));
                DataView _dtv3 = dt1.DefaultView;
                _dtv3.Sort = "SortedData DESC";

                dt1 = _dtv3.ToTable();


                ds.Tables.Add(dt1);

                GenerateChartExecutive(dt, 740, 170, "chart4");

                dt = new DataTable();
                dt.Columns.Add("Y1", typeof(long));
                dt.Columns.Add("X");

                decimal total5 = ReturnValue(objReport.lstOutput[0].col37a) + ReturnValue(objReport.lstOutput[0].col36d)
            + ReturnValue(objReport.lstOutput[0].col36e)
            + ReturnValue(objReport.lstOutput[0].col36b) + ReturnValue(objReport.lstOutput[0].col37h) + ReturnValue(objReport.lstOutput[0].col37d);
                decimal total51 = Math.Round(total5 / 6, 0);

                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col37d), "% of Total New Patients Attracted by Practice Website");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col37h), "Recall Staff Minutes per Complete Eye Exam");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col36b), "Marketing Spending % of Gross Revenue");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col36e), "New Patient Exams % of Total Exams");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col36d), "Annual Marketing Spending per Complete Exam");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col37a), "Website Expense");

                dt.Rows.Add(Convert.ToInt64(total51), "Marketing Average Percentile Ranking");
                DataView _dv4 = dt.DefaultView;
                _dv4.Sort = "Y1 ASC";

                dt = _dv4.ToTable();


                dt1 = new DataTable();
                dt1.Columns.Add("SortedData", typeof(long));
                dt1.Columns.Add("DisplayText");
                dt1.Columns.Add("DisplayValue");

                dt1.Columns.Add("Pageno");
                dt1.Columns.Add("75percentile");
                dt1.Columns.Add("75percentilesorted", typeof(long));
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col37d), "% of Total New Patients Attracted by Practice Website", objReport.lstOutput[0].col37d, "40", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col37h), "Recall Staff Minutes per Complete Eye Exam", objReport.lstOutput[0].col37h, "40", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col36b), "Marketing Spending % of Gross Revenue", objReport.lstOutput[0].col36b, "39", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col36e), "New Patient Exams % of Total Exams", objReport.lstOutput[0].col36e, "39", "0", 0); //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col36d), "Annual Marketing Spending per Complete Exam", objReport.lstOutput[0].col36d, "39", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col37a), "Website Expense", objReport.lstOutput[0].col37a, "40", "0");  //0
                dt1.AsEnumerable().CopyToDataTable(dt2, LoadOption.PreserveChanges);
                dt1.Rows.Add(Convert.ToInt64(total51), "Marketing Average Percentile Ranking", Returnaverage(total51.ToString()));
                DataView _dtv4 = dt1.DefaultView;
                _dtv4.Sort = "SortedData DESC";

                dt1 = _dtv4.ToTable();


                ds.Tables.Add(dt1);

                GenerateChartExecutive(dt, 740, 238, "chart5");


                dt = new DataTable();
                dt.Columns.Add("Y1", typeof(long));
                dt.Columns.Add("X");


                decimal total6 = ReturnValue(objReport.lstOutput[0].col45d) + ReturnValue(objReport.lstOutput[0].col44b)
            + ReturnValue(objReport.lstOutput[0].col41b)
            + ReturnValue(objReport.lstOutput[0].col40b) + ReturnValue(objReport.lstOutput[0].col44d) +
            ReturnValue(objReport.lstOutput[0].col42b)
            + ReturnValue(objReport.lstOutput[0].col42b)
            + ReturnValue(objReport.lstOutput[0].col40c)
            + ReturnValue(objReport.lstOutput[0].col41a)
            + ReturnValue(objReport.lstOutput[0].col40a)
            + ReturnValue(objReport.lstOutput[0].col45g)
            + ReturnValue(objReport.lstOutput[0].col43e)
            + ReturnValue(objReport.lstOutput[0].col43g)
            + ReturnValue(objReport.lstOutput[0].col45g);
                decimal total61 = Math.Round(total6 / 13, 0);

                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col43g), "Cost-of Goods % of Gross Revenue");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col43e), "Accounts Receivables Days Outstanding");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col45g), "Chair Cost per Complete Exam");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col40a), "Non-contact Lens Exam Fee");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col41a), "Contact Lens New Fit Exam Fee –Soft Multi-focal");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col40c), "Contact Lens New Fit Exam Fee –Soft Toric");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col42b), "Average Collected Revenue per Complete Exam");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col44d), "Occupancy % of Gross Revenue");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col40b), "Contact Lens New Fit Exam Fee –Sphere");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col41b), "Contact Lens Exam Fee –No Refitting");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col44b), "Staffing % of Gross Revenue");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col45d), "Net Income % of Gross Revenue");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col45g), "% of Exams Provided with Managed Care Discount");


                dt.Rows.Add(Convert.ToInt64(total61), "Financial Average Percentile Ranking");
                DataView _dv5 = dt.DefaultView;
                _dv5.Sort = "Y1 ASC";

                dt = _dv5.ToTable();



                dt1 = new DataTable();
                dt1.Columns.Add("SortedData", typeof(long));
                dt1.Columns.Add("DisplayText");
                dt1.Columns.Add("DisplayValue");
                dt1.Columns.Add("Pageno");
                dt1.Columns.Add("75percentile");
                dt1.Columns.Add("75percentilesorted", typeof(long));
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col43g), "Cost-of Goods % of Gross Revenue", objReport.lstOutput[0].col43g, "47", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col43e), "Accounts Receivables Days Outstanding", objReport.lstOutput[0].col43e, "46", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col45g), "Chair Cost per Complete Exam", objReport.lstOutput[0].col45g, "49", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col40a), "Non-contact Lens Exam Fee", objReport.lstOutput[0].col40a, "43", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col41a), "Contact Lens New Fit Exam Fee –Soft Multi-focal", objReport.lstOutput[0].col41a, "44", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col40c), "Contact Lens New Fit Exam Fee –Soft Toric", objReport.lstOutput[0].col40c, "43", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col42b), "Average Collected Revenue per Complete Exam", objReport.lstOutput[0].col42b, "45", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col44d), "Occupancy % of Gross Revenue", objReport.lstOutput[0].col44d, "48", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col40b), "Contact Lens New Fit Exam Fee –Sphere", objReport.lstOutput[0].col40b, "43", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col41b), "Contact Lens Exam Fee –No Refitting", objReport.lstOutput[0].col41b, "44", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col44b), "Staffing % of Gross Revenue", objReport.lstOutput[0].col44b, "47", "0", 0);  //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col45d), "Net Income % of Gross Revenue", objReport.lstOutput[0].col45d, "48", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col45g), "% of Exams Provided with Managed Care Discount", objReport.lstOutput[0].col45g, "45", "0", 0);//0
                dt1.AsEnumerable().CopyToDataTable(dt2, LoadOption.PreserveChanges);

                dt1.Rows.Add(Convert.ToInt64(total61), "Financial Average Percentile Ranking", Returnaverage(total61.ToString()));
                DataView _dtv5 = dt1.DefaultView;
                _dtv5.Sort = "SortedData DESC";

                dt1 = _dtv5.ToTable();
                total6 = total6 + ReturnValue(objReport.lstOutput[0].col45g);


                ds.Tables.Add(dt1);



                GenerateChartExecutive(dt, 740, 476, "chart6");

                dt = new DataTable();
                dt.Columns.Add("Y1", typeof(long));
                dt.Columns.Add("X");
                decimal total7 = ReturnValue(objReport.lstOutput[0].col50d) + ReturnValue(objReport.lstOutput[0].col50b)
        + ReturnValue(objReport.lstOutput[0].col49b)
        + ReturnValue(objReport.lstOutput[0].col49b);
                decimal total71 = Math.Round(total7 / 4, 0);




                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col49b), "Financial Management Score");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col49b), "Total “Best Practices” Score");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col50b), "Marketing Management");
                dt.Rows.Add(ReturnValue(objReport.lstOutput[0].col50d), "Staff Management");

                dt.Rows.Add(Convert.ToInt64(total71), "Average Percentile Ranking");
                DataView _dv6 = dt.DefaultView;
                _dv6.Sort = "Y1 ASC";

                dt = _dv6.ToTable();


                dt1 = new DataTable();
                dt1.Columns.Add("SortedData", typeof(long));
                dt1.Columns.Add("DisplayText");
                dt1.Columns.Add("DisplayValue");
                dt1.Columns.Add("Pageno");

                dt1.Columns.Add("75percentile");
                dt1.Columns.Add("75percentilesorted", typeof(long));

                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col49b), "Financial Management Score", objReport.lstOutput[0].col49b, "52", "0", 0);   //0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col49b), "Total “Best Practices” Score", objReport.lstOutput[0].col49b, "52", "0", 0);//0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col50b), "Marketing Management", objReport.lstOutput[0].col50b, "53", "0", 0);//0
                dt1.Rows.Add(ReturnValue(objReport.lstOutput[0].col50d), "Staff Management", objReport.lstOutput[0].col50d, "53", "0", 0);  //0
                dt1.AsEnumerable().CopyToDataTable(dt2, LoadOption.PreserveChanges);
                dt1.Rows.Add(Convert.ToInt64(total71), "Average Percentile Ranking", Returnaverage(total71.ToString()));

                DataView _dtv6 = dt1.DefaultView;
                _dtv6.Sort = "SortedData DESC";

                dt1 = _dtv6.ToTable();

                ds.Tables.Add(dt1);

                DataView _dfinal = dt2.DefaultView;
                _dfinal.Sort = "SortedData DESC";

                dt2 = _dfinal.ToTable();


                GenerateChartExecutive(dt, 740, 170, "chart7");



                dt = new DataTable();
                dt.Columns.Add("Y1");
                dt.Columns.Add("X");

                dt.Rows.Add(Convert.ToInt64(total11), "Total Practice Productivity");
                dt.Rows.Add(Convert.ToInt64(total21), "Eyewear");
                dt.Rows.Add(Convert.ToInt64(total31), "Contact Lenses");
                dt.Rows.Add(Convert.ToInt64(total41), "Medical Eye Care");
                dt.Rows.Add(Convert.ToInt64(total51), "Marketing");
                dt.Rows.Add(Convert.ToInt64(total61), "Financial");
                dt.Rows.Add(Convert.ToInt64(total71), "Management 'Best Practices'");

                GenerateChart(dt);
                decimal x1 = 0;
                x1 = x1 + total11 + total21 + total31 + total41 + total51 + total61 + total71;

                x = Convert.ToInt64(Math.Round(x1 / 7, 0));
                Dictionary<Dictionary<string, DataTable>, DataSet> dc = new Dictionary<Dictionary<string, DataTable>, DataSet>();
                Dictionary<string, DataTable> dct = new Dictionary<string, DataTable>();
                dct.Add(x.ToString(), dt2);

                dc.Add(dct, ds);
                return dc;

            }
            catch (Exception ex)
            {

                string filePath = null;
                if (ConfigurationManager.AppSettings["ErrorFilePath"] != null)
                {
                    filePath = ConfigurationManager.AppSettings["ErrorFilePath"].ToString();// @"C:\Error.txt";
                }
                if (filePath != null)
                {
                    using (StreamWriter writer = new StreamWriter(filePath, true))
                    {
                        writer.WriteLine("ChartCreation :" + ex.Message);
                        writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                    }
                }
                Dictionary<Dictionary<string, DataTable>, DataSet> dc = new Dictionary<Dictionary<string, DataTable>, DataSet>();
                return dc;
            }
        }

        protected void GenerateChart(DataTable dtChartDataSource)
        {
            Chart chart = new Chart()
            {
                Width = 700,
                Height = 350
            };

            // chart.Legends.Add(new Legend(){Name = "Legend"});
            // chart.Legends[0].Docking = Docking.Bottom;
            ChartArea chartArea = new ChartArea() { Name = "ChartArea" };
            //Remove X-axis grid lines
            chartArea.AxisX.MajorGrid.LineWidth = 0;
            //Remove Y-axis grid lines
            chartArea.AxisY.MajorGrid.LineWidth = 0;
            //Chart Area Back Color
            chartArea.BackColor = System.Drawing.Color.FromName("White");
            chart.ChartAreas.Add(chartArea);
            chart.Palette = ChartColorPalette.BrightPastel;
            chart.RightToLeft = RightToLeft.Yes;
            string series = string.Empty;
            chart.BorderlineWidth = 0;
            //chart.Location = new System.Drawing.Point(322, 208);
            chart.Width= 650;
            chart.Height = 260;
            chart.Palette = System.Web.UI.DataVisualization.Charting.ChartColorPalette.None;
        //    chart.PaletteCustomColors = new System.Drawing.Color[] {
        //System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(143)))), ((int)(((byte)(209)))))};

            chart.PaletteCustomColors = new System.Drawing.Color[] { System.Drawing.ColorTranslator.FromHtml("#002776") };
            chartArea.AxisX.TitleAlignment = StringAlignment.Near;
            chartArea.AxisX.LabelAutoFitStyle = LabelAutoFitStyles.IncreaseFont;

            chartArea.AxisX.MajorGrid.Enabled = false;
            chartArea.AxisY.MajorGrid.Enabled = false;
            chartArea.AxisX.TitleAlignment = System.Drawing.StringAlignment.Far;
            chartArea.AxisX.LineWidth = 0;
            chartArea.AxisY.LineWidth = 0;
            chartArea.AxisY.Enabled = AxisEnabled.False;
            chartArea.AxisX.IsMarginVisible = false;

            chart.RightToLeft = RightToLeft.Yes;
            chart.BackGradientStyle = System.Web.UI.DataVisualization.Charting.GradientStyle.TopBottom;
            chartArea.AxisX.IsStartedFromZero = true;
            //chartArea.AxisX.LabelStyle.ForeColor=Color

            //create series and add data points to the series
            if (dtChartDataSource != null)
            {
                foreach (DataColumn dc in dtChartDataSource.Columns)
                {
                    //a series to the chart
                    if (chart.Series.FindByName(dc.ColumnName) == null)
                    {
                        series = dc.ColumnName;
                        chart.Series.Add(series);
                        chart.Series[series].ChartType = SeriesChartType.Column;
                        chart.Series[series]["PixelPointWidth"] = "100";
                        chart.Series[series]["PointWidth"] = "50";
                        chart.Series[series]["DrawingStyle"] = "Cylinder";

                        chart.Series[series].IsXValueIndexed = true;
                        chart.Series[series].LabelForeColor = System.Drawing.ColorTranslator.FromHtml("#002776");

                    }
                    //Add data points to the series
                    foreach (DataRow dr in dtChartDataSource.Rows)
                    {
                        double dataPoint = 0;
                        double.TryParse(dr[dc.ColumnName].ToString(), out dataPoint);
                        DataPoint objDataPoint = new DataPoint() { AxisLabel = "series", YValues = new double[] { dataPoint } };

                        chart.Series[series].Points.Add(dataPoint);




                    }

                }

                chart.Series[0].MarkerStyle = MarkerStyle.None;
                chart.Series[0].IsValueShownAsLabel = true;
                chart.Series[0]["BarLabelStyle"] = "Center";


                chartArea.ShadowColor = System.Drawing.Color.White;

                chartArea.AxisX.IsMarginVisible = true;


                //chartArea.AxisX.LabelStyle.Interval = 1;
                chartArea.AxisX.LabelStyle.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);


                //LabelStyle yAxisStyle = new LabelStyle();
                //yAxisStyle.ForeColor = System.Drawing.ColorTranslator.FromHtml("#444444");
                //yAxisStyle.Font = new System.Drawing.Font("Arial", 11, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);


                chartArea.Position = new System.Web.UI.DataVisualization.Charting.ElementPosition(0, 0, 100, 100);
                chartArea.AxisX.Minimum = 0;

                chartArea.AxisX.LabelStyle.ForeColor = System.Drawing.ColorTranslator.FromHtml("#002776");
                int ii = 0;
                for (int i = 0; i <= dtChartDataSource.Rows.Count - 1; i++)
                {
                    chart.ChartAreas["ChartArea"].AxisX.CustomLabels.Add(i, i + 1.2, dtChartDataSource.Rows[i]["X"].ToString());
                    //chart1.ChartAreas[0].AxisX.CustomLabels.Add(xval[i]);
                    chart.Series[0].Points[ii].BackGradientStyle = GradientStyle.TopBottom;
                    ii++;

                }
                chartArea.AxisX.Title = "";

            }
            chart.SaveImage(Server.MapPath("~/ImageChart")+"\\Chart.png", System.Web.UI.DataVisualization.Charting.ChartImageFormat.Png);
        }

        protected void GenerateChartExecutive(DataTable dtChartDataSource, int width, int height, string filename)
        {
            Chart chart = new Chart()
            {
                Width = width,  // 700,
                Height = height   // 400
            };

            //chart.Legends.Add(new Legend() { Name = "Legend" });
            //chart.Legends[0].Docking = Docking.Bottom;
            ChartArea chartArea = new ChartArea() { Name = "ChartArea" };
            //Remove X-axis grid lines
            chartArea.AxisX.MajorGrid.LineWidth = 0;
            //Remove Y-axis grid lines
            chartArea.AxisY.MajorGrid.LineWidth = 0;
            //Chart Area Back Color
            chartArea.BackColor = System.Drawing.Color.FromName("White");
            chart.ChartAreas.Add(chartArea);
            chart.Palette = ChartColorPalette.BrightPastel;
            chart.RightToLeft = RightToLeft.Yes;
            string series = string.Empty;
            chart.BorderlineWidth = 0;
            //chart.Location = new System.Drawing.Point(322, 208);
            chart.Width = 780;// new System.Drawing.Size(740, height);
            chart.Height = height;
            chart.Palette = System.Web.UI.DataVisualization.Charting.ChartColorPalette.None;
       //     chart.PaletteCustomColors = new System.Drawing.Color[] {
       //System.Drawing.ColorTranslator.FromHtml("#082d79")};

            chart.PaletteCustomColors = new System.Drawing.Color[] { System.Drawing.ColorTranslator.FromHtml("#002776") };

            //     chart.PaletteCustomColors = new System.Drawing.Color[] {
            //System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(143)))), ((int)(((byte)(209)))))};
            chartArea.AxisX.TitleAlignment = StringAlignment.Center;
            chartArea.AxisX.LabelAutoFitStyle = LabelAutoFitStyles.IncreaseFont;

            chartArea.AxisX.MajorGrid.Enabled = false;
            chartArea.AxisY.MajorGrid.Enabled = false;
            chartArea.AxisX.TitleAlignment = System.Drawing.StringAlignment.Far;
            chartArea.AxisX.LineWidth = 1;
            chartArea.AxisY.LineWidth = 0;
            chartArea.AxisY.Enabled = AxisEnabled.False;

            chartArea.AxisX.IsMarginVisible = false;

            chart.RightToLeft = RightToLeft.Yes;
            chart.BackGradientStyle = System.Web.UI.DataVisualization.Charting.GradientStyle.TopBottom;
            chartArea.AxisX.IsStartedFromZero = true;
            //chartArea.AxisX.LabelStyle.ForeColor=Color

            StripLine stripline = new StripLine();
            stripline.Interval = 1;
            stripline.IntervalOffset = 60; //average value of the y axis; eg:35
            stripline.StripWidth = 1;
            stripline.BackColor = System.Drawing.Color.Yellow;
            chartArea.AxisX.StripLines.Add(stripline);

            //create series and add data points to the series
            if (dtChartDataSource != null)
            {
                foreach (DataColumn dc in dtChartDataSource.Columns)
                {
                    //a series to the chart
                    if (chart.Series.FindByName(dc.ColumnName) == null)
                    {

                        series = dc.ColumnName;

                        chart.Series.Add(series);
                        chart.Series[series].ChartType = SeriesChartType.Bar;
                        chart.Series[series]["PixelPointWidth"] = "35";
                        chart.Series[series]["DrawingStyle"] = "Cylinder";
                        chart.Series[series]["PointWidth"] = "50";
                        chart.Series[series].IsXValueIndexed = true;
                        chart.Series[series].LabelForeColor = System.Drawing.ColorTranslator.FromHtml("#002776");

                    }
                    //Add data points to the series
                    int i = 0;
                    foreach (DataRow dr in dtChartDataSource.Rows)
                    {

                        double dataPoint = 0;
                        double.TryParse(dr[dc.ColumnName].ToString(), out dataPoint);
                        DataPoint objDataPoint = new DataPoint() { AxisLabel = "series", YValues = new double[] { dataPoint } };
                        chart.Series[series].Points.Add(dataPoint);

                        if (dr["X"].ToString() == "Practice Productivity Metrics Average Percentile Ranking")
                        {

                            chart.Series[series].Points[i].Color = System.Drawing.ColorTranslator.FromHtml("#00ff00");
                        }

                        else if (dr["X"].ToString() == "Eyewear Metrics Average Percentile Ranking")
                        {

                            chart.Series[series].Points[i].Color = System.Drawing.ColorTranslator.FromHtml("#00ff00");
                        }
                        else if (dr["X"].ToString() == "Contact Lens Metrics Average Percentile")
                        {

                            chart.Series[series].Points[i].Color = System.Drawing.ColorTranslator.FromHtml("#00ff00");
                        }
                        else if (dr["X"].ToString() == "Medical Eye Care Metrics Average Percentile Ranking")
                        {

                            chart.Series[series].Points[i].Color = System.Drawing.ColorTranslator.FromHtml("#00ff00");
                        }
                        else if (dr["X"].ToString() == "Marketing Average Percentile Ranking")
                        {

                            chart.Series[series].Points[i].Color = System.Drawing.ColorTranslator.FromHtml("#00ff00");
                        }
                        else if (dr["X"].ToString() == "Financial Average Percentile Ranking")
                        {

                            chart.Series[series].Points[i].Color = System.Drawing.ColorTranslator.FromHtml("#00ff00");
                        }
                        else if (dr["X"].ToString() == "Average Percentile Ranking")
                        {

                            chart.Series[series].Points[i].Color = System.Drawing.ColorTranslator.FromHtml("#00ff00");
                        }
                        else
                        {

                            chart.Series[series].Points[i].Color = System.Drawing.ColorTranslator.FromHtml("#002776");

                        }



                        i++;





                    }

                }

                double mean = chart.DataManipulator.Statistics.Mean("Y1");

                //StripLine stripline = new StripLine();
                //stripline.Interval = 0;
                //stripline.IntervalOffset = mean; //average value of the y axis; eg:35
                //stripline.StripWidth = 1;
                //stripline.BackColor = Color.Red;
                //chartArea.AxisY.StripLines.Add(stripline);

                chart.Series[0].MarkerStyle = MarkerStyle.None;
                chart.Series[0].IsValueShownAsLabel = true;
                chart.Series[0]["BarLabelStyle"] = "Outside";
                chart.Series[0]["BoxPlotShowMedian"] = "true";


                chartArea.ShadowColor = System.Drawing.Color.White;

                chartArea.AxisX.IsMarginVisible = true;


                //chartArea.AxisX.LabelStyle.Interval = 1;
                chartArea.AxisX.LabelStyle.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);


                //LabelStyle yAxisStyle = new LabelStyle();
                //yAxisStyle.ForeColor = System.Drawing.ColorTranslator.FromHtml("#444444");
                //yAxisStyle.Font = new System.Drawing.Font("Arial", 11, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);


                chartArea.Position = new System.Web.UI.DataVisualization.Charting.ElementPosition(0, 0, 100, 100);
                chartArea.AxisX.Minimum = 0.5;
                //chartArea.AxisY.Minimum =10 ;



                chartArea.AxisX.LabelStyle.ForeColor = System.Drawing.ColorTranslator.FromHtml("#002776");
                int ii = 0;
                for (int i = 0; i <= dtChartDataSource.Rows.Count - 1; i++)
                {
                    if (dtChartDataSource.Rows[i]["X"].ToString().StartsWith("% of"))
                    {
                        string text = dtChartDataSource.Rows[i]["X"].ToString().Split('%')[1];
                        chart.ChartAreas["ChartArea"].AxisX.CustomLabels.Add(i + 1, i + 2, dtChartDataSource.Rows[i]["X"].ToString().Replace("%", "Percentage"));
                    }

                    else
                    {
                        chart.ChartAreas["ChartArea"].AxisX.CustomLabels.Add(i + 1, i + 2, dtChartDataSource.Rows[i]["X"].ToString());
                    }
                    chart.Series[0].Points[ii].BackGradientStyle = GradientStyle.LeftRight;

                    ii++;

                }


            }
            chart.SaveImage(Server.MapPath("~/ImageChart")+"\\" + filename + ".png", System.Web.UI.DataVisualization.Charting.ChartImageFormat.Png);
        }

        protected void btnprint_Click(object sender, EventArgs e)
        {
            WebClient req = new WebClient();
            HttpResponse response = HttpContext.Current.Response;
            response.Clear();
            response.ClearContent();
            response.ClearHeaders();
            response.Buffer = true;
            string pdf = Server.MapPath("/") + @"bin\Debug" + @"\mba2014 questionnaire.pdf";
            response.AddHeader("Content-Disposition", "attachment;filename=\"" + "mba2014 questionnaire.pdf" + "\"");
            byte[] data = req.DownloadData(pdf);
            response.BinaryWrite(data);
            response.End();
        }
       
    
}

   
}