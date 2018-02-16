using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SurveyApp
{
	public partial class Result : System.Web.UI.Page
	{
		static string finalpath = string.Empty;
		static string filenamestr = string.Empty;
		static string filenameinfographicstr = string.Empty;
		static string filenameexecutivestr = string.Empty;
		static string Username = string.Empty;

		protected void Page_Load(object sender, EventArgs e)
		{
			finalpath = System.Web.HttpContext.Current.Session["finalpath"].ToString();
			filenamestr = System.Web.HttpContext.Current.Session["PPAName"].ToString();
			filenameinfographicstr = System.Web.HttpContext.Current.Session["infographic"].ToString();
			filenameexecutivestr = System.Web.HttpContext.Current.Session["executive"].ToString();
			Username = System.Web.HttpContext.Current.Session["Username"].ToString();

			if (Request.QueryString["Result"] != null)
			{
				string result = Request.QueryString["Result"];
				if (result == "Y")
				{
					lblmsg.Text = "The Practice Performance detailed report, executive summary,one page infographic and KeyMetrics-Spectacle Report are in process and will be available soon.<br>In case of any queries, you can write to " + "info@essilor.com";
					lblmsg.Attributes.Remove("class");
					lblmsg.Attributes.Add("class", "validation-errorsuccess");
				}

				else if (result == "N/I")
				{
					lblmsg.Text = "Error Occured While Generating Infographic and Executive PDF.Detailed pdf is generated successfully.";
					lblmsg.Attributes.Remove("class");
					lblmsg.Attributes.Add("class", "validation-errorfailed");
				}
				else if (result == "N/E")
				{
					lblmsg.Text = "Error Occured While Generating Executive PDF.Detailed and Infographic pdf is generated successfully.";

					lblmsg.Attributes.Remove("class");
					lblmsg.Attributes.Add("class", "validation-errorfailed");
				}
				else if (result == "N/D")
				{
					lblmsg.Text = "Error Occured While Generating Detailed,Infographic and Executive PDF.";

					lblmsg.Attributes.Remove("class");
					lblmsg.Attributes.Add("class", "validation-errorfailed");
				}
				else if (result == "N/K")
				{
					lblmsg.Text = "Error Occured While Generating KeyMeterics PDF.";

					lblmsg.Attributes.Remove("class");
					lblmsg.Attributes.Add("class", "validation-errorfailed");
				}

				else if (result == "G")
				{

					// string  filePath = ConfigurationManager.AppSettings["ErrorFilePath"].ToString();// @"C:\Error.txt";
					lblmsg.Text = "Error Occured While Generating Detailed,Infographic and Executive PDF.";

					lblmsg.Attributes.Remove("class");
					lblmsg.Attributes.Add("class", "validation-errorfailed");
				}

				else if (result == "S")
				{
					lblmsg.Text = "Error Occured While submitting Survey.";

					lblmsg.Attributes.Remove("class");
					lblmsg.Attributes.Add("class", "validation-error");
				}

				else
				{
					lblmsg.Text = result.ToString();
				}
			}

		}

		protected void Homepagebutton_Click(object sender, EventArgs e)
		{
			Response.Redirect("Homepage.aspx?id=6");

		}

		protected void btnDownload_Click(object sender, EventArgs e)
		{
			//string finalpath = Session["finalpath"].ToString();
			//string filenamestr = System.Web.HttpContext.Current.Session["varName"].ToString();
			//string filenameinfographicstr = System.Web.HttpContext.Current.Session["varNameinfographic"].ToString();
			//string filenameexecutivestr = System.Web.HttpContext.Current.Session["varNameexecutive"].ToString();
			//string filenameekeymetricsstr = System.Web.HttpContext.Current.Session["varNameKeyMetrics"].ToString();

			using (ZipFile zip = new ZipFile())
			{
				zip.AlternateEncodingUsage = ZipOption.AsNecessary;
				zip.AddDirectoryByName("surveypdf1");
				zip.AddFile(finalpath + filenamestr + ".pdf");
				zip.AddFile(finalpath + filenameinfographicstr + ".pdf", "surveypdf1");
				zip.AddFile(finalpath + filenameexecutivestr + ".pdf", "surveypdf1");
				//zip.AddFile(finalpath + filenameekeymetricsstr + ".pdf", "surveypdf1");

				Response.Clear();
				Response.BufferOutput = false;
				//string zipName = String.Format("Practice Performance Assessment{0}.zip", DateTime.Now.ToString("yyyy-MMM-dd-HHmmss"));
				string zipName = "PracticePerformanceAssessment_" + Username + "_" + DateTime.Now.ToString("MM/dd/yyyy") + ".zip";
				Response.ContentType = "application/zip";
				Response.AddHeader("content-disposition", "attachment; filename=" + zipName);
				zip.Save(Response.OutputStream);
				//Response.End();
				HttpContext.Current.ApplicationInstance.CompleteRequest();
			}
		}
	}
}