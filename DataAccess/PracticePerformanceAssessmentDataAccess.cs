using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Diagnostics;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Transactions;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Configuration;
using BusinessObjects;
using System.Data.Objects;
using System.ComponentModel;
using System.Data.Common;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;
using System.Drawing.Drawing2D;



using System.Drawing.Imaging;
using System.Web;
using System.Data.SqlClient;


namespace PracticePerformanceAssessmentDataAccess
{
	public class PracticePerformanceAssessmentDataAccess
	{
		//PPASurvey_DBEntities context = new PPASurvey_DBEntities();
		//ObjectContext context2 = ((IObjectContextAdapter)context).ObjectContext;
		private readonly PPASurvey_DBEntities db;
		// ObjectContext db;
		//dynamic stopwatch;
		int NoOfDocumentsLimit = 0;
		int lastReportGenerated = 0;

		string finalfilename = "";
		private bool compare;
		private bool display3d;
		private bool showValues;
		private string charType = "Column";
		public PracticePerformanceAssessmentDataAccess(string StartupPath, int NoOfRowsLimit)
		{
			//stopwatch = Stopwatch.StartNew();
			string strConnection = ConfigurationManager.ConnectionStrings["PPASurvey_DBEntities"].ConnectionString;
			// strConnection = strConnection.Replace("AptaraPath", StartupPath);
			db = new PPASurvey_DBEntities(strConnection);
			//db.Database.Connection.ConnectionString = strConnection;


			//NoOfDocumentsLimit = Int32.Parse(ConfigurationManager.AppSettings["DocumentsLimit"]); //here we are getting NoOfDocumentsLimit from App.config.
			NoOfDocumentsLimit = NoOfRowsLimit;
		}

		public static ObjectContext ConvertContext(DbContext db)

		{

			return ((IObjectContextAdapter)db).ObjectContext;

		}


		public BusinessObjects.Report GetOutputData(string practiceid)
		{
			BusinessObjects.Report objReport = new BusinessObjects.Report();

			List<BusinessObjects.Output> lstOutput = new List<Output>();


			//var stopwatch = Stopwatch.StartNew();


			try
			{
				int rowid = db.Source_InputData.Where(r => r.IDname == practiceid && r.IDformat == "Y").Select(r => r.RowId).ToList().Last();
				if (db.Target_OutputData.Where(x => x.SourceDataRefId == rowid).Select(r => r.SourceDataRefId).Count() > 0)
				{
					// lastReportGenerated = db.Target_OutputData.Select(r => r.SourceDataRefId).ToList().Last();


					lastReportGenerated = db.Target_OutputData.Where(x => x.SourceDataRefId == rowid).Select(r => r.SourceDataRefId).ToList().Last();
				}

				//Only run the reports that are not ran earlier.
				List<int> reportToBeGenerated = db.Source_InputData.Where(r => r.RowId > lastReportGenerated && r.IDformat == "Y" && r.IDname == practiceid).Select(r => r.RowId).ToList();
				//int reportToBeGenerated =db.Source_InputData.Where(r => r.IDname == practiceid.ToString() && r.IDformat == "Y").Select(r => r.RowId).ToList().Last();
				//for (int i = 0; i < reportToBeGenerated.Count(); i++) //
				if (reportToBeGenerated.Count < NoOfDocumentsLimit)
				{
					NoOfDocumentsLimit = 1; //reportToBeGenerated.Count();
				}

				//for (int i = reportToBeGenerated.Count()-1; i < NoOfDocumentsLimit; i++)
				//{

				int sourceRowId = reportToBeGenerated[reportToBeGenerated.Count() - 1];

				var SourceInputData = (from r in db.Source_InputData
									   where r.RowId == sourceRowId
									   select r).ToList().FirstOrDefault();

				#region  Comments         
				//Mapping the [Source.InputData] table data to the BusinessObjects.Input Properties.
				/* #region ReadInputDataFromDatabase

                 BusinessObjects.Input objInputData = new BusinessObjects.Input();


                 BusinessObjects.Output objOutputData = new Output();

                 //objInputData.colIDformat = Convert.ToChar(SourceInputData.IDformat);
                 objInputData.RowId = SourceInputData.RowId;
                 objInputData.colIDformat = SourceInputData.IDformat;
                 objInputData.colIDendDate = Convert.ToDateTime(SourceInputData.IDendDate);
                 objInputData.colIDend = SourceInputData.IDend;
                 objInputData.colIDstart = SourceInputData.IDstart;
                 objInputData.colIDdate = Convert.ToDateTime(SourceInputData.IDdate);
                 objInputData.colIDname = SourceInputData.IDname;
                 objInputData.colQ1 = Convert.ToDecimal(SourceInputData.Q1);
                 objInputData.colQ2 = Convert.ToDecimal(SourceInputData.Q2);
                 objInputData.colQ3 = Convert.ToDecimal(SourceInputData.Q3);
                 objInputData.colQ4 = Convert.ToDecimal(SourceInputData.Q4);
                 objInputData.colQ5 = Convert.ToDecimal(SourceInputData.Q5);
                 objInputData.colQ6 = Convert.ToDecimal(SourceInputData.Q6);
                 objInputData.colQ7 = Convert.ToDecimal(SourceInputData.Q7);
                 objInputData.colQ8 = Convert.ToDecimal(SourceInputData.Q8);
                 objInputData.colQ9 = Convert.ToDecimal(SourceInputData.Q9);
                 objInputData.colQ10 = Convert.ToDecimal(SourceInputData.Q10);
                 objInputData.colQ11 = Convert.ToDecimal(SourceInputData.Q11);
                 objInputData.colQ12 = Convert.ToDecimal(SourceInputData.Q12);
                 objInputData.colQ13a = Convert.ToDecimal(SourceInputData.Q13a);
                 objInputData.colQ13b = Convert.ToDecimal(SourceInputData.Q13b);
                 objInputData.colQ13c = Convert.ToDecimal(SourceInputData.Q13c);
                 objInputData.colQ13d = Convert.ToDecimal(SourceInputData.Q13d);
                 objInputData.colQ14 = Convert.ToDecimal(SourceInputData.Q14);
                 objInputData.colQ15a = Convert.ToDecimal(SourceInputData.Q15a);
                 objInputData.colQ15b = Convert.ToDecimal(SourceInputData.Q15b);
                 objInputData.colQ15c = Convert.ToDecimal(SourceInputData.Q15c);
                 objInputData.colQ15d = Convert.ToDecimal(SourceInputData.Q15d);
                 objInputData.colQ16 = Convert.ToDecimal(SourceInputData.Q16);
                 objInputData.colQ17 = Convert.ToDecimal(SourceInputData.Q17);
                 objInputData.colQ18 = Convert.ToDecimal(SourceInputData.Q18);
                 objInputData.colQ19 = Convert.ToDecimal(SourceInputData.Q19);
                 objInputData.colQ20a = Convert.ToDecimal(SourceInputData.Q20a);
                 objInputData.colQ20b = Convert.ToDecimal(SourceInputData.Q20b);
                 objInputData.colQ20c = Convert.ToDecimal(SourceInputData.Q20c);
                 objInputData.colQ20d = Convert.ToDecimal(SourceInputData.Q20d);
                 objInputData.colQ20e = Convert.ToDecimal(SourceInputData.Q20e);
                 objInputData.colQ20f = Convert.ToDecimal(SourceInputData.Q20f);
                 objInputData.colQ20g = Convert.ToDecimal(SourceInputData.Q20g);
                 objInputData.colQ21a = Convert.ToDecimal(SourceInputData.Q21a);
                 objInputData.colQ21b = Convert.ToDecimal(SourceInputData.Q21b);
                 objInputData.colQ21c = Convert.ToDecimal(SourceInputData.Q21c);
                 objInputData.colQ21d = Convert.ToDecimal(SourceInputData.Q21d);
                 objInputData.colQ22 = Convert.ToDecimal(SourceInputData.Q22);
                 objInputData.colQ23 = Convert.ToDecimal(SourceInputData.Q23);
                 objInputData.colQ24 = Convert.ToDecimal(SourceInputData.Q24);
                 objInputData.colQ25 = Convert.ToDecimal(SourceInputData.Q25);
                 objInputData.colQ26 = Convert.ToDecimal(SourceInputData.Q26);
                 objInputData.colQ26a = Convert.ToDecimal(SourceInputData.Q26a);
                 objInputData.colQ26b = Convert.ToDecimal(SourceInputData.Q26b);
                 objInputData.colQ26c = Convert.ToDecimal(SourceInputData.Q26c);
                 objInputData.colQ26d = Convert.ToDecimal(SourceInputData.Q26d);
                 objInputData.colQ26e = Convert.ToDecimal(SourceInputData.Q26e);
                 objInputData.colQ26f = Convert.ToDecimal(SourceInputData.Q26f);
                 objInputData.colQ26g = Convert.ToDecimal(SourceInputData.Q26g);

                 //objInputData.colQ26g = (SourceInputData.Q26g)==null?-11:Convert.ToDecimal(SourceInputData.Q26g);

                 objInputData.colQ26h = Convert.ToDecimal(SourceInputData.Q26h);
                 objInputData.colQ26i = Convert.ToDecimal(SourceInputData.Q26i);
                 objInputData.colQ27 = Convert.ToDecimal(SourceInputData.Q27);
                 objInputData.colQ27a = Convert.ToDecimal(SourceInputData.Q27a);
                 objInputData.colQ27b = Convert.ToDecimal(SourceInputData.Q27b);
                 objInputData.colQ27c = Convert.ToDecimal(SourceInputData.Q27c);
                 objInputData.colQ27d = Convert.ToDecimal(SourceInputData.Q27d);
                 objInputData.colQ27e = Convert.ToDecimal(SourceInputData.Q27e);
                 objInputData.colQ28 = Convert.ToDecimal(SourceInputData.Q28);
                 objInputData.colQ29 = Convert.ToDecimal(SourceInputData.Q29);
                 objInputData.colQ30 = Convert.ToDecimal(SourceInputData.Q30);
                 objInputData.colQ31a = Convert.ToDecimal(SourceInputData.Q31a);
                 objInputData.colQ31b = Convert.ToDecimal(SourceInputData.Q31b);
                 objInputData.colQ31c = Convert.ToDecimal(SourceInputData.Q31c);
                 objInputData.colQ32a = Convert.ToDecimal(SourceInputData.Q32a);
                 objInputData.colQ32b = Convert.ToDecimal(SourceInputData.Q32b);
                 objInputData.colQ32c = Convert.ToDecimal(SourceInputData.Q32c);
                 objInputData.colQ32d = Convert.ToDecimal(SourceInputData.Q32d);
                 objInputData.colQ32e = Convert.ToDecimal(SourceInputData.Q32e);
                 objInputData.colQ33a = Convert.ToDecimal(SourceInputData.Q33a);
                 objInputData.colQ33b = Convert.ToDecimal(SourceInputData.Q33b);
                 objInputData.colQ33c = Convert.ToDecimal(SourceInputData.Q33c);
                 objInputData.colQ33d = Convert.ToDecimal(SourceInputData.Q33d);
                 objInputData.colQ33e = Convert.ToDecimal(SourceInputData.Q33e);
                 objInputData.colQ34 = Convert.ToDecimal(SourceInputData.Q34);
                 objInputData.colQ35 = Convert.ToDecimal(SourceInputData.Q35);
                 objInputData.colQ36 = Convert.ToDecimal(SourceInputData.Q36);
                 objInputData.colQ37 = Convert.ToDecimal(SourceInputData.Q37);
                 objInputData.colQ38 = Convert.ToDecimal(SourceInputData.Q38);
                 objInputData.colQ39a = Convert.ToDecimal(SourceInputData.Q39a);
                 objInputData.colQ39b = Convert.ToDecimal(SourceInputData.Q39b);
                 objInputData.colQ39c = Convert.ToDecimal(SourceInputData.Q39c);
                 objInputData.colQ39d = Convert.ToDecimal(SourceInputData.Q39d);
                 objInputData.colQ39e = Convert.ToDecimal(SourceInputData.Q39e);
                 objInputData.colQ39f = Convert.ToDecimal(SourceInputData.Q39f);
                 objInputData.colQ40a = Convert.ToDecimal(SourceInputData.Q40a);
                 objInputData.colQ40b = Convert.ToDecimal(SourceInputData.Q40b);
                 objInputData.colQ40c = Convert.ToDecimal(SourceInputData.Q40c);
                 objInputData.colQ40d = Convert.ToDecimal(SourceInputData.Q40d);
                 objInputData.colQ40e = Convert.ToDecimal(SourceInputData.Q40e);
                 objInputData.colQ40f = Convert.ToDecimal(SourceInputData.Q40f);
                 objInputData.colQ41a = Convert.ToDecimal(SourceInputData.Q41a);
                 objInputData.colQ42a = Convert.ToDecimal(SourceInputData.Q42a);
                 objInputData.colQ43a = Convert.ToDecimal(SourceInputData.Q43a);
                 objInputData.colQ44 = Convert.ToDecimal(SourceInputData.Q44);
                 objInputData.colQ45a = Convert.ToDecimal(SourceInputData.Q45a);
                 objInputData.colQ45b = Convert.ToDecimal(SourceInputData.Q45b);
                 objInputData.colQ46a = Convert.ToDecimal(SourceInputData.Q46a);
                 objInputData.colQ47 = Convert.ToDecimal(SourceInputData.Q47);
                 objInputData.colQ48 = Convert.ToDecimal(SourceInputData.Q48);
                 objInputData.colQ49 = Convert.ToDecimal(SourceInputData.Q49);
                 objInputData.colQ50 = Convert.ToDecimal(SourceInputData.Q50);
                 objInputData.colQ51 = Convert.ToDecimal(SourceInputData.Q51);
                 objInputData.colQ52a = Convert.ToDecimal(SourceInputData.Q52a);
                 objInputData.colQ52b = Convert.ToDecimal(SourceInputData.Q52b);
                 objInputData.colQ52c = Convert.ToDecimal(SourceInputData.Q52c);
                 objInputData.colQ52d = Convert.ToDecimal(SourceInputData.Q52d);
                 objInputData.colQ52e = Convert.ToDecimal(SourceInputData.Q52e);
                 objInputData.colQ52f = Convert.ToDecimal(SourceInputData.Q52f);
                 objInputData.colQ52k = Convert.ToDecimal(SourceInputData.Q52k);
                 objInputData.colQ52h = Convert.ToDecimal(SourceInputData.Q52h);
                 objInputData.colQ52i = Convert.ToDecimal(SourceInputData.Q52i);
                 objInputData.colQ52j = Convert.ToDecimal(SourceInputData.Q52j);
                 objInputData.colQ53 = Convert.ToDecimal(SourceInputData.Q53);
                 objInputData.colQ54 = Convert.ToDecimal(SourceInputData.Q54);
                 objInputData.colQ55 = Convert.ToDecimal(SourceInputData.Q55);
                 objInputData.colQ56 = Convert.ToDecimal(SourceInputData.Q56);
                 objInputData.colQ57 = Convert.ToDecimal(SourceInputData.Q57);
                 objInputData.colQ58 = Convert.ToDecimal(SourceInputData.Q58);
                 objInputData.colQ59 = Convert.ToDecimal(SourceInputData.Q59);
                 objInputData.colQ60 = Convert.ToDecimal(SourceInputData.Q60);
                 objInputData.colQ61 = Convert.ToDecimal(SourceInputData.Q61);
                 objInputData.colQ62 = Convert.ToDecimal(SourceInputData.Q62);
                 objInputData.colQ63 = Convert.ToDecimal(SourceInputData.Q63);
                 objInputData.colQ64a = SourceInputData.Q64a;
                 objInputData.colQ64b = SourceInputData.Q64b;
                 objInputData.colQ64c = SourceInputData.Q64c;
                 objInputData.colQ64d = SourceInputData.Q64d;
                 objInputData.colQ64e = SourceInputData.Q64e;
                 objInputData.colQ64f = SourceInputData.Q64f;
                 objInputData.colQ64g = SourceInputData.Q64g;
                 objInputData.colQ64h = SourceInputData.Q64h;
                 objInputData.colQ64i = SourceInputData.Q64i;
                 objInputData.colQ64j = SourceInputData.Q64j;
                 objInputData.colQ64k = SourceInputData.Q64k;
                 objInputData.colQ64l = SourceInputData.Q64l;
                 objInputData.colQ64m = SourceInputData.Q64m;
                 objInputData.colQ64n = SourceInputData.Q64n;
                 objInputData.colQ64o = SourceInputData.Q64o;
                 objInputData.colQ65a = SourceInputData.Q65a;
                 objInputData.colQ65b = SourceInputData.Q65b;
                 objInputData.colQ65c = SourceInputData.Q65c;
                 objInputData.colQ65d = SourceInputData.Q65d;
                 objInputData.colQ65e = SourceInputData.Q65e;
                 objInputData.colQ65f = SourceInputData.Q65f;
                 objInputData.colQ65g = SourceInputData.Q65g;
                 objInputData.colQ65h = SourceInputData.Q65h;
                 objInputData.colQ65i = SourceInputData.Q65i;
                 objInputData.colQ65j = SourceInputData.Q65j;
                 objInputData.colQ65k = SourceInputData.Q65k;
                 objInputData.colQ65l = SourceInputData.Q65l;
                 objInputData.colQ65m = SourceInputData.Q65m;
                 objInputData.colQ65n = SourceInputData.Q65n;
                 objInputData.colQ66a = SourceInputData.Q66a;
                 objInputData.colQ66b = SourceInputData.Q66b;
                 objInputData.colQ66c = SourceInputData.Q66c;
                 objInputData.colQ66d = SourceInputData.Q66d;
                 objInputData.colQ66e = SourceInputData.Q66e;
                 objInputData.colQ66f = SourceInputData.Q66f;
                 objInputData.colQ66g = SourceInputData.Q66g;
                 objInputData.colQ66h = SourceInputData.Q66h;
                 objInputData.colQ66i = SourceInputData.Q66i;
                 objInputData.colQ66j = SourceInputData.Q66j;
                 objInputData.colQ66k = SourceInputData.Q66k;
                 objInputData.colQ66l = SourceInputData.Q66l;
                 objInputData.colQ66m = SourceInputData.Q66m;
                 objInputData.colQ66n = SourceInputData.Q66n;
                 objInputData.colQ66o = SourceInputData.Q66o;
                 objInputData.colQ66p = SourceInputData.Q66p;
                 objInputData.colQ66q = SourceInputData.Q66q;
                 objInputData.colQ66r = SourceInputData.Q66r;
                 objInputData.colQ66s = SourceInputData.Q66s;
                 objInputData.colQ66t = SourceInputData.Q66t;
                 objInputData.colQ66u = SourceInputData.Q66u;
                 objInputData.colQ66v = SourceInputData.Q66v;
                 objInputData.colQ67a = Convert.ToDecimal(SourceInputData.Q67a);
                 objInputData.colQ67b = Convert.ToDecimal(SourceInputData.Q67b);
                 objInputData.colQ67c = Convert.ToDecimal(SourceInputData.Q67c);
                 //objInputData.colQ68 = Convert.ToChar(SourceInputData.Q68);
                 objInputData.colQ68 = SourceInputData.Q68;
                 objInputData.colQ69a = Convert.ToDecimal(SourceInputData.Q69a);
                 objInputData.colQ70a = Convert.ToDecimal(SourceInputData.Q70a);
                 objInputData.colQ71a = Convert.ToDecimal(SourceInputData.Q71a);
                 objInputData.colQ72 = SourceInputData.Q72;
                 objInputData.colQ73 = SourceInputData.Q73;
                 objInputData.colQ74 = SourceInputData.Q74;
                 objInputData.colQ75 = SourceInputData.Q75;
                 objInputData.colQ75a = SourceInputData.Q75a;
                 objInputData.colQ75b = SourceInputData.Q75b;
                 objInputData.colQ76 = SourceInputData.Q76;

                 //praveenk-Release2
                 objInputData.AdditionalInfo = SourceInputData.AdditionalInfo;

                 #endregion ReadInputDataFromDatabase*/

				#endregion

				#region ReadInputDataFromDatabase

				BusinessObjects.Input objInputData = new BusinessObjects.Input();


				BusinessObjects.Output objOutputData = new Output();

				//objInputData.colIDformat = Convert.ToChar(SourceInputData.IDformat);
				objInputData.RowId = SourceInputData.RowId;
				objInputData.colIDformat = SourceInputData.IDformat;
				objInputData.colIDendDate = Convert.ToDateTime(SourceInputData.IDendDate);
				objInputData.colIDend = SourceInputData.IDend;
				objInputData.colIDstart = SourceInputData.IDstart;
				objInputData.colIDdate = Convert.ToDateTime(SourceInputData.IDdate);
				objInputData.colIDname = SourceInputData.IDname;
				objInputData.colQ1 = (SourceInputData.Q1);
				objInputData.colQ2 = (SourceInputData.Q2);
				objInputData.colQ3 = (SourceInputData.Q3);
				objInputData.colQ4 = (SourceInputData.Q4);
				objInputData.colQ5 = (SourceInputData.Q5);
				objInputData.colQ6 = (SourceInputData.Q6);
				objInputData.colQ7 = (SourceInputData.Q7);
				objInputData.colQ8 = (SourceInputData.Q8);
				objInputData.colQ9 = (SourceInputData.Q9);
				objInputData.colQ10 = (SourceInputData.Q10);
				objInputData.colQ11 = (SourceInputData.Q11);
				objInputData.colQ12 = (SourceInputData.Q12);
				objInputData.colQ13a = (SourceInputData.Q13a);
				objInputData.colQ13b = (SourceInputData.Q13b);
				objInputData.colQ13c = (SourceInputData.Q13c);
				objInputData.colQ13d = (SourceInputData.Q13d);
				objInputData.colQ14 = (SourceInputData.Q14);
				objInputData.colQ15a = (SourceInputData.Q15a);
				objInputData.colQ15b = (SourceInputData.Q15b);
				objInputData.colQ15c = (SourceInputData.Q15c);
				objInputData.colQ15d = (SourceInputData.Q15d);
				objInputData.colQ16 = (SourceInputData.Q16);
				objInputData.colQ17 = (SourceInputData.Q17);
				objInputData.colQ18 = (SourceInputData.Q18);
				objInputData.colQ19 = (SourceInputData.Q19);
				objInputData.colQ20a = (SourceInputData.Q20a);
				objInputData.colQ20b = (SourceInputData.Q20b);
				objInputData.colQ20c = (SourceInputData.Q20c);
				objInputData.colQ20d = (SourceInputData.Q20d);
				objInputData.colQ20e = (SourceInputData.Q20e);
				objInputData.colQ20f = (SourceInputData.Q20f);
				objInputData.colQ20g = (SourceInputData.Q20g);
				objInputData.colQ21a = (SourceInputData.Q21a);
				objInputData.colQ21b = (SourceInputData.Q21b);
				objInputData.colQ21c = (SourceInputData.Q21c);
				objInputData.colQ21d = (SourceInputData.Q21d);
				objInputData.colQ22 = (SourceInputData.Q22);
				objInputData.colQ23 = (SourceInputData.Q23);
				objInputData.colQ24 = (SourceInputData.Q24);
				objInputData.colQ25 = (SourceInputData.Q25);
				objInputData.colQ26 = (SourceInputData.Q26);
				objInputData.colQ26a = (SourceInputData.Q26a);
				objInputData.colQ26b = (SourceInputData.Q26b);
				objInputData.colQ26c = (SourceInputData.Q26c);
				objInputData.colQ26d = (SourceInputData.Q26d);
				objInputData.colQ26e = (SourceInputData.Q26e);

				if (SourceInputData.Q26f == 0)
				{
					objInputData.colQ26f = null;
				}

				else

				{
					objInputData.colQ26f = (SourceInputData.Q26f);
				}
				objInputData.colQ26g = (SourceInputData.Q26g);

				//objInputData.colQ26g = (SourceInputData.Q26g)==null?-11:(SourceInputData.Q26g);

				objInputData.colQ26h = (SourceInputData.Q26h);
				objInputData.colQ26i = (SourceInputData.Q26i);
				objInputData.colQ27 = (SourceInputData.Q27);
				if (SourceInputData.Q27a == 0)
				{
					objInputData.colQ27a = null;
				}
				else
				{
					objInputData.colQ27a = (SourceInputData.Q27a);
				}
				objInputData.colQ27b = (SourceInputData.Q27b);
				objInputData.colQ27c = (SourceInputData.Q27c);
				objInputData.colQ27d = (SourceInputData.Q27d);
				objInputData.colQ27e = (SourceInputData.Q27e);
				objInputData.colQ28 = (SourceInputData.Q28);

				if (SourceInputData.Q29 == 0)
				{
					objInputData.colQ29 = null;
				}

				else
				{
					objInputData.colQ29 = (SourceInputData.Q29);
				}
				objInputData.colQ30 = (SourceInputData.Q30);
				objInputData.colQ31a = (SourceInputData.Q31a);
				objInputData.colQ31b = (SourceInputData.Q31b);
				objInputData.colQ31c = (SourceInputData.Q31c);
				objInputData.colQ32a = (SourceInputData.Q32a);
				objInputData.colQ32b = (SourceInputData.Q32b);
				objInputData.colQ32c = (SourceInputData.Q32c);
				objInputData.colQ32d = (SourceInputData.Q32d);
				objInputData.colQ32e = (SourceInputData.Q32e);
				objInputData.colQ33a = (SourceInputData.Q33a);
				objInputData.colQ33b = (SourceInputData.Q33b);
				objInputData.colQ33c = (SourceInputData.Q33c);
				objInputData.colQ33d = (SourceInputData.Q33d);
				objInputData.colQ33e = (SourceInputData.Q33e);
				objInputData.colQ34 = (SourceInputData.Q34);
				objInputData.colQ35 = (SourceInputData.Q35);
				objInputData.colQ36 = (SourceInputData.Q36);
				objInputData.colQ37 = (SourceInputData.Q37);
				objInputData.colQ38 = (SourceInputData.Q38);
				objInputData.colQ39a = (SourceInputData.Q39a);
				objInputData.colQ39b = (SourceInputData.Q39b);
				objInputData.colQ39c = (SourceInputData.Q39c);
				objInputData.colQ39d = (SourceInputData.Q39d);
				objInputData.colQ39e = (SourceInputData.Q39e);
				objInputData.colQ39f = (SourceInputData.Q39f);
				objInputData.colQ40a = (SourceInputData.Q40a);
				objInputData.colQ40b = (SourceInputData.Q40b);
				objInputData.colQ40c = (SourceInputData.Q40c);
				objInputData.colQ40d = (SourceInputData.Q40d);
				objInputData.colQ40e = (SourceInputData.Q40e);
				objInputData.colQ40f = (SourceInputData.Q40f);
				objInputData.colQ41a = (SourceInputData.Q41a);
				objInputData.colQ42a = (SourceInputData.Q42a);
				objInputData.colQ43a = (SourceInputData.Q43a);
				objInputData.colQ44 = (SourceInputData.Q44);
				objInputData.colQ45a = (SourceInputData.Q45a);
				objInputData.colQ45b = (SourceInputData.Q45b);
				objInputData.colQ46a = (SourceInputData.Q46a);
				objInputData.colQ47 = (SourceInputData.Q47);
				objInputData.colQ48 = (SourceInputData.Q48);
				objInputData.colQ49 = (SourceInputData.Q49);
				objInputData.colQ50 = (SourceInputData.Q50);
				objInputData.colQ51 = (SourceInputData.Q51);
				objInputData.colQ52a = (SourceInputData.Q52a);
				objInputData.colQ52b = (SourceInputData.Q52b);
				objInputData.colQ52c = (SourceInputData.Q52c);
				objInputData.colQ52d = (SourceInputData.Q52d);
				objInputData.colQ52e = (SourceInputData.Q52e);
				objInputData.colQ52f = (SourceInputData.Q52f);
				objInputData.colQ52k = (SourceInputData.Q52k);
				objInputData.colQ52h = (SourceInputData.Q52h);
				objInputData.colQ52i = (SourceInputData.Q52i);
				objInputData.colQ52j = (SourceInputData.Q52j);
				objInputData.colQ53 = (SourceInputData.Q53);

				if (objInputData.colQ54 == 0)
				{
					objInputData.colQ54 = null;
				}

				else
				{
					objInputData.colQ54 = (SourceInputData.Q54);
				}
				objInputData.colQ55 = (SourceInputData.Q55);
				objInputData.colQ56 = (SourceInputData.Q56);
				objInputData.colQ57 = (SourceInputData.Q57);
				objInputData.colQ58 = (SourceInputData.Q58);
				objInputData.colQ59 = (SourceInputData.Q59);
				objInputData.colQ60 = (SourceInputData.Q60);
				objInputData.colQ61 = (SourceInputData.Q61);
				objInputData.colQ62 = (SourceInputData.Q62);
				objInputData.colQ63 = (SourceInputData.Q63);
				objInputData.colQ64a = SourceInputData.Q64a;
				objInputData.colQ64b = SourceInputData.Q64b;
				objInputData.colQ64c = SourceInputData.Q64c;
				objInputData.colQ64d = SourceInputData.Q64d;
				objInputData.colQ64e = SourceInputData.Q64e;
				objInputData.colQ64f = SourceInputData.Q64f;
				objInputData.colQ64g = SourceInputData.Q64g;
				objInputData.colQ64h = SourceInputData.Q64h;
				objInputData.colQ64i = SourceInputData.Q64i;
				objInputData.colQ64j = SourceInputData.Q64j;
				objInputData.colQ64k = SourceInputData.Q64k;
				objInputData.colQ64l = SourceInputData.Q64l;
				objInputData.colQ64m = SourceInputData.Q64m;
				objInputData.colQ64n = SourceInputData.Q64n;
				objInputData.colQ64o = SourceInputData.Q64o;
				objInputData.colQ65a = SourceInputData.Q65a;
				objInputData.colQ65b = SourceInputData.Q65b;
				objInputData.colQ65c = SourceInputData.Q65c;
				objInputData.colQ65d = SourceInputData.Q65d;
				objInputData.colQ65e = SourceInputData.Q65e;
				objInputData.colQ65f = SourceInputData.Q65f;
				objInputData.colQ65g = SourceInputData.Q65g;
				objInputData.colQ65h = SourceInputData.Q65h;
				objInputData.colQ65i = SourceInputData.Q65i;
				objInputData.colQ65j = SourceInputData.Q65j;
				objInputData.colQ65k = SourceInputData.Q65k;
				objInputData.colQ65l = SourceInputData.Q65l;
				objInputData.colQ65m = SourceInputData.Q65m;
				objInputData.colQ65n = SourceInputData.Q65n;
				objInputData.colQ66a = SourceInputData.Q66a;
				objInputData.colQ66b = SourceInputData.Q66b;
				objInputData.colQ66c = SourceInputData.Q66c;
				objInputData.colQ66d = SourceInputData.Q66d;
				objInputData.colQ66e = SourceInputData.Q66e;
				objInputData.colQ66f = SourceInputData.Q66f;
				objInputData.colQ66g = SourceInputData.Q66g;
				objInputData.colQ66h = SourceInputData.Q66h;
				objInputData.colQ66i = SourceInputData.Q66i;
				objInputData.colQ66j = SourceInputData.Q66j;
				objInputData.colQ66k = SourceInputData.Q66k;
				objInputData.colQ66l = SourceInputData.Q66l;
				objInputData.colQ66m = SourceInputData.Q66m;
				objInputData.colQ66n = SourceInputData.Q66n;
				objInputData.colQ66o = SourceInputData.Q66o;
				objInputData.colQ66p = SourceInputData.Q66p;
				objInputData.colQ66q = SourceInputData.Q66q;
				objInputData.colQ66r = SourceInputData.Q66r;
				objInputData.colQ66s = SourceInputData.Q66s;
				objInputData.colQ66t = SourceInputData.Q66t;
				objInputData.colQ66u = SourceInputData.Q66u;
				objInputData.colQ66v = SourceInputData.Q66v;
				objInputData.colQ67a = (SourceInputData.Q67a);
				objInputData.colQ67b = (SourceInputData.Q67b);
				objInputData.colQ67c = (SourceInputData.Q67c);
				//objInputData.colQ68 = Convert.ToChar(SourceInputData.Q68);
				objInputData.colQ68 = SourceInputData.Q68;
				objInputData.colQ69a = (SourceInputData.Q69a);
				objInputData.colQ70a = (SourceInputData.Q70a);
				objInputData.colQ71a = (SourceInputData.Q71a);
				objInputData.colQ72 = SourceInputData.Q72;
				objInputData.colQ73 = SourceInputData.Q73;
				objInputData.colQ74 = SourceInputData.Q74;
				objInputData.colQ75 = SourceInputData.Q75;
				objInputData.colQ75a = SourceInputData.Q75a;
				objInputData.colQ75b = SourceInputData.Q75b;
				objInputData.colQ76 = SourceInputData.Q76;

				//praveenk-Release2
				objInputData.AdditionalInfo = SourceInputData.AdditionalInfo;

				//msinghai - KeyMetrics Release - Newly added columns
				objInputData.colQ89a = SourceInputData.Q89a;
				objInputData.colQ89b = SourceInputData.Q89b;
				objInputData.colQ89c = SourceInputData.Q89c;
				objInputData.colQ89d = SourceInputData.Q89d;
				objInputData.colQ90a = SourceInputData.Q90a;
				objInputData.colQ90b = SourceInputData.Q90b;
				objInputData.colQ90c = SourceInputData.Q90c;
				objInputData.colQ91 = SourceInputData.Q91;
				objInputData.colQ92a = SourceInputData.Q92a;
				objInputData.colQ92b = SourceInputData.Q92b;
				objInputData.colQ92c = SourceInputData.Q92c;
				objInputData.colQ92d = SourceInputData.Q92d;
				objInputData.colQ92e = SourceInputData.Q92e;
				objInputData.colQ92f = SourceInputData.Q92f;
				objInputData.colQ92g = SourceInputData.Q92g;
				objInputData.colQ93a = SourceInputData.Q93a;
				objInputData.colQ93b = SourceInputData.Q93b;
				objInputData.colQ94 = SourceInputData.Q94;
				objInputData.colQ95a = SourceInputData.Q95a;
				objInputData.colQ95b = SourceInputData.Q95b;
				objInputData.colQ95c = SourceInputData.Q95c;
				objInputData.colQ95d = SourceInputData.Q95d;
				objInputData.colQ95e = SourceInputData.Q95e;
				objInputData.colQ95f = SourceInputData.Q95f;
				objInputData.colQ95g = SourceInputData.Q95g;
				objInputData.colQ95h = SourceInputData.Q95h;
				objInputData.colQ96a = SourceInputData.Q96a;
				objInputData.colQ96b = SourceInputData.Q96b;
				objInputData.colQ96c = SourceInputData.Q96c;
				objInputData.colQ96d = SourceInputData.Q96d;
				objInputData.colQ96e = SourceInputData.Q96e;
				objInputData.colQ96f = SourceInputData.Q96f;
				objInputData.colQ96g = SourceInputData.Q96g;
				objInputData.colQ96h = SourceInputData.Q96h;
				objInputData.colQ96i = SourceInputData.Q96i;
				objInputData.colQ97 = SourceInputData.Q97;


				#endregion ReadInputDataFromDatabase

				#region CalculatingAllFormulas

				#region StraightForwardFormula

				//1. Check DIV/0;
				//if (objInputData.colQ14 == 0)
				//    objOutputData.col3a = 0;
				//else
				//    objOutputData.col3a = Math.Round((objInputData.colQ24 / objInputData.colQ14), 2);
				if (objInputData.colQ24 == null && objInputData.colQ14 == null)
					objOutputData.col3a = null;
				else if (objInputData.colQ14 == 0 || objInputData.colQ14 == null)
					objOutputData.col3a = 0;
				else
					objOutputData.col3a = Math.Round((Convert.ToDecimal(objInputData.colQ24) / Convert.ToDecimal(objInputData.colQ14)), 2);

				//2.Check DIV/0;
				//if (objInputData.colQ11 == 0)
				//    objOutputData.col4a = 0;
				//else
				//    objOutputData.col4a = Math.Round((objInputData.colQ14 / objInputData.colQ11), 2);
				if (objInputData.colQ14 == null && objInputData.colQ11 == null)
					objOutputData.col4a = null;
				else if (objInputData.colQ11 == 0 || objInputData.colQ11 == null)
					objOutputData.col4a = 0;
				else
					objOutputData.col4a = Math.Round((Convert.ToDecimal(objInputData.colQ14) / Convert.ToDecimal(objInputData.colQ11)), 2);

				//3.Check DIV/0;
				//if (objInputData.colQ12 == 0)
				//    objOutputData.col5a = 0;
				//else
				//    objOutputData.col5a = Math.Round((objInputData.colQ24 / objInputData.colQ12), 2);
				if (objInputData.colQ24 == null && objInputData.colQ12 == null)
					objOutputData.col5a = null;
				else if (objInputData.colQ12 == 0 || objInputData.colQ12 == null)
					objOutputData.col5a = 0;
				else
					objOutputData.col5a = Math.Round((Convert.ToDecimal(objInputData.colQ24) / Convert.ToDecimal(objInputData.colQ12)), 2);


				//4. Check DIV/0;
				//if (objInputData.colQ12 == 0)
				//    objOutputData.col6a = 0;
				//else
				//    objOutputData.col6a = Math.Round((objInputData.colQ14 / objInputData.colQ12) * 100, 2);
				if (objInputData.colQ14 == null && objInputData.colQ12 == null)
					objOutputData.col6a = null;
				else if (objInputData.colQ12 == 0 || objInputData.colQ12 == null)
					objOutputData.col6a = 0;
				else
					objOutputData.col6a = Math.Round((Convert.ToDecimal(objInputData.colQ14) / Convert.ToDecimal(objInputData.colQ12)) * 100, 2);


				//5. Check DIV/0;
				//if (objInputData.colQ11 == 0)
				//    objOutputData.col7a = 0;
				//else
				//    objOutputData.col7a = Math.Round((objInputData.colQ24 / objInputData.colQ11), 2);
				if (objInputData.colQ24 == null && objInputData.colQ11 == null)
					objOutputData.col7a = null;
				else if (objInputData.colQ11 == 0 || objInputData.colQ11 == null)
					objOutputData.col7a = 0;
				else
					objOutputData.col7a = Math.Round((Convert.ToDecimal(objInputData.colQ24) / Convert.ToDecimal(objInputData.colQ11)), 2);


				//6.
				//objOutputData.col8a = Math.Round((objInputData.colQ11 / 2080), 2); //2080 is Constant.
				if (objInputData.colQ11 == null)
					objOutputData.col8a = null;
				else
					objOutputData.col8a = Math.Round((Convert.ToDecimal(objInputData.colQ11) / 2080), 2); //2080 is Constant.


				//7. Check DIV/0;
				//if (objInputData.colQ11 == 0)
				//    objOutputData.col8b = 0;
				//else
				//    objOutputData.col8b = Math.Round((objInputData.colQ24 / (objInputData.colQ11 / 2080)), 2);
				if (objInputData.colQ24 == null && objInputData.colQ11 == null)
					objOutputData.col8b = null;
				else if (objInputData.colQ11 == 0 || objInputData.colQ11 == null)
					objOutputData.col8b = 0;
				else
					objOutputData.col8b = Math.Round((Convert.ToDecimal(objInputData.colQ24) / (Convert.ToDecimal(objInputData.colQ11) / 2080)), 2);



				//8. Check DIV/0;
				if (objInputData.colQ24 == null && objInputData.colQ7 == null)
					objOutputData.col9a = null;
				else if (objInputData.colQ7 == 0 || objInputData.colQ7 == null)
					objOutputData.col9a = 0;
				else
					objOutputData.col9a = Math.Round((Convert.ToDecimal(objInputData.colQ24) / Convert.ToDecimal(objInputData.colQ7)), 2);

				//9. Check DIV/0;
				//if (objInputData.colQ2 == 0)
				//    objOutputData.col9c = 0;
				//else
				//    objOutputData.col9c = Math.Round((objInputData.colQ24 / objInputData.colQ2), 2);
				if (objInputData.colQ24 == null && objInputData.colQ2 == null)
					objOutputData.col9c = null;
				else if (objInputData.colQ2 == 0 || objInputData.colQ2 == null)
					objOutputData.col9c = 0;
				else
					objOutputData.col9c = Math.Round((Convert.ToDecimal(objInputData.colQ24) / Convert.ToDecimal(objInputData.colQ2)), 2);


				//10. Check DIV/0;
				//if (objInputData.colQ24 == 0)
				//    objOutputData.col12a = 0;
				//else
				//    objOutputData.col12a = Math.Round((objInputData.colQ26f / objInputData.colQ24), 2);
				//issue1-praveenk-fixed.
				if (objInputData.colQ26f == null && objInputData.colQ24 == null)
					objOutputData.col12a = null;
				else if (objInputData.colQ24 == 0 || objInputData.colQ24 == null)
					objOutputData.col12a = 0;
				else
					objOutputData.col12a = Math.Round((Convert.ToDecimal(objInputData.colQ26f) / Convert.ToDecimal(objInputData.colQ24)), 2);


				//11.
				//objOutputData.col13a = Math.Round((objInputData.colQ28 + objInputData.colQ29), 2);
				if (objInputData.colQ28 == null && objInputData.colQ29 == null)
					objOutputData.col13a = null;
				else
					objOutputData.col13a = Math.Round((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)), 2);

				//12. Check DIV/0;
				//if (objInputData.colQ14 == 0)
				//    objOutputData.col13b = 0;
				//else
				//    objOutputData.col13b = Math.Round(((objInputData.colQ28 + objInputData.colQ29) / (objInputData.colQ14 / 100)), 2);
				if (objInputData.colQ28 == null && objInputData.colQ29 == null && objInputData.colQ14 == null)
					objOutputData.col13b = null;
				else if (objInputData.colQ14 == 0 || objInputData.colQ14 == null)
					objOutputData.col13b = 0;
				else
					objOutputData.col13b = Math.Round(((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) / (Convert.ToDecimal(objInputData.colQ14) / 100)), 2);

				//13. Check DIV/0;
				//if ((objInputData.colQ28 + objInputData.colQ29) == 0)
				//    objOutputData.col13e = 0;
				//else
				//    objOutputData.col13e = Math.Round((objInputData.colQ26f / (objInputData.colQ28 + objInputData.colQ29)), 2);
				if (objInputData.colQ26f == null && objInputData.colQ28 == null && objInputData.colQ29 == null)
					objOutputData.col13e = null;
				else if ((objInputData.colQ28 + objInputData.colQ29) == 0 || (objInputData.colQ28 == null && objInputData.colQ29 == null))
					objOutputData.col13e = 0;
				else
					objOutputData.col13e = Math.Round((Convert.ToDecimal(objInputData.colQ26f) / (Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29))), 2);



				//14. Check DIV/0;
				//if ((objInputData.colQ28 + objInputData.colQ29) == 0)
				//    objOutputData.col14a = 0;
				//else
				//    objOutputData.col14a = Math.Round((objInputData.colQ26f / (objInputData.colQ28 + objInputData.colQ29)), 2);
				if (objInputData.colQ26f == null && objInputData.colQ28 == null && objInputData.colQ29 == null)
					objOutputData.col14a = null;
				else if ((objInputData.colQ28 + objInputData.colQ29) == 0 || (objInputData.colQ28 == null && objInputData.colQ29 == null))
					objOutputData.col14a = 0;
				else
					objOutputData.col14a = Math.Round((Convert.ToDecimal(objInputData.colQ26f) / (Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29))), 2);


				//15.
				//objOutputData.col15a = Math.Round((objInputData.colQ52a + objInputData.colQ52b + objInputData.colQ52c + objInputData.colQ52d + objInputData.colQ52e), 2);
				if (objInputData.colQ52a == null && objInputData.colQ52b == null && objInputData.colQ52c == null && objInputData.colQ52d == null && objInputData.colQ52d == null && objInputData.colQ52e == null)
					objOutputData.col15a = null;
				else
					objOutputData.col15a = Math.Round((Convert.ToDecimal(objInputData.colQ52a) + Convert.ToDecimal(objInputData.colQ52b) + Convert.ToDecimal(objInputData.colQ52c) + Convert.ToDecimal(objInputData.colQ52d) + Convert.ToDecimal(objInputData.colQ52e)), 2);


				//16.
				//objOutputData.col15b = Math.Round((objInputData.colQ26f - objInputData.colQ52a - objInputData.colQ52b - objInputData.colQ52c - objInputData.colQ52d - objInputData.colQ52e), 2);
				if (objInputData.colQ26f == null && objInputData.colQ52a == null && objInputData.colQ52b == null && objInputData.colQ52c == null && objInputData.colQ52d == null && objInputData.colQ52e == null)
					objOutputData.col15b = null;
				else
					objOutputData.col15b = Math.Round((Convert.ToDecimal(objInputData.colQ26f) - Convert.ToDecimal(objInputData.colQ52a) - Convert.ToDecimal(objInputData.colQ52b) - Convert.ToDecimal(objInputData.colQ52c) - Convert.ToDecimal(objInputData.colQ52d) - Convert.ToDecimal(objInputData.colQ52e)), 2);


				//17. Check DIV/0;
				//if (objInputData.colQ26f == 0)
				//    objOutputData.col15c = 0;
				//else
				//    objOutputData.col15c = Math.Round((objOutputData.col15b / objInputData.colQ26f), 2);
				//issue2-praveenk-fixed.
				if (objInputData.colQ26f == null && objOutputData.col15b == null)
					objOutputData.col15c = null;
				else if (objInputData.colQ26f == 0 || objInputData.colQ26f == null)
					objOutputData.col15c = 0;
				else
					objOutputData.col15c = Math.Round((Convert.ToDecimal(objOutputData.col15b) / Convert.ToDecimal(objInputData.colQ26f)), 2);


				//18.
				//objOutputData.col16a = Math.Round((objInputData.colQ31b / 100) * (objInputData.colQ28 + objInputData.colQ29), 2);
				if (objInputData.colQ31b == null && objInputData.colQ28 == null && objInputData.colQ29 == null)
					objOutputData.col16a = null;
				else
					objOutputData.col16a = Math.Round((Convert.ToDecimal(objInputData.colQ31b) / 100) * (Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)), 2);

				//19.
				//objOutputData.col16b = Math.Round(((objInputData.colQ31b / 100) * (objInputData.colQ28 + objInputData.colQ29)) * (objInputData.colQ32c / 100), 2);
				if (objInputData.colQ31b == null && objInputData.colQ28 == null && objInputData.colQ29 == null && objInputData.colQ32c == null)
					objOutputData.col16b = null;
				else
					objOutputData.col16b = Math.Round(((Convert.ToDecimal(objInputData.colQ31b) / 100) * (Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29))) * (Convert.ToDecimal(objInputData.colQ32c) / 100), 2);


				//20. Check DIV/0;
				//if (objOutputData.col16a == 0)
				//    objOutputData.col16c = 0;
				//else
				//    objOutputData.col16c = Math.Round((objOutputData.col16b / objOutputData.col16a), 2);
				if (objOutputData.col16b == null && objOutputData.col16a == null)
					objOutputData.col16c = null;
				else if (objOutputData.col16a == 0 || objOutputData.col16a == null)
					objOutputData.col16c = 0;
				else
					objOutputData.col16c = Math.Round((Convert.ToDecimal(objOutputData.col16b) / Convert.ToDecimal(objOutputData.col16a)), 2);


				//21.
				//objOutputData.col17a = Math.Round((objInputData.colQ28 + objInputData.colQ29) * (objInputData.colQ33b / 100), 2);
				if (objInputData.colQ28 == null && objInputData.colQ29 == null && objInputData.colQ33b == null)
					objOutputData.col17a = null;
				else
					objOutputData.col17a = Math.Round((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * (Convert.ToDecimal(objInputData.colQ33b) / 100), 2);

				//22.
				//objOutputData.col18a = Math.Round((objInputData.colQ28 + objInputData.colQ29) * (objInputData.colQ33a / 100), 2);
				if (objInputData.colQ28 == null && objInputData.colQ29 == null && objInputData.colQ33a == null)
					objOutputData.col18a = null;
				else
					objOutputData.col18a = Math.Round((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * (Convert.ToDecimal(objInputData.colQ33a) / 100), 2);

				//23.
				//objOutputData.col19a = Math.Round((objInputData.colQ28 + objInputData.colQ29) * (objInputData.colQ33c / 100), 2);
				if (objInputData.colQ28 == null && objInputData.colQ29 == null && objInputData.colQ33c == null)
					objOutputData.col19a = null;
				else
					objOutputData.col19a = Math.Round((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * (Convert.ToDecimal(objInputData.colQ33c) / 100), 2);

				//24.
				//objOutputData.col20b = Math.Round(((objOutputData.col13a) * ((100 - objInputData.colQ30) / 100)) + (objInputData.colQ30 / 100) * Convert.ToDecimal(2.1), 2);
				if (objOutputData.col13a == null && objInputData.colQ30 == null)
					objOutputData.col20b = null;
				else
					objOutputData.col20b = Math.Round(((Convert.ToDecimal(objOutputData.col13a)) * ((100 - Convert.ToDecimal(objInputData.colQ30)) / 100)) + (Convert.ToDecimal(objInputData.colQ30) / 100) * Convert.ToDecimal(2.1), 2);

				//25.
				//objOutputData.col20c = Math.Round((objOutputData.col20b * (objInputData.colQ30 / 100)), 2);

				if (objOutputData.col20b == null && objInputData.colQ30 == null)
					objOutputData.col20c = null;
				else
					objOutputData.col20c = Math.Round((Convert.ToDecimal(objOutputData.col20b) * (Convert.ToDecimal(objInputData.colQ30) / 100)), 2);

				//26. Check DIV/0;
				//if ((objInputData.colQ28 + objInputData.colQ29) == 0)
				//    objOutputData.col20e = 0;
				//else
				//    objOutputData.col20e = Math.Round((objInputData.colQ26f / (objInputData.colQ28 + objInputData.colQ29)), 2);
				if (objInputData.colQ26f == null && objInputData.colQ28 == null && objInputData.colQ29 == null)
					objOutputData.col20e = null;
				else if ((objInputData.colQ28 + objInputData.colQ29) == 0 || (objInputData.colQ28 == null && objInputData.colQ29 == null))
					objOutputData.col20e = 0;
				else
					objOutputData.col20e = Math.Round((Convert.ToDecimal(objInputData.colQ26f) / (Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29))), 2);


				//27.
				//objOutputData.col21b = Math.Round((objInputData.colQ15b * (objInputData.colQ37 / 100)), 2);
				if (objInputData.colQ15b == null && objInputData.colQ37 == null)
					objOutputData.col21b = null;
				else
					objOutputData.col21b = Math.Round((Convert.ToDecimal(objInputData.colQ15b) * (Convert.ToDecimal(objInputData.colQ37) / 100)), 2);

				//28.
				//objOutputData.col21c = Math.Round((objOutputData.col21b * 150), 2);
				if (objOutputData.col21b == null)
					objOutputData.col21c = null;
				else
					objOutputData.col21c = Math.Round((Convert.ToDecimal(objOutputData.col21b) * 150), 2);

				//29. Check DIV/0;
				//if (objInputData.colQ24 == 0)
				//    objOutputData.col24a = 0;
				//else
				//    objOutputData.col24a = Math.Round((objInputData.colQ26g / objInputData.colQ24), 2);
				if (objInputData.colQ26g == null && objInputData.colQ24 == null)
					objOutputData.col24a = null;
				else if (objInputData.colQ24 == 0 || objInputData.colQ24 == null)
					objOutputData.col24a = 0;
				else
					objOutputData.col24a = Math.Round((Convert.ToDecimal(objInputData.colQ26g) / Convert.ToDecimal(objInputData.colQ24)), 2);

				//30.
				//objOutputData.col25a = Math.Round((objInputData.colQ26g - objInputData.colQ52f), 2);
				if (objInputData.colQ26g == null && objInputData.colQ52f == null)
					objOutputData.col25a = null;
				else
					objOutputData.col25a = Math.Round((Convert.ToDecimal(objInputData.colQ26g) - Convert.ToDecimal(objInputData.colQ52f)), 2);

				//31. Check DIV/0;
				//if (objInputData.colQ26g == 0)
				//    objOutputData.col25b = 0;
				//else
				//    objOutputData.col25b = Math.Round((objOutputData.col25a / objInputData.colQ26g), 2);
				if (objOutputData.col25a == null && objInputData.colQ26g == null)
					objOutputData.col25b = null;
				else if (objInputData.colQ26g == 0 || objInputData.colQ26g == null)
					objOutputData.col25b = 0;
				else
					objOutputData.col25b = Math.Round((Convert.ToDecimal(objOutputData.col25a) / Convert.ToDecimal(objInputData.colQ26g)), 2);



				//32. Check DIV/0;
				//if (objInputData.colQ14 == 0)
				//    objOutputData.col26b = 0;
				//else
				//    objOutputData.col26b = Math.Round((objInputData.colQ15b / objInputData.colQ14), 2);
				if (objInputData.colQ15b == null && objInputData.colQ14 == null)
					objOutputData.col26b = null;
				else if (objInputData.colQ14 == 0 || objInputData.colQ14 == null)
					objOutputData.col26b = 0;
				else
					objOutputData.col26b = Math.Round((Convert.ToDecimal(objInputData.colQ15b) / Convert.ToDecimal(objInputData.colQ14)), 2);


				//33. Check DIV/0;
				//if (objInputData.colQ15b == 0)
				//    objOutputData.col27a = 0;
				//else
				//    objOutputData.col27a = Math.Round((objInputData.colQ26g / objInputData.colQ15b), 2);
				if (objInputData.colQ26g == null && objInputData.colQ15b == null)
					objOutputData.col27a = null;
				else if (objInputData.colQ15b == 0 || objInputData.colQ15b == null)
					objOutputData.col27a = 0;
				else
					objOutputData.col27a = Math.Round((Convert.ToDecimal(objInputData.colQ26g) / Convert.ToDecimal(objInputData.colQ15b)), 2);


				//34. Check DIV/0;
				//if (objInputData.colQ15b == 0)
				//    objOutputData.col28a = 0;
				//else
				//    objOutputData.col28a = Math.Round((objInputData.colQ42a / (objInputData.colQ15b / 100)), 2);
				if (objInputData.colQ42a == null && objInputData.colQ15b == null)
					objOutputData.col28a = null;
				else if (objInputData.colQ15b == 0 || objInputData.colQ15b == null)
					objOutputData.col28a = 0;
				else
					objOutputData.col28a = Math.Round((Convert.ToDecimal(objInputData.colQ42a) / (Convert.ToDecimal(objInputData.colQ15b) / 100)), 2);

				//35.
				//objOutputData.col33a = Math.Round((objInputData.colQ26b + objInputData.colQ26c), 2);
				if (objInputData.colQ26b == null && objInputData.colQ26c == null)
					objOutputData.col33a = null;
				else
					objOutputData.col33a = Math.Round((Convert.ToDecimal(objInputData.colQ26b) + Convert.ToDecimal(objInputData.colQ26c)), 2);

				//36. Check DIV/0;
				//if (objInputData.colQ24 == 0)
				//    objOutputData.col33b = 0;
				//else
				//    objOutputData.col33b = Math.Round((objOutputData.col33a / objInputData.colQ24), 2);
				if (objOutputData.col33a == null && objInputData.colQ24 == null)
					objOutputData.col33b = null;
				else if (objInputData.colQ24 == 0 || objInputData.colQ24 == null)
					objOutputData.col33b = 0;
				else
					objOutputData.col33b = Math.Round((Convert.ToDecimal(objOutputData.col33a) / Convert.ToDecimal(objInputData.colQ24)), 2);


				//37.
				//objOutputData.col33d = Math.Round((objInputData.colQ20a + objInputData.colQ20b + objInputData.colQ20c + objInputData.colQ20d + objInputData.colQ20e + objInputData.colQ20f + objInputData.colQ20g), 2);
				if (objInputData.colQ20a == null && objInputData.colQ20b == null && objInputData.colQ20c == null && objInputData.colQ20d == null && objInputData.colQ20e == null && objInputData.colQ20f == null && objInputData.colQ20g == null)
					objOutputData.col33d = null;
				else
					objOutputData.col33d = Math.Round((Convert.ToDecimal(objInputData.colQ20a) + Convert.ToDecimal(objInputData.colQ20b) + Convert.ToDecimal(objInputData.colQ20c) + Convert.ToDecimal(objInputData.colQ20d) + Convert.ToDecimal(objInputData.colQ20e) + Convert.ToDecimal(objInputData.colQ20f) + Convert.ToDecimal(objInputData.colQ20g)), 2);

				//38.
				//objOutputData.col33e = Math.Round((objOutputData.col33d + objInputData.colQ14), 2);
				if (objOutputData.col33d == null && objInputData.colQ14 == null)
					objOutputData.col33e = null;
				else
					objOutputData.col33e = Math.Round((Convert.ToDecimal(objOutputData.col33d) + Convert.ToDecimal(objInputData.colQ14)), 2);

				//39.                   
				//if ((objOutputData.col33d + objInputData.colQ14) == 0)
				//    objOutputData.col33f = 0;
				//else
				//    objOutputData.col33f = Math.Round((objOutputData.col33d / (objOutputData.col33d + objInputData.colQ14)), 2);
				if (objOutputData.col33d == null && objOutputData.col33d == null && objInputData.colQ14 == null)
					objOutputData.col33f = null;
				else if ((objOutputData.col33d + objInputData.colQ14) == 0 || (objOutputData.col33d == null && objInputData.colQ14 == null))
					objOutputData.col33f = 0;
				else
					objOutputData.col33f = Math.Round((Convert.ToDecimal(objOutputData.col33d) / (Convert.ToDecimal(objOutputData.col33d) + Convert.ToDecimal(objInputData.colQ14))), 2);


				//40.
				//if (objInputData.colQ12 == 0)
				//    objOutputData.col34a = 0;
				//else
				//    objOutputData.col34a = Math.Round((objOutputData.col33d / (objInputData.colQ12 / 1000)), 2);
				if (objOutputData.col33d == null && objInputData.colQ12 == null)
					objOutputData.col34a = null;
				else if (objInputData.colQ12 == 0 || objInputData.colQ12 == null)
					objOutputData.col34a = 0;
				else
					objOutputData.col34a = Math.Round((Convert.ToDecimal(objOutputData.col33d) / (Convert.ToDecimal(objInputData.colQ12) / 1000)), 2);

				//41.
				//objOutputData.col34c = Math.Round((objInputData.colQ21a + objInputData.colQ21b + objInputData.colQ21c + objInputData.colQ21d), 2);
				if (objInputData.colQ21a == null && objInputData.colQ21b == null && objInputData.colQ21c == null && objInputData.colQ21d == null)
					objOutputData.col34c = null;
				else
					objOutputData.col34c = Math.Round((Convert.ToDecimal(objInputData.colQ21a) + Convert.ToDecimal(objInputData.colQ21b) + Convert.ToDecimal(objInputData.colQ21c) + Convert.ToDecimal(objInputData.colQ21d)), 2);

				//42.
				//if (objInputData.colQ12 == 0)
				//    objOutputData.col34d = 0;
				//else
				//    objOutputData.col34d = Math.Round((objOutputData.col34c / (objInputData.colQ12 / 1000)), 2);
				if (objOutputData.col34c == null && objInputData.colQ12 == null)
					objOutputData.col34d = null;
				else if (objInputData.colQ12 == 0 || objInputData.colQ12 == null)
					objOutputData.col34d = 0;
				else
					objOutputData.col34d = Math.Round((Convert.ToDecimal(objOutputData.col34c) / (Convert.ToDecimal(objInputData.colQ12) / 1000)), 2);

				//43.
				//if (objInputData.colQ24 == 0)
				//    objOutputData.col36a = 0;
				//else
				//    objOutputData.col36a = Math.Round((objInputData.colQ56 / objInputData.colQ24), 2);
				if (objInputData.colQ56 == null && objInputData.colQ24 == null)
					objOutputData.col36a = null;
				else if (objInputData.colQ24 == 0 || objInputData.colQ24 == null)
					objOutputData.col36a = 0;
				else
					objOutputData.col36a = Math.Round((Convert.ToDecimal(objInputData.colQ56) / Convert.ToDecimal(objInputData.colQ24)), 2);

				//44.                    
				//if (objInputData.colQ14 == 0)
				//    objOutputData.col36c = 0;
				//else
				//    objOutputData.col36c = Math.Round((objInputData.colQ56 / objInputData.colQ14), 2);
				if (objInputData.colQ56 == null && objInputData.colQ14 == null)
					objOutputData.col36c = null;
				else if (objInputData.colQ14 == 0 || objInputData.colQ14 == null)
					objOutputData.col36c = 0;
				else
					objOutputData.col36c = Math.Round((Convert.ToDecimal(objInputData.colQ56) / Convert.ToDecimal(objInputData.colQ14)), 2);

				//45.
				//if (objInputData.colQ16 == 0)
				//    objOutputData.col37b = 0;
				//else
				//    objOutputData.col37b = Math.Round((objInputData.colQ14 * (objInputData.colQ16 / 100)), 2);
				if (objInputData.colQ14 == null && objInputData.colQ16 == null)
					objOutputData.col37b = null;
				else if (objInputData.colQ16 == 0 || objInputData.colQ16 == null)
					objOutputData.col37b = 0;
				else
					objOutputData.col37b = Math.Round((Convert.ToDecimal(objInputData.colQ14) * (Convert.ToDecimal(objInputData.colQ16) / 100)), 2);


				//46.
				//if (objInputData.colQ14 == 0 || objInputData.colQ16 == 0)
				//    objOutputData.col37c = 0;
				//else
				//    objOutputData.col37c = Math.Round((objInputData.colQ17 / (objInputData.colQ14 * (objInputData.colQ16 / 100))), 2);
				if (objInputData.colQ17 == null && objInputData.colQ14 == null && objInputData.colQ16 == null)
					objOutputData.col37c = null;
				else if (objInputData.colQ14 == 0 || objInputData.colQ16 == 0 || (objInputData.colQ14 == null || objInputData.colQ16 == null))
					objOutputData.col37c = 0;
				else
					objOutputData.col37c = Math.Round((Convert.ToDecimal(objInputData.colQ17) / (Convert.ToDecimal(objInputData.colQ14) * (Convert.ToDecimal(objInputData.colQ16) / 100))), 2);


				//47.
				//objOutputData.col37e = Math.Round((objInputData.colQ8 * 52), 2);
				if (objInputData.colQ8 == null)
					objOutputData.col37e = null;
				else
					objOutputData.col37e = Math.Round((Convert.ToDecimal(objInputData.colQ8) * 52), 2);

				//48.
				//objOutputData.col37f = Math.Round((objInputData.colQ8 * 52 * 60), 2);
				if (objInputData.colQ8 == null)
					objOutputData.col37f = null;
				else
					objOutputData.col37f = Math.Round((Convert.ToDecimal(objInputData.colQ8) * 52 * 60), 2);

				//49.
				//if (objInputData.colQ14 == 0)
				//    objOutputData.col37g = 0;
				//else
				//    objOutputData.col37g = Math.Round(((objInputData.colQ8 * 52 * 60) / objInputData.colQ14), 2);
				if (objInputData.colQ8 == null && objInputData.colQ14 == null)
					objOutputData.col37g = null;
				else if (objInputData.colQ14 == 0 || objInputData.colQ14 == null)
					objOutputData.col37g = 0;
				else
					objOutputData.col37g = Math.Round(((Convert.ToDecimal(objInputData.colQ8) * 52 * 60) / Convert.ToDecimal(objInputData.colQ14)), 2);

				//50.
				//if (objInputData.colQ14 == 0)
				//    objOutputData.col42a = 0;
				//else
				//    objOutputData.col42a = Math.Round((objInputData.colQ26 / objInputData.colQ14), 2);
				if (objInputData.colQ26 == null & objInputData.colQ14 == null)
					objOutputData.col42a = null;
				else if (objInputData.colQ14 == 0 || objInputData.colQ14 == null)
					objOutputData.col42a = 0;
				else
					objOutputData.col42a = Math.Round((Convert.ToDecimal(objInputData.colQ26) / Convert.ToDecimal(objInputData.colQ14)), 2);


				//51.
				//objOutputData.col43a = Math.Round((objInputData.colQ24 / 12), 2);
				if (objInputData.colQ24 == null)
					objOutputData.col43a = null;
				else
					objOutputData.col43a = Math.Round((Convert.ToDecimal(objInputData.colQ24) / 12), 2);

				//52.
				//decimal Q24DivResult = 0;
				//Q24DivResult = Math.Round((objInputData.colQ24 / 12), 2);
				//objOutputData.col43a = Q24DivResult;
				decimal Q24DivResult = 0;
				Q24DivResult = Math.Round((Convert.ToDecimal(objInputData.colQ24) / 12), 2);
				if (objInputData.colQ24 == null)
					objOutputData.col43a = null;
				else
					objOutputData.col43a = Q24DivResult;



				//53. Check DIV/0;
				//if (Q24DivResult == 0)
				//    objOutputData.col43b = 0;
				//else
				//    objOutputData.col43b = Math.Round((objInputData.colQ25 / Q24DivResult), 2);
				if (objInputData.colQ25 == null)
					objOutputData.col43b = null;
				else if (Q24DivResult == 0)
					objOutputData.col43b = 0;
				else
					objOutputData.col43b = Math.Round((Convert.ToDecimal(objInputData.colQ25) / Convert.ToDecimal(Q24DivResult)), 2);




				//54. Check DIV/0;
				//if (Q24DivResult == 0)
				//    objOutputData.col43d = 0;
				//else
				//    objOutputData.col43d = Math.Round(((objInputData.colQ25 / Q24DivResult) * 30), 2);
				if (objInputData.colQ25 == null)
					objOutputData.col43d = null;
				else if (Q24DivResult == 0)
					objOutputData.col43d = 0;
				else
					objOutputData.col43d = Math.Round(((Convert.ToDecimal(objInputData.colQ25) / Convert.ToDecimal(Q24DivResult)) * 30), 2);



				//55. Check DIV/0;
				//if (objInputData.colQ24 == 0)
				//    objOutputData.col43f = 0;
				//else
				//    objOutputData.col43f = Math.Round((objInputData.colQ52j / objInputData.colQ24), 2);
				if (objInputData.colQ52j == null && objInputData.colQ24 == null)
					objOutputData.col43f = null;
				else if (objInputData.colQ24 == 0 || objInputData.colQ24 == null)
					objOutputData.col43f = 0;
				else
					objOutputData.col43f = Math.Round((Convert.ToDecimal(objInputData.colQ52j) / Convert.ToDecimal(objInputData.colQ24)), 2);


				//56. Check DIV/0;
				//if (objInputData.colQ24 == 0)
				//    objOutputData.col44a = 0;
				//else
				//    objOutputData.col44a = Math.Round((objInputData.colQ53 / objInputData.colQ24), 2);
				if (objInputData.colQ53 == null && objInputData.colQ24 == null)
					objOutputData.col44a = null;
				else if (objInputData.colQ24 == 0 || objInputData.colQ24 == null)
					objOutputData.col44a = 0;
				else
					objOutputData.col44a = Math.Round((Convert.ToDecimal(objInputData.colQ53) / Convert.ToDecimal(objInputData.colQ24)), 2);



				//57. Check DIV/0;
				//if (objInputData.colQ24 == 0)
				//    objOutputData.col44c = 0;
				//else
				//    objOutputData.col44c = Math.Round((objInputData.colQ54 / objInputData.colQ24), 2);
				if (objInputData.colQ54 == null && objInputData.colQ24 == null)
					objOutputData.col44c = null;
				else if (objInputData.colQ24 == 0 || objInputData.colQ24 == null)
					objOutputData.col44c = 0;
				else
					objOutputData.col44c = Math.Round((Convert.ToDecimal(objInputData.colQ54) / Convert.ToDecimal(objInputData.colQ24)), 2);


				//58.
				//objOutputData.col45a = Math.Round((objInputData.colQ52j + objInputData.colQ53 + objInputData.colQ54 + objInputData.colQ55 + objInputData.colQ56 + objInputData.colQ57 + objInputData.colQ58 + objInputData.colQ59 + objInputData.colQ60), 2);
				if (objInputData.colQ52j == null && objInputData.colQ53 == null && objInputData.colQ54 == null && objInputData.colQ55 == null && objInputData.colQ56 == null && objInputData.colQ57 == null && objInputData.colQ58 == null && objInputData.colQ59 == null && objInputData.colQ60 == null)
					objOutputData.col45a = null;
				else
					objOutputData.col45a = Math.Round((Convert.ToDecimal(objInputData.colQ52j) + Convert.ToDecimal(objInputData.colQ53) + Convert.ToDecimal(objInputData.colQ54) + Convert.ToDecimal(objInputData.colQ55) + Convert.ToDecimal(objInputData.colQ56) + Convert.ToDecimal(objInputData.colQ57) + Convert.ToDecimal(objInputData.colQ58) + Convert.ToDecimal(objInputData.colQ59) + Convert.ToDecimal(objInputData.colQ60)), 2);

				//59.
				//objOutputData.col45b = Math.Round((objInputData.colQ24 - objOutputData.col45a), 2);
				if (objInputData.colQ24 == null && objOutputData.col45a == null)
					objOutputData.col45b = null;
				else
					objOutputData.col45b = Math.Round((Convert.ToDecimal(objInputData.colQ24) - Convert.ToDecimal(objOutputData.col45a)), 2);


				//60. Check DIV/0;
				//if (objInputData.colQ24 == 0)
				//    objOutputData.col45c = 0;
				//else
				//    objOutputData.col45c = Math.Round((objOutputData.col45b / objInputData.colQ24), 2);
				if (objOutputData.col45b == null && objInputData.colQ24 == null)
					objOutputData.col45c = null;
				else if (objInputData.colQ24 == 0 || objInputData.colQ24 == null)
					objOutputData.col45c = 0;
				else
					objOutputData.col45c = Math.Round((Convert.ToDecimal(objOutputData.col45b) / Convert.ToDecimal(objInputData.colQ24)), 2);

				//61.
				//objOutputData.col45e = Math.Round((objInputData.colQ53 + objInputData.colQ54 + objInputData.colQ55 + objInputData.colQ56 + objInputData.colQ57 + objInputData.colQ58 + objInputData.colQ59 + objInputData.colQ60), 2);
				if (objInputData.colQ53 == null && objInputData.colQ54 == null && objInputData.colQ55 == null && objInputData.colQ56 == null && objInputData.colQ57 == null && objInputData.colQ58 == null && objInputData.colQ59 == null && objInputData.colQ60 == null)
					objOutputData.col45e = null;
				else
					objOutputData.col45e = Math.Round((Convert.ToDecimal(objInputData.colQ53) + Convert.ToDecimal(objInputData.colQ54) + Convert.ToDecimal(objInputData.colQ55) + Convert.ToDecimal(objInputData.colQ56) + Convert.ToDecimal(objInputData.colQ57) + Convert.ToDecimal(objInputData.colQ58) + Convert.ToDecimal(objInputData.colQ59) + Convert.ToDecimal(objInputData.colQ60)), 2);

				//62. Check DIV/0;
				//if (objInputData.colQ14 == 0)
				//    objOutputData.col45f = 0;
				//else
				//    objOutputData.col45f = Math.Round((objOutputData.col45e / objInputData.colQ14), 2);
				if (objOutputData.col45e == null && objInputData.colQ14 == null)
					objOutputData.col45f = null;
				else if (objInputData.colQ14 == 0 || objInputData.colQ14 == null)
					objOutputData.col45f = 0;
				else
					objOutputData.col45f = Math.Round((Convert.ToDecimal(objOutputData.col45e) / Convert.ToDecimal(objInputData.colQ14)), 2);


				List<bool?> lstQ66nTo66v = new List<bool?>();
				lstQ66nTo66v.Add(objInputData.colQ66n);
				lstQ66nTo66v.Add(objInputData.colQ66o);
				lstQ66nTo66v.Add(objInputData.colQ66p);
				lstQ66nTo66v.Add(objInputData.colQ66q);
				lstQ66nTo66v.Add(objInputData.colQ66r);
				lstQ66nTo66v.Add(objInputData.colQ66s);
				lstQ66nTo66v.Add(objInputData.colQ66t);
				lstQ66nTo66v.Add(objInputData.colQ66u);
				lstQ66nTo66v.Add(objInputData.colQ66v);
				int countOftrueInQ66nTo66v = 0;         //This count variable used in two formulas.
				countOftrueInQ66nTo66v = lstQ66nTo66v.Count(item => EqualityComparer<object>.Default.Equals(item, true));

				List<bool?> lstQ64aTo66k = new List<bool?>();
				lstQ64aTo66k.Add(objInputData.colQ64a);
				lstQ64aTo66k.Add(objInputData.colQ64b);
				lstQ64aTo66k.Add(objInputData.colQ64c);
				lstQ64aTo66k.Add(objInputData.colQ64d);
				lstQ64aTo66k.Add(objInputData.colQ64e);
				lstQ64aTo66k.Add(objInputData.colQ64f);
				lstQ64aTo66k.Add(objInputData.colQ64g);
				lstQ64aTo66k.Add(objInputData.colQ64h);
				lstQ64aTo66k.Add(objInputData.colQ64i);
				lstQ64aTo66k.Add(objInputData.colQ64j);
				lstQ64aTo66k.Add(objInputData.colQ64k);
				lstQ64aTo66k.Add(objInputData.colQ64l);
				lstQ64aTo66k.Add(objInputData.colQ64m);
				lstQ64aTo66k.Add(objInputData.colQ64n);
				lstQ64aTo66k.Add(objInputData.colQ64o);
				lstQ64aTo66k.Add(objInputData.colQ65a);
				lstQ64aTo66k.Add(objInputData.colQ65b);
				lstQ64aTo66k.Add(objInputData.colQ65c);
				lstQ64aTo66k.Add(objInputData.colQ65d);
				lstQ64aTo66k.Add(objInputData.colQ65e);
				lstQ64aTo66k.Add(objInputData.colQ65f);
				lstQ64aTo66k.Add(objInputData.colQ65g);
				lstQ64aTo66k.Add(objInputData.colQ65h);
				lstQ64aTo66k.Add(objInputData.colQ65i);
				lstQ64aTo66k.Add(objInputData.colQ65j);
				lstQ64aTo66k.Add(objInputData.colQ65k);
				lstQ64aTo66k.Add(objInputData.colQ65l);
				lstQ64aTo66k.Add(objInputData.colQ65m);
				lstQ64aTo66k.Add(objInputData.colQ65n);
				lstQ64aTo66k.Add(objInputData.colQ66a);
				lstQ64aTo66k.Add(objInputData.colQ66b);
				lstQ64aTo66k.Add(objInputData.colQ66c);
				lstQ64aTo66k.Add(objInputData.colQ66d);
				lstQ64aTo66k.Add(objInputData.colQ66e);
				lstQ64aTo66k.Add(objInputData.colQ66f);
				lstQ64aTo66k.Add(objInputData.colQ66g);
				lstQ64aTo66k.Add(objInputData.colQ66h);
				lstQ64aTo66k.Add(objInputData.colQ66i);
				lstQ64aTo66k.Add(objInputData.colQ66j);
				lstQ64aTo66k.Add(objInputData.colQ66k);
				int countOftrueInQ64aTo66k = 0;
				countOftrueInQ64aTo66k = lstQ64aTo66k.Count(item => EqualityComparer<object>.Default.Equals(item, true));
				//63.
				objOutputData.col49a = Math.Round(((Convert.ToDecimal(countOftrueInQ64aTo66k) * 10) + countOftrueInQ66nTo66v), 2);

				List<bool?> lstQ64aTo64o = new List<bool?>();
				lstQ64aTo64o.Add(objInputData.colQ64a);
				lstQ64aTo64o.Add(objInputData.colQ64b);
				lstQ64aTo64o.Add(objInputData.colQ64c);
				lstQ64aTo64o.Add(objInputData.colQ64d);
				lstQ64aTo64o.Add(objInputData.colQ64e);
				lstQ64aTo64o.Add(objInputData.colQ64f);
				lstQ64aTo64o.Add(objInputData.colQ64g);
				lstQ64aTo64o.Add(objInputData.colQ64h);
				lstQ64aTo64o.Add(objInputData.colQ64i);
				lstQ64aTo64o.Add(objInputData.colQ64j);
				lstQ64aTo64o.Add(objInputData.colQ64k);
				lstQ64aTo64o.Add(objInputData.colQ64l);
				lstQ64aTo64o.Add(objInputData.colQ64m);
				lstQ64aTo64o.Add(objInputData.colQ64n);
				lstQ64aTo64o.Add(objInputData.colQ64o);
				int countOftrueInQ64aTo64o = 0;
				countOftrueInQ64aTo64o = lstQ64aTo64o.Count(item => EqualityComparer<object>.Default.Equals(item, true));
				//64.
				objOutputData.col49c = Math.Round(((Convert.ToDecimal(countOftrueInQ64aTo64o) * 10)), 2);

				List<bool?> lstQ65aTo65n = new List<bool?>();
				lstQ65aTo65n.Add(objInputData.colQ65a);
				lstQ65aTo65n.Add(objInputData.colQ65b);
				lstQ65aTo65n.Add(objInputData.colQ65c);
				lstQ65aTo65n.Add(objInputData.colQ65d);
				lstQ65aTo65n.Add(objInputData.colQ65e);
				lstQ65aTo65n.Add(objInputData.colQ65f);
				lstQ65aTo65n.Add(objInputData.colQ65g);
				lstQ65aTo65n.Add(objInputData.colQ65h);
				lstQ65aTo65n.Add(objInputData.colQ65i);
				lstQ65aTo65n.Add(objInputData.colQ65j);
				lstQ65aTo65n.Add(objInputData.colQ65k);
				lstQ65aTo65n.Add(objInputData.colQ65l);
				lstQ65aTo65n.Add(objInputData.colQ65m);
				lstQ65aTo65n.Add(objInputData.colQ65n);
				int countOftrueInQ65aTo65n = 0;
				countOftrueInQ65aTo65n = lstQ65aTo65n.Count(item => EqualityComparer<object>.Default.Equals(item, true));
				//65.
				objOutputData.col50a = Math.Round(((Convert.ToDecimal(countOftrueInQ65aTo65n) * 10)), 2);

				List<bool?> lstQ66aTo66k = new List<bool?>();
				lstQ66aTo66k.Add(objInputData.colQ66a);
				lstQ66aTo66k.Add(objInputData.colQ66b);
				lstQ66aTo66k.Add(objInputData.colQ66c);
				lstQ66aTo66k.Add(objInputData.colQ66d);
				lstQ66aTo66k.Add(objInputData.colQ66e);
				lstQ66aTo66k.Add(objInputData.colQ66f);
				lstQ66aTo66k.Add(objInputData.colQ66g);
				lstQ66aTo66k.Add(objInputData.colQ66h);
				lstQ66aTo66k.Add(objInputData.colQ66i);
				lstQ66aTo66k.Add(objInputData.colQ66j);
				lstQ66aTo66k.Add(objInputData.colQ66k);
				int countOftrueInQ66aTo66k = 0;
				countOftrueInQ66aTo66k = lstQ66aTo66k.Count(item => EqualityComparer<object>.Default.Equals(item, true));
				//66.
				objOutputData.col50c = Math.Round(((Convert.ToDecimal(countOftrueInQ66aTo66k) * 10) + countOftrueInQ66nTo66v), 2);

				//67.
				//objOutputData.coln12a = Math.Round((objOutputData.col12a * 100), 2);
				if (objOutputData.col12a == null)
					objOutputData.coln12a = null;
				else
					objOutputData.coln12a = Math.Round((Convert.ToDecimal(objOutputData.col12a) * 100), 2);

				//68.
				//objOutputData.coln15c = Math.Round((objOutputData.col15c * 100), 2);
				if (objOutputData.col15c == null)
					objOutputData.coln15c = null;
				else
					objOutputData.coln15c = Math.Round((Convert.ToDecimal(objOutputData.col15c) * 100), 2);

				//69.
				//objOutputData.coln16c = Math.Round((objOutputData.col16c * 100), 2);
				if (objOutputData.col16c == null)
					objOutputData.coln16c = null;
				else
					objOutputData.coln16c = Math.Round((Convert.ToDecimal(objOutputData.col16c) * 100), 2);

				//70.
				//objOutputData.coln24a = Math.Round((objOutputData.col24a * 100), 2);
				if (objOutputData.col24a == null)
					objOutputData.coln24a = null;
				else
					objOutputData.coln24a = Math.Round((Convert.ToDecimal(objOutputData.col24a) * 100), 2);

				//71.
				//objOutputData.coln25b = Math.Round((objOutputData.col25b * 100), 2);
				if (objOutputData.col25b == null)
					objOutputData.coln25b = null;
				else
					objOutputData.coln25b = Math.Round((Convert.ToDecimal(objOutputData.col25b) * 100), 2);


				//72.
				//objOutputData.coln26b = Math.Round((objOutputData.col26b * 100), 2);
				if (objOutputData.col26b == null)
					objOutputData.coln26b = null;
				else
					objOutputData.coln26b = Math.Round((Convert.ToDecimal(objOutputData.col26b) * 100), 2);


				//73.
				//objOutputData.coln33b = Math.Round((objOutputData.col33b * 100), 2);
				if (objOutputData.col33b == null)
					objOutputData.coln33b = null;
				else
					objOutputData.coln33b = Math.Round((Convert.ToDecimal(objOutputData.col33b) * 100), 2);

				//74.
				//objOutputData.coln33f = Math.Round((objOutputData.col33f * 100), 2);
				if (objOutputData.col33f == null)
					objOutputData.coln33f = null;
				else
					objOutputData.coln33f = Math.Round((Convert.ToDecimal(objOutputData.col33f) * 100), 2);

				//75.
				//objOutputData.coln36a = Math.Round((objOutputData.col36a * 100), 2);
				if (objOutputData.col36a == null)
					objOutputData.coln36a = null;
				else
					objOutputData.coln36a = Math.Round((Convert.ToDecimal(objOutputData.col36a) * 100), 2);

				//76.
				//objOutputData.coln37c = Math.Round((objOutputData.col37c * 100), 2);
				if (objOutputData.col37c == null)
					objOutputData.coln37c = null;
				else
					objOutputData.coln37c = Math.Round((Convert.ToDecimal(objOutputData.col37c) * 100), 2);

				//77.
				//objOutputData.coln43b = Math.Round((objOutputData.col43b * 100), 2);
				if (objOutputData.col43b == null)
					objOutputData.coln43b = null;
				else
					objOutputData.coln43b = Math.Round((Convert.ToDecimal(objOutputData.col43b) * 100), 2);

				//78.
				//objOutputData.coln43f = Math.Round((objOutputData.col43f * 100), 2);
				if (objOutputData.col43f == null)
					objOutputData.coln43f = null;
				else
					objOutputData.coln43f = Math.Round((Convert.ToDecimal(objOutputData.col43f) * 100), 2);

				//79.
				//objOutputData.coln44a = Math.Round((objOutputData.col44a * 100), 2);
				if (objOutputData.col44a == null)
					objOutputData.coln44a = null;
				else
					objOutputData.coln44a = Math.Round((Convert.ToDecimal(objOutputData.col44a) * 100), 2);

				//80.
				//objOutputData.coln44c = Math.Round((objOutputData.col44c * 100), 2);
				if (objOutputData.col44c == null)
					objOutputData.coln44c = null;
				else
					objOutputData.coln44c = Math.Round((Convert.ToDecimal(objOutputData.col44c) * 100), 2);

				//81.
				//objOutputData.coln45c = Math.Round((objOutputData.col45c * 100), 2);
				if (objOutputData.col45c == null)
					objOutputData.coln45c = null;
				else
					objOutputData.coln45c = Math.Round((Convert.ToDecimal(objOutputData.col45c) * 100), 2);

				//Radio button output starts here.
				//82.
				objOutputData.coln64a = objInputData.colQ64a;
				//83.
				objOutputData.coln64b = objInputData.colQ64b;
				//84.
				objOutputData.coln64c = objInputData.colQ64c;
				//85.
				objOutputData.coln64d = objInputData.colQ64d;
				//86.
				objOutputData.coln64e = objInputData.colQ64e;
				//87.
				objOutputData.coln64f = objInputData.colQ64f;
				//88.
				objOutputData.coln64g = objInputData.colQ64g;
				//89.
				objOutputData.coln64h = objInputData.colQ64h;
				//90.
				objOutputData.coln64i = objInputData.colQ64i;
				//91.
				objOutputData.coln64j = objInputData.colQ64j;
				//92.
				objOutputData.coln64k = objInputData.colQ64k;
				//93.
				objOutputData.coln64l = objInputData.colQ64l;
				//94.
				objOutputData.coln64m = objInputData.colQ64m;
				//95.
				objOutputData.coln64n = objInputData.colQ64n;
				//96.
				objOutputData.coln64o = objInputData.colQ64o;
				//97.
				objOutputData.coln65a = objInputData.colQ65a;
				//98.
				objOutputData.coln65b = objInputData.colQ65b;
				//99.
				objOutputData.coln65c = objInputData.colQ65c;
				//100.
				objOutputData.coln65d = objInputData.colQ65d;
				//101.
				objOutputData.coln65e = objInputData.colQ65e;
				//102.
				objOutputData.coln65f = objInputData.colQ65f;
				//103.
				objOutputData.coln65g = objInputData.colQ65g;
				//104.
				objOutputData.coln65h = objInputData.colQ65h;
				//105.
				objOutputData.coln65i = objInputData.colQ65i;
				//106.
				objOutputData.coln65j = objInputData.colQ65j;
				//107.
				objOutputData.coln65k = objInputData.colQ65k;
				//108.
				objOutputData.coln65l = objInputData.colQ65l;
				//109.
				objOutputData.coln65m = objInputData.colQ65m;
				//110.
				objOutputData.coln65n = objInputData.colQ65n;
				//111.
				objOutputData.coln66a = objInputData.colQ66a;
				//112.
				objOutputData.coln66b = objInputData.colQ66b;
				//113.
				objOutputData.coln66c = objInputData.colQ66c;
				//114.
				objOutputData.coln66d = objInputData.colQ66d;
				//115.
				objOutputData.coln66e = objInputData.colQ66e;
				//116.
				objOutputData.coln66f = objInputData.colQ66f;
				//117.
				objOutputData.coln66g = objInputData.colQ66g;
				//118.
				objOutputData.coln66h = objInputData.colQ66h;
				//119.
				objOutputData.coln66i = objInputData.colQ66i;
				//120.
				objOutputData.coln66j = objInputData.colQ66j;
				//121.
				objOutputData.coln66k = objInputData.colQ66k;
				//122.
				objOutputData.coln66l = objInputData.colQ66l;
				//123.
				objOutputData.coln66m = objInputData.colQ66m;
				//124.
				objOutputData.coln66n = objInputData.colQ66n;
				//125.
				objOutputData.coln66o = objInputData.colQ66o;
				//126.
				objOutputData.coln66p = objInputData.colQ66p;
				//127.
				objOutputData.coln66q = objInputData.colQ66q;
				//128.
				objOutputData.coln66r = objInputData.colQ66r;
				//129.
				objOutputData.coln66s = objInputData.colQ66s;
				//130.
				objOutputData.coln66t = objInputData.colQ66t;
				//131.
				objOutputData.coln66u = objInputData.colQ66u;
				//132.
				objOutputData.coln66v = objInputData.colQ66v;
				//Radio button output ends here.

				//133.
				if (objInputData.colQ68 == "1")
					objOutputData.coln68 = "Male";
				else if (objInputData.colQ68 == "2")
					objOutputData.coln68 = "Female";
				else
					objOutputData.coln68 = "";

				//134.
				//objOutputData.colt20 = Math.Round((objInputData.colQ20a + objInputData.colQ20b + objInputData.colQ20c + objInputData.colQ20d + objInputData.colQ20e + objInputData.colQ20f + objInputData.colQ20g), 2);
				if (objInputData.colQ20a == null && objInputData.colQ20b == null && objInputData.colQ20c == null && objInputData.colQ20d == null && objInputData.colQ20e == null && objInputData.colQ20f == null && objInputData.colQ20g == null)
					objOutputData.colt20 = null;
				else
					objOutputData.colt20 = Math.Round((Convert.ToDecimal(objInputData.colQ20a) + Convert.ToDecimal(objInputData.colQ20b) + Convert.ToDecimal(objInputData.colQ20c) + Convert.ToDecimal(objInputData.colQ20d) + Convert.ToDecimal(objInputData.colQ20e) + Convert.ToDecimal(objInputData.colQ20f) + Convert.ToDecimal(objInputData.colQ20g)), 2);

				//135.
				//objOutputData.colt21 = Math.Round((objInputData.colQ21a + objInputData.colQ21b + objInputData.colQ21c + objInputData.colQ21d), 2);
				if (objInputData.colQ21a == null && objInputData.colQ21b == null && objInputData.colQ21c == null && objInputData.colQ21d == null)
					objOutputData.colt21 = null;
				else
					objOutputData.colt21 = Math.Round((Convert.ToDecimal(objInputData.colQ21a) + Convert.ToDecimal(objInputData.colQ21b) + Convert.ToDecimal(objInputData.colQ21c) + Convert.ToDecimal(objInputData.colQ21d)), 2);

				//137.
				//objOutputData.colt27 = Math.Round((objInputData.colQ27a + objInputData.colQ27c), 2);
				if (objInputData.colQ27a == null && objInputData.colQ27c == null)
					objOutputData.colt27 = null;
				else
					objOutputData.colt27 = Math.Round((Convert.ToDecimal(objInputData.colQ27a) + Convert.ToDecimal(objInputData.colQ27c)), 2);

				#endregion StraightForwardFormulas

				#region LookupFormulas

				//1.
				//objOutputData.col3b = GetLookUpLable("Lookup.GrossRevenuePerCompleteExam", objOutputData.col3a, "$");
				if (objOutputData.col3a == null)
					objOutputData.col3b = null;
				else
					objOutputData.col3b = GetLookUpLable("Lookup.GrossRevenuePerCompleteExam", Convert.ToDecimal(objOutputData.col3a), "$");
				//2.
				//objOutputData.col4b = GetLookUpLable("Lookup.CompleteExamsPerODHour", objOutputData.col4a);
				if (objOutputData.col4a == null)
					objOutputData.col4b = null;
				else
					objOutputData.col4b = GetLookUpLable("Lookup.CompleteExamsPerODHour", Convert.ToDecimal(objOutputData.col4a));

				//3.
				//objOutputData.col5b = GetLookUpLable("Lookup.GrossRevPerActivePatient", objOutputData.col5a, "$");
				if (objOutputData.col5a == null)
					objOutputData.col5b = null;
				else
					objOutputData.col5b = GetLookUpLable("Lookup.GrossRevPerActivePatient", Convert.ToDecimal(objOutputData.col5a), "$");

				//4.
				//objOutputData.col6b = GetLookUpLable("Lookup.CompleteExamsPer100Active", objOutputData.col6a);
				if (objOutputData.col6a == null)
					objOutputData.col6b = null;
				else
					objOutputData.col6b = GetLookUpLable("Lookup.CompleteExamsPer100Active", Convert.ToDecimal(objOutputData.col6a));

				//5.
				//objOutputData.col7b = GetLookUpLable("Lookup.GrossRevenuePerODHour", objOutputData.col7a, "$");
				if (objOutputData.col7a == null)
					objOutputData.col7b = null;
				else
					objOutputData.col7b = GetLookUpLable("Lookup.GrossRevenuePerODHour", Convert.ToDecimal(objOutputData.col7a), "$");

				//6.
				//objOutputData.col8c = GetLookUpLable("Lookup.AnnGrossRevPerFTEOD($000)", objOutputData.col8b, "$");
				if (objOutputData.col8b == null)
					objOutputData.col8c = null;
				else
					objOutputData.col8c = GetLookUpLable("Lookup.AnnGrossRevPerFTEOD($000)", Convert.ToDecimal(objOutputData.col8b), "$");

				//7.
				//objOutputData.col9b = GetLookUpLable("Lookup.GrossRevPerNonODStaffHr", objOutputData.col9a, "$");
				if (objOutputData.col9a == null)
					objOutputData.col9b = null;
				else
					objOutputData.col9b = GetLookUpLable("Lookup.GrossRevPerNonODStaffHr", Convert.ToDecimal(objOutputData.col9a), "$");

				//8.
				//objOutputData.col9d = GetLookUpLable("Lookup.GrossRevPerSqFt", objOutputData.col9c, "$");
				if (objOutputData.col9c == null)
					objOutputData.col9d = null;
				else
					objOutputData.col9d = GetLookUpLable("Lookup.GrossRevPerSqFt", Convert.ToDecimal(objOutputData.col9c), "$");

				//9.
				//objOutputData.col12b = GetLookUpLable("Lookup.EyewearSalePercentageOfGrossRev", objOutputData.col12a);
				if (objOutputData.col12a == null)
					objOutputData.col12b = null;
				else
					objOutputData.col12b = GetLookUpLable("Lookup.EyewearSalePercentageOfGrossRev", Convert.ToDecimal(objOutputData.col12a));

				//10.
				//objOutputData.col13c = GetLookUpLable("Lookup.EyewearRxPer100ComplExam", objOutputData.col13b);
				if (objOutputData.col13b == null)
					objOutputData.col13c = null;
				else
					objOutputData.col13c = GetLookUpLable("Lookup.EyewearRxPer100ComplExam", Convert.ToDecimal(objOutputData.col13b));

				//11.
				//objOutputData.col14b = GetLookUpLable("Lookup.GrossRevPerEyewearRx", objOutputData.col14a, "$");
				if (objOutputData.col14a == null)
					objOutputData.col14b = null;
				else
					objOutputData.col14b = GetLookUpLable("Lookup.GrossRevPerEyewearRx", Convert.ToDecimal(objOutputData.col14a), "$");

				//12.
				//objOutputData.col15d = GetLookUpLable("Lookup.EyewearGrossProfitMargin", objOutputData.col15c);
				if (objOutputData.col15c == null)
					objOutputData.col15d = null;
				else
					objOutputData.col15d = GetLookUpLable("Lookup.EyewearGrossProfitMargin", Convert.ToDecimal(objOutputData.col15c));

				//13.
				//objOutputData.col16d = GetLookUpLable("Lookup.ProgressiveLensAndPresbyopRx", objOutputData.col16c * 100, "%");//Here the Input value is in percent.But in database,Lookup tables are saved as without percent.Hence it required to multiply the Input by 100.
				if (objOutputData.col16c == null)
					objOutputData.col16d = null;
				else
					objOutputData.col16d = GetLookUpLable("Lookup.ProgressiveLensAndPresbyopRx", Convert.ToDecimal(objOutputData.col16c) * 100, "%");
				//14.
				//objOutputData.col17b = GetLookUpLable("Lookup.NoGlareLensPercentSpecLensRx", objInputData.colQ33b, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ33b == null)
					objOutputData.col17b = null;
				else
					objOutputData.col17b = GetLookUpLable("Lookup.NoGlareLensPercentSpecLensRx", Convert.ToDecimal(objInputData.colQ33b), "%");

				//15.
				//objOutputData.col18b = GetLookUpLable("Lookup.HighIndexLensPercentSpecLensRx", objInputData.colQ33a, "%"); //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ33a == null)
					objOutputData.col18b = null;
				else
					objOutputData.col18b = GetLookUpLable("Lookup.HighIndexLensPercentSpecLensRx", Convert.ToDecimal(objInputData.colQ33a), "%");
				//16.
				//objOutputData.col19b = GetLookUpLable("Lookup.PhotochrLensPercentofSpecLensRx", objInputData.colQ33c, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ33c == null)
					objOutputData.col19b = null;
				else
					objOutputData.col19b = GetLookUpLable("Lookup.PhotochrLensPercentofSpecLensRx", Convert.ToDecimal(objInputData.colQ33c), "%");

				//17.
				//objOutputData.col20a = GetLookUpLable("Lookup.MultipleEyewearPurchasePercent", objInputData.colQ30, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ30 == null)
					objOutputData.col20a = null;
				else
					objOutputData.col20a = GetLookUpLable("Lookup.MultipleEyewearPurchasePercent", Convert.ToDecimal(objInputData.colQ30), "%");

				//18.
				//objOutputData.col21a = GetLookUpLable("Lookup.PercentPatientsCLExamPurchEyewea", objInputData.colQ37, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ37 == null)
					objOutputData.col21a = null;
				else
					objOutputData.col21a = GetLookUpLable("Lookup.PercentPatientsCLExamPurchEyewea", Convert.ToDecimal(objInputData.colQ37), "%");

				//19.
				//objOutputData.col24b = GetLookUpLable("Lookup.CLSalesPercentGrossRev", objOutputData.col24a);
				if (objOutputData.col24a == null)
					objOutputData.col24b = null;
				else
					objOutputData.col24b = GetLookUpLable("Lookup.CLSalesPercentGrossRev", Convert.ToDecimal(objOutputData.col24a));

				//20.
				//objOutputData.col25c = GetLookUpLable("Lookup.CLGrossProfitMargin", objOutputData.col25b);
				if (objOutputData.col25b == null)
					objOutputData.col25c = null;
				else
					objOutputData.col25c = GetLookUpLable("Lookup.CLGrossProfitMargin", Convert.ToDecimal(objOutputData.col25b));

				//21.
				//objOutputData.col26a = GetLookUpLable("Lookup.CLWearerPercentActivePatients", objInputData.colQ13b, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ13b == null)
					objOutputData.col26a = null;
				else
					objOutputData.col26a = GetLookUpLable("Lookup.CLWearerPercentActivePatients", Convert.ToDecimal(objInputData.colQ13b), "%");

				//22.
				//objOutputData.col26c = GetLookUpLable("Lookup.CLExamPercentTotalExam", objOutputData.col26b);
				if (objOutputData.col26b == null)
					objOutputData.col26c = null;
				else
					objOutputData.col26c = GetLookUpLable("Lookup.CLExamPercentTotalExam", Convert.ToDecimal(objOutputData.col26b));

				//23.
				//objOutputData.col27b = GetLookUpLable("Lookup.AnnCLSalesPerCLExam", objOutputData.col27a, "$");
				if (objOutputData.col27a == null)
					objOutputData.col27b = null;
				else
					objOutputData.col27b = GetLookUpLable("Lookup.AnnCLSalesPerCLExam", Convert.ToDecimal(objOutputData.col27a), "$");

				//24.
				//objOutputData.col28b = GetLookUpLable("Lookup.CLNewFitsPer100CLExam", objOutputData.col28a);
				if (objOutputData.col28a == null)
					objOutputData.col28b = null;
				else
					objOutputData.col28b = GetLookUpLable("Lookup.CLNewFitsPer100CLExam", Convert.ToDecimal(objOutputData.col28a));

				//25.
				//objOutputData.col28c = GetLookUpLable("Lookup.CLRefitPercentCLExam", objInputData.colQ43a, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ43a == null)
					objOutputData.col28c = null;
				else
					objOutputData.col28c = GetLookUpLable("Lookup.CLRefitPercentCLExam", Convert.ToDecimal(objInputData.colQ43a), "%");

				//26.
				//objOutputData.col29a = GetLookUpLable("Lookup.SiliconeHydroLensWearPercentSoft", objInputData.colQ41a, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ41a == null)
					objOutputData.col29a = null;
				else
					objOutputData.col29a = GetLookUpLable("Lookup.SiliconeHydroLensWearPercentSoft", Convert.ToDecimal(objInputData.colQ41a), "%");

				//27.
				//objOutputData.col29b = GetLookUpLable("Lookup.DailyDisposableLensPercentSoft", objInputData.colQ39a, "%");//Here no need to divide the input by 100 because in database Lookup tables are saved as without percent.
				if (objInputData.colQ39a == null)
					objOutputData.col29b = null;
				else
					objOutputData.col29b = GetLookUpLable("Lookup.DailyDisposableLensPercentSoft", Convert.ToDecimal(objInputData.colQ39a), "%");

				//28.
				//objOutputData.col29c = GetLookUpLable("Lookup.MonthlySoftLensPercentWearers", objInputData.colQ39c, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ39c == null)
					objOutputData.col29c = null;
				else
					objOutputData.col29c = GetLookUpLable("Lookup.MonthlySoftLensPercentWearers", Convert.ToDecimal(objInputData.colQ39c), "%");

				//29.
				//objOutputData.col30a = GetLookUpLable("Lookup.SoftToricPercentSoftLens", objInputData.colQ40b, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ40b == null)
					objOutputData.col30a = null;
				else
					objOutputData.col30a = GetLookUpLable("Lookup.SoftToricPercentSoftLens", Convert.ToDecimal(objInputData.colQ40b), "%");

				//30.
				//objOutputData.col30b = GetLookUpLable("Lookup.SoftMultiFocPercentSoftLens", objInputData.colQ40d, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ40d == null)
					objOutputData.col30b = null;
				else
					objOutputData.col30b = GetLookUpLable("Lookup.SoftMultiFocPercentSoftLens", Convert.ToDecimal(objInputData.colQ40d), "%");

				//31.
				//objOutputData.col33c = GetLookUpLable("Lookup.NonRefrFeePercentGrossRev", objOutputData.col33b);
				if (objOutputData.col33b == null)
					objOutputData.col33c = null;
				else
					objOutputData.col33c = GetLookUpLable("Lookup.NonRefrFeePercentGrossRev", Convert.ToDecimal(objOutputData.col33b));

				//32.
				//objOutputData.col33g = GetLookUpLable("Lookup.MedicalEyeCareVisitPercentTotal", objOutputData.col33f);
				if (objOutputData.col33f == null)
					objOutputData.col33g = null;
				else
					objOutputData.col33g = GetLookUpLable("Lookup.MedicalEyeCareVisitPercentTotal", Convert.ToDecimal(objOutputData.col33f));

				//33.
				//objOutputData.col34b = GetLookUpLable("Lookup.AnnMedEyeCareVisitPer1000", objOutputData.col34a);
				if (objOutputData.col34a == null)
					objOutputData.col34b = null;
				else
					objOutputData.col34b = GetLookUpLable("Lookup.AnnMedEyeCareVisitPer1000", Convert.ToDecimal(objOutputData.col34a));

				//34.
				//objOutputData.col34e = GetLookUpLable("Lookup.AnnPharmRxPer1000", objOutputData.col34d);
				if (objOutputData.col34d == null)
					objOutputData.col34e = null;
				else
					objOutputData.col34e = GetLookUpLable("Lookup.AnnPharmRxPer1000", Convert.ToDecimal(objOutputData.col34d));

				//35.
				//objOutputData.col36b = GetLookUpLable("Lookup.MrktSpendPercentGrossRev", objOutputData.col36a * 100, "%");//Here the Input value is in percent.But in database,Lookup tables are saved as without percent.Hence it required to multiply the Input by 100.
				if (objOutputData.col36a == null)
					objOutputData.col36b = null;
				else
					objOutputData.col36b = GetLookUpLable("Lookup.MrktSpendPercentGrossRev", Convert.ToDecimal(objOutputData.col36a) * 100, "%");

				//36.
				//objOutputData.col36d = GetLookUpLable("Lookup.AnnMrktSpendPerComplExam", objOutputData.col36c, "$");
				if (objOutputData.col36c == null)
					objOutputData.col36d = null;
				else
					objOutputData.col36d = GetLookUpLable("Lookup.AnnMrktSpendPerComplExam", Convert.ToDecimal(objOutputData.col36c), "$");

				//37.
				//objOutputData.col36e = GetLookUpLable("Lookup.NewPatientExamPercentTotalExam", objInputData.colQ16, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ16 == null)
					objOutputData.col36e = null;
				else
					objOutputData.col36e = GetLookUpLable("Lookup.NewPatientExamPercentTotalExam", Convert.ToDecimal(objInputData.colQ16), "%");

				//38.
				//objOutputData.col37a = GetLookUpLable("Lookup.WebsiteAnnualExpense", objInputData.colQ63, "$");
				if (objInputData.colQ63 == null)
					objOutputData.col37a = null;
				else
					objOutputData.col37a = GetLookUpLable("Lookup.WebsiteAnnualExpense", Convert.ToDecimal(objInputData.colQ63), "$");

				//39.
				//objOutputData.col37d = GetLookUpLable("Lookup.PercentOfNewPatientsAttracted", objOutputData.col37c);
				if (objOutputData.col37c == null)
					objOutputData.col37d = null;
				else
					objOutputData.col37d = GetLookUpLable("Lookup.PercentOfNewPatientsAttracted", Convert.ToDecimal(objOutputData.col37c));

				//40.
				//objOutputData.col37h = GetLookUpLable("Lookup.RecallMinPerComplExam", objOutputData.col37g);
				if (objOutputData.col37g == null)
					objOutputData.col37h = null;
				else
					objOutputData.col37h = GetLookUpLable("Lookup.RecallMinPerComplExam", Convert.ToDecimal(objOutputData.col37g));

				//41.
				//objOutputData.col40a = GetLookUpLable("Lookup.ExamFeeNonCL", objInputData.colQ47, "$");
				if (objInputData.colQ47 == null)
					objOutputData.col40a = null;
				else
					objOutputData.col40a = GetLookUpLable("Lookup.ExamFeeNonCL", Convert.ToDecimal(objInputData.colQ47), "$");

				//42.
				//objOutputData.col40b = GetLookUpLable("Lookup.ExamFeeSoftNewFitSPHERE", objInputData.colQ48, "$");
				if (objInputData.colQ48 == null)
					objOutputData.col40b = null;
				else
					objOutputData.col40b = GetLookUpLable("Lookup.ExamFeeSoftNewFitSPHERE", Convert.ToDecimal(objInputData.colQ48), "$");

				//43.
				//objOutputData.col40c = GetLookUpLable("Lookup.ExamFeeSoftNewFitTORIC", objInputData.colQ49, "$");
				if (objInputData.colQ49 == null)
					objOutputData.col40c = null;
				else
					objOutputData.col40c = GetLookUpLable("Lookup.ExamFeeSoftNewFitTORIC", Convert.ToDecimal(objInputData.colQ49), "$");

				//44.
				//objOutputData.col41a = GetLookUpLable("Lookup.ExamFeeSoftNewFitMULTIFO", objInputData.colQ50, "$");
				if (objInputData.colQ50 == null)
					objOutputData.col41a = null;
				else
					objOutputData.col41a = GetLookUpLable("Lookup.ExamFeeSoftNewFitMULTIFO", Convert.ToDecimal(objInputData.colQ50), "$");

				//45.
				//objOutputData.col41b = GetLookUpLable("Lookup.ExamFeeSoftLensNOREFITT", objInputData.colQ51, "$");
				if (objInputData.colQ51 == null)
					objOutputData.col41b = null;
				else
					objOutputData.col41b = GetLookUpLable("Lookup.ExamFeeSoftLensNOREFITT", Convert.ToDecimal(objInputData.colQ51), "$");

				//46.
				//objOutputData.col42b = GetLookUpLable("Lookup.AvgCollectFeeRevPerCompl", objOutputData.col42a, "$");
				if (objOutputData.col42a == null)
					objOutputData.col42b = null;
				else
					objOutputData.col42b = GetLookUpLable("Lookup.AvgCollectFeeRevPerCompl", Convert.ToDecimal(objOutputData.col42a), "$");

				//47.
				//objOutputData.col42c = GetLookUpLable("Lookup.PercentExamsProvideWMangCareDis", objInputData.colQ18, "%");//Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				if (objInputData.colQ18 == null)
					objOutputData.col42c = null;
				else
					objOutputData.col42c = GetLookUpLable("Lookup.PercentExamsProvideWMangCareDis", Convert.ToDecimal(objInputData.colQ18), "%");

				//48.
				//objOutputData.col43e = GetLookUpLable("Lookup.AcctRecDaysOutstanding", objOutputData.col43d);
				if (objOutputData.col43d == null)
					objOutputData.col43e = null;
				else
					objOutputData.col43e = GetLookUpLable("Lookup.AcctRecDaysOutstanding", Convert.ToDecimal(objOutputData.col43d));

				//49.
				//objOutputData.col43g = GetLookUpLable("Lookup.CostOfGoodsPercentOfGrossRev", objOutputData.col43f);
				if (objOutputData.col43f == null)
					objOutputData.col43g = null;
				else
					objOutputData.col43g = GetLookUpLable("Lookup.CostOfGoodsPercentOfGrossRev", Convert.ToDecimal(objOutputData.col43f));

				//50.
				//objOutputData.col44b = GetLookUpLable("Lookup.StaffExpensePercentOfGrossRev", objOutputData.col44a);
				if (objOutputData.col44a == null)
					objOutputData.col44b = null;
				else
					objOutputData.col44b = GetLookUpLable("Lookup.StaffExpensePercentOfGrossRev", Convert.ToDecimal(objOutputData.col44a));

				//51.
				//objOutputData.col44d = GetLookUpLable("Lookup.OccupancyExpensePercentGrossRev", objOutputData.col44c);
				if (objOutputData.col44c == null)
					objOutputData.col44d = null;
				else
					objOutputData.col44d = GetLookUpLable("Lookup.OccupancyExpensePercentGrossRev", Convert.ToDecimal(objOutputData.col44c));

				//52.
				//objOutputData.col45d = GetLookUpLable("Lookup.NetIncomePercentGrossRev", objOutputData.col45c);
				if (objOutputData.col45c == null)
					objOutputData.col45d = null;
				else
					objOutputData.col45d = GetLookUpLable("Lookup.NetIncomePercentGrossRev", Convert.ToDecimal(objOutputData.col45c));

				//53.
				//objOutputData.col45g = GetLookUpLable("Lookup.ChairCostPerComplExam", objOutputData.col45f, "$");
				if (objOutputData.col45f == null)
					objOutputData.col45g = null;
				else
					objOutputData.col45g = GetLookUpLable("Lookup.ChairCostPerComplExam", Convert.ToDecimal(objOutputData.col45f), "$");

				//54.
				//objOutputData.col49b = GetLookUpLable("Lookup.BestPracticeTOTAL", objOutputData.col49a);
				if (objOutputData.col49a == null)
					objOutputData.col49b = null;
				else
					objOutputData.col49b = GetLookUpLable("Lookup.BestPracticeTOTAL", Convert.ToDecimal(objOutputData.col49a));

				//55.
				//objOutputData.col49d = GetLookUpLable("Lookup.BestPracticeFINANCE", objOutputData.col49c);
				if (objOutputData.col49c == null)
					objOutputData.col49d = null;
				else
					objOutputData.col49d = GetLookUpLable("Lookup.BestPracticeFINANCE", Convert.ToDecimal(objOutputData.col49c));

				//56.
				//objOutputData.col50b = GetLookUpLable("Lookup.BestPracticeMARKETING", objOutputData.col50a);
				if (objOutputData.col50a == null)
					objOutputData.col50b = null;
				else
					objOutputData.col50b = GetLookUpLable("Lookup.BestPracticeMARKETING", Convert.ToDecimal(objOutputData.col50a));

				//57.
				//objOutputData.col50d = GetLookUpLable("Lookup.BestPracticeSTAFF", objOutputData.col50c);
				if (objOutputData.col50c == null)
					objOutputData.col50d = null;
				else
					objOutputData.col50d = GetLookUpLable("Lookup.BestPracticeSTAFF", Convert.ToDecimal(objOutputData.col50c));

				#endregion LookupFormulas

				#region PercentileFormulas
				decimal AdditionOfcolQ28andQ29 = 0;

				//1.

				//if (GetLookUpPercentile("Lookup.GrossRevenuePerCompleteExam", objOutputData.col3a, "$") > 74)
				//    objOutputData.col3c = "Performance achieved";
				//else
				//    objOutputData.col3c = Convert.ToString(Math.Round((objInputData.colQ14 * 371), 2));

				if (objOutputData.col3a == null)
					objOutputData.col3c = null;
				else
				{
					if (GetLookUpPercentile("Lookup.GrossRevenuePerCompleteExam", Convert.ToDecimal(objOutputData.col3a), "$") > 74)
						objOutputData.col3c = "Performance achieved";
					else
					{
						if (objInputData.colQ14 == null)
							objOutputData.col3c = null;
						else
							objOutputData.col3c = Convert.ToString(Math.Round((Convert.ToDecimal(objInputData.colQ14) * 371), 2));
					}
				}



				//2.
				//if (GetLookUpPercentile("Lookup.GrossRevenuePerCompleteExam", objOutputData.col3a, "$") > 74)
				//    objOutputData.col3d = 0;
				//else
				//    objOutputData.col3d = Math.Round(((objInputData.colQ14 * 371) - objInputData.colQ24), 2);
				if (objOutputData.col3a == null)
					objOutputData.col3d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.GrossRevenuePerCompleteExam", Convert.ToDecimal(objOutputData.col3a), "$") > 74)
						objOutputData.col3d = 0;
					else
					{
						if (objInputData.colQ14 == null && objInputData.colQ24 == null)
							objOutputData.col3d = null;
						else
							objOutputData.col3d = Math.Round(((Convert.ToDecimal(objInputData.colQ14) * 371) - Convert.ToDecimal(objInputData.colQ24)), 2);
					}
				}


				//3.
				//if (GetLookUpPercentile("Lookup.CompleteExamsPerODHour", objOutputData.col4a) > 74)
				//    objOutputData.col4c = "Performance achieved";
				//else
				//    objOutputData.col4c = Convert.ToString(Math.Round((objInputData.colQ11 * 1.44M), 2));
				if (objOutputData.col4a == null)
					objOutputData.col4c = null;
				else
				{
					if (GetLookUpPercentile("Lookup.CompleteExamsPerODHour", Convert.ToDecimal(objOutputData.col4a)) > 74)
						objOutputData.col4c = "Performance achieved";
					else
					{
						if (objInputData.colQ11 == null)
							objOutputData.col4c = null;
						else
							objOutputData.col4c = Convert.ToString(Math.Round((Convert.ToDecimal(objInputData.colQ11) * 1.44M), 2));
					}
				}



				//4.
				//if (GetLookUpPercentile("Lookup.CompleteExamsPerODHour", objOutputData.col4a) > 74)
				//    objOutputData.col4d = "Performance achieved";
				//else
				//    objOutputData.col4d = Convert.ToString(Math.Round((objInputData.colQ11 * 1.44M * objOutputData.col3a), 2));
				if (objOutputData.col4a == null)
					objOutputData.col4d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.CompleteExamsPerODHour", Convert.ToDecimal(objOutputData.col4a)) > 74)
						objOutputData.col4d = "Performance achieved";
					else
					{
						if (objInputData.colQ11 == null && objOutputData.col3a == null)
							objOutputData.col4d = null;
						else
							objOutputData.col4d = Convert.ToString(Math.Round((Convert.ToDecimal(objInputData.colQ11) * 1.44M * Convert.ToDecimal(objOutputData.col3a)), 2));
					}
				}


				//5.
				//if (GetLookUpPercentile("Lookup.CompleteExamsPerODHour", objOutputData.col4a) > 74)
				//    objOutputData.col4e = 0;
				//else
				//    objOutputData.col4e = Math.Round(((objInputData.colQ11 * 1.44M * objOutputData.col3a) - objInputData.colQ24), 2);
				if (objOutputData.col4a == null)
					objOutputData.col4e = null;
				else
				{
					if (GetLookUpPercentile("Lookup.CompleteExamsPerODHour", Convert.ToDecimal(objOutputData.col4a)) > 74)
						objOutputData.col4e = 0;
					else
					{
						if (objInputData.colQ11 == null && objOutputData.col3a == null && objInputData.colQ24 == null)
							objOutputData.col4e = null;
						else
							objOutputData.col4e = Math.Round(((Convert.ToDecimal(objInputData.colQ11) * 1.44M * Convert.ToDecimal(objOutputData.col3a)) - Convert.ToDecimal(objInputData.colQ24)), 2);
					}
				}

				//6.
				//if (GetLookUpPercentile("Lookup.GrossRevPerActivePatient", objOutputData.col5a, "$") > 74)
				//    objOutputData.col5c = "Performance achieved";
				//else
				//    objOutputData.col5c = Convert.ToString(Math.Round((objInputData.colQ12 * 176), 2));
				if (objOutputData.col5a == null)
					objOutputData.col5c = null;
				else
				{
					if (GetLookUpPercentile("Lookup.GrossRevPerActivePatient", Convert.ToDecimal(objOutputData.col5a), "$") > 74)
						objOutputData.col5c = "Performance achieved";
					else
					{
						if (objInputData.colQ12 == null)
							objOutputData.col5c = null;
						else
							objOutputData.col5c = Convert.ToString(Math.Round((Convert.ToDecimal(objInputData.colQ12) * 176), 2));
					}
				}


				//7.
				//if (GetLookUpPercentile("Lookup.GrossRevPerActivePatient", objOutputData.col5a, "$") > 74)
				//    objOutputData.col5d = 0;
				//else
				//    objOutputData.col5d = Math.Round(((objInputData.colQ12 * 176) - objInputData.colQ24), 2);
				if (objOutputData.col5a == null)
					objOutputData.col5d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.GrossRevPerActivePatient", Convert.ToDecimal(objOutputData.col5a), "$") > 74)
						objOutputData.col5d = 0;
					else
					{
						if (objInputData.colQ12 == null && objInputData.colQ24 == null)
							objOutputData.col5d = null;
						else
							objOutputData.col5d = Math.Round(((Convert.ToDecimal(objInputData.colQ12) * 176) - Convert.ToDecimal(objInputData.colQ24)), 2);
					}
				}


				//8.
				//if (GetLookUpPercentile("Lookup.CompleteExamsPer100Active", objOutputData.col6a) > 74)
				//    objOutputData.col6c = "Performance achieved";
				//else
				//    objOutputData.col6c = Convert.ToString(Math.Round(((54 * (objInputData.colQ12) / 100)), 2));
				if (objOutputData.col6a == null)
					objOutputData.col6c = null;
				else
				{
					if (GetLookUpPercentile("Lookup.CompleteExamsPer100Active", Convert.ToDecimal(objOutputData.col6a)) > 74)
						objOutputData.col6c = "Performance achieved";
					else
					{
						if (objInputData.colQ12 == null)
							objOutputData.col6c = null;
						else
							objOutputData.col6c = Convert.ToString(Math.Round(((54 * (Convert.ToDecimal(objInputData.colQ12)) / 100)), 2));
					}
				}


				//9.
				//if (GetLookUpPercentile("Lookup.CompleteExamsPer100Active", objOutputData.col6a) > 74)
				//    objOutputData.col6d = "Performance achieved";
				//else
				//    objOutputData.col6d = Convert.ToString(Math.Round(((objOutputData.col3a * 54 * (objInputData.colQ12 / 100))), 2));
				if (objOutputData.col6a == null)
					objOutputData.col6d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.CompleteExamsPer100Active", Convert.ToDecimal(objOutputData.col6a)) > 74)
						objOutputData.col6d = "Performance achieved";
					else
					{
						if (objOutputData.col3a == null && objInputData.colQ12 == null)
							objOutputData.col6d = null;
						else
							objOutputData.col6d = Convert.ToString(Math.Round(((Convert.ToDecimal(objOutputData.col3a) * 54 * (Convert.ToDecimal(objInputData.colQ12) / 100))), 2));
					}
				}



				//10.
				//if (GetLookUpPercentile("Lookup.CompleteExamsPer100Active", objOutputData.col6a) > 74)
				//    objOutputData.col6e = 0;
				//else
				//{
				//    // Check DIV/0;
				//    if (objInputData.colQ14 == 0)  
				//        objOutputData.col6e = 0;
				//    else
				//        objOutputData.col6e = Math.Round(((objInputData.colQ24 / objInputData.colQ14) * 54 * (objInputData.colQ12 / 100) - objInputData.colQ24), 2);
				//}
				if (objOutputData.col6a == null)
					objOutputData.col6e = null;
				else
				{
					if (GetLookUpPercentile("Lookup.CompleteExamsPer100Active", Convert.ToDecimal(objOutputData.col6a)) > 74)
						objOutputData.col6e = 0;
					else
					{
						// Check DIV/0;
						if ((objInputData.colQ14 == 0 && objInputData.colQ14 == null) || (objInputData.colQ24 == null))
							objOutputData.col6e = 0;
						else
							objOutputData.col6e = Math.Round(((Convert.ToDecimal(objInputData.colQ24) / Convert.ToDecimal(objInputData.colQ14)) * 54 * (Convert.ToDecimal(objInputData.colQ12) / 100) - Convert.ToDecimal(objInputData.colQ24)), 2);
					}
				}


				//11.
				//if (GetLookUpPercentile("Lookup.GrossRevenuePerODHour", objOutputData.col7a, "$") > 74)
				//    objOutputData.col7c = "Performance achieved";
				//else
				//    objOutputData.col7c = Convert.ToString(Math.Round((objInputData.colQ11 * 426), 2));
				if (objOutputData.col7a == null)
					objOutputData.col7c = null;
				else
				{
					if (GetLookUpPercentile("Lookup.GrossRevenuePerODHour", Convert.ToDecimal(objOutputData.col7a), "$") > 74)
						objOutputData.col7c = "Performance achieved";
					else
					{
						if (objInputData.colQ11 == null)
							objOutputData.col7c = null;
						else
							objOutputData.col7c = Convert.ToString(Math.Round((Convert.ToDecimal(objInputData.colQ11) * 426), 2));
					}
				}


				//12.
				//if (GetLookUpPercentile("Lookup.GrossRevenuePerODHour", objOutputData.col7a, "$") > 74)
				//    objOutputData.col7d = 0;
				//else
				//    objOutputData.col7d = Math.Round(((objInputData.colQ11 * 426) - objInputData.colQ24), 2);
				if (objOutputData.col7a == null)
					objOutputData.col7d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.GrossRevenuePerODHour", Convert.ToDecimal(objOutputData.col7a), "$") > 74)
						objOutputData.col7d = 0;
					else
					{
						if (objInputData.colQ11 == null && objInputData.colQ24 == null)
							objOutputData.col7d = null;
						else
							objOutputData.col7d = Math.Round(((Convert.ToDecimal(objInputData.colQ11) * 426) - Convert.ToDecimal(objInputData.colQ24)), 2);
					}
				}

				//13.
				//if (GetLookUpPercentile("Lookup.AnnGrossRevPerFTEOD($000)", objOutputData.col8b, "$") > 74)
				//    objOutputData.col8d = "Performance achieved";
				//else
				//    objOutputData.col8d = Convert.ToString(Math.Round(((objInputData.colQ11 / 2080) * 881000), 2));
				if (objOutputData.col8b == null)
					objOutputData.col8d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.AnnGrossRevPerFTEOD($000)", Convert.ToDecimal(objOutputData.col8b), "$") > 74)
						objOutputData.col8d = "Performance achieved";
					else
					{
						if (objInputData.colQ11 == null)
							objOutputData.col8d = null;
						else
							objOutputData.col8d = Convert.ToString(Math.Round(((Convert.ToDecimal(objInputData.colQ11) / 2080) * 881000), 2));
					}
				}


				//14.
				//if (GetLookUpPercentile("Lookup.AnnGrossRevPerFTEOD($000)", objOutputData.col8b, "$") > 74)
				//    objOutputData.col8e = 0;
				//else
				//    objOutputData.col8e = Math.Round((((objInputData.colQ11 / 2080) * 881000) - objInputData.colQ24), 2);
				if (objOutputData.col8b == null)
					objOutputData.col8e = null;
				else
				{
					if (GetLookUpPercentile("Lookup.AnnGrossRevPerFTEOD($000)", Convert.ToDecimal(objOutputData.col8b), "$") > 74)
						objOutputData.col8e = 0;
					else
					{
						if (objInputData.colQ11 == null && objInputData.colQ24 == null)
							objOutputData.col8e = null;
						else
							objOutputData.col8e = Math.Round((((Convert.ToDecimal(objInputData.colQ11) / 2080) * 881000) - Convert.ToDecimal(objInputData.colQ24)), 2);
					}

				}


				//15.
				//if (GetLookUpPercentile("Lookup.EyewearRxPer100ComplExam", objOutputData.col13b) > 74)
				//    objOutputData.col13d = "Performance achieved";
				//else
				//    objOutputData.col13d = Convert.ToString(Math.Round(((objInputData.colQ14 / 100) * 76), 2));
				if (objOutputData.col13b == null)
					objOutputData.col13d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.EyewearRxPer100ComplExam", Convert.ToDecimal(objOutputData.col13b)) > 74)
						objOutputData.col13d = "Performance achieved";
					else
					{
						if (objInputData.colQ14 == null)
							objOutputData.col13d = null;
						else
							objOutputData.col13d = Convert.ToString(Math.Round(((Convert.ToDecimal(objInputData.colQ14) / 100) * 76), 2));
					}
				}




				//16.
				//if (GetLookUpPercentile("Lookup.EyewearRxPer100ComplExam", objOutputData.col13b) > 74)
				//    objOutputData.col13f = "Performance achieved";
				//else
				//{
				//    AdditionOfcolQ28andQ29 = objInputData.colQ28 + objInputData.colQ29;
				//    // Check DIV/0;
				//    if (AdditionOfcolQ28andQ29 == 0)
				//        objOutputData.col13f = 0.ToString();
				//    else
				//        objOutputData.col13f = Convert.ToString(Math.Round(((objInputData.colQ26f / AdditionOfcolQ28andQ29) * ((objInputData.colQ14 / 100) * 76)), 2));
				//}
				if (objOutputData.col13b == null)
					objOutputData.col13f = null;
				else
				{
					if (GetLookUpPercentile("Lookup.EyewearRxPer100ComplExam", Convert.ToDecimal(objOutputData.col13b)) > 74)
						objOutputData.col13f = "Performance achieved";
					else
					{
						AdditionOfcolQ28andQ29 = Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29);
						if (AdditionOfcolQ28andQ29 == 0 || (objInputData.colQ28 == null && objInputData.colQ29 == null))
							objOutputData.col13f = 0.ToString();
						else
							objOutputData.col13f = Convert.ToString(Math.Round(((Convert.ToDecimal(objInputData.colQ26f) / Convert.ToDecimal(AdditionOfcolQ28andQ29)) * ((Convert.ToDecimal(objInputData.colQ14) / 100) * 76)), 2));
					}
				}



				//17.
				//if (GetLookUpPercentile("Lookup.EyewearRxPer100ComplExam", objOutputData.col13b) > 74)
				//    objOutputData.col13g = 0;
				//else
				//{
				//    AdditionOfcolQ28andQ29 = objInputData.colQ28 + objInputData.colQ29;
				//    // Check DIV/0;
				//    if (AdditionOfcolQ28andQ29 == 0)
				//        objOutputData.col13g = 0;
				//    else
				//        objOutputData.col13g = Math.Round((((objInputData.colQ26f / AdditionOfcolQ28andQ29) * ((objInputData.colQ14 / 100) * 76)) - objInputData.colQ26f), 2);
				//}
				if (objOutputData.col13b == null)
					objOutputData.col13g = null;
				else
				{
					if (GetLookUpPercentile("Lookup.EyewearRxPer100ComplExam", Convert.ToDecimal(objOutputData.col13b)) > 74)
						objOutputData.col13g = 0;
					else
					{
						AdditionOfcolQ28andQ29 = Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29);
						// Check DIV/0;
						if (AdditionOfcolQ28andQ29 == 0 || (objInputData.colQ28 == null && objInputData.colQ29 == null))
							objOutputData.col13g = 0;
						else
							objOutputData.col13g = Math.Round((((Convert.ToDecimal(objInputData.colQ26f) / Convert.ToDecimal(AdditionOfcolQ28andQ29)) * ((Convert.ToDecimal(objInputData.colQ14) / 100) * 76)) - Convert.ToDecimal(objInputData.colQ26f)), 2);
					}
				}


				//18.
				//if (GetLookUpPercentile("Lookup.GrossRevPerEyewearRx", objOutputData.col14a, "$") > 74)
				//    objOutputData.col14c = "Performance achieved";
				//else
				//    objOutputData.col14c = Convert.ToString(Math.Round(((objInputData.colQ28 + objInputData.colQ29) * 288), 2));
				if (objOutputData.col14a == null)
					objOutputData.col14c = null;
				else
				{
					if (GetLookUpPercentile("Lookup.GrossRevPerEyewearRx", Convert.ToDecimal(objOutputData.col14a), "$") > 74)
						objOutputData.col14c = "Performance achieved";
					else
					{
						if (objInputData.colQ28 == null && objInputData.colQ29 == null)
							objOutputData.col14c = null;
						else
							objOutputData.col14c = Convert.ToString(Math.Round(((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * 288), 2));
					}
				}


				//19.

				//if (GetLookUpPercentile("Lookup.GrossRevPerEyewearRx", objOutputData.col14a, "$") > 74)
				//    objOutputData.col14d = 0;
				//else
				//    objOutputData.col14d = Math.Round((((objInputData.colQ28 + objInputData.colQ29) * 288) - objInputData.colQ26f), 2);

				if (objOutputData.col14a == null)
					objOutputData.col14d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.GrossRevPerEyewearRx", Convert.ToDecimal(objOutputData.col14a), "$") > 74)
						objOutputData.col14d = 0;
					else
					{
						if (objInputData.colQ28 == null && objInputData.colQ29 == null && objInputData.colQ26f == null)
							objOutputData.col14d = null;
						else
							objOutputData.col14d = Math.Round((((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * 288) - Convert.ToDecimal(objInputData.colQ26f)), 2);
					}
				}


				//20.
				//if (GetLookUpPercentile("Lookup.EyewearGrossProfitMargin", objOutputData.col15c) > 74)
				//    objOutputData.col15e = "Performance achieved";
				//else
				//    objOutputData.col15e = Convert.ToString(Math.Round((objInputData.colQ26f * 0.66M), 2));
				if (objOutputData.col15c == null)
					objOutputData.col15e = null;
				else
				{
					if (GetLookUpPercentile("Lookup.EyewearGrossProfitMargin", Convert.ToDecimal(objOutputData.col15c)) > 74)
						objOutputData.col15e = "Performance achieved";
					else
					{
						if (objInputData.colQ26f == null)
							objOutputData.col15e = null;
						else
							objOutputData.col15e = Convert.ToString(Math.Round((Convert.ToDecimal(objInputData.colQ26f) * 0.66M), 2));
					}
				}


				//21.
				//if (GetLookUpPercentile("Lookup.EyewearGrossProfitMargin", objOutputData.col15c) > 74)
				//    objOutputData.col15f = 0;
				//else
				//    objOutputData.col15f = Math.Round(((objInputData.colQ26f * 0.66M) - objOutputData.col15b), 2);
				if (objOutputData.col15c == null)
					objOutputData.col15f = null;
				else
				{
					if (GetLookUpPercentile("Lookup.EyewearGrossProfitMargin", Convert.ToDecimal(objOutputData.col15c)) > 74)
						objOutputData.col15f = 0;
					else
					{
						if (objInputData.colQ26f == null && objOutputData.col15b == null)
							objOutputData.col15f = null;
						else
							objOutputData.col15f = Math.Round(((Convert.ToDecimal(objInputData.colQ26f) * 0.66M) - Convert.ToDecimal(objOutputData.col15b)), 2);
					}
				}

				//22.
				//if (GetLookUpPercentile("Lookup.ProgressiveLensAndPresbyopRx", objOutputData.col16c * 100, "%") > 74)
				//    objOutputData.col16e = "Performance achieved";
				//else
				//    objOutputData.col16e = Convert.ToString(Math.Round((((objInputData.colQ31b / 100) * (objInputData.colQ28 + objInputData.colQ29)) * 0.75M), 2));

				if (objOutputData.col16c == null)
					objOutputData.col16e = null;
				else
				{
					if (GetLookUpPercentile("Lookup.ProgressiveLensAndPresbyopRx", Convert.ToDecimal(objOutputData.col16c) * 100, "%") > 74)
						objOutputData.col16e = "Performance achieved";
					else
					{
						if (objInputData.colQ31b == null && objInputData.colQ28 == null && objInputData.colQ29 == null)
							objOutputData.col16e = null;
						else
							objOutputData.col16e = Convert.ToString(Math.Round((((Convert.ToDecimal(objInputData.colQ31b) / 100) * (Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29))) * 0.75M), 2));
					}
				}


				//23.
				//if (GetLookUpPercentile("Lookup.ProgressiveLensAndPresbyopRx", objOutputData.col16c * 100, "%") > 74)
				//    objOutputData.col16f = 0;
				//else
				//    objOutputData.col16f = Math.Round((((((objInputData.colQ31b / 100) * (objInputData.colQ28 + objInputData.colQ29)) * 0.75M) - (((objInputData.colQ31b / 100) * (objInputData.colQ28 + objInputData.colQ29)) * (objInputData.colQ32c / 100))) * 106), 2);
				if (objOutputData.col16c == null)
					objOutputData.col16f = null;
				else
				{
					if (GetLookUpPercentile("Lookup.ProgressiveLensAndPresbyopRx", Convert.ToDecimal(objOutputData.col16c) * 100, "%") > 74)
						objOutputData.col16f = 0;
					else
					{
						if (objInputData.colQ31b == null && objInputData.colQ28 == null && objInputData.colQ29 == null && objInputData.colQ32c == null)
							objOutputData.col16f = null;
						else
							objOutputData.col16f = Math.Round((((((Convert.ToDecimal(objInputData.colQ31b) / 100) * (Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29))) * 0.75M) - (((Convert.ToDecimal(objInputData.colQ31b) / 100) * (Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29))) * (Convert.ToDecimal(objInputData.colQ32c) / 100))) * 106), 2);
					}
				}


				//24.
				//if (GetLookUpPercentile("Lookup.NoGlareLensPercentSpecLensRx", objInputData.colQ33b, "%") > 74)  //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col17c = "Performance achieved";
				//else
				//    objOutputData.col17c = Convert.ToString(Math.Round(((objInputData.colQ28 + objInputData.colQ29) * 0.75M), 2));

				if (objInputData.colQ33b == null)
					objOutputData.col17c = null;
				else
				{
					if (GetLookUpPercentile("Lookup.NoGlareLensPercentSpecLensRx", Convert.ToDecimal(objInputData.colQ33b), "%") > 74)  //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col17c = "Performance achieved";
					else
					{
						if (objInputData.colQ28 == null && objInputData.colQ29 == null)
							objOutputData.col17c = null;
						else
							objOutputData.col17c = Convert.ToString(Math.Round(((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * 0.75M), 2));
					}
				}


				//25.
				//if (GetLookUpPercentile("Lookup.NoGlareLensPercentSpecLensRx", objInputData.colQ33b, "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col17d = 0;
				//else
				//    objOutputData.col17d = Math.Round(((((objInputData.colQ28 + objInputData.colQ29) * 0.75M) - ((objInputData.colQ28 + objInputData.colQ29) * (objInputData.colQ33b / 100))) * 80), 2);

				if (objInputData.colQ33b == null)
					objOutputData.col17d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.NoGlareLensPercentSpecLensRx", Convert.ToDecimal(objInputData.colQ33b), "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col17d = 0;
					else
					{
						if (objInputData.colQ28 == null && objInputData.colQ29 == null && objInputData.colQ33b == null)
							objOutputData.col17d = null;
						else
							objOutputData.col17d = Math.Round(((((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * 0.75M) - ((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * (Convert.ToDecimal(objInputData.colQ33b) / 100))) * 80), 2);
					}
				}


				//26.
				//if (GetLookUpPercentile("Lookup.HighIndexLensPercentSpecLensRx", objInputData.colQ33a, "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col18c = "Performance achieved";
				//else
				//    objOutputData.col18c = Convert.ToString(Math.Round(((objInputData.colQ28 + objInputData.colQ29) * 0.2M), 2));
				if (objInputData.colQ33a == null)
					objOutputData.col18c = null;
				else
				{
					if (GetLookUpPercentile("Lookup.HighIndexLensPercentSpecLensRx", Convert.ToDecimal(objInputData.colQ33a), "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col18c = "Performance achieved";
					else
					{
						if (objInputData.colQ28 == null && objInputData.colQ29 == null)
							objOutputData.col18c = null;
						else
							objOutputData.col18c = Convert.ToString(Math.Round(((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * 0.2M), 2));
					}
				}


				//27.
				//if (GetLookUpPercentile("Lookup.HighIndexLensPercentSpecLensRx", objInputData.colQ33a, "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col18d = 0;
				//else
				//    objOutputData.col18d = Math.Round(((((objInputData.colQ28 + objInputData.colQ29) * 0.2M) - ((objInputData.colQ28 + objInputData.colQ29) * (objInputData.colQ33a / 100))) * 45), 2);

				if (objInputData.colQ33a == null)
					objOutputData.col18d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.HighIndexLensPercentSpecLensRx", Convert.ToDecimal(objInputData.colQ33a), "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col18d = 0;
					else
					{
						if (objInputData.colQ28 == null && objInputData.colQ29 == null && objInputData.colQ33a == null)
							objOutputData.col18d = null;
						else
							objOutputData.col18d = Math.Round(((((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * 0.2M) - ((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * (Convert.ToDecimal(objInputData.colQ33a) / 100))) * 45), 2);
					}
				}

				//28.
				//if (GetLookUpPercentile("Lookup.PhotochrLensPercentofSpecLensRx", objInputData.colQ33c, "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col19c = "Performance achieved";
				//else
				//    objOutputData.col19c = Convert.ToString(Math.Round(((objInputData.colQ28 + objInputData.colQ29) * 0.27M), 2));

				if (objInputData.colQ33c == null)
					objOutputData.col19c = null;
				else
				{
					if (GetLookUpPercentile("Lookup.PhotochrLensPercentofSpecLensRx", Convert.ToDecimal(objInputData.colQ33c), "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col19c = "Performance achieved";
					else
					{
						if (objInputData.colQ28 == null && objInputData.colQ29 == null)
							objOutputData.col19c = null;
						else
							objOutputData.col19c = Convert.ToString(Math.Round(((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * 0.27M), 2));
					}

				}

				//29.
				//if (GetLookUpPercentile("Lookup.PhotochrLensPercentofSpecLensRx", objInputData.colQ33c, "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col19d = 0;
				//else
				//    objOutputData.col19d = Math.Round(((((objInputData.colQ28 + objInputData.colQ29) * 0.27M) - ((objInputData.colQ28 + objInputData.colQ29) * (objInputData.colQ33c / 100))) * 108), 2);
				if (objInputData.colQ33c == null)
					objOutputData.col19d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.PhotochrLensPercentofSpecLensRx", Convert.ToDecimal(objInputData.colQ33c), "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col19d = 0;
					else
					{
						if (objInputData.colQ28 == null && objInputData.colQ29 == null && objInputData.colQ33c == null)
							objOutputData.col19d = null;
						else
							objOutputData.col19d = Math.Round(((((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * 0.27M) - ((Convert.ToDecimal(objInputData.colQ28) + Convert.ToDecimal(objInputData.colQ29)) * (Convert.ToDecimal(objInputData.colQ33c) / 100))) * 108), 2);
					}
				}


				//30.
				//if (GetLookUpPercentile("Lookup.MultipleEyewearPurchasePercent", objInputData.colQ30, "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col20d = "Performance achieved";
				//else
				//    objOutputData.col20d = Convert.ToString(Math.Round((objOutputData.col20b * 0.15M), 2));
				if (objInputData.colQ30 == null)
					objOutputData.col20d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.MultipleEyewearPurchasePercent", Convert.ToDecimal(objInputData.colQ30), "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col20d = "Performance achieved";
					else
					{
						if (objOutputData.col20b == null)
							objOutputData.col20d = null;
						else
							objOutputData.col20d = Convert.ToString(Math.Round((Convert.ToDecimal(objOutputData.col20b) * 0.15M), 2));

					}
				}

				//31.
				//if (GetLookUpPercentile("Lookup.MultipleEyewearPurchasePercent", objInputData.colQ30, "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col20f = 0;
				//else
				//    if (objOutputData.col20d == "Performance achieved")  // To be discussed.
				//        objOutputData.col20f = 0;
				//    else
				//        objOutputData.col20f = Math.Round(((Convert.ToDecimal(objOutputData.col20d) - objOutputData.col20c) * objOutputData.col20e), 2);
				if (objInputData.colQ30 == null)
					objOutputData.col20f = null;
				else
				{
					if (GetLookUpPercentile("Lookup.MultipleEyewearPurchasePercent", Convert.ToDecimal(objInputData.colQ30), "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col20f = 0;
					else
					{
						if (objOutputData.col20d == "Performance achieved")
							objOutputData.col20f = 0;
						else
						{
							if (objOutputData.col20d == null && objOutputData.col20c == null && objOutputData.col20e == null)
								objOutputData.col20f = null;
							else
								objOutputData.col20f = Math.Round(((Convert.ToDecimal(objOutputData.col20d) - Convert.ToDecimal(objOutputData.col20c)) * Convert.ToDecimal(objOutputData.col20e)), 2);
						}
					}
				}


				//32.
				//if (GetLookUpPercentile("Lookup.PercentPatientsCLExamPurchEyewea", objInputData.colQ37, "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col21d = "Performance achieved";
				//else
				//    objOutputData.col21d = Convert.ToString(Math.Round((objInputData.colQ15b * 0.37M), 2));
				if (objInputData.colQ37 == null)
					objOutputData.col21d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.PercentPatientsCLExamPurchEyewea", Convert.ToDecimal(objInputData.colQ37), "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col21d = "Performance achieved";
					else
					{
						if (objInputData.colQ15b == null)
							objOutputData.col21d = null;
						else
							objOutputData.col21d = Convert.ToString(Math.Round((Convert.ToDecimal(objInputData.colQ15b) * 0.37M), 2));
					}
				}

				//33.
				//if (GetLookUpPercentile("Lookup.PercentPatientsCLExamPurchEyewea", objInputData.colQ37, "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col21e = "Performance achieved";
				//else
				//    objOutputData.col21e = Convert.ToString(Math.Round(((objInputData.colQ15b * 0.37M) * 150), 2));
				if (objInputData.colQ37 == null)
					objOutputData.col21e = null;
				else
				{
					if (GetLookUpPercentile("Lookup.PercentPatientsCLExamPurchEyewea", Convert.ToDecimal(objInputData.colQ37), "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col21e = "Performance achieved";
					else
					{
						if (objInputData.colQ15b == null)
							objOutputData.col21e = null;
						else
							objOutputData.col21e = Convert.ToString(Math.Round(((Convert.ToDecimal(objInputData.colQ15b) * 0.37M) * 150), 2));
					}
				}


				//34.
				//if (GetLookUpPercentile("Lookup.PercentPatientsCLExamPurchEyewea", objInputData.colQ37, "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
				//    objOutputData.col21f = 0;
				//else
				//    objOutputData.col21f = Math.Round((((objInputData.colQ15b * 0.37M) * 150) - ((objInputData.colQ15b * (objInputData.colQ37 / 100)) * 150)), 2);
				if (objInputData.colQ37 == null)
					objOutputData.col21f = null;
				else
				{
					if (GetLookUpPercentile("Lookup.PercentPatientsCLExamPurchEyewea", Convert.ToDecimal(objInputData.colQ37), "%") > 74) //Here no need to divide the Input by 100 ,because in database Lookup tables are saved as without percent.
						objOutputData.col21f = 0;
					else
					{
						if (objInputData.colQ15b == null && objInputData.colQ37 == null)
							objOutputData.col21f = null;
						else
							objOutputData.col21f = Math.Round((((Convert.ToDecimal(objInputData.colQ15b) * 0.37M) * 150) - ((Convert.ToDecimal(objInputData.colQ15b) * (Convert.ToDecimal(objInputData.colQ37) / 100)) * 150)), 2);

					}
				}


				//35.
				//if (GetLookUpPercentile("Lookup.CLGrossProfitMargin", objOutputData.col25b) > 74)
				//    objOutputData.col25d = "Performance achieved";
				//else
				//    objOutputData.col25d = Convert.ToString(Math.Round((objInputData.colQ26g * 0.54M), 2));
				if (objOutputData.col25b == null)
					objOutputData.col25d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.CLGrossProfitMargin", Convert.ToDecimal(objOutputData.col25b)) > 74)
						objOutputData.col25d = "Performance achieved";
					else
					{
						if (objInputData.colQ26g == null)
							objOutputData.col25d = null;
						else
							objOutputData.col25d = Convert.ToString(Math.Round((Convert.ToDecimal(objInputData.colQ26g) * 0.54M), 2));
					}
				}



				//36.
				//if (GetLookUpPercentile("Lookup.CLGrossProfitMargin", objOutputData.col25b) > 74)
				//    objOutputData.col25e = 0;
				//else
				//    objOutputData.col25e = Math.Round(((objInputData.colQ26g * 0.54M) - (objInputData.colQ26g - objInputData.colQ52f)), 2);
				if (objOutputData.col25b == null)
					objOutputData.col25e = null;
				else
				{
					if (GetLookUpPercentile("Lookup.CLGrossProfitMargin", Convert.ToDecimal(objOutputData.col25b)) > 74)
						objOutputData.col25e = 0;
					else
					{
						if (objInputData.colQ26g == null && objInputData.colQ52f == null)
							objOutputData.col25e = null;
						else
							objOutputData.col25e = Math.Round(((Convert.ToDecimal(objInputData.colQ26g) * 0.54M) - (Convert.ToDecimal(objInputData.colQ26g) - Convert.ToDecimal(objInputData.colQ52f))), 2);
					}
				}


				//37.
				//if (GetLookUpPercentile("Lookup.AnnCLSalesPerCLExam", objOutputData.col27a, "$") > 74)
				//    objOutputData.col27c = "Performance achieved";
				//else
				//    objOutputData.col27c = Convert.ToString(Math.Round((objInputData.colQ15b * 203), 2));
				if (objOutputData.col27a == null)
					objOutputData.col27c = null;
				else
				{
					if (GetLookUpPercentile("Lookup.AnnCLSalesPerCLExam", Convert.ToDecimal(objOutputData.col27a), "$") > 74)
						objOutputData.col27c = "Performance achieved";
					else
					{
						if (objInputData.colQ15b == null)
							objOutputData.col27c = null;
						else
							objOutputData.col27c = Convert.ToString(Math.Round((Convert.ToDecimal(objInputData.colQ15b) * 203), 2));
					}
				}


				//38.
				//if (GetLookUpPercentile("Lookup.AnnCLSalesPerCLExam", objOutputData.col27a, "$") > 74)
				//    objOutputData.col27d = 0;
				//else
				//    objOutputData.col27d = Math.Round((objInputData.colQ15b * 203) - objInputData.colQ26g, 2);

				if (objOutputData.col27a == null)
					objOutputData.col27d = null;
				else
				{
					if (GetLookUpPercentile("Lookup.AnnCLSalesPerCLExam", Convert.ToDecimal(objOutputData.col27a), "$") > 74)
						objOutputData.col27d = 0;
					else
					{
						if (objInputData.colQ15b == null && objInputData.colQ26g == null)
							objOutputData.col27d = null;
						else
							objOutputData.col27d = Math.Round((Convert.ToDecimal(objInputData.colQ15b) * 203) - Convert.ToDecimal(objInputData.colQ26g), 2);
					}
				}


				#endregion PercentileFormulas

				#region AverageFormulas
				//1.
				//objOutputData.colav122a = Math.Round(((GetLookUpPercentile("Lookup.PercentPatientsCLExamPurchEyewea", objInputData.colQ37, "%")
				//                            + GetLookUpPercentile("Lookup.MultipleEyewearPurchasePercent", objInputData.colQ30, "%")
				//                            + GetLookUpPercentile("Lookup.PhotochrLensPercentofSpecLensRx", objInputData.colQ33c, "%")
				//                            + GetLookUpPercentile("Lookup.HighIndexLensPercentSpecLensRx", objInputData.colQ33a, "%")
				//                            + GetLookUpPercentile("Lookup.NoGlareLensPercentSpecLensRx", objInputData.colQ33b, "%")
				//                            + GetLookUpPercentile("Lookup.ProgressiveLensAndPresbyopRx", objOutputData.col16c * 100, "%")
				//                            + GetLookUpPercentile("Lookup.EyewearGrossProfitMargin", objOutputData.col15c)
				//                            + GetLookUpPercentile("Lookup.GrossRevPerEyewearRx", objOutputData.col14a, "$")
				//                            + GetLookUpPercentile("Lookup.EyewearRxPer100ComplExam", objOutputData.col13b)) / 9),2);

				if (objInputData.colQ37 == null && objInputData.colQ30 == null && objInputData.colQ33c == null && objInputData.colQ33a == null &&
					objInputData.colQ33b == null && objOutputData.col16c == null && objOutputData.col15c == null && objOutputData.col14a == null && objOutputData.col13b == null)
				{
					objOutputData.colav122a = null;
				}
				else
				{
					objOutputData.colav122a = Math.Round(((GetLookUpPercentile("Lookup.PercentPatientsCLExamPurchEyewea", Convert.ToDecimal(objInputData.colQ37), "%")
											+ GetLookUpPercentile("Lookup.MultipleEyewearPurchasePercent", Convert.ToDecimal(objInputData.colQ30), "%")
											+ GetLookUpPercentile("Lookup.PhotochrLensPercentofSpecLensRx", Convert.ToDecimal(objInputData.colQ33c), "%")
											+ GetLookUpPercentile("Lookup.HighIndexLensPercentSpecLensRx", Convert.ToDecimal(objInputData.colQ33a), "%")
											+ GetLookUpPercentile("Lookup.NoGlareLensPercentSpecLensRx", Convert.ToDecimal(objInputData.colQ33b), "%")
											+ GetLookUpPercentile("Lookup.ProgressiveLensAndPresbyopRx", Convert.ToDecimal(objOutputData.col16c) * 100, "%")
											+ GetLookUpPercentile("Lookup.EyewearGrossProfitMargin", Convert.ToDecimal(objOutputData.col15c))
											+ GetLookUpPercentile("Lookup.GrossRevPerEyewearRx", Convert.ToDecimal(objOutputData.col14a), "$")
											+ GetLookUpPercentile("Lookup.EyewearRxPer100ComplExam", Convert.ToDecimal(objOutputData.col13b))) / 9), 2);

				}

				//2.
				//objOutputData.colav131a = Math.Round(((GetLookUpPercentile("Lookup.SoftMultiFocPercentSoftLens", objInputData.colQ40d, "%")
				//                        + GetLookUpPercentile("Lookup.SoftToricPercentSoftLens", objInputData.colQ40b, "%")
				//                        + GetLookUpPercentile("Lookup.MonthlySoftLensPercentWearers", objInputData.colQ39c, "%")
				//                        + GetLookUpPercentile("Lookup.DailyDisposableLensPercentSoft", objInputData.colQ39a, "%")
				//                        + GetLookUpPercentile("Lookup.SiliconeHydroLensWearPercentSoft", objInputData.colQ41a, "%")
				//                        + GetLookUpPercentile("Lookup.CLRefitPercentCLExam", objInputData.colQ43a, "%")
				//                        + GetLookUpPercentile("Lookup.CLNewFitsPer100CLExam", objOutputData.col28a)
				//                        + GetLookUpPercentile("Lookup.AnnCLSalesPerCLExam", objOutputData.col27a, "$")
				//                        + GetLookUpPercentile("Lookup.CLExamPercentTotalExam", objOutputData.col26b)
				//                        + GetLookUpPercentile("Lookup.CLWearerPercentActivePatients", objInputData.colQ13b, "%")
				//                        + GetLookUpPercentile("Lookup.CLGrossProfitMargin", objOutputData.col25b)
				//                        + GetLookUpPercentile("Lookup.CLSalesPercentGrossRev", objOutputData.col24a)) / 12),2);
				if (objInputData.colQ40d == null && objInputData.colQ40b == null && objInputData.colQ39c == null && objInputData.colQ39a == null &&
				   objInputData.colQ41a == null && objInputData.colQ43a == null && objOutputData.col28a == null && objOutputData.col27a == null &&
				   objOutputData.col26b == null && objInputData.colQ13b == null && objOutputData.col25b == null && objOutputData.col24a == null)
				{
					objOutputData.colav131a = null;
				}
				else
				{
					objOutputData.colav131a = Math.Round(((GetLookUpPercentile("Lookup.SoftMultiFocPercentSoftLens", Convert.ToDecimal(objInputData.colQ40d), "%")
											+ GetLookUpPercentile("Lookup.SoftToricPercentSoftLens", Convert.ToDecimal(objInputData.colQ40b), "%")
											+ GetLookUpPercentile("Lookup.MonthlySoftLensPercentWearers", Convert.ToDecimal(objInputData.colQ39c), "%")
											+ GetLookUpPercentile("Lookup.DailyDisposableLensPercentSoft", Convert.ToDecimal(objInputData.colQ39a), "%")
											+ GetLookUpPercentile("Lookup.SiliconeHydroLensWearPercentSoft", Convert.ToDecimal(objInputData.colQ41a), "%")
											+ GetLookUpPercentile("Lookup.CLRefitPercentCLExam", Convert.ToDecimal(objInputData.colQ43a), "%")
											+ GetLookUpPercentile("Lookup.CLNewFitsPer100CLExam", Convert.ToDecimal(objOutputData.col28a))
											+ GetLookUpPercentile("Lookup.AnnCLSalesPerCLExam", Convert.ToDecimal(objOutputData.col27a), "$")
											+ GetLookUpPercentile("Lookup.CLExamPercentTotalExam", Convert.ToDecimal(objOutputData.col26b))
											+ GetLookUpPercentile("Lookup.CLWearerPercentActivePatients", Convert.ToDecimal(objInputData.colQ13b), "%")
											+ GetLookUpPercentile("Lookup.CLGrossProfitMargin", Convert.ToDecimal(objOutputData.col25b))
											+ GetLookUpPercentile("Lookup.CLSalesPercentGrossRev", Convert.ToDecimal(objOutputData.col24a))) / 12), 2);
				}


				//objOutputData.colav134a = Math.Round(((GetLookUpPercentile("Lookup.NonRefrFeePercentGrossRev", objOutputData.col33b)
				//                            + GetLookUpPercentile("Lookup.MedicalEyeCareVisitPercentTotal", objOutputData.col33f)
				//                            + GetLookUpPercentile("Lookup.AnnMedEyeCareVisitPer1000", objOutputData.col34a)
				//                            + GetLookUpPercentile("Lookup.AnnPharmRxPer1000", objOutputData.col34d)) / 4),2);
				if (objOutputData.col33b == null && objOutputData.col33f == null && objOutputData.col34a == null && objOutputData.col34d == null)
				{
					objOutputData.colav134a = null;
				}
				else
				{
					objOutputData.colav134a = Math.Round(((GetLookUpPercentile("Lookup.NonRefrFeePercentGrossRev", Convert.ToDecimal(objOutputData.col33b))
												+ GetLookUpPercentile("Lookup.MedicalEyeCareVisitPercentTotal", Convert.ToDecimal(objOutputData.col33f))
												+ GetLookUpPercentile("Lookup.AnnMedEyeCareVisitPer1000", Convert.ToDecimal(objOutputData.col34a))
												+ GetLookUpPercentile("Lookup.AnnPharmRxPer1000", Convert.ToDecimal(objOutputData.col34d))) / 4), 2);

				}


				//4.
				//objOutputData.colav146a = Math.Round(((GetLookUpPercentile("Lookup.ExamFeeNonCL", objInputData.colQ47, "$")
				//                        + GetLookUpPercentile("Lookup.ExamFeeSoftNewFitSPHERE", objInputData.colQ48, "$")
				//                        + GetLookUpPercentile("Lookup.ExamFeeSoftNewFitTORIC", objInputData.colQ49, "$")
				//                        + GetLookUpPercentile("Lookup.ExamFeeSoftNewFitMULTIFO", objInputData.colQ50, "$")
				//                        + GetLookUpPercentile("Lookup.ExamFeeSoftLensNOREFITT", objInputData.colQ51, "$")) / 5),2);
				if (objInputData.colQ47 == null && objInputData.colQ48 == null && objInputData.colQ49 == null && objInputData.colQ50 == null && objInputData.colQ51 == null)
				{
					objOutputData.colav146a = null;
				}
				else
				{
					objOutputData.colav146a = Math.Round(((GetLookUpPercentile("Lookup.ExamFeeNonCL", Convert.ToDecimal(objInputData.colQ47), "$")
											+ GetLookUpPercentile("Lookup.ExamFeeSoftNewFitSPHERE", Convert.ToDecimal(objInputData.colQ48), "$")
											+ GetLookUpPercentile("Lookup.ExamFeeSoftNewFitTORIC", Convert.ToDecimal(objInputData.colQ49), "$")
											+ GetLookUpPercentile("Lookup.ExamFeeSoftNewFitMULTIFO", Convert.ToDecimal(objInputData.colQ50), "$")
											+ GetLookUpPercentile("Lookup.ExamFeeSoftLensNOREFITT", Convert.ToDecimal(objInputData.colQ51), "$")) / 5), 2);

				}


				#endregion AverageFormulas 

				#endregion CalculatingAllFormulas

				#region SaveFinalValuesToDatabase

				Target_OutputData finalOutputData = new Target_OutputData();

				objOutputData.SourceDataRefId = sourceRowId;
				finalOutputData.SourceDataRefId = sourceRowId;
				finalOutputData.C3a = objOutputData.col3a;
				finalOutputData.C3b = objOutputData.col3b;
				finalOutputData.C3c = objOutputData.col3c;
				finalOutputData.C3d = objOutputData.col3d;
				finalOutputData.C4a = objOutputData.col4a;
				finalOutputData.C4b = objOutputData.col4b;
				finalOutputData.C4c = objOutputData.col4c;
				finalOutputData.C4d = objOutputData.col4d;
				finalOutputData.C4e = objOutputData.col4e;
				finalOutputData.C5a = objOutputData.col5a;
				finalOutputData.C5b = objOutputData.col5b;
				finalOutputData.C5c = objOutputData.col5c;
				finalOutputData.C5d = objOutputData.col5d;
				finalOutputData.C6a = objOutputData.col6a;
				finalOutputData.C6b = objOutputData.col6b;
				finalOutputData.C6c = objOutputData.col6c;
				finalOutputData.C6d = objOutputData.col6d;
				finalOutputData.C6e = objOutputData.col6e;
				finalOutputData.C7a = objOutputData.col7a;
				finalOutputData.C7b = objOutputData.col7b;
				finalOutputData.C7c = objOutputData.col7c;
				finalOutputData.C7d = objOutputData.col7d;
				finalOutputData.C8a = objOutputData.col8a;
				finalOutputData.C8b = objOutputData.col8b;
				finalOutputData.C8c = objOutputData.col8c;
				finalOutputData.C8d = objOutputData.col8d;
				finalOutputData.C8e = objOutputData.col8e;
				finalOutputData.C9a = objOutputData.col9a;
				finalOutputData.C9b = objOutputData.col9b;
				finalOutputData.C9c = objOutputData.col9c;
				finalOutputData.C9d = objOutputData.col9d;
				finalOutputData.C12a = objOutputData.col12a;
				finalOutputData.C12b = objOutputData.col12b;
				finalOutputData.C13a = objOutputData.col13a;
				finalOutputData.C13b = objOutputData.col13b;
				finalOutputData.C13c = objOutputData.col13c;
				finalOutputData.C13d = objOutputData.col13d;
				finalOutputData.C13e = objOutputData.col13e;
				finalOutputData.C13f = objOutputData.col13f;
				finalOutputData.C13g = objOutputData.col13g;
				finalOutputData.C14a = objOutputData.col14a;
				finalOutputData.C14b = objOutputData.col14b;
				finalOutputData.C14c = objOutputData.col14c;
				finalOutputData.C14d = objOutputData.col14d;
				finalOutputData.C15a = objOutputData.col15a;
				finalOutputData.C15b = objOutputData.col15b;
				finalOutputData.C15c = objOutputData.col15c;
				finalOutputData.C15d = objOutputData.col15d;
				finalOutputData.C15e = objOutputData.col15e;
				finalOutputData.C15f = objOutputData.col15f;
				finalOutputData.C16a = objOutputData.col16a;
				finalOutputData.C16b = objOutputData.col16b;
				finalOutputData.C16c = objOutputData.col16c;
				finalOutputData.C16d = objOutputData.col16d;
				finalOutputData.C16e = objOutputData.col16e;
				finalOutputData.C16f = objOutputData.col16f;
				finalOutputData.C17a = objOutputData.col17a;
				finalOutputData.C17b = objOutputData.col17b;
				finalOutputData.C17c = objOutputData.col17c;
				finalOutputData.C17d = objOutputData.col17d;
				finalOutputData.C18a = objOutputData.col18a;
				finalOutputData.C18b = objOutputData.col18b;
				finalOutputData.C18c = objOutputData.col18c;
				finalOutputData.C18d = objOutputData.col18d;
				finalOutputData.C19a = objOutputData.col19a;
				finalOutputData.C19b = objOutputData.col19b;
				finalOutputData.C19c = objOutputData.col19c;
				finalOutputData.C19d = objOutputData.col19d;
				finalOutputData.C20a = objOutputData.col20a;
				finalOutputData.C20b = objOutputData.col20b;
				finalOutputData.C20c = objOutputData.col20c;
				finalOutputData.C20d = objOutputData.col20d;
				finalOutputData.C20e = objOutputData.col20e;
				finalOutputData.C20f = objOutputData.col20f;
				finalOutputData.C21a = objOutputData.col21a;
				finalOutputData.C21b = objOutputData.col21b;
				finalOutputData.C21c = objOutputData.col21c;
				finalOutputData.C21d = objOutputData.col21d;
				finalOutputData.C21e = objOutputData.col21e;
				finalOutputData.C21f = objOutputData.col21f;
				finalOutputData.C24a = objOutputData.col24a;
				finalOutputData.C24b = objOutputData.col24b;
				finalOutputData.C25a = objOutputData.col25a;
				finalOutputData.C25b = objOutputData.col25b;
				finalOutputData.C25c = objOutputData.col25c;
				finalOutputData.C25d = objOutputData.col25d;
				finalOutputData.C25e = objOutputData.col25e;
				finalOutputData.C26a = objOutputData.col26a;
				finalOutputData.C26b = objOutputData.col26b;
				finalOutputData.C26c = objOutputData.col26c;
				finalOutputData.C27a = objOutputData.col27a;
				finalOutputData.C27b = objOutputData.col27b;
				finalOutputData.C27c = objOutputData.col27c;
				finalOutputData.C27d = objOutputData.col27d;
				finalOutputData.C28a = objOutputData.col28a;
				finalOutputData.C28b = objOutputData.col28b;
				finalOutputData.C28c = objOutputData.col28c;
				finalOutputData.C29a = objOutputData.col29a;
				finalOutputData.C29b = objOutputData.col29b;
				finalOutputData.C29c = objOutputData.col29c;
				finalOutputData.C30a = objOutputData.col30a;
				finalOutputData.C30b = objOutputData.col30b;
				finalOutputData.C33a = objOutputData.col33a;
				finalOutputData.C33b = objOutputData.col33b;
				finalOutputData.C33c = objOutputData.col33c;
				finalOutputData.C33d = objOutputData.col33d;
				finalOutputData.C33e = objOutputData.col33e;
				finalOutputData.C33f = objOutputData.col33f;
				finalOutputData.C33g = objOutputData.col33g;
				finalOutputData.C34a = objOutputData.col34a;
				finalOutputData.C34b = objOutputData.col34b;
				finalOutputData.C34c = objOutputData.col34c;
				finalOutputData.C34d = objOutputData.col34d;
				finalOutputData.C34e = objOutputData.col34e;
				finalOutputData.C36a = objOutputData.col36a;
				finalOutputData.C36b = objOutputData.col36b;
				finalOutputData.C36c = objOutputData.col36c;
				finalOutputData.C36d = objOutputData.col36d;
				finalOutputData.C36e = objOutputData.col36e;
				finalOutputData.C37a = objOutputData.col37a;
				finalOutputData.C37b = objOutputData.col37b;
				finalOutputData.C37c = objOutputData.col37c;
				finalOutputData.C37d = objOutputData.col37d;
				finalOutputData.C37e = objOutputData.col37e;
				finalOutputData.C37f = objOutputData.col37f;
				finalOutputData.C37g = objOutputData.col37g;
				finalOutputData.C37h = objOutputData.col37h;
				finalOutputData.C40a = objOutputData.col40a;
				finalOutputData.C40b = objOutputData.col40b;
				finalOutputData.C40c = objOutputData.col40c;
				finalOutputData.C41a = objOutputData.col41a;
				finalOutputData.C41b = objOutputData.col41b;
				finalOutputData.C42a = objOutputData.col42a;
				finalOutputData.C42b = objOutputData.col42b;
				finalOutputData.C42c = objOutputData.col42c;
				finalOutputData.C43a = objOutputData.col43a;
				finalOutputData.C43b = objOutputData.col43b;
				finalOutputData.C43d = objOutputData.col43d;
				finalOutputData.C43e = objOutputData.col43e;
				finalOutputData.C43f = objOutputData.col43f;
				finalOutputData.C43g = objOutputData.col43g;
				finalOutputData.C44a = objOutputData.col44a;
				finalOutputData.C44b = objOutputData.col44b;
				finalOutputData.C44c = objOutputData.col44c;
				finalOutputData.C44d = objOutputData.col44d;
				finalOutputData.C45a = objOutputData.col45a;
				finalOutputData.C45b = objOutputData.col45b;
				finalOutputData.C45c = objOutputData.col45c;
				finalOutputData.C45d = objOutputData.col45d;
				finalOutputData.C45e = objOutputData.col45e;
				finalOutputData.C45f = objOutputData.col45f;
				finalOutputData.C45g = objOutputData.col45g;
				finalOutputData.C49a = objOutputData.col49a;
				finalOutputData.C49b = objOutputData.col49b;
				finalOutputData.C49c = objOutputData.col49c;
				finalOutputData.C49d = objOutputData.col49d;
				finalOutputData.C50a = objOutputData.col50a;
				finalOutputData.C50b = objOutputData.col50b;
				finalOutputData.C50c = objOutputData.col50c;
				finalOutputData.C50d = objOutputData.col50d;
				finalOutputData.av122a = objOutputData.colav122a;
				finalOutputData.av131a = objOutputData.colav131a;
				finalOutputData.av134a = objOutputData.colav134a;
				finalOutputData.av146a = objOutputData.colav146a;
				finalOutputData.n12a = objOutputData.coln12a;
				finalOutputData.n15c = objOutputData.coln15c;
				finalOutputData.n16c = objOutputData.coln16c;
				finalOutputData.n24a = objOutputData.coln24a;
				finalOutputData.n25b = objOutputData.coln25b;
				finalOutputData.n26b = objOutputData.coln26b;
				finalOutputData.n33b = objOutputData.coln33b;
				finalOutputData.n33f = objOutputData.coln33f;
				finalOutputData.n36a = objOutputData.coln36a;
				finalOutputData.n37c = objOutputData.coln37c;
				finalOutputData.n43b = objOutputData.coln43b;
				finalOutputData.n43f = objOutputData.coln43f;
				finalOutputData.n44a = objOutputData.coln44a;
				finalOutputData.n44c = objOutputData.coln44c;
				finalOutputData.n45c = objOutputData.coln45c;
				finalOutputData.n64a = objOutputData.coln64a;
				finalOutputData.n64b = objOutputData.coln64b;
				finalOutputData.n64c = objOutputData.coln64c;
				finalOutputData.n64d = objOutputData.coln64d;
				finalOutputData.n64e = objOutputData.coln64e;
				finalOutputData.n64f = objOutputData.coln64f;
				finalOutputData.n64g = objOutputData.coln64g;
				finalOutputData.n64h = objOutputData.coln64h;
				finalOutputData.n64i = objOutputData.coln64i;
				finalOutputData.n64j = objOutputData.coln64j;
				finalOutputData.n64k = objOutputData.coln64k;
				finalOutputData.n64l = objOutputData.coln64l;
				finalOutputData.n64m = objOutputData.coln64m;
				finalOutputData.n64n = objOutputData.coln64n;
				finalOutputData.n64o = objOutputData.coln64o;
				finalOutputData.n65a = objOutputData.coln65a;
				finalOutputData.n65b = objOutputData.coln65b;
				finalOutputData.n65c = objOutputData.coln65c;
				finalOutputData.n65d = objOutputData.coln65d;
				finalOutputData.n65e = objOutputData.coln65e;
				finalOutputData.n65f = objOutputData.coln65f;
				finalOutputData.n65g = objOutputData.coln65g;
				finalOutputData.n65h = objOutputData.coln65h;
				finalOutputData.n65i = objOutputData.coln65i;
				finalOutputData.n65j = objOutputData.coln65j;
				finalOutputData.n65k = objOutputData.coln65k;
				finalOutputData.n65l = objOutputData.coln65l;
				finalOutputData.n65m = objOutputData.coln65m;
				finalOutputData.n65n = objOutputData.coln65n;
				finalOutputData.n66a = objOutputData.coln66a;
				finalOutputData.n66b = objOutputData.coln66b;
				finalOutputData.n66c = objOutputData.coln66c;
				finalOutputData.n66d = objOutputData.coln66d;
				finalOutputData.n66e = objOutputData.coln66e;
				finalOutputData.n66f = objOutputData.coln66f;
				finalOutputData.n66g = objOutputData.coln66g;
				finalOutputData.n66h = objOutputData.coln66h;
				finalOutputData.n66i = objOutputData.coln66i;
				finalOutputData.n66j = objOutputData.coln66j;
				finalOutputData.n66k = objOutputData.coln66k;
				finalOutputData.n66l = objOutputData.coln66l;
				finalOutputData.n66m = objOutputData.coln66m;
				finalOutputData.n66n = objOutputData.coln66n;
				finalOutputData.n66o = objOutputData.coln66o;
				finalOutputData.n66p = objOutputData.coln66p;
				finalOutputData.n66q = objOutputData.coln66q;
				finalOutputData.n66r = objOutputData.coln66r;
				finalOutputData.n66s = objOutputData.coln66s;
				finalOutputData.n66t = objOutputData.coln66t;
				finalOutputData.n66u = objOutputData.coln66u;
				finalOutputData.n66v = objOutputData.coln66v;
				finalOutputData.n68 = objOutputData.coln68;
				finalOutputData.t20 = objOutputData.colt20;
				finalOutputData.t21 = objOutputData.colt21;
				finalOutputData.t27 = objOutputData.colt27;
				//praveenk-Release2
				finalOutputData.AdditionalInfo = objInputData.AdditionalInfo;

				//lstOutput.Add(objOutputData);
				objReport.lstOutput.Add(objOutputData);
				objReport.lstInput.Add(objInputData);
				db.AddToTarget_OutputData(finalOutputData);
				// }

				db.SaveChanges();
				objReport.lstInput = objReport.lstInput.OrderBy(r => r.RowId).ToList();


				#endregion SaveFinalValuesToDatabase

				//var milliseconds = stopwatch.ElapsedMilliseconds;

			}
			catch (Exception ex)
			{
				throw ex;
			}
			return objReport;

		}

		public string GetLookUpLable(string tblName, decimal lookUpValue, string valueType = "")
		{
			string lookUpLabel = "";
			try
			{
				List<string> relatedLookUpLables = db.ExecuteStoreQuery<string>("select LookupLable from [" + tblName + "] where [LookupValue" + valueType + "]<=" + lookUpValue).ToList();
				if (relatedLookUpLables.Count() > 0)
					lookUpLabel = relatedLookUpLables.LastOrDefault();
				else
					lookUpLabel = db.ExecuteStoreQuery<string>("select LookupLable from [" + tblName + "]").FirstOrDefault();
			}
			catch (Exception ex)
			{
				return ex.Message;
			}
			return lookUpLabel;
		}

		public decimal GetLookUpPercentile(string tblName, decimal lookUpValue, string valueType = "")
		{
			decimal lookUpPercentile = 0;
			try
			{
				List<decimal> relatedLookUpLables = db.ExecuteStoreQuery<decimal>("select Percentile from [" + tblName + "] where [LookupValue" + valueType + "]<=" + lookUpValue).ToList();
				if (relatedLookUpLables.Count() > 0)
					lookUpPercentile = relatedLookUpLables.LastOrDefault();
				else
					lookUpPercentile = db.ExecuteStoreQuery<decimal>("select Percentile from [" + tblName + "]").FirstOrDefault();
			}
			catch (Exception ex)
			{
				throw ex;
			}
			return lookUpPercentile;

		}

		public decimal GetLookUpValue(string tblName, int lookUpPercentile, string colNameFuffix = null)
		{
			decimal lookUpValue = 0;
			try
			{
				List<decimal> relatedLookUpValues = db.ExecuteStoreQuery<decimal>("SELECT TOP 1 COALESCE(NULLIF([LookupValue" + colNameFuffix + "], 0), [ResponseMin" + colNameFuffix + "]) from [" + tblName + "] where [Percentile] =" + lookUpPercentile).ToList();
				if (relatedLookUpValues.Count() > 0)
					lookUpValue = relatedLookUpValues.FirstOrDefault();
			}
			catch (Exception ex)
			{
				throw ex;
			}
			return lookUpValue;

		}

		public List<decimal> GetQuintileLookUpValues(string tblName, string colNameFuffix = null)
		{
			List<decimal> relatedLookUpValues = new List<decimal>();
			try
			{
				relatedLookUpValues = db.ExecuteStoreQuery<decimal>("SELECT [LookupValue" + colNameFuffix + "] from [" + tblName + "] where [Percentile] IN (11,31,51,71,91)").ToList();
			}
			catch (Exception ex)
			{
				throw ex;
			}
			return relatedLookUpValues;
		}

		public List<decimal> GetAllLookUpValues(string tblName, string colNameFuffix = null)
		{
			List<decimal> relatedLookUpValues = new List<decimal>();
			try
			{
				relatedLookUpValues = db.ExecuteStoreQuery<decimal>("SELECT COALESCE(NULLIF([LookupValue" + colNameFuffix + "], 0), [ResponseMin" + colNameFuffix + "]) from [" + tblName + "]").ToList();
			}
			catch (Exception ex)
			{
				throw ex;
			}
			return relatedLookUpValues;
		}

		public List<decimal> GetMinAndMaxLookUpValue(string tblName, int lookUpPercentile, string colNameFuffix = null)
		{
			List<decimal> minMaxResponseValues = new List<decimal>();
			try
			{
				System.Data.DataTable dt = new System.Data.DataTable();
				DataSet ds = new DataSet();
				string connStr = ConfigurationSettings.AppSettings["myConnectionString"];

				SqlConnection con = new SqlConnection(connStr);

				con.Open();

				var cmd = new SqlCommand();
				cmd.Connection = con;
				string tableName = "[PPASurvey_DBProd].[dbo].[" + tblName + "]";
				String strQuery = "select  top 1 CAST ([ResponseMin" + colNameFuffix + "] as decimal(18,2)), CAST ([ResponseMax" + colNameFuffix + "] as decimal(18,2)) from " + tableName + "where [Percentile] =" + lookUpPercentile;
				cmd.CommandText = strQuery;
				cmd.CommandType = CommandType.Text;

				SqlDataAdapter adp = new SqlDataAdapter(cmd);

				adp.Fill(ds);
				dt = ds.Tables[0];
				if (dt != null)
				{
					minMaxResponseValues.Add(Convert.ToDecimal(dt.Rows[0][0])); // min value
					minMaxResponseValues.Add(Convert.ToDecimal(dt.Rows[0][1])); // max value

				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
			return minMaxResponseValues;

		}

		public Dictionary<string, string> GetExternalTableValues(string esTable)
		{
			Dictionary<string, string> externalTblData = new Dictionary<string, string>();
			try
			{
				System.Data.DataTable dt = new System.Data.DataTable();
				DataSet ds = new DataSet();
				string connStr = ConfigurationSettings.AppSettings["myConnectionString"];

				SqlConnection con = new SqlConnection(connStr);

				con.Open();

				var cmd = new SqlCommand();
				cmd.Connection = con;
				string tableName = "[PPASurvey_DBProd].[dbo].[" + esTable + "]";
				String strQuery = "select  * from " + tableName;
				cmd.CommandText = strQuery;
				cmd.CommandType = CommandType.Text;

				SqlDataAdapter adp = new SqlDataAdapter(cmd);

				adp.Fill(ds);
				dt = ds.Tables[0];
				if (dt != null)
				{
					for (int i = 0; i < dt.Columns.Count; i++)
						externalTblData.Add(dt.Columns[i].ColumnName, Convert.ToString(dt.Rows[0][i]));

				}
			}
			catch (Exception ex)
			{
				throw ex;
			}
			return externalTblData;

		}

		private static bool IsFileInUse(string path)
		{
			FileStream stream = null;
			try
			{
				stream = new FileStream(path, FileMode.OpenOrCreate);
			}
			catch (IOException)
			{
				//the file is unavailable because it is:
				//still being written to
				//or being processed by another thread
				//or does not exist (has already been processed)
				return true;
			}
			finally
			{
				if (stream != null)
					stream.Close();
			}

			//file is not locked
			return false;
		}

		public string ReadCSVAndInsertToSQL(string strFilePath, string fileName, string strWordSourcePath)
		{

			try
			{
				if (File.Exists(strFilePath + @"\" + fileName) == false)
					return "Input File is missing, please check and retry.";

				if (IsFileInUse(strFilePath + @"\" + fileName))
				{
					return "File " + fileName + " already in use please close it first.";
				}

				//string strdirepath = @"D:\Shahbaz\Practice Performance Assessment\";
				string connectionString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Extended Properties=""text;HDR=YES;FMT=Delimited""", strFilePath);
				OleDbConnection myConnection = new OleDbConnection(connectionString);
				myConnection.Open();
				string connStr = ConfigurationSettings.AppSettings["myConnectionString"];

				SqlConnection con = new SqlConnection(connStr);
				var cmd = new SqlCommand();
				cmd.Connection = con;
				//cmd.CommandText = "SynchNewYearSurveyData";
				cmd.CommandType = CommandType.Text;

				con.Open();

				string StrUpdateQuery = "Update [Source.InputDataSave] set Active =0 Where Q38=" + "'" + System.Web.HttpContext.Current.Session["namesave"].ToString() + "'" + " and Q47=" + "'" + System.Web.HttpContext.Current.Session["practicenamesave"].ToString() + "'" + " and IDname=" + "'" + System.Web.HttpContext.Current.Session["practiceid"].ToString() + "'" + "and Year=" + "'" + System.Web.HttpContext.Current.Session["YearName"].ToString() + "'";
				// always read  from the sheet1.
				OleDbCommand oledbCommand = new OleDbCommand("SELECT  * FROM [" + fileName + "]");
				oledbCommand.Connection = myConnection;
				OleDbDataReader myReader = oledbCommand.ExecuteReader();
				int CountOfRows_Y = 0;
				while (myReader.Read())
				{
					StringBuilder Sb = new StringBuilder();
					bool IsFirstFiveColhasvalue = true;
					bool IsValidRow = true;
					for (int j = 0; j < 6; j++)  //here we are checking is any column is blank from B to F
					{
						if (j == 0)
						{
							if (myReader[j].ToString().ToUpper() != "Y")
							{
								IsValidRow = false;
							}
						}
						if (myReader[j].ToString() == "")
						{
							IsFirstFiveColhasvalue = false;
						}
					}
					if (IsFirstFiveColhasvalue == true && IsValidRow == true)
					{
						CountOfRows_Y++;
						// it can read upto 193 columns means A to GK. 
						//praveenk-Release2
						//for (int i = 0; i < 193; i++)
						for (int i = 0; i < 229; i++) //new column added.Hence it need to read from A to GL - Updated to read until HU (msinghai).
						{
							//MessageBox.Show(myReader[i].ToString() + " ");

							if (i == 1 || i == 4) //here we are taking Date only
							{
								int fromIndex = myReader[i].ToString().IndexOf(" ");
								Sb.Append("'" + myReader[i].ToString().Remove(fromIndex) + "'" + ",");
							}
							else if (i == 2 || i == 3) //here we are taking Time Only
							{
								int fromIndex = myReader[i].ToString().IndexOf(" ");
								Sb.Append("'" + myReader[i].ToString().Substring(fromIndex) + "'" + ",");
							}
							//else if (i > 185 || i == 0 || i == 5)
							//praveenk-Release2
							else if ((i > 185 && i < 193) || i == 0 || i == 5 || i == 229)
							{
								Sb.Append("'" + myReader[i].ToString() + "'" + ",");
							}
							else if ((i > 127 && i < 179) || (i > 220 && i < 229)) //in this range question contains of type Yes/No.
							{
								if (myReader[i].ToString() == "2") //if i value=2 means user selected Answer is'No' 
									Sb.Append(0.ToString() + ",");
								else if (myReader[i].ToString() == "") //if i value="" means question not attempted than we are putting NULL in database
									Sb.Append("NULL" + ",");
								else
									Sb.Append(myReader[i].ToString() + ",");//else user selected Answer is 'Yes' 
							}
							else if (i == 193)
							{
								string tempStr = myReader[i].ToString() == "" ? "NULL" : myReader[i].ToString();
								Sb.Append("'" + DateTime.Now + "'," + tempStr + ",");
							}
							else
							{
								if (myReader[i].ToString() == "") //if i value="" means question not answered than we are putting NULL in database
									Sb.Append("NULL" + ",");
								else
									Sb.Append(myReader[i].ToString() + ",");
							}

						}
						if (CountOfRows_Y <= NoOfDocumentsLimit)
						{
							Sb.Remove(Sb.Length - 1, 1);
							string StrInsertQuery = "Insert into [Source.InputData] values(" + Sb + ")";
							db.ExecuteStoreCommand(StrInsertQuery);
							db.SaveChanges();

							//Also save in Input Data Bench Mark Source - msinghai 12/18
							Sb.Insert(0, System.Web.HttpContext.Current.Session["YearValue"].ToString() + ",");
							StrInsertQuery = "Insert into [Source.InputDataBenchMarkSource] values(" + Sb + ")";
							db.ExecuteStoreCommand(StrInsertQuery);
							db.SaveChanges();

							//Populate '_J' Lookup tables
							string year = System.Web.HttpContext.Current.Session["YearName"].ToString();
							char type = 'F';
							PopulateBenchMarks(null, year, type);
							//ObjectParameter[] PracticeIdParameter = new ObjectParameter[3];
							//ObjectParameter lookupTableName = new ObjectParameter("LookupName", DBNull.Value); //to update all tables, pass null
							//PracticeIdParameter[0] = lookupTableName;
							//PracticeIdParameter[1] = new ObjectParameter("year", year);
							//PracticeIdParameter[2] = new ObjectParameter("type", 'F');

							//db.ExecuteFunction("SP_PopulateAllBenchmarks", PracticeIdParameter);
							db.SaveChanges();


						}
						else //this else condition for exiting from while loop if CountOfRows_Y is reach to its limit(i.e. NoOfDocumentsLimit).
						{


							return "success";
						}
					}

				}

				myConnection.Close();
				if (CountOfRows_Y == 0)
					return "In Input File, no record exists with Generate Report? = Y. Please modify and retry.";
				else

					cmd.CommandText = StrUpdateQuery;
				cmd.ExecuteNonQuery();
				con.Close();
				return "success";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}

		}

		public string ReadCSVAndInsertToSQLForSave(System.Data.DataTable dt)
		{

			try
			{

				int CountOfRows_Y = 0;

				string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
				SqlConnection con = new SqlConnection(connStr);
				var cmd = new SqlCommand();
				cmd.Connection = con;
				//cmd.CommandText = "SynchNewYearSurveyData";
				cmd.CommandType = CommandType.Text;

				con.Open();
				//building stril builder
				// int count = dt.Columns.Count;
				for (int j1 = 0; j1 < dt.Rows.Count; j1++)
				{
					StringBuilder Sb = new StringBuilder();
					bool IsFirstFiveColhasvalue = true;
					bool IsValidRow = true;

					for (int j = 0; j < 6; j++)  //here we are checking is any column is blank from B to F
					{
						if (j == 0)
						{
							if (dt.Rows[0][j].ToString().ToUpper() != "Y")
							{
								IsValidRow = false;
							}
						}
						if (dt.Rows[0][j].ToString() == "")
						{
							IsFirstFiveColhasvalue = false;
						}
					}


					if (IsFirstFiveColhasvalue == true && IsValidRow == true)
					{
						CountOfRows_Y++;
						// it can read upto 193 columns means A to GK. 
						//praveenk-Release2
						//for (int i = 0; i < 193; i++)
						for (int i = 0; i < 194; i++) //new column added.Hence it need to read from A to GL.
						{
							//MessageBox.Show(myReader[i].ToString() + " ");

							if (i == 1 || i == 4) //here we are taking Date only
							{
								int fromIndex = dt.Rows[0][i].ToString().IndexOf(" ");
								Sb.Append("'" + dt.Rows[0][i].ToString().Remove(fromIndex) + "'" + ",");
							}
							else if (i == 2 || i == 3) //here we are taking Time Only
							{
								int fromIndex = dt.Rows[0][i].ToString().IndexOf(" ");
								Sb.Append("'" + dt.Rows[0][i].ToString().Substring(fromIndex) + "'" + ",");
							}
							//else if (i > 185 || i == 0 || i == 5)
							//praveenk-Release2
							else if (i > 185 || i == 0 || i == 5 || i == 193)
							{
								Sb.Append("'" + dt.Rows[0][i].ToString() + "'" + ",");
							}
							else if (i > 127 && i < 179) //in this range question contains of type Yes/No.
							{
								if (dt.Rows[0][i].ToString() == "2") //if i value=2 means user selected Answer is'No' 
									Sb.Append(0.ToString() + ",");
								else if (dt.Rows[0][i].ToString() == "") //if i value="" means question not attempted than we are putting NULL in database
									Sb.Append("NULL" + ",");
								else
									Sb.Append(dt.Rows[0][i].ToString() + ",");//else user selected Answer is 'Yes' 
							}
							else
							{
								if (dt.Rows[0][i].ToString() == "") //if i value="" means question not answered than we are putting NULL in database
									Sb.Append("NULL" + ",");
								else
									Sb.Append(dt.Rows[0][i].ToString() + ",");
							}

						}
						if (CountOfRows_Y <= NoOfDocumentsLimit)
						{
							string year = System.Web.HttpContext.Current.Session["YearName"].ToString();
							Sb.Append(year + ",");
							Sb.Append("1" + ",");
							Sb.Remove(Sb.Length - 1, 1);


							// string StrDeleteQuery = "Delete from [Source.InputDataSave] Where Q38=" + "'" + dt.Rows[0]["Q38"] + "'" + " and Q47=" + "'" + dt.Rows[0]["Q47"] + "'" + " and IDname=" + "'" + dt.Rows[0]["ID.name"] + "'";

							string StrDeleteQuery = "Delete from [Source.InputDataSave] Where Q38=" + "'" + dt.Rows[0]["Q38"] + "'" + " and Q47=" + "'" + dt.Rows[0]["Q47"] + "'" + " and IDname=" + "'" + dt.Rows[0]["ID.name"] + "'" + "and Year=" + "'" + year + "'" + " and Active=1";
							cmd.CommandText = StrDeleteQuery;
							cmd.ExecuteNonQuery();

							//string StrDeleteQuery = "Delete from [Source.InputDataSave] Where Q38=" + "'" + dt.Rows[0]["Q38"] + "'" + " and Q47=" + "'" + dt.Rows[0]["Q47"] + "'" + " and IDname=" + "'" + dt.Rows[0]["ID.name"] + "'" + "and Year=" + "'" + year + "'";

							string StrInsertQuery = "Insert into [Source.InputDataSave] values(" + Sb + ")";
							cmd.CommandText = StrInsertQuery;
							cmd.ExecuteNonQuery();




							//db.SaveChanges();
						}
						else //this else condition for exiting from while loop if CountOfRows_Y is reach to its limit(i.e. NoOfDocumentsLimit).
						{
							return "success";
						}
					}

				}


				//building complete


				con.Close();
				if (CountOfRows_Y == 0)
					return "In Input File, no record exists with Generate Report? = Y. Please modify and retry.";
				else
					return "success";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}

		}

		public string InsertSurveyTranscation(int SurveyId, string YearId, int PracticeId, string UserName, DateTime EntryDate, string DetailedPath, string InfographicPath, string ExecutivePath, string CSVPath)
		{
			try
			{
				ObjectParameter[] PracticeIdParameter = new ObjectParameter[9];
				ObjectParameter PracticeIdParameter1 = new ObjectParameter("SurveyId", SurveyId);
				PracticeIdParameter[0] = PracticeIdParameter1;
				PracticeIdParameter1 = new ObjectParameter("YearId", YearId);
				PracticeIdParameter[1] = PracticeIdParameter1;
				PracticeIdParameter1 = new ObjectParameter("PracticeId", PracticeId);
				PracticeIdParameter[2] = PracticeIdParameter1;
				PracticeIdParameter1 = new ObjectParameter("UserName", UserName);
				PracticeIdParameter[3] = PracticeIdParameter1;
				PracticeIdParameter1 = new ObjectParameter("EntryDate", EntryDate);
				PracticeIdParameter[4] = PracticeIdParameter1;
				PracticeIdParameter1 = new ObjectParameter("DetailedPath", DetailedPath);
				PracticeIdParameter[5] = PracticeIdParameter1;
				PracticeIdParameter1 = new ObjectParameter("InfographicPath", InfographicPath);
				PracticeIdParameter[6] = PracticeIdParameter1;
				PracticeIdParameter1 = new ObjectParameter("ExecutivePath", ExecutivePath);
				PracticeIdParameter[7] = PracticeIdParameter1;
				PracticeIdParameter1 = new ObjectParameter("CSVPath", CSVPath);
				PracticeIdParameter[8] = PracticeIdParameter1;


				var result = db.ExecuteFunction<InsertTranscation_Result>("InsertTranscation", MergeOption.OverwriteChanges, PracticeIdParameter).FirstOrDefault();

				return "success";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}


		}

		private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceText)
		{
			object matchCase = true;
			object matchWholeWord = true;
			object matchWildCards = false;
			object matchSoundsLike = false;
			object matchAllWordForms = false;
			object forward = true;
			object format = false;
			object matchKashida = false;
			object matchDiacritics = false;
			object matchAlefHamza = false;
			object matchControl = false;
			object read_only = true;
			object visible = true;
			object replace = 2;
			object wrap = 1;
			try
			{
				wordApp.Selection.Find.Execute(ref findText, ref matchCase,
					ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
					ref matchAllWordForms, ref forward, ref wrap, ref format,
					ref replaceText, ref replace, ref matchKashida,
							ref matchDiacritics,
					ref matchAlefHamza, ref matchControl);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public string GenerateWordReport(string year1, string strSourcePath, string strTargetPath, Report objReport, string practiceid, DataSet ds, System.Data.DataTable dtsort)
		{
			try
			{

				if (IsFileInUse(strSourcePath))
				{
					return "File already in use, please close and retry.";
				}
				//Only run the reports that are not ran earlier.
				// List<int> reportToBeGenerated = db.Source_InputData.Where(r => r.RowId > lastReportGenerated && r.IDformat == "Y").Select(r => r.RowId).ToList();

				List<int> reportToBeGenerated = db.Source_InputData.Where(r => r.IDname == practiceid && r.IDformat == "Y").Select(r => r.RowId).ToList();

				if (reportToBeGenerated.Count() < NoOfDocumentsLimit)
				{
					NoOfDocumentsLimit = 1; // reportToBeGenerated.Count();
				}

				for (int i = 0; i < NoOfDocumentsLimit; i++)
				{

					DateTime CurrentDateTime = DateTime.Now;
					string filenamestr = objReport.lstInput[i].colQ73 + "_" + objReport.lstInput[i].colIDname + "_" + CurrentDateTime.ToString("MMddyyyy-hhmmss");

					System.Web.HttpContext.Current.Session["varName"] = filenamestr;

					/*Muntajib-Remove tempPath For both word & pdf*/
					string tempPath = AppDomain.CurrentDomain.BaseDirectory + filenamestr + ".doc";
					// string tempPath = @"D:\generate" + "\\" + filenamestr + ".doc";
					filenamestr = strTargetPath + "\\" + filenamestr + ".doc";


					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					FileStream fs = new FileStream(tempPath, FileMode.Create, FileAccess.ReadWrite);
					fs.Close();
					//  Just to kill WINWORD.EXE if it is running
					//  killprocess("winword");     
					//  copy letter format to temp.doc

					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					File.Copy(strSourcePath, tempPath, true);

					//  create missing object
					object missing = Missing.Value;
					//  create Word application object
					Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();


					//  create Word document object
					Microsoft.Office.Interop.Word.Document aDoc = null;
					//  create & define filename object with temp.doc
					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					//System.Threading.Thread.Sleep(5000);
					object filename = tempPath;


					if (File.Exists((string)filename))
					{
						/*//--Shahbaz.Need to enable when we go with Word to Pdf Report.*/
						object SaveChanges = false;

						object readOnly = false;
						object isVisible = false;
						//  make visible Word application
						wordApp.Visible = false;
						//  open Word document named temp.doc
						aDoc = wordApp.Documents.Open(ref filename, ref missing,
					   ref readOnly, ref missing, ref missing, ref missing,
					   ref missing, ref missing, ref missing, ref missing,
					   ref missing, ref isVisible, ref missing, ref missing,
					   ref missing, ref missing);
						// System.Threading.Thread.Sleep(2000);

						aDoc.Activate();
						// System.Threading.Thread.Sleep(2000);


						/*//--Shahbaz.Need to enable this code,when we go with Word to Pdf Report.
                        //--Shahbaz.Before enabling this we need to install the "SaveAsPDFandXPS.exe" from the network drive.
                        //--Shahbaz.If we want to put it as web app,we just need to install "SaveAsPDFandXPS.exe" on the server.*/

						object outputFileName = filenamestr.Replace(".doc", ".pdf");
						object fileFormat = WdSaveFormat.wdFormatPDF;

						//  Call FindAndReplace()function for each change

						#region Replace Word Documnet Tempalte's content.

						//To replace the Page Header.
						foreach (Section aSection in wordApp.ActiveDocument.Sections)
						{
							//It contains multiple headers in blank template.
							foreach (HeaderFooter aHeader in aSection.Headers)
							{
								//Only Replace the header contains "Prepared exclusively for «Q73», «Q74»" in it.
								if (aHeader.Range.Text == "Prepared exclusively for «Q73», «Q74»\r")
								{
									aHeader.Range.Text = "Prepared exclusively for " + objReport.lstInput[i].colQ73 + ", " + objReport.lstInput[i].colQ74;
								}

								if (aHeader.Range.Text == "Prepared exclusively for «Q74»\r")
								{
									aHeader.Range.Text = "Prepared exclusively for " + objReport.lstInput[i].colQ74;
								}
							}
						}

						this.FindAndReplace(wordApp, "«Q73»", objReport.lstInput[i].colQ73);
						this.FindAndReplace(wordApp, "«Q74»", objReport.lstInput[i].colQ74);
						// objReport.lstInput[i].
						//praveenk-Rlease2
						this.FindAndReplace(wordApp, "«AddInfo»", objReport.lstInput[i].AdditionalInfo);

						//5.Gross Revenue per Complete Exam					
						//this.FindAndReplace(wordApp, "«Q24»", "$" + objReport.lstInput[i].colQ24.ToString("#,0.##"));//show blank if they have not answered that question.
						this.FindAndReplace(wordApp, "«Q24»", objReport.lstInput[i].colQ24 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ24)).ToString("#,0"));

						//this.FindAndReplace(wordApp, "«Q14»", objReport.lstInput[i].colQ14.ToString("#,0.##")); //show blank if they have not answered that question.
						this.FindAndReplace(wordApp, "«Q14»", objReport.lstInput[i].colQ14 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ14)).ToString("#,0"));
						this.FindAndReplace(wordApp, "«M_3a»", objReport.lstOutput[i].col3a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col3a)).ToString("#,0"));
						this.FindAndReplace(wordApp, "«M_3b»", objReport.lstOutput[i].col3b == null ? "" : objReport.lstOutput[i].col3b);

						//this.FindAndReplace(wordApp, "«Q14»", objReport.lstInput[i].colQ14.ToString("#,0.##")); //Repeated.
						if (objReport.lstOutput[i].col3c == null)
							this.FindAndReplace(wordApp, "«M_3c»", "");
						else
							this.FindAndReplace(wordApp, "«M_3c»", (objReport.lstOutput[i].col3c == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col3c)).ToString("#,0"));


						this.FindAndReplace(wordApp, "«M_3d»", objReport.lstOutput[i].col3d == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col3d)).ToString("#,0"));

						//6.Complete Exams per OD Hour					
						this.FindAndReplace(wordApp, "«Q11»", objReport.lstInput[i].colQ11 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ11)).ToString("#,0"));
						//this.FindAndReplace(wordApp, "«Q14»", objReport.lstInput[i].colQ14.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_4a»", Convert.ToDecimal(objReport.lstOutput[i].col4a).ToString());
						this.FindAndReplace(wordApp, "«M_4b»", objReport.lstOutput[i].col4b == null ? "" : objReport.lstOutput[i].col4b.ToString());
						//this.FindAndReplace(wordApp, "«Q11»", objReport.lstInput[i].colQ11.ToString("#,0.##")); //Repeated.
						if (objReport.lstOutput[i].col4c == null)
							this.FindAndReplace(wordApp, "«M_4c»", "");
						else
							this.FindAndReplace(wordApp, "«M_4c»", (objReport.lstOutput[i].col4c == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col4c)).ToString("#,0"));

						//this.FindAndReplace(wordApp,	"«M_3a»"	,	"$"+objReport.lstOutput[i].col3a);//Repeated.
						if (objReport.lstOutput[i].col4d == null)
							this.FindAndReplace(wordApp, "«M_4d»", "");
						else
							this.FindAndReplace(wordApp, "«M_4d»", (objReport.lstOutput[i].col4d == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col4d)).ToString("#,0"));

						this.FindAndReplace(wordApp, "«M_4e»", objReport.lstOutput[i].col4e == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col4e)).ToString("#,0"));

						//7.Annual Gross Revenue per Active Patient					
						//this.FindAndReplace(wordApp, "«Q24»", "$"+objReport.lstInput[i].colQ24.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q12»", objReport.lstInput[i].colQ12 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ12)).ToString("#,0"));
						this.FindAndReplace(wordApp, "«M_5a»", objReport.lstOutput[i].col5a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col5a)).ToString("#,0"));
						this.FindAndReplace(wordApp, "«M_5b»", objReport.lstOutput[i].col5b);
						//this.FindAndReplace(wordApp,	"«Q12»"	,	objReport.lstInput[i].colQ12.ToString("#,0.##")); //Repeated.
						if (objReport.lstOutput[i].col5c == null)
							this.FindAndReplace(wordApp, "«M_5c»", "");
						else
							this.FindAndReplace(wordApp, "«M_5c»", (objReport.lstOutput[i].col5c == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col5c)).ToString("#,0"));

						this.FindAndReplace(wordApp, "«M_5d»", objReport.lstOutput[i].col5d == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col5d)).ToString("#,0"));

						//8.Annual Complete Exams per 100 Active Patients					
						//this.FindAndReplace(wordApp, "«Q14»", objReport.lstInput[i].colQ14.ToString("#,0.##")); //Repeated.
						//this.FindAndReplace(wordApp, "«Q12»", objReport.lstInput[i].colQ12.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_6a»", Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col6a)).ToString("#,0"));
						this.FindAndReplace(wordApp, "«M_6b»", objReport.lstOutput[i].col6b);
						//this.FindAndReplace(wordApp, "«Q12»", objReport.lstInput[i].colQ12.ToString("#,0.##"));//Repeated.
						//this.FindAndReplace(wordApp, "«Q24»", "$"+objReport.lstInput[i].colQ24.ToString("#,0.##")); //Repeated.
						if (objReport.lstOutput[i].col6c == null)
							this.FindAndReplace(wordApp, "«M_6c»", "");
						else
							this.FindAndReplace(wordApp, "«M_6c»", (objReport.lstOutput[i].col6c == "Performance achieved") ? "Performance achieved" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col6c)).ToString("#,0"));
						//this.FindAndReplace(wordApp,	"«M_3a»"	,	"$"+objReport.lstOutput[i].col3a);//Repeated.
						if (objReport.lstOutput[i].col6d == null)
							this.FindAndReplace(wordApp, "«M_6d»", "");
						else
							this.FindAndReplace(wordApp, "«M_6d»", (objReport.lstOutput[i].col6d == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col6d)).ToString("#,0"));
						this.FindAndReplace(wordApp, "«M_6e»", objReport.lstOutput[i].col6e == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col6e)).ToString("#,0"));

						//9.Gross Revenue per OD Hour					
						//this.FindAndReplace(wordApp, "«Q24»", "$" + objReport.lstInput[i].colQ24.ToString("#,0.##")); //Repeated.
						//this.FindAndReplace(wordApp, "«Q11»", objReport.lstInput[i].colQ11.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_7a»", objReport.lstOutput[i].col7a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col7a)).ToString("#,0"));
						this.FindAndReplace(wordApp, "«M_7b»", objReport.lstOutput[i].col7b);
						//this.FindAndReplace(wordApp,	"«Q11»"	,	objReport.lstInput[i].colQ11.ToString("#,0.##")); //Repeated
						//this.FindAndReplace(wordApp, "«Q24»", "$"+objReport.lstInput[i].colQ24.ToString("#,0.##"));//Repeated.
						if (objReport.lstOutput[i].col7c == null)
							this.FindAndReplace(wordApp, "«M_7c»", "");
						else
							this.FindAndReplace(wordApp, "«M_7c»", (objReport.lstOutput[i].col7c == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col7c)).ToString("#,0"));

						this.FindAndReplace(wordApp, "«M_7d»", objReport.lstOutput[i].col7d == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col7d)).ToString("#,0"));

						//10.Annual Gross Revenue per FTE OD					
						//this.FindAndReplace(wordApp, "«Q24»", "$" + objReport.lstInput[i].colQ24.ToString("#,0.##")); //Repeated.
						//this.FindAndReplace(wordApp, "«Q11»", objReport.lstInput[i].colQ11.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_8a»", Convert.ToDecimal(objReport.lstOutput[i].col8a));
						this.FindAndReplace(wordApp, "«M_8b»", objReport.lstOutput[i].col8b == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col8b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_8c»", objReport.lstOutput[i].col8c);
						//this.FindAndReplace(wordApp,	"«M_8a»"	,	objReport.lstOutput[i].col8a	);//Repeated.
						if (objReport.lstOutput[i].col8d == null)
							this.FindAndReplace(wordApp, "«M_8d»", "");
						else
							this.FindAndReplace(wordApp, "«M_8d»", (objReport.lstOutput[i].col8d == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col8d)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_8e»", objReport.lstOutput[i].col8e == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col8e)).ToString("#,0.##"));

						//11.Gross Revenue per Non-OD Staff Hour					
						//this.FindAndReplace(wordApp,	"«Q24»"	,	"$"+objReport.lstInput[i].colQ24.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q7»", objReport.lstInput[i].colQ7 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ7)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_9a»", objReport.lstOutput[i].col9a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col9a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_9b»", objReport.lstOutput[i].col9b);

						//11.Gross Revenue per Square Foot of Office Space					
						//this.FindAndReplace(wordApp,    "«Q24»", "$" + objReport.lstInput[i].colQ24.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q2»", objReport.lstInput[i].colQ2 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ2)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_9c»", objReport.lstOutput[i].col9c == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col9c)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_9d»", objReport.lstOutput[i].col9d);

						//12.Your Total Practice Productivity Metrics:Best to Worst Percentile Rankings
						//this.FindAndReplace(wordApp,	"«M_9b»"	,	objReport.lstOutput[i].col9b); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_3b»"	,	objReport.lstOutput[i].col3b); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_5b»"	,	objReport.lstOutput[i].col5b); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_6b»"	,	objReport.lstOutput[i].col6b); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_7b»"	,	objReport.lstOutput[i].col7b); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_8c»"	,	objReport.lstOutput[i].col8c); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_9d»"	,	objReport.lstOutput[i].col9d); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_4b»"	,	objReport.lstOutput[i].col4b); //Repeated.

						//13.Eyewear Metrics					
						//14.Eyewear Sales % of Gross Revenue 					
						//this.FindAndReplace(wordApp, "«Q24»", "$"+objReport.lstInput[i].colQ24.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q26f»", objReport.lstInput[i].colQ26f == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ26f)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n12a»", objReport.lstOutput[i].coln12a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln12a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_12b»", objReport.lstOutput[i].col12b);

						//15.Eyewear Rxes per 100 Complete Exams					
						this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col13a)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q14»", objReport.lstInput[i].colQ14.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_13b»", Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col13b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_13c»", objReport.lstOutput[i].col13c);
						//this.FindAndReplace(wordApp, "«Q26f»", "$" + objReport.lstInput[i].colQ26f.ToString("#,0.##")); //Repeated.
						//this.FindAndReplace(wordApp, "«Q14»", objReport.lstInput[i].colQ14.ToString("#,0.##")); //Repeated.
						if (objReport.lstOutput[i].col13d == null)
							this.FindAndReplace(wordApp, "«M_13d»", "");
						else
							this.FindAndReplace(wordApp, "«M_13d»", (objReport.lstOutput[i].col13d == "Performance achieved") ? "Performance achieved" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col13d)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_13e»", objReport.lstOutput[i].col13e == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col13e)).ToString("#,0.##"));
						if (objReport.lstOutput[i].col13f == null)
							this.FindAndReplace(wordApp, "«M_13f»", "");
						else
							this.FindAndReplace(wordApp, "«M_13f»", (objReport.lstOutput[i].col13f == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col13f)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_13g»", objReport.lstOutput[i].col13g == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col13g)).ToString("#,0.##"));

						//16.Eyewear Gross Revenue per Eyewear Rx					
						//this.FindAndReplace(wordApp,	"«Q26f»"	,	objReport.lstInput[i].colQ26f	); //Repeated.
						//this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_14a»", objReport.lstOutput[i].col14a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col14a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_14b»", objReport.lstOutput[i].col14b);
						//this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a.ToString("#,0.##")); //Repeated.
						if (objReport.lstOutput[i].col14c == null)
							this.FindAndReplace(wordApp, "«M_14c»", "");
						else
							this.FindAndReplace(wordApp, "«M_14c»", (objReport.lstOutput[i].col14c == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col14c)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_14d»", objReport.lstOutput[i].col14d == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col14d)).ToString("#,0.##"));

						//17.Eyewear Gross Profit Margin % 					
						//this.FindAndReplace(wordApp,	"«Q26f»"	,	objReport.lstInput[i].colQ26f	); //Repeated.
						this.FindAndReplace(wordApp, "«M_15a»", objReport.lstOutput[i].col15a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col15a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_15b»", objReport.lstOutput[i].col15b == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col15b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n15c»", objReport.lstOutput[i].coln15c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln15c)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_15d»", objReport.lstOutput[i].col15d);
						//this.FindAndReplace(wordApp,	"«Q26f»"	,	objReport.lstInput[i].colQ26f	); //Repeated.
						if (objReport.lstOutput[i].col15e == null)
							this.FindAndReplace(wordApp, "«M_15e»", "");
						else
							this.FindAndReplace(wordApp, "«M_15e»", (objReport.lstOutput[i].col15e == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col15e)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_15f»", objReport.lstOutput[i].col15f == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col15f)).ToString("#,0.##"));

						//18.Progressive Lens % of Presbyopic Rxes					
						//this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_16a»", objReport.lstOutput[i].col16a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col16a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_16b»", objReport.lstOutput[i].col16b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col16b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n16c»", objReport.lstOutput[i].coln16c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln16c)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_16d»", objReport.lstOutput[i].col16d);
						//this.FindAndReplace(wordApp, "«M_16a»", objReport.lstOutput[i].col16a.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_16e»", objReport.lstOutput[i].col16e);

						//this.FindAndReplace(wordApp, "«M_16e»", objReport.lstOutput[i].col16e == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col16e)).ToString("#,0.##"));
						//Math.Round(Convert.ToDecimal((objReport.lstOutput[i].col16e))));
						this.FindAndReplace(wordApp, "«M_16f»", objReport.lstOutput[i].col16f == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col16f)).ToString("#,0.##"));

						//19.No-Glare (AR) Lens % of Eyewear Rxes					
						//this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_17a»", objReport.lstOutput[i].col17a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col17a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q33b»", objReport.lstInput[i].colQ33b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ33b)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_17b»", objReport.lstOutput[i].col17b);
						//this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a.ToString("#,0.##")); //Repeated.
						if (objReport.lstOutput[i].col17c == null)
							this.FindAndReplace(wordApp, "«M_17c»", "");
						else
							this.FindAndReplace(wordApp, "«M_17c»", (objReport.lstOutput[i].col17c == "Performance achieved") ? "Performance achieved" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col17c)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_17d»", objReport.lstOutput[i].col17d == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col17d)).ToString("#,0.##"));

						//20.High Index Lens % of Eyewear Rxes					
						//this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_18a»", objReport.lstOutput[i].col18a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col18a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q33a»", objReport.lstInput[i].colQ33a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ33a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_18b»", objReport.lstOutput[i].col18b);
						//this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a.ToString("#,0.##")); //Repeated.
						if (objReport.lstOutput[i].col18c == null)
							this.FindAndReplace(wordApp, "«M_18c»", "");
						else
							this.FindAndReplace(wordApp, "«M_18c»", (objReport.lstOutput[i].col18c == "Performance achieved") ? "Performance achieved" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col18c)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_18d»", objReport.lstOutput[i].col18d == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col18d)).ToString("#,0.##"));

						//21.Photochromic Lens % of Eyewear Rxes					
						//this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_19a»", objReport.lstOutput[i].col19a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col19a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q33c»", objReport.lstInput[i].colQ33c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ33c)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_19b»", objReport.lstOutput[i].col19b);
						//this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a.ToString("#,0.##")); //Repeated.
						if (objReport.lstOutput[i].col19c == null)
							this.FindAndReplace(wordApp, "«M_19c»", "");
						else
							this.FindAndReplace(wordApp, "«M_19c»", (objReport.lstOutput[i].col19c == "Performance achieved") ? "Performance achieved" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col19c)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_19d»", objReport.lstOutput[i].col19d == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col19d)).ToString("#,0.##"));

						//22.Eyewear Multiple Pair Sales % Eyewear Buyers					
						this.FindAndReplace(wordApp, "«Q30»", objReport.lstInput[i].colQ30 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ30)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_20a»", objReport.lstOutput[i].col20a);
						//this.FindAndReplace(wordApp, "«M_13a»", objReport.lstOutput[i].col13a.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_20b»", objReport.lstOutput[i].col20b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col20b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_20c»", objReport.lstOutput[i].col20c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col20c)).ToString("#,0.##"));
						if (objReport.lstOutput[i].col20d == null)
							this.FindAndReplace(wordApp, "«M_20d»", "");
						else
							this.FindAndReplace(wordApp, "«M_20d»", (objReport.lstOutput[i].col20d == "Performance achieved") ? "Performance achieved" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col20d)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_20e»", objReport.lstOutput[i].col20e == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col20e)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_20f»", objReport.lstOutput[i].col20f == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col20f)).ToString("#,0.##"));

						//23.% of Contact Lens Patients Purchasing Eyewear During Exam Visit 					
						this.FindAndReplace(wordApp, "«Q15b»", objReport.lstInput[i].colQ15b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ15b)).ToString("#,0.##"));
						// this.FindAndReplace(wordApp, "«Q15b»", objReport.lstInput[i].colQ15b == null ? "" : Math.Round((Convert.ToDecimal(objReport.lstInput[i].colQ15b) * Convert.ToDecimal(objReport.lstInput[i].colQ14)) / 100).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q37»", objReport.lstInput[i].colQ37 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ37)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_21a»", objReport.lstOutput[i].col21a);
						this.FindAndReplace(wordApp, "«M_21b»", objReport.lstOutput[i].col21b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col21b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_21c»", objReport.lstOutput[i].col21c == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col21c)).ToString("#,0.##"));
						if (objReport.lstOutput[i].col21d == null)
							this.FindAndReplace(wordApp, "«M_21d»", "");
						else
							this.FindAndReplace(wordApp, "«M_21d»", (objReport.lstOutput[i].col21d == "Performance achieved") ? "Performance achieved" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col21d)).ToString("#,0.##"));
						if (objReport.lstOutput[i].col21e == null)
							this.FindAndReplace(wordApp, "«M_21e»", "");
						else
							this.FindAndReplace(wordApp, "«M_21e»", (objReport.lstOutput[i].col21e == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col21e)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_21f»", objReport.lstOutput[i].col21f == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col21f)).ToString("#,0.##"));

						//24.Your Eyewear Metrics Performance Summary					
						//this.FindAndReplace(wordApp,	"«M_12b»"	,	objReport.lstOutput[i].col12b	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_13c»"	,	objReport.lstOutput[i].col13c	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_14b»"	,	objReport.lstOutput[i].col14b	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_15d»"	,	objReport.lstOutput[i].col15d	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_16d»"	,	objReport.lstOutput[i].col16d	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_17b»"	,	objReport.lstOutput[i].col17b	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_18b»"	,	objReport.lstOutput[i].col18b	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_19b»"	,	objReport.lstOutput[i].col19b	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_20a»"	,	objReport.lstOutput[i].col20a	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_21a»"	,	objReport.lstOutput[i].col21a	); //Repeated.
						this.FindAndReplace(wordApp, "«av122a»", objReport.lstOutput[i].colav122a == null ? "" : GetOrdinal(Convert.ToInt32(Math.Round(Convert.ToDecimal(objReport.lstOutput[i].colav122a)).ToString())));

						//25.Contact Lens Metrics:					
						//26.Contact Lens Sales % of Gross Revenue					
						//this.FindAndReplace(wordApp,	"«Q24»"	,	objReport.lstInput[i].colQ24	); //Repeated.
						this.FindAndReplace(wordApp, "«Q26g»", objReport.lstInput[i].colQ26g == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ26g)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n24a»", objReport.lstOutput[i].coln24a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln24a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_24b»", objReport.lstOutput[i].col24b);


						//27.Contact Lens Gross Profit Margin %					
						//this.FindAndReplace(wordApp, "«Q26g»", "$"+objReport.lstInput[i].colQ26g.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q52f»", objReport.lstInput[i].colQ52f == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ52f)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_25a»", objReport.lstOutput[i].col25a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col25a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n25b»", objReport.lstOutput[i].coln25b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln25b)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_25c»", objReport.lstOutput[i].col25c);
						//this.FindAndReplace(wordApp,	"«Q26g»"	,	objReport.lstInput[i].colQ26g	); //Repeated.
						if (objReport.lstOutput[i].col25d == null)
							this.FindAndReplace(wordApp, "«M_25d»", "");
						else
							this.FindAndReplace(wordApp, "«M_25d»", (objReport.lstOutput[i].col25d == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col25d)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_25e»", objReport.lstOutput[i].col25e == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col25e)).ToString("#,0.##"));

						//28.Contact Lens Wearer % of Active Patients					
						// this.FindAndReplace(wordApp,	"«Q12»"	,	objReport.lstInput[i].colQ12	); //Repeated.
						this.FindAndReplace(wordApp, "«Q13b»", objReport.lstInput[i].colQ13b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ13b)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_26a»", objReport.lstOutput[i].col26a);

						//28.Contact Lens Exams % of Total Complete Eye Exams					
						//this.FindAndReplace(wordApp,	"«Q14»"	,	objReport.lstInput[i].colQ14); //Repeated.
						//this.FindAndReplace(wordApp,	"«Q15b»"	,	objReport.lstInput[i].colQ15b); //Repeated.
						this.FindAndReplace(wordApp, "«n26b»", objReport.lstOutput[i].coln26b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln26b)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_26c»", objReport.lstOutput[i].col26c);

						//29.Annual Contact Lens Sales per Contact Lens Exam					
						//this.FindAndReplace(wordApp, "«Q26g»", "$"+objReport.lstInput[i].colQ26g.ToString("#,0.##")); //Repeated.
						//this.FindAndReplace(wordApp,	"«Q15b»"	,	objReport.lstInput[i].colQ15b	); //Repeated.
						this.FindAndReplace(wordApp, "«M_27a»", objReport.lstOutput[i].col27a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col27a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_27b»", objReport.lstOutput[i].col27b);
						//this.FindAndReplace(wordApp,	"«Q15b»"	,	objReport.lstInput[i].colQ15b	); //Repeated.
						if (objReport.lstOutput[i].col27c == null)
							this.FindAndReplace(wordApp, "«M_27c»", "");
						else
							this.FindAndReplace(wordApp, "«M_27c»", (objReport.lstOutput[i].col27c == "Performance achieved") ? "Performance achieved" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col27c)).ToString("#,0.##"));

						this.FindAndReplace(wordApp, "«M_27d»", objReport.lstOutput[i].col27d == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col27d)).ToString("#,0.##"));

						//30.Contact Lens New Fits per 100 Contact Lens Exams					
						this.FindAndReplace(wordApp, "«Q42a»", objReport.lstInput[i].colQ42a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ42a)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp,	"«Q15b»"	,	objReport.lstInput[i].colQ15b	); //Repeated.
						this.FindAndReplace(wordApp, "«M_28a»", objReport.lstOutput[i].col28a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col28a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_28b»", objReport.lstOutput[i].col28b);

						//Contact Lens Refits % of Contact Lens Exams					
						this.FindAndReplace(wordApp, "«Q43a»", objReport.lstInput[i].colQ43a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ43a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_28c»", objReport.lstOutput[i].col28c);

						//31.Silicone Hydrogel Wearer % of Soft Lens Wearers					
						this.FindAndReplace(wordApp, "«Q41a»", objReport.lstInput[i].colQ41a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ41a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_29a»", objReport.lstOutput[i].col29a);

						//Daily Disposable Wearer % of Soft Lens Wearers					
						this.FindAndReplace(wordApp, "«Q39a»", objReport.lstInput[i].colQ39a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ39a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_29b»", objReport.lstOutput[i].col29b);

						//Monthly Lens % of Soft Lens Wearers					
						this.FindAndReplace(wordApp, "«Q39c»", objReport.lstInput[i].colQ39c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ39c)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_29c»", objReport.lstOutput[i].col29c);

						//32.Soft Toric Lens Wearer % of Soft Lens Wearers					
						this.FindAndReplace(wordApp, "«Q40b»", objReport.lstInput[i].colQ40b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ40b)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_30a»", objReport.lstOutput[i].col30a);

						//Soft Multi-focal Lens Wearer % of Soft Lens Wearers					
						this.FindAndReplace(wordApp, "«Q40d»", objReport.lstInput[i].colQ40d == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ40d)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_30b»", objReport.lstOutput[i].col30b);

						//33.Your Contact Lens Metrics Performance Summary					
						//this.FindAndReplace(wordApp,	"«M_24b»"	,	objReport.lstOutput[i].col24b	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_25c»"	,	objReport.lstOutput[i].col25c	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_26a»"	,	objReport.lstOutput[i].col26a	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_26c»"	,	objReport.lstOutput[i].col26c	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_27b»"	,	objReport.lstOutput[i].col27b	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_28b»"	,	objReport.lstOutput[i].col28b	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_28c»"	,	objReport.lstOutput[i].col28c	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_29a»"	,	objReport.lstOutput[i].col29a	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_29c»"	,	objReport.lstOutput[i].col29c	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_30a»"	,	objReport.lstOutput[i].col30a	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_30b»"	,	objReport.lstOutput[i].col30b	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_29b»"	,	objReport.lstOutput[i].col29b	); //Repeated.
						this.FindAndReplace(wordApp, "«av131a»", objReport.lstOutput[i].colav131a == null ? "" : GetOrdinal(Convert.ToInt32(Math.Round(Convert.ToDecimal(objReport.lstOutput[i].colav131a)))));

						//34.Medical Eye Care Metrics					
						//35.Non-refractive Fee Revenue % of Total Gross Revenue					
						// this.FindAndReplace(wordApp,	"«Q24»"	,	objReport.lstInput[i].colQ24	); //Repeated.
						this.FindAndReplace(wordApp, "«M_33a»", objReport.lstOutput[i].col33a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col33a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n33b»", objReport.lstOutput[i].coln33b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln33b)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_33c»", objReport.lstOutput[i].col33c);

						//Medical Eye Care Visits % of Total Patient Visits					
						this.FindAndReplace(wordApp, "«M_33d»", objReport.lstOutput[i].col33d == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col33d)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp,	"«Q14»"	,	objReport.lstInput[i].colQ14); //Repeated.
						this.FindAndReplace(wordApp, "«M_33e»", objReport.lstOutput[i].col33e == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col33e)).ToString());
						this.FindAndReplace(wordApp, "«n33f»", objReport.lstOutput[i].coln33f == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln33f)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_33g»", objReport.lstOutput[i].col33g);
						//Annual Medical Eye Care Visits per 1,000 Active Patients					
						this.FindAndReplace(wordApp, "«M_33d»", objReport.lstOutput[i].col33d == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col33d)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp,	"«Q12»"	,	objReport.lstInput[i].colQ12	); //Repeated.
						this.FindAndReplace(wordApp, "«M_34a»", objReport.lstOutput[i].col34a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col34a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_34b»", objReport.lstOutput[i].col34b);
						//36.Annual Pharmaceutical Rxes per 1,000 Active Patients					
						this.FindAndReplace(wordApp, "«M_34c»", objReport.lstOutput[i].col34c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col34c)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp,	"«Q12»"	,	objReport.lstInput[i].colQ12	); //Repeated.
						this.FindAndReplace(wordApp, "«M_34d»", objReport.lstOutput[i].col34d == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col34d)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_34e»", objReport.lstOutput[i].col34e);
						//Your Medical Eye Care Metrics Performance Summary
						//this.FindAndReplace(wordApp,	"«M_33c»"	,	objReport.lstOutput[i].col33c	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_33g»"	,	objReport.lstOutput[i].col33g	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_34b»"	,	objReport.lstOutput[i].col34b	); //Repeated.
						//this.FindAndReplace(wordApp,	"«M_34e»"	,	objReport.lstOutput[i].col34e	); //Repeated.
						this.FindAndReplace(wordApp, "«av134a»", objReport.lstOutput[i].colav134a == null ? "" : GetOrdinal(Convert.ToInt32(Math.Round(Convert.ToDecimal(objReport.lstOutput[i].colav134a)))));
						//37.Marketing Metrics                                 
						//38.Marketing Spending % of Gross Revenue                              
						//this.FindAndReplace(wordApp, "«Q24»", "$" + objReport.lstInput[i].colQ24.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q56»", objReport.lstInput[i].colQ56 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ56)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n36a»", objReport.lstOutput[i].coln36a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln36a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_36b»", objReport.lstOutput[i].col36b);
						//Annual Marketing Spending per Complete Exam                                  
						//this.FindAndReplace(wordApp, "«Q56»"       ,      objReport.lstInput[i].colQ56       ); //Repeated
						//this.FindAndReplace(wordApp, "«Q14»", objReport.lstInput[i].colQ14.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«M_36c»", "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col36c)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_36d»", objReport.lstOutput[i].col36d);
						//New Patient Exams % of Total Exams                             
						this.FindAndReplace(wordApp, "«Q16»", objReport.lstInput[i].colQ16 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ16)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_36e»", objReport.lstOutput[i].col36e);
						//39.Website Expense                               
						this.FindAndReplace(wordApp, "«Q63»", objReport.lstInput[i].colQ6 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ63)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_37a»", objReport.lstOutput[i].col37a);
						//39.% of Total New Patients Attracted by Practice Website                                   
						//this.FindAndReplace(wordApp, "«Q14»", objReport.lstInput[i].colQ14.ToString("#,0.##"));//Repeated
						//this.FindAndReplace(wordApp, "«Q16»", objReport.lstInput[i].colQ16 + "%");//Repeated
						this.FindAndReplace(wordApp, "«M_37b»", objReport.lstOutput[i].col37b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col37b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q17»", objReport.lstInput[i].colQ17 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ17)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n37c»", objReport.lstOutput[i].coln37c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln37c)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_37d»", objReport.lstOutput[i].col37d);
						//39.Recall Staff Minutes per Complete Eye Exam                                
						this.FindAndReplace(wordApp, "«Q8»", objReport.lstInput[i].colQ8);
						this.FindAndReplace(wordApp, "«M_37e»", objReport.lstOutput[i].col37e);
						this.FindAndReplace(wordApp, "«M_37f»", objReport.lstOutput[i].col37f == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col37f)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q14»"       ,      objReport.lstInput[i].colQ14       ); //Repeated
						this.FindAndReplace(wordApp, "«M_37g»", Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col37g)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_37h»", objReport.lstOutput[i].col37h);
						//40.Your Marketing Metrics Performance Summary                                
						//this.FindAndReplace(wordApp, "«M_36b»"     ,      objReport.lstOutput[i].col36b       );//Repeated
						//this.FindAndReplace(wordApp, "«M_36d»"     ,      objReport.lstOutput[i].col36d       );//Repeated
						//this.FindAndReplace(wordApp, "«M_36e»"     ,      objReport.lstOutput[i].col36e       );//Repeated
						// this.FindAndReplace(wordApp,       "«M_37a»"     ,       objReport.lstOutput[i].col37a     );//Repeated
						//this.FindAndReplace(wordApp, "«M_37d»"     ,      objReport.lstOutput[i].col37d       );//Repeated
						//this.FindAndReplace(wordApp, "«M_37h»"     ,      objReport.lstOutput[i].col37h       );//Repeated
						//41.Financial Metrics                             
						//42.Professional Exam Fees:                              
						//Non-contact Lens Exam Fee                               
						this.FindAndReplace(wordApp, "«Q47»", objReport.lstInput[i].colQ47 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ47)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_40a»", objReport.lstOutput[i].col40a);
						//Contact Lens New Fit Exam Fee -- Sphere                               
						this.FindAndReplace(wordApp, "«Q48»", objReport.lstInput[i].colQ48 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ48)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_40b»", objReport.lstOutput[i].col40b);
						//Contact Lens New Fit Exam Fee – Soft Toric                            
						this.FindAndReplace(wordApp, "«Q49»", objReport.lstInput[i].colQ49 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ49)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_40c»", objReport.lstOutput[i].col40c);
						//43.Contact Lens New Fit Exam Fee – Soft Multi-focal                                 
						this.FindAndReplace(wordApp, "«Q50»", objReport.lstInput[i].colQ50 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ50)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_41a»", objReport.lstOutput[i].col41a);
						//Contact Lens Exam Fee – No Refitting                                  
						this.FindAndReplace(wordApp, "«Q51»", objReport.lstInput[i].colQ51 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ51)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_41b»", objReport.lstOutput[i].col41b);
						//44.Average Collected Fee Revenue per Complete Exam                                  
						this.FindAndReplace(wordApp, "«Q26»", objReport.lstInput[i].colQ26 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ26)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q14»"       ,      objReport.lstInput[i].colQ14       );//Repeated
						this.FindAndReplace(wordApp, "«M_42a»", objReport.lstOutput[i].col42a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col42a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_42b»", objReport.lstOutput[i].col42b);
						//% of Exams Provided with Managed Care Discount                               
						this.FindAndReplace(wordApp, "«Q18»", Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ18)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_42c»", objReport.lstOutput[i].col42c);
						//45.Accounts Receivables Metrics                                
						//this.FindAndReplace(wordApp, "«Q24»"       ,      objReport.lstInput[i].colQ24       );//Repeated
						this.FindAndReplace(wordApp, "«M_43a»", objReport.lstOutput[i].col43a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col43a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q25»", objReport.lstInput[i].colQ25 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ25)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n43b»", objReport.lstOutput[i].coln43b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln43b)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_43d»", Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col43d)));
						this.FindAndReplace(wordApp, "«M_43e»", objReport.lstOutput[i].col43e);
						//46.Practice Expense/Net Income Ratios :                               
						//Cost-of Goods                              
						//this.FindAndReplace(wordApp, "«Q24»"       ,      objReport.lstInput[i].colQ24       );//Repeated
						this.FindAndReplace(wordApp, "«Q52j»", objReport.lstInput[i].colQ52j == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ52j)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n43f»", objReport.lstOutput[i].coln43f == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln43f)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_43g»", objReport.lstOutput[i].col43g);
						//Staffing                            
						//this.FindAndReplace(wordApp, "«Q24»"       ,      objReport.lstInput[i].colQ24       );//Repeated
						this.FindAndReplace(wordApp, "«Q53»", objReport.lstInput[i].colQ53 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ53)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n44a»", objReport.lstOutput[i].coln44a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln44a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_44b»", objReport.lstOutput[i].col44b);
						//General Overhead                                 
						//47.Occupancy                               
						//this.FindAndReplace(wordApp, "«Q24»"       ,      objReport.lstInput[i].colQ24       );//Repeated
						this.FindAndReplace(wordApp, "«Q54»", objReport.lstInput[i].colQ54 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ54)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n44c»", objReport.lstOutput[i].coln44c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln44c)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_44d»", objReport.lstOutput[i].col44d);
						//Net Income % of Gross Revenue                                  
						//this.FindAndReplace(wordApp, "«Q24»"       ,      objReport.lstInput[i].colQ24       ); //Repeated
						this.FindAndReplace(wordApp, "«M_45a»", objReport.lstOutput[i].col45a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col45a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_45b»", objReport.lstOutput[i].col45b == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col45b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«n45c»", objReport.lstOutput[i].coln45c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].coln45c)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«M_45d»", objReport.lstOutput[i].col45d);
						//Chair Cost per Complete Exam                            
						//this.FindAndReplace(wordApp, "«M_45a»"     ,      objReport.lstOutput[i].col45a       );//Repeated
						//this.FindAndReplace(wordApp, "«Q52j»"      ,      objReport.lstInput[i].colQ52j       );//Repeated
						this.FindAndReplace(wordApp, "«M_45e»", objReport.lstOutput[i].col45e == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col45e)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q14»"       ,      objReport.lstInput[i].colQ14       );//Repeated
						this.FindAndReplace(wordApp, "«M_45f»", objReport.lstOutput[i].col45f == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstOutput[i].col45f)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«M_45g»", objReport.lstOutput[i].col45g);
						//48.Your Financial Metrics Performance Summary                                
						this.FindAndReplace(wordApp, "«M_40a»", objReport.lstOutput[i].col40a);
						this.FindAndReplace(wordApp, "«M_40b»", objReport.lstOutput[i].col40b);
						this.FindAndReplace(wordApp, "«M_40c»", objReport.lstOutput[i].col40c);
						this.FindAndReplace(wordApp, "«M_41a»", objReport.lstOutput[i].col41a);
						this.FindAndReplace(wordApp, "«M_41b»", objReport.lstOutput[i].col41b);
						this.FindAndReplace(wordApp, "«av146a»", objReport.lstOutput[i].colav146a == null ? "" : GetOrdinal(Convert.ToInt32(Math.Round(Convert.ToDecimal(objReport.lstOutput[i].colav146a)))));
						//this.FindAndReplace(wordApp, "«M_42b»", objReport.lstOutput[i].col42b); //Repeated.
						this.FindAndReplace(wordApp, "«M_42c»", objReport.lstOutput[i].col42c);
						//this.FindAndReplace(wordApp, "«M_43e»", objReport.lstOutput[i].col43e); //Repeated.
						this.FindAndReplace(wordApp, "«M_43g»", objReport.lstOutput[i].col43g);
						this.FindAndReplace(wordApp, "«M_44b»", objReport.lstOutput[i].col44b);
						this.FindAndReplace(wordApp, "«M_44d»", objReport.lstOutput[i].col44d);
						//this.FindAndReplace(wordApp, "«M_45d»", objReport.lstOutput[i].col45d); //Repeated.
						//this.FindAndReplace(wordApp, "«M_45g»", objReport.lstOutput[i].col45g);//Repeated
						//49.Best Practices                                
						//Financial Management                             
						if (objReport.lstOutput[i].coln64m == null)
							this.FindAndReplace(wordApp, "«n64m»", "");
						else
							this.FindAndReplace(wordApp, "«n64m»", (objReport.lstOutput[i].coln64m == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64l == null)
							this.FindAndReplace(wordApp, "«n64l»", "");
						else
							this.FindAndReplace(wordApp, "«n64l»", (objReport.lstOutput[i].coln64l == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64i == null)
							this.FindAndReplace(wordApp, "«n64i»", "");
						else
							this.FindAndReplace(wordApp, "«n64i»", (objReport.lstOutput[i].coln64i == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64c == null)
							this.FindAndReplace(wordApp, "«n64c»", "");
						else
							this.FindAndReplace(wordApp, "«n64c»", (objReport.lstOutput[i].coln64c == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64j == null)
							this.FindAndReplace(wordApp, "«n64j»", "");
						else
							this.FindAndReplace(wordApp, "«n64j»", (objReport.lstOutput[i].coln64j == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64g == null)
							this.FindAndReplace(wordApp, "«n64g»", "");
						else
							this.FindAndReplace(wordApp, "«n64g»", (objReport.lstOutput[i].coln64g == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64h == null)
							this.FindAndReplace(wordApp, "«n64h»", "");
						else
							this.FindAndReplace(wordApp, "«n64h»", (objReport.lstOutput[i].coln64h == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64b == null)
							this.FindAndReplace(wordApp, "«n64b»", "");
						else
							this.FindAndReplace(wordApp, "«n64b»", (objReport.lstOutput[i].coln64b == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64k == null)
							this.FindAndReplace(wordApp, "«n64k»", "");
						else
							this.FindAndReplace(wordApp, "«n64k»", (objReport.lstOutput[i].coln64k == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64o == null)
							this.FindAndReplace(wordApp, "«n64o»", "");
						else
							this.FindAndReplace(wordApp, "«n64o»", (objReport.lstOutput[i].coln64o == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64n == null)
							this.FindAndReplace(wordApp, "«n64n»", "");
						else
							this.FindAndReplace(wordApp, "«n64n»", (objReport.lstOutput[i].coln64n == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64d == null)
							this.FindAndReplace(wordApp, "«n64d»", "");
						else
							this.FindAndReplace(wordApp, "«n64d»", (objReport.lstOutput[i].coln64d == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64f == null)
							this.FindAndReplace(wordApp, "«n64f»", "");
						else
							this.FindAndReplace(wordApp, "«n64f»", (objReport.lstOutput[i].coln64f == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64e == null)
							this.FindAndReplace(wordApp, "«n64e»", "");
						else
							this.FindAndReplace(wordApp, "«n64e»", (objReport.lstOutput[i].coln64e == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln64a == null)
							this.FindAndReplace(wordApp, "«n64a»", "");
						else
							this.FindAndReplace(wordApp, "«n64a»", (objReport.lstOutput[i].coln64a == true) ? "Yes" : "No");
						//49.Marketing Management                                 
						if (objReport.lstOutput[i].coln65h == null)
							this.FindAndReplace(wordApp, "«n65h»", "");
						else
							this.FindAndReplace(wordApp, "«n65h»", (objReport.lstOutput[i].coln65h == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65i == null)
							this.FindAndReplace(wordApp, "«n65i»", "");
						else
							this.FindAndReplace(wordApp, "«n65i»", (objReport.lstOutput[i].coln65i == true) ? "Yes" : "No");
						//50.     
						if (objReport.lstOutput[i].coln65e == null)
							this.FindAndReplace(wordApp, "«n65e»", "");
						else
							this.FindAndReplace(wordApp, "«n65e»", (objReport.lstOutput[i].coln65e == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65f == null)
							this.FindAndReplace(wordApp, "«n65f»", "");
						else
							this.FindAndReplace(wordApp, "«n65f»", (objReport.lstOutput[i].coln65f == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65n == null)
							this.FindAndReplace(wordApp, "«n65n»", "");
						else
							this.FindAndReplace(wordApp, "«n65n»", (objReport.lstOutput[i].coln65n == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65j == null)
							this.FindAndReplace(wordApp, "«n65j»", "");
						else
							this.FindAndReplace(wordApp, "«n65j»", (objReport.lstOutput[i].coln65j == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65l == null)
							this.FindAndReplace(wordApp, "«n65l»", "");
						else
							this.FindAndReplace(wordApp, "«n65l»", (objReport.lstOutput[i].coln65l == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65b == null)
							this.FindAndReplace(wordApp, "«n65b»", "");
						else
							this.FindAndReplace(wordApp, "«n65b»", (objReport.lstOutput[i].coln65b == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65g == null)
							this.FindAndReplace(wordApp, "«n65g»", "");
						else
							this.FindAndReplace(wordApp, "«n65g»", (objReport.lstOutput[i].coln65g == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65c == null)
							this.FindAndReplace(wordApp, "«n65c»", "");
						else
							this.FindAndReplace(wordApp, "«n65c»", (objReport.lstOutput[i].coln65c == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65k == null)
							this.FindAndReplace(wordApp, "«n65k»", "");
						else
							this.FindAndReplace(wordApp, "«n65k»", (objReport.lstOutput[i].coln65k == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65m == null)
							this.FindAndReplace(wordApp, "«n65m»", "");
						else
							this.FindAndReplace(wordApp, "«n65m»", (objReport.lstOutput[i].coln65m == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65d == null)
							this.FindAndReplace(wordApp, "«n65d»", "");
						else
							this.FindAndReplace(wordApp, "«n65d»", (objReport.lstOutput[i].coln65d == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln65a == null)
							this.FindAndReplace(wordApp, "«n65a»", "");
						else
							this.FindAndReplace(wordApp, "«n65a»", (objReport.lstOutput[i].coln65a == true) ? "Yes" : "No");

						//50.Staff Management  
						if (objReport.lstOutput[i].coln66b == null)
							this.FindAndReplace(wordApp, "«n66b»", "");
						else
							this.FindAndReplace(wordApp, "«n66b»", (objReport.lstOutput[i].coln66b == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66i == null)
							this.FindAndReplace(wordApp, "«n66i»", "");
						else
							this.FindAndReplace(wordApp, "«n66i»", (objReport.lstOutput[i].coln66i == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66e == null)
							this.FindAndReplace(wordApp, "«n66e»", "");
						else
							this.FindAndReplace(wordApp, "«n66e»", (objReport.lstOutput[i].coln66e == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66a == null)
							this.FindAndReplace(wordApp, "«n66a»", "");
						else
							this.FindAndReplace(wordApp, "«n66a»", (objReport.lstOutput[i].coln66a == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66g == null)
							this.FindAndReplace(wordApp, "«n66g»", "");
						else
							this.FindAndReplace(wordApp, "«n66g»", (objReport.lstOutput[i].coln66g == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66k == null)
							this.FindAndReplace(wordApp, "«n66k»", "");
						else
							this.FindAndReplace(wordApp, "«n66k»", (objReport.lstOutput[i].coln66k == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66c == null)
							this.FindAndReplace(wordApp, "«n66c»", "");
						else
							this.FindAndReplace(wordApp, "«n66c»", (objReport.lstOutput[i].coln66c == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66d == null)
							this.FindAndReplace(wordApp, "«n66d»", "");
						else
							this.FindAndReplace(wordApp, "«n66d»", (objReport.lstOutput[i].coln66d == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66f == null)
							this.FindAndReplace(wordApp, "«n66f»", "");
						else
							this.FindAndReplace(wordApp, "«n66f»", (objReport.lstOutput[i].coln66f == true) ? "Yes" : "No");

						//51.     
						if (objReport.lstOutput[i].coln66j == null)
							this.FindAndReplace(wordApp, "«n66j»", "");
						else
							this.FindAndReplace(wordApp, "«n66j»", (objReport.lstOutput[i].coln66j == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66h == null)
							this.FindAndReplace(wordApp, "«n66h»", "");
						else
							this.FindAndReplace(wordApp, "«n66h»", (objReport.lstOutput[i].coln66h == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66t == null)
							this.FindAndReplace(wordApp, "«n66t»", "");
						else
							this.FindAndReplace(wordApp, "«n66t»", (objReport.lstOutput[i].coln66t == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66u == true)
							this.FindAndReplace(wordApp, "«n66u»", "");
						else
							this.FindAndReplace(wordApp, "«n66u»", (objReport.lstOutput[i].coln66u == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66o == null)
							this.FindAndReplace(wordApp, "«n66o»", "");
						else
							this.FindAndReplace(wordApp, "«n66o»", (objReport.lstOutput[i].coln66o == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66p == null)
							this.FindAndReplace(wordApp, "«n66p»", "");
						else
							this.FindAndReplace(wordApp, "«n66p»", (objReport.lstOutput[i].coln66p == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66s == null)
							this.FindAndReplace(wordApp, "«n66s»", "");
						else
							this.FindAndReplace(wordApp, "«n66s»", (objReport.lstOutput[i].coln66s == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66m == null)
							this.FindAndReplace(wordApp, "«n66m»", "");
						else
							this.FindAndReplace(wordApp, "«n66m»", (objReport.lstOutput[i].coln66m == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66n == null)
							this.FindAndReplace(wordApp, "«n66n»", "");
						else
							this.FindAndReplace(wordApp, "«n66n»", (objReport.lstOutput[i].coln66n == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66r == null)
							this.FindAndReplace(wordApp, "«n66r»", "");
						else
							this.FindAndReplace(wordApp, "«n66r»", (objReport.lstOutput[i].coln66r == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66v == null)
							this.FindAndReplace(wordApp, "«n66v»", "");
						else
							this.FindAndReplace(wordApp, "«n66v»", (objReport.lstOutput[i].coln66v == true) ? "Yes" : "No");

						if (objReport.lstOutput[i].coln66q == null)
							this.FindAndReplace(wordApp, "«n66q»", "");
						else
							this.FindAndReplace(wordApp, "«n66q»", (objReport.lstOutput[i].coln66q == true) ? "Yes" : "No");

						//Total “Best Practices” Score (10 points per “best practice”)                               
						this.FindAndReplace(wordApp, "«M_49a»", objReport.lstOutput[i].col49a);
						this.FindAndReplace(wordApp, "«M_49b»", objReport.lstOutput[i].col49b);
						//Financial Management Score                              
						this.FindAndReplace(wordApp, "«M_49c»", objReport.lstOutput[i].col49c);
						this.FindAndReplace(wordApp, "«M_49d»", objReport.lstOutput[i].col49d);
						//52.Marketing Management Score                                  
						this.FindAndReplace(wordApp, "«M_50a»", objReport.lstOutput[i].col50a);
						this.FindAndReplace(wordApp, "«M_50b»", objReport.lstOutput[i].col50b);
						//Staff Management Score                                  
						this.FindAndReplace(wordApp, "«M_50c»", objReport.lstOutput[i].col50c);
						this.FindAndReplace(wordApp, "«M_50d»", objReport.lstOutput[i].col50d);

						//53.Your Percentile Rankings: Best to Worst                            
						//this.FindAndReplace(wordApp, "«M_13c»", objReport.lstOutput[i].col13c); //Repeated.
						//this.FindAndReplace(wordApp, "«M_15d»", objReport.lstOutput[i].col15d); //Repeated.
						//this.FindAndReplace(wordApp, "«M_33c»", objReport.lstOutput[i].col33c); //Repeated.
						//this.FindAndReplace(wordApp, "«M_20a»", objReport.lstOutput[i].col20a); //Repeated.
						//this.FindAndReplace(wordApp, "«M_27b»", objReport.lstOutput[i].col27b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_24b»", objReport.lstOutput[i].col24b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_30b»", objReport.lstOutput[i].col30b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_26c»", objReport.lstOutput[i].col26c); //Repeated.
						//this.FindAndReplace(wordApp, "«av146a»", objReport.lstOutput[i].colav146a); //Repeated.

						//this.FindAndReplace(wordApp, "«M_25c»", objReport.lstOutput[i].col25c); //Repeated.
						//this.FindAndReplace(wordApp, "«av122a»", objReport.lstOutput[i].colav122a); //Repeated.
						//this.FindAndReplace(wordApp, "«av131a»", objReport.lstOutput[i].colav131a); //Repeated.
						//this.FindAndReplace(wordApp, "«M_28b»", objReport.lstOutput[i].col28b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_12b»", objReport.lstOutput[i].col12b); //Repeated.
						//this.FindAndReplace(wordApp, "«av134a»", objReport.lstOutput[i].colav134a); //Repeated.
						//this.FindAndReplace(wordApp, "«M_49b»", objReport.lstOutput[i].col49b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_29b»", objReport.lstOutput[i].col29b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_21a»", objReport.lstOutput[i].col21a); //Repeated.
						//this.FindAndReplace(wordApp, "«M_34b»", objReport.lstOutput[i].col34b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_33g»", objReport.lstOutput[i].col33g); //Repeated.
						//this.FindAndReplace(wordApp, "«M_43e»", objReport.lstOutput[i].col43e); //Repeated.
						//this.FindAndReplace(wordApp, "«M_34e»", objReport.lstOutput[i].col34e); //Repeated.
						//this.FindAndReplace(wordApp, "«M_42b»", objReport.lstOutput[i].col42b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_14b»", objReport.lstOutput[i].col14b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_16d»", objReport.lstOutput[i].col16d); //Repeated.
						//this.FindAndReplace(wordApp, "«M_19b»", objReport.lstOutput[i].col19b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_30a»", objReport.lstOutput[i].col30a); //Repeated.
						//this.FindAndReplace(wordApp, "«M_17b»", objReport.lstOutput[i].col17b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_28c»", objReport.lstOutput[i].col28c); //Repeated.
						//this.FindAndReplace(wordApp, "«M_18b»", objReport.lstOutput[i].col18b); //Repeated.
						//this.FindAndReplace(wordApp, "«M_26a»", objReport.lstOutput[i].col26a); //Repeated.
						//this.FindAndReplace(wordApp, "«M_29a»", objReport.lstOutput[i].col29a); //Repeated.
						//this.FindAndReplace(wordApp, "«M_45d»", objReport.lstOutput[i].col45d); //Repeated.
						//54.Section 3:                              
						//Questionnaire Responses                                 
						//About Your Facilities                            
						this.FindAndReplace(wordApp, "«Q1»", objReport.lstInput[i].colQ1);
						//this.FindAndReplace(wordApp, "«Q2»", objReport.lstInput[i].colQ2.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q3»", objReport.lstInput[i].colQ3 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ3)).ToString("#,0.##"));
						//About Your Manpower   
						this.FindAndReplace(wordApp, "«Q4»", objReport.lstInput[i].colQ4 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ4)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q5»", objReport.lstInput[i].colQ5 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ5)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q6»", objReport.lstInput[i].colQ6 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ6)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q7»", objReport.lstInput[i].colQ7.ToString("#,0.##")); //Repeated.
						// this.FindAndReplace(wordApp, "«Q8»", objReport.lstInput[i].colQ8.ToString("#,0.##"));//Repeated
						this.FindAndReplace(wordApp, "«Q9»", objReport.lstInput[i].colQ9);
						this.FindAndReplace(wordApp, "«Q10»", objReport.lstInput[i].colQ10 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ10)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q11»", objReport.lstInput[i].colQ11.ToString("#,0.##")); //Repeated.
						//54.About Your Patient Base   
						//this.FindAndReplace(wordApp, "«Q12»", objReport.lstInput[i].colQ12); //Repeated.
						//55.Of your total active patients, approximately what percentage falls into each of the following three groups?                         
						this.FindAndReplace(wordApp, "«Q13a»", objReport.lstInput[i].colQ13a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ13a)).ToString("#,0.##") + "%");
						//this.FindAndReplace(wordApp, "«Q13b»", objReport.lstInput[i].colQ13b + "%"); //Repeated.
						this.FindAndReplace(wordApp, "«Q13c»", objReport.lstInput[i].colQ13c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ13c)).ToString("#,0.##") + "%");
						//About Your Patient Visits                               
						//this.FindAndReplace(wordApp, "«Q14»", objReport.lstInput[i].colQ14.ToString("#,0.##"));//Repeated
						this.FindAndReplace(wordApp, "«Q15a»", objReport.lstInput[i].colQ15a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ15a)).ToString("#,0.##"));
						// this.FindAndReplace(wordApp, "«Q15a»", objReport.lstInput[i].colQ15a == null ? "" : Math.Round((Convert.ToDecimal(objReport.lstInput[i].colQ15a) * Convert.ToDecimal(objReport.lstInput[i].colQ14)) / 100).ToString("#,0.##"));

						//this.FindAndReplace(wordApp, "«Q15b»", objReport.lstInput[i].colQ15b.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q15c»", objReport.lstInput[i].colQ15c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ15c)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q15c»", objReport.lstInput[i].colQ15c == null ? "" : Math.Round((Convert.ToDecimal(objReport.lstInput[i].colQ15c) * Convert.ToDecimal(objReport.lstInput[i].colQ14)) / 100).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q16»", objReport.lstInput[i].colQ16 + "%"); //Repeated.
						//this.FindAndReplace(wordApp, "«Q17»", objReport.lstInput[i].colQ17.ToString("#,0.##"));//Repeated
						//this.FindAndReplace(wordApp, "«Q18»", objReport.lstInput[i].colQ18 + "%"); //Repeated.
						this.FindAndReplace(wordApp, "«Q19»", objReport.lstInput[i].colQ19 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ19)).ToString("#,0.##"));
						//56.                                 
						this.FindAndReplace(wordApp, "«Q20a»", objReport.lstInput[i].colQ20a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ20a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q20b»", objReport.lstInput[i].colQ20b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ20b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q20c»", objReport.lstInput[i].colQ20c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ20c)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q20d»", objReport.lstInput[i].colQ20d == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ20d)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q20e»", objReport.lstInput[i].colQ20e == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ20e)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q20f»", objReport.lstInput[i].colQ20f == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ20f)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q20g»", objReport.lstInput[i].colQ20g == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ20g)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«t20»", objReport.lstOutput[i].colt20 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].colt20)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q21a»", objReport.lstInput[i].colQ21a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ21a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q21b»", objReport.lstInput[i].colQ21b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ21b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q21c»", objReport.lstInput[i].colQ21c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ21c)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q21d»", objReport.lstInput[i].colQ21d == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ21d)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«t21»", objReport.lstOutput[i].colt21 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstOutput[i].colt21)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q22»", objReport.lstInput[i].colQ22 == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ22)).ToString("#,0.##"));
						//About Your Practice Revenue                             
						this.FindAndReplace(wordApp, "«Q23»", objReport.lstInput[i].colQ23 == null ? "" : "$" + Convert.ToDecimal(objReport.lstInput[i].colQ23).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q24»", objReport.lstInput[i].colQ24.ToString("#,0.##"));//Repeated
						//this.FindAndReplace(wordApp, "«Q25»", objReport.lstInput[i].colQ25.ToString("#,0.##"));//Repeated
						//57.                                 
						//this.FindAndReplace(wordApp, "«Q26»"       ,      objReport.lstInput[i].colQ26       );//Repeated
						this.FindAndReplace(wordApp, "«Q26a»", objReport.lstInput[i].colQ26a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ26a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q26b»", objReport.lstInput[i].colQ26b == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ26b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q26c»", objReport.lstInput[i].colQ26c == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ26c)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q26d»", objReport.lstInput[i].colQ26d == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ26d)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q26f»", "$" + objReport.lstInput[i].colQ26f.ToString("#,0.##")); //Repeated.
						//this.FindAndReplace(wordApp, "«Q26g»", "$" + objReport.lstInput[i].colQ26g.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q26h»", objReport.lstInput[i].colQ26h == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ26h)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q26i»", objReport.lstInput[i].colQ26i == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ26i)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q27a»", objReport.lstInput[i].colQ27a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ27a)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q27a»", "$" + objReport.lstInput[i].colQ27a.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q27c»", objReport.lstInput[i].colQ27c == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ27c)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q27e»", objReport.lstInput[i].colQ27e == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ27e)).ToString("#,0.##"));
						//About Your Eyewear Dispensing                                  
						this.FindAndReplace(wordApp, "«Q28»", objReport.lstInput[i].colQ28 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ28)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q29»", objReport.lstInput[i].colQ29 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ29)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q30»", objReport.lstInput[i].colQ30 + "%"); //Repeated.
						this.FindAndReplace(wordApp, "«Q31a»", objReport.lstInput[i].colQ31a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ31a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«Q31b»", objReport.lstInput[i].colQ31b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ31b)).ToString("#,0.##") + "%");
						//58.                                 
						this.FindAndReplace(wordApp, "«Q32a»", objReport.lstInput[i].colQ32a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ32a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«Q32b»", objReport.lstInput[i].colQ32b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ32b)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«Q32c»", objReport.lstInput[i].colQ32c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ32c)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«Q32d»", objReport.lstInput[i].colQ32d == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ32d)).ToString("#,0.##") + "%");
						//this.FindAndReplace(wordApp, "«Q33a»", objReport.lstInput[i].colQ33a + "%"); //Repeated.
						//this.FindAndReplace(wordApp, "«Q33b»", objReport.lstInput[i].colQ33b + "%"); //Repeated.
						//this.FindAndReplace(wordApp, "«Q33c»", objReport.lstInput[i].colQ33c + "%"); //Repeated.
						this.FindAndReplace(wordApp, "«Q33d»", objReport.lstInput[i].colQ33d == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ33d)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«Q33e»", objReport.lstInput[i].colQ33e == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ33e)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«Q34»", objReport.lstInput[i].colQ34);
						this.FindAndReplace(wordApp, "«Q35»", objReport.lstInput[i].colQ35);
						this.FindAndReplace(wordApp, "«Q36»", objReport.lstInput[i].colQ36 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ36)).ToString("#,0.##"));
						// this.FindAndReplace(wordApp, "«Q37»", objReport.lstInput[i].colQ37 + "%"); //Repeated.
						this.FindAndReplace(wordApp, "«Q38»", objReport.lstInput[i].colQ38);
						//58.About Your Contact Lens Dispensing                                 
						//this.FindAndReplace(wordApp, "«Q39a»", objReport.lstInput[i].colQ39a + "%"); //Repeated.
						this.FindAndReplace(wordApp, "«Q39b»", objReport.lstInput[i].colQ39b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ39b)).ToString("#,0.##") + "%");
						//this.FindAndReplace(wordApp, "«Q39c»", objReport.lstInput[i].colQ39c + "%"); //Repeated.
						this.FindAndReplace(wordApp, "«Q39d»", objReport.lstInput[i].colQ39d == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ39d)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«Q39e»", objReport.lstInput[i].colQ39e == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ39e)).ToString("#,0.##") + "%");
						//59.                                 
						this.FindAndReplace(wordApp, "«Q40a»", objReport.lstInput[i].colQ40a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ40a)).ToString("#,0.##") + "%");
						//this.FindAndReplace(wordApp, "«Q40b»", objReport.lstInput[i].colQ40b + "%"); //Repeated.
						this.FindAndReplace(wordApp, "«Q40c»", objReport.lstInput[i].colQ40c == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ40c)).ToString("#,0.##") + "%");
						//this.FindAndReplace(wordApp, "«Q40d»", objReport.lstInput[i].colQ40d + "%"); //Repeated.
						this.FindAndReplace(wordApp, "«Q40e»", objReport.lstInput[i].colQ40e == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ40e)).ToString("#,0.##") + "%");
						//this.FindAndReplace(wordApp, "«Q41a»", objReport.lstInput[i].colQ41a + "%"); //Repeated.
						//this.FindAndReplace(wordApp, "«Q42a»", objReport.lstInput[i].colQ42a.ToString("#,0.##")); //Repeated.
						//this.FindAndReplace(wordApp, "«Q43a»", objReport.lstInput[i].colQ43a + "%"); //Repeated.
						this.FindAndReplace(wordApp, "«Q44»", objReport.lstInput[i].colQ44 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ44)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q45a»", objReport.lstInput[i].colQ45a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ45a)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«Q45b»", objReport.lstInput[i].colQ45b == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ45b)).ToString("#,0.##") + "%");
						this.FindAndReplace(wordApp, "«Q46a»", objReport.lstInput[i].colQ46a == null ? "" : Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ46a)).ToString("#,0.##").ToString());
						//59.About Your Professional Fees                                
						//this.FindAndReplace(wordApp, "«Q47»", "$" + objReport.lstInput[i].colQ47.ToString("#,0.##"));//Repeated
						// this.FindAndReplace(wordApp, "«Q48»", "$" + objReport.lstInput[i].colQ48.ToString("#,0.##"));//Repeated
						//this.FindAndReplace(wordApp, "«Q49»", "$" + objReport.lstInput[i].colQ49.ToString("#,0.##"));//Repeated
						//60      
						//this.FindAndReplace(wordApp, "«Q50»", "$" + objReport.lstInput[i].colQ50.ToString("#,0.##"));//Repeated
						//this.FindAndReplace(wordApp, "«Q51»", "$" + objReport.lstInput[i].colQ51.ToString("#,0.##"));//Repeated
						//About Your Practice Expenses                            
						this.FindAndReplace(wordApp, "«Q52a»", objReport.lstInput[i].colQ52a == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ52a)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q52b»", objReport.lstInput[i].colQ52b == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ52b)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q52c»", objReport.lstInput[i].colQ52c == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ52c)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q52d»", objReport.lstInput[i].colQ52d == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ52d)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q52e»", objReport.lstInput[i].colQ52e == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ52e)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q52f»", "$" + objReport.lstInput[i].colQ52f.ToString("#,0.##")); //Repeated.
						this.FindAndReplace(wordApp, "«Q52k»", objReport.lstInput[i].colQ52k == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ52k)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q52h»", objReport.lstInput[i].colQ52h == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ52h)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q52i»", objReport.lstInput[i].colQ52i == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ52i)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q52j»", "$" + objReport.lstInput[i].colQ52j.ToString("#,0.##"));//Repeated
						//this.FindAndReplace(wordApp, "«Q53»", "$" + objReport.lstInput[i].colQ53.ToString("#,0.##"));//Repeated.
						//this.FindAndReplace(wordApp, "«Q54»", "$" + objReport.lstInput[i].colQ54.ToString("#,0.##"));//Repeated
						this.FindAndReplace(wordApp, "«Q55»", objReport.lstInput[i].colQ55 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ55)).ToString("#,0.##"));
						// this.FindAndReplace(wordApp, "«Q56»", "$" + objReport.lstInput[i].colQ56.ToString("#,0.##"));//Repeated
						this.FindAndReplace(wordApp, "«Q57»", objReport.lstInput[i].colQ57 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ57)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q58»", objReport.lstInput[i].colQ58 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ58)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q59»", objReport.lstInput[i].colQ59 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ59)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q60»", objReport.lstInput[i].colQ60 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ60)).ToString("#,0.##"));
						this.FindAndReplace(wordApp, "«Q61»", objReport.lstInput[i].colQ61 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ61)).ToString("#,0.##"));
						//61.  
						this.FindAndReplace(wordApp, "«Q62»", objReport.lstInput[i].colQ62 == null ? "" : "$" + Math.Round(Convert.ToDecimal(objReport.lstInput[i].colQ62)).ToString("#,0.##"));
						//this.FindAndReplace(wordApp, "«Q63»", "$" + objReport.lstInput[i].colQ63.ToString("#,0.##"));//Repeated
						this.FindAndReplace(wordApp, "«Q67a»", objReport.lstInput[i].colQ67a);
						this.FindAndReplace(wordApp, "«Q67b»", objReport.lstInput[i].colQ67b);
						this.FindAndReplace(wordApp, "«Q67c»", objReport.lstInput[i].colQ67c);
						//61.Classification                              
						this.FindAndReplace(wordApp, "«n68»", objReport.lstOutput[i].coln68);
						this.FindAndReplace(wordApp, "«Q69a»", objReport.lstInput[i].colQ69a);
						this.FindAndReplace(wordApp, "«Q70a»", objReport.lstInput[i].colQ70a);
						this.FindAndReplace(wordApp, "«Q71a»", objReport.lstInput[i].colQ71a);
						DateTime dt = DateTime.Now;
						string day = dt.Day.ToString();
						string month = dt.ToString("MMMM", CultureInfo.InvariantCulture);
						string year = dt.Year.ToString();
						this.FindAndReplace(wordApp, "<<dd>>", day);
						this.FindAndReplace(wordApp, "<<mm>>", month);
						this.FindAndReplace(wordApp, "<<yyyy>>", year);
						#endregion Replace Word Documnet Tempalte's content.

						//  Call sorting for each change
						if (ds.Tables[0] != null)
						{
							System.Data.DataTable dt1 = ds.Tables[0];
							int i1 = 1;
							foreach (DataRow dr in dt1.Rows)
							{

								//  this.FindAndReplace(wordApp, "« P" + i1 + "»", dr["DisplayText"]);
								// this.FindAndReplace(wordApp, "« PV" + i1 + "»", dr["DisplayValue"]);
								string name = "P" + i1;
								string name1 = "PV" + i1;
								if (dr["DisplayText"].ToString() == "Practice Productivity Metrics Average Percentile Ranking")
								{
									aDoc.Bookmarks[name1].Range.Text = ReturnAvg(dr["DisplayValue"].ToString());
									aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


								}
								else
								{
									aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									//aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;

									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									//aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;



								}




								i1++;
							}

							//aDoc.Bookmarks["P1"].Range.Words[1].Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;

							//aDoc.Bookmarks["P1"].Range.Text = "TESTING";

							//aDoc.Bookmarks["PV1"].Range.Text = "10.5";
							//aDoc.Bookmarks["PV1"].Range.Words[1].Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;

						}

						if (ds.Tables[1] != null)
						{
							System.Data.DataTable dt1 = ds.Tables[1];
							int i1 = 1;
							foreach (DataRow dr in dt1.Rows)
							{

								//this.FindAndReplace(wordApp, "« E" + i1 + "»", dr["DisplayText"]);
								//this.FindAndReplace(wordApp, "« EV" + i1 + "»", dr["DisplayValue"]);
								string name = "E" + i1;
								string name1 = "EV" + i1;
								if (dr["DisplayText"].ToString() == "Eyewear Metrics Average Percentile Ranking")
								{
									aDoc.Bookmarks[name1].Range.Text = ReturnAvg(dr["DisplayValue"].ToString());
									aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


								}
								else
								{
									if (dr["DisplayValue"] != null)
									{
										aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									}

									else
									{
										aDoc.Bookmarks[name1].Range.Text = "0";

									}
									// aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									//aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;

									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									//aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;



								}
								i1++;
							}

						}

						if (ds.Tables[2] != null)
						{
							System.Data.DataTable dt1 = ds.Tables[2];
							int i1 = 1;
							foreach (DataRow dr in dt1.Rows)
							{

								//this.FindAndReplace(wordApp, "« C" + i1 + "»", dr["DisplayText"]);
								//this.FindAndReplace(wordApp, "« CV" + i1 + "»", dr["DisplayValue"]);
								string name = "C" + i1;
								string name1 = "CV" + i1;
								if (dr["DisplayText"].ToString() == "Contact Lens Metrics Average Percentile")
								{
									aDoc.Bookmarks[name1].Range.Text = ReturnAvg(dr["DisplayValue"].ToString());
									aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


								}
								else
								{
									if (dr["DisplayValue"] != null)
									{
										aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									}

									else
									{
										aDoc.Bookmarks[name1].Range.Text = "0";

									}
									// aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									//aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;

									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									//aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;



								}
								i1++;
							}

						}
						if (ds.Tables[3] != null)
						{
							System.Data.DataTable dt1 = ds.Tables[3];
							int i1 = 1;
							foreach (DataRow dr in dt1.Rows)
							{

								//   this.FindAndReplace(wordApp, "« M" + i1 + "»", dr["DisplayText"]);
								//  this.FindAndReplace(wordApp, "« MV" + i1 + "»", dr["DisplayValue"]);
								string name = "M" + i1;
								string name1 = "MV" + i1;
								if (dr["DisplayText"].ToString() == "Medical Eye Care Metrics Average Percentile Ranking")
								{
									aDoc.Bookmarks[name1].Range.Text = ReturnAvg(dr["DisplayValue"].ToString());
									aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


								}
								else
								{
									if (dr["DisplayValue"] != null)
									{
										aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									}

									else
									{
										aDoc.Bookmarks[name1].Range.Text = "0";

									}
									// aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									//aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;

									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									//aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;



								}
								i1++;
							}

						}

						if (ds.Tables[4] != null)
						{
							System.Data.DataTable dt1 = ds.Tables[4];
							int i1 = 1;
							foreach (DataRow dr in dt1.Rows)
							{

								// this.FindAndReplace(wordApp, "« MM" + i1 + "»", dr["DisplayText"]);
								//  this.FindAndReplace(wordApp, "« MMV" + i1 + "»", dr["DisplayValue"]);
								string name = "MM" + i1;
								string name1 = "MMV" + i1;
								if (dr["DisplayText"].ToString() == "Marketing Average Percentile Ranking")
								{
									aDoc.Bookmarks[name1].Range.Text = ReturnAvg(dr["DisplayValue"].ToString());
									aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


								}
								else
								{
									if (dr["DisplayValue"] != null)
									{
										aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									}

									else
									{
										aDoc.Bookmarks[name1].Range.Text = "0";

									}
									// aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									//aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;

									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									//aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;



								}
								i1++;
							}

						}
						if (ds.Tables[5] != null)
						{
							System.Data.DataTable dt1 = ds.Tables[5];
							int i1 = 1;
							foreach (DataRow dr in dt1.Rows)
							{

								// this.FindAndReplace(wordApp, "« F" + i1 + "»", dr["DisplayText"]);
								// this.FindAndReplace(wordApp, "« FV" + i1 + "»", dr["DisplayValue"]);
								string name = "F" + i1;
								string name1 = "FV" + i1;
								if (dr["DisplayText"].ToString() == "Financial Average Percentile Ranking")
								{
									aDoc.Bookmarks[name1].Range.Text = ReturnAvg(dr["DisplayValue"].ToString());
									aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


								}
								else
								{
									if (dr["DisplayValue"] != null)
									{
										aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									}

									else
									{
										aDoc.Bookmarks[name1].Range.Text = "0";

									}
									// aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									//aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;

									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									//aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;



								}
								i1++;
							}

						}

						if (ds.Tables[6] != null)
						{
							System.Data.DataTable dt1 = ds.Tables[6];
							int i1 = 1;
							foreach (DataRow dr in dt1.Rows)
							{

								// this.FindAndReplace(wordApp, "« B" + i1 + "»", dr["DisplayText"]);
								// this.FindAndReplace(wordApp, "« BV" + i1 + "»", dr["DisplayValue"]);
								string name = "B" + i1;
								string name1 = "BV" + i1;
								if (dr["DisplayText"].ToString() == "Average Percentile Ranking")
								{
									aDoc.Bookmarks[name1].Range.Text = ReturnAvg(dr["DisplayValue"].ToString());
									aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlueGray;


								}
								else
								{
									if (dr["DisplayValue"] != null)
									{
										aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									}

									else
									{
										aDoc.Bookmarks[name1].Range.Text = "0";

									}
									//aDoc.Bookmarks[name1].Range.Text = dr["DisplayValue"].ToString();
									//aDoc.Bookmarks[name1].Range.Words.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;

									aDoc.Bookmarks[name].Range.Text = dr["DisplayText"].ToString();
									//aDoc.Bookmarks[name].Range.Sentences.First.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;



								}
								i1++;
							}

						}

						if (dtsort != null)
						{
							int i1 = 1;
							foreach (DataRow dr in dtsort.Rows)
							{

								this.FindAndReplace(wordApp, "« A" + i1 + "»", dr["DisplayText"]);

								if (dr["DisplayValue"] != null)
								{
									if (!string.IsNullOrEmpty(dr["DisplayValue"].ToString()))
									{

										this.FindAndReplace(wordApp, "« AV" + i1 + "»", dr["DisplayValue"]);
									}
									else
									{
										this.FindAndReplace(wordApp, "« AV" + i1 + "»", "0");

									}
								}

								else
								{
									this.FindAndReplace(wordApp, "« AV" + i1 + "»", "0");

								}

								i1++;
							}
						}

						DateTime dtw = (DateTime)objReport.lstInput[i].colIDendDate;
						if (System.Web.HttpContext.Current.Session["YearName"] != null)
						{
							string activeyear = (Convert.ToInt32(System.Web.HttpContext.Current.Session["YearName"]) - 1).ToString();
							string previousyear = (Convert.ToInt32(activeyear) - 1).ToString();

							this.FindAndReplace(wordApp, "<a>", activeyear);
							this.FindAndReplace(wordApp, "<p>", previousyear);

						}

						else
						{
							string year11 = dtw.Year.ToString();
							string activeyear = (Convert.ToInt32(year11) - 1).ToString();
							string previousyear = (Convert.ToInt32(activeyear) - 1).ToString();
							this.FindAndReplace(wordApp, "<a>", activeyear);
							this.FindAndReplace(wordApp, "<p>", previousyear);

						}
						//  Save temp.doc after modified
						aDoc.Save();

						/*--Shahbaz.Need to enable when we go with Word to Pdf Report.*/
						//  Save document into PDF Format
						aDoc.SaveAs(ref outputFileName,
						ref fileFormat, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing);

						((_Application)wordApp).Quit(SaveChanges, ref missing, ref missing);
						// wordApp.Quit(SaveChanges, ref missing, ref missing);

						//wordApp.Quit(ref missing, ref missing, ref missing);

						/*Muntajib-Remove to below if block to For both word & pdf*/

						//systesession["filepath"] = outputFileName;
						System.Threading.Thread.Sleep(5000);
						if (File.Exists(tempPath))
						{
							File.Delete(tempPath);
						}

					}
					else
						return "File does not exist, please check and retry.";
				}
				string updatedStatus = UpdateReportGenerateStatus(objReport.lstOutput);
				if (updatedStatus == "success")
				{
					//var milliseconds = stopwatch.ElapsedMilliseconds;
					return "success";
				}
				else
				{
					if (ConfigurationManager.AppSettings["ErrorFilePath"] != null)
					{
						string filePath = ConfigurationManager.AppSettings["ErrorFilePath"].ToString();    //"C:\\pics\\";//@"C:\Error.txt";

						using (StreamWriter writer = new StreamWriter(filePath, true))
						{
							writer.WriteLine("Detailed1 :" + updatedStatus);
							writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
						}
					}
					return updatedStatus;
				}

			}
			catch (Exception ex)
			{
				if (ConfigurationManager.AppSettings["ErrorFilePath"] != null)
				{
					string filePath = ConfigurationManager.AppSettings["ErrorFilePath"].ToString();    //"C:\\pics\\";//@"C:\Error.txt";

					using (StreamWriter writer = new StreamWriter(filePath, true))
					{
						writer.WriteLine("Detailed2 :" + ex.Message);
						writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
					}
				}
				return ex.Message;
			}
		}
		public string GenerateInfographicWordReport(string imagelocation, string strSourcePath, string strTargetPath, Report objReport, string avg, string path, System.Data.DataTable dtsort)
		{
			try
			{

				if (IsFileInUse(strSourcePath))
				{
					return "File already in use, please close and retry.";
				}
				//Only run the reports that are not ran earlier.
				// List<int> reportToBeGenerated = db.Source_InputData.Where(r => r.RowId > lastReportGenerated && r.IDformat == "Y").Select(r => r.RowId).ToList();

				List<int> reportToBeGenerated = db.Source_InputData.Where(r => r.RowId > lastReportGenerated && r.IDformat == "Y").Select(r => r.RowId).ToList();

				if (reportToBeGenerated.Count() < NoOfDocumentsLimit)
				{
					NoOfDocumentsLimit = 1; // reportToBeGenerated.Count();
				}

				//for (int i = 0; i < NoOfDocumentsLimit; i++)
				//{

				for (int i = reportToBeGenerated.Count - 1; i <= reportToBeGenerated.Count - 1; i++)
				{

					DateTime CurrentDateTime = DateTime.Now;
					string filenamestr = "Infographic_" + objReport.lstInput[0].colQ73 + "_" + objReport.lstInput[0].colIDname + "_" + CurrentDateTime.ToString("MMddyyyy-hhmmss");

					System.Web.HttpContext.Current.Session["varNameinfographic"] = filenamestr;

					/*Muntajib-Remove tempPath For both word & pdf*/
					string tempPath = AppDomain.CurrentDomain.BaseDirectory + filenamestr + ".doc";
					filenamestr = strTargetPath + "\\" + filenamestr + ".doc";

					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					FileStream fs = new FileStream(tempPath, FileMode.Create, FileAccess.ReadWrite);
					fs.Close();
					//  Just to kill WINWORD.EXE if it is running
					//  killprocess("winword");     
					//  copy letter format to temp.doc

					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					File.Copy(strSourcePath, tempPath, true);

					//  create missing object
					object missing = Missing.Value;
					//  create Word application object
					Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
					//  create Word document object
					Microsoft.Office.Interop.Word.Document aDoc = null;
					//  create & define filename object with temp.doc
					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					object filename = tempPath;
					if (File.Exists((string)filename))
					{
						/*//--Shahbaz.Need to enable when we go with Word to Pdf Report.*/
						object SaveChanges = false;

						object readOnly = false;
						object isVisible = false;
						//  make visible Word application
						wordApp.Visible = false;
						//  open Word document named temp.doc
						aDoc = wordApp.Documents.Open(ref filename, ref missing,
					   ref readOnly, ref missing, ref missing, ref missing,
					   ref missing, ref missing, ref missing, ref missing,
					   ref missing, ref isVisible, ref missing, ref missing,
					   ref missing, ref missing);
						aDoc.Activate();

						/*//--Shweta.Need to enable this code,when we go with Word to Pdf Report.
                       
                        //--Shweta.If we want to put it as web app,we just need to install "SaveAsPDFandXPS.exe" on the server.*/

						object outputFileName = filenamestr.Replace(".doc", ".pdf");
						object fileFormat = WdSaveFormat.wdFormatPDF;

						//  Call FindAndReplace()function for each change

						#region Replace Word Documnet Tempalte's content.


						this.FindAndReplace(wordApp, "ag", GetOrdinal(Convert.ToInt32(avg)));


						if (objReport.lstInput[0].colQ74 != null)
						{
							this.FindAndReplace(wordApp, "<<name>>", objReport.lstInput[0].colQ74);
						}

						else
						{
							this.FindAndReplace(wordApp, "<<name>>", "");
						}


						string pictureName = imagelocation + @"\Chart.png";
						string picName1 = imagelocation + @"\yellow.png";  // @"C:\yellow.png";
						string pictureName1 = imagelocation + @"\1.png";
						string pictureName2 = imagelocation + @"\2.png"; //@"C:\yellowblue14.png";
						string pictureName3 = imagelocation + @"\3.png"; //@"C:\yellowblue23.png";
						string pictureName4 = imagelocation + @"\4.png";// @"C:\yellowblue32.png";
						string pictureName5 = imagelocation + @"\5.png";// @"C:\yellowblue41.png";

						string pictureName6 = imagelocation + @"\6.png";
						string pictureName7 = imagelocation + @"\7.png";
						string pictureName8 = imagelocation + @"\8.png";
						string pictureName9 = imagelocation + @"\9.png";
						string picName6 = imagelocation + @"\blue.png";   // @"C:\blue.png";//imagelocation + "/Chart.png";    // @"C:\Visual Studio 2010\Projects\SurveyApp\SurveyApp\ImageChart\Chart.png";
						int count = aDoc.Bookmarks.Count;
						Bookmarks bs;   // aDoc.Bookmarks;
						List<string> lst = new List<string>();
						lst.Add("chart");
						lst.Add("img1");
						lst.Add("img2");
						lst.Add("img3");
						lst.Add("img4");
						lst.Add("img5");
						lst.Add("img6");
						lst.Add("img7");
						lst.Add("img8");
						lst.Add("img9");
						lst.Add("img91");


						bool flag = false;
						for (int i1 = 1; i1 < lst.Count + 1; i1++)
						{


							//if (lst.Contains(aDoc.Bookmarks[i1].Name))
							//{
							object oRange = aDoc.Bookmarks[lst[i1 - 1]].Range;
							object saveWithDocument = true;
							object missing1 = Type.Missing;
							// string pictureName = @"C:\Documents and Settings\shweta.singh\My Documents\visual studio 2010\Projects\WebApplication3\WebApplication3\Pics\ChartImg.png";
							//aDoc.InlineShapes.AddPicture(pictureName, ref missing1, ref saveWithDocument, ref oRange);
							if (i1 == 1)
							{
								var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName, false, true);
								shape.Width = 650;
								shape.Height = 260;
							}

							else
							{

								string x = string.Empty;
								string x1 = string.Empty;
								if (avg.Length > 1)
								{
									x = avg.Substring(0, 1);
									x1 = avg.Substring(1, 1);

									if (i1 <= Convert.ToInt32(x) + 1)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(picName1, false, true);
										shape.Width = 40;
										shape.Height = 40;
										// flag = true;
									}
									else if (x1.ToString() == "1" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName1, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (x1.ToString() == "2" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName2, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (x1.ToString() == "3" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName3, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (x1.ToString() == "4" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName4, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}

									else if (x1.ToString() == "5" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName5, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (x1.ToString() == "6" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName6, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (x1.ToString() == "7" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName7, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (x1.ToString() == "8" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName8, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (x1.ToString() == "9" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName9, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}

									else if (x1.ToString() == "0" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(picName6, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (i1 > Convert.ToInt32(x) + 1 && flag == true)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(picName6, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}


								}

								else if (avg.Length <= 1)
								{

									if (avg.ToString() == "1" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName1, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (avg.ToString() == "2" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName2, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}

									else if (avg.ToString() == "3" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName3, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (avg.ToString() == "4" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName4, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (avg.ToString() == "5" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName5, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (avg.ToString() == "6" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName6, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (avg.ToString() == "7" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName7, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (avg.ToString() == "8" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName8, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									else if (avg.ToString() == "9" && flag == false)
									{
										var shape = aDoc.Bookmarks[lst[i1 - 1]].Range.InlineShapes.AddPicture(pictureName9, false, true);
										shape.Width = 40;
										shape.Height = 40;
										flag = true;
									}
									//else if (i1 > Convert.ToInt32(avg) + 1 && flag == true)
									//{
									//    var shape = aDoc.Bookmarks[lst[i1-1]].Range.InlineShapes.AddPicture(pictureName6, false, true);
									//    shape.Width = 40;
									//    shape.Height = 40;
									//    flag = true;
									//}

								}
								// }
							}
						}



						#endregion Replace Word Documnet Tempalte's content.

						//  Save temp.doc after modified




						string picturegreen = path + @"\greenbox.png";
						string expression = "SortedData >= '90' ";
						string sortOrder = "SortedData DESC";
						// DataRow[] foundRows;
						System.Data.DataTable dt2 = new System.Data.DataTable();
						dt2.Columns.Add("SortedData", typeof(long));
						dt2.Columns.Add("DisplayText");
						dt2.Columns.Add("DisplayValue");
						dt2.Columns.Add("Pageno");
						dt2.Columns.Add("75percentile");
						dt2.Columns.Add("75percentilesorted", typeof(long));
						dtsort.Select(expression, sortOrder).CopyToDataTable(dt2, LoadOption.OverwriteChanges);
						int length = dt2.Rows.Count;




						if (length > 0)
						{
							int n = 0;

							foreach (DataRow dr in dt2.Rows)
							{
								int x = n + 1;
								string nam = "G" + x;
								string H = "H" + x;
								string picturename = imagelocation + @"\im" + x + ".png";
								// this.FindAndReplace(wordApp, "<<val" + x + "hel>>", "Read more about this metric on page no." + dtsort.Rows[n]["Pageno"] + " in the detailed report");
								var shape = aDoc.Bookmarks[nam].Range.InlineShapes.AddPicture(picturename, false, true);


								shape.Width = 80;
								shape.Height = 90;
								Bitmap bmp = new Bitmap(imagelocation + @"\greenbox.png", false);
								Graphics graphicsobj = Graphics.FromImage(bmp);
								Brush brush = new SolidBrush(Color.White);
								System.Drawing.Point postionWaterMark = new System.Drawing.Point((bmp.Width), (bmp.Height * 9 / 10));
								RectangleF rf = new RectangleF(2, 4, 201, 71);
								graphicsobj.DrawString("Read more about this metric on page no. " + dtsort.Rows[n]["Pageno"] + " in the detailed report", new System.Drawing.Font("Arial", 15, FontStyle.Regular, GraphicsUnit.Pixel), brush, rf);
								string filepath = imagelocation + @"\picturegreen1.png";
								bmp.Save(filepath);
								var shape1 = aDoc.Bookmarks[H].Range.InlineShapes.AddPicture(filepath, false, true);
								this.FindAndReplace(wordApp, "<G" + x + ">", dtsort.Rows[n]["DisplayText"]);

								this.FindAndReplace(wordApp, "<<val" + x + ">>", dtsort.Rows[n]["DisplayValue"]);


								n++;

								if (n > 5)
								{
									break;
								}
							}
							if (n < 6)
							{
								for (int j = n + 1; j <= 6; j++)
								{


									this.FindAndReplace(wordApp, "<<val" + j + ">>", "");
									this.FindAndReplace(wordApp, "<G" + j + ">", "");
								}
							}

						}

						else
						{
							if (length == 0)
							{
								for (int h = 1; h <= 6; h++)
								{
									this.FindAndReplace(wordApp, "<<val" + h + ">>", "");
									this.FindAndReplace(wordApp, "<G" + h + ">", "");
								}
							}
							else
							{

								for (int j = length + 1; j <= 6; j++)
								{


									this.FindAndReplace(wordApp, "<<val" + j + ">>", "");
									this.FindAndReplace(wordApp, "<G" + j + ">", "");
								}
							}

						}

						string expression1 = "75percentilesorted >= '20000' ";
						string sortOrder1 = "75percentilesorted DESC";
						System.Data.DataTable dt75 = new System.Data.DataTable();
						dt75.Columns.Add("SortedData", typeof(long));
						dt75.Columns.Add("DisplayText");
						dt75.Columns.Add("DisplayValue");
						dt75.Columns.Add("Pageno");
						dt75.Columns.Add("75percentile");
						dt75.Columns.Add("75percentilesorted", typeof(long));

						DataView _dtv1 = dtsort.DefaultView;
						_dtv1.Sort = "75percentilesorted DESC";

						dtsort = _dtv1.ToTable();


						// dtsort.Select(expression1, sortOrder1).CopyToDataTable(dt75, LoadOption.OverwriteChanges);

						if (dtsort.Rows.Count > 0)
						{
							int n = 0;
							int x = 0;
							foreach (DataRow dr in dtsort.Rows)
							{
								if (dr["75percentilesorted"] != null && !string.IsNullOrEmpty(dr["75percentilesorted"].ToString()))
								{
									if (Convert.ToInt64(dr["75percentilesorted"]) >= 20000)
									{
										x = n + 1;
										if (x > 6)
										{
											break;
										}
										string nam = "M" + x;
										string H = "V" + x;
										string picturename = imagelocation + @"\img" + x + ".png";
										//  this.FindAndReplace(wordApp, "<<val" + x + "help>>", "Read more about this metric on page no." + dtsort.Rows[n]["Pageno"] + " in the detailed report");
										var shape = aDoc.Bookmarks[nam].Range.InlineShapes.AddPicture(picturename, false, true);
										//  var shape1 = aDoc.Bookmarks[H].Range.InlineShapes.AddPicture(picturegreen, false, true);
										shape.Width = 80;
										shape.Height = 90;
										Bitmap bmp = new Bitmap(imagelocation + @"\greenbox.png", false);
										Graphics graphicsobj = Graphics.FromImage(bmp);
										Brush brush = new SolidBrush(Color.White);
										System.Drawing.Point postionWaterMark = new System.Drawing.Point((bmp.Width), (bmp.Height * 9 / 10));
										RectangleF rf = new RectangleF(2, 4, 201, 71);
										graphicsobj.DrawString("Read more about this metric on page no. " + dtsort.Rows[n]["Pageno"] + " in the detailed report", new System.Drawing.Font("Arial", 15, FontStyle.Regular, GraphicsUnit.Pixel), brush, rf);
										string filepath = imagelocation + @"\picturegreen1.png";
										bmp.Save(filepath);
										var shape1 = aDoc.Bookmarks[H].Range.InlineShapes.AddPicture(filepath, false, true);
										this.FindAndReplace(wordApp, "<M" + x + ">", dtsort.Rows[n]["DisplayText"]);
										if (dtsort.Rows[n]["75percentile"] != null && !string.IsNullOrEmpty(dtsort.Rows[n]["75percentile"].ToString()))
										{
											if (dtsort.Rows[n]["75percentile"].ToString().Contains('.'))
											{
												string text = dtsort.Rows[n]["75percentile"].ToString().Split('.')[0];
												this.FindAndReplace(wordApp, "<<val" + x + "desc>>", text);
											}

											else
											{
												this.FindAndReplace(wordApp, "<<val" + x + "desc>>", dtsort.Rows[n]["75percentile"]);
											}
										}

										n++;
									}

								}

							}

							if (x < 6)
							{
								for (int y = x + 1; y <= 6; y++)
								{
									this.FindAndReplace(wordApp, "<M" + y + ">", "");

									this.FindAndReplace(wordApp, "<<val" + y + "desc>>", "");
								}

							}

						}
						aDoc.Save();



						/*--Shahbaz.Need to enable when we go with Word to Pdf Report.*/
						//  Save document into PDF Format
						aDoc.SaveAs(ref outputFileName,
						ref fileFormat, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing);
						// wordApp.Quit(SaveChanges, ref missing, ref missing);

						// wordApp.Quit(ref missing, ref missing, ref missing);

						/*Muntajib-Remove to below if block to For both word & pdf*/

						//systesession["filepath"] = outputFileName;
						((_Application)wordApp).Quit(SaveChanges, ref missing, ref missing);

						//((_Application)wordApp).Quit(ref missing, ref missing, ref missing);

						/*Muntajib-Remove to below if block to For both word & pdf*/

						System.Threading.Thread.Sleep(5000);
						if (File.Exists(tempPath))
						{
							File.Delete(tempPath);
						}

					}
					else
						return "File does not exist, please check and retry.";
				}
				string updatedStatus = UpdateReportGenerateStatus(objReport.lstOutput);
				if (updatedStatus == "success")
				{
					//var milliseconds = stopwatch.ElapsedMilliseconds;
					return "success";
				}
				else
				{
					return updatedStatus;
				}

			}
			catch (Exception ex)
			{

				// string filePath = @"C:\Error.txt";
				string filePath = null;
				if (ConfigurationManager.AppSettings["ErrorFilePath"] != null)
				{
					filePath = ConfigurationManager.AppSettings["ErrorFilePath"].ToString();// @"C:\Error.txt";
				}
				if (filePath != null)
				{

					using (StreamWriter writer = new StreamWriter(filePath, true))
					{
						writer.WriteLine("Messageinfographic :" + ex.Message + " stacktrace" + ex.StackTrace);
						writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
					}
				}
				return ex.Message;
			}
		}
		public string GenerateExecutiveWordReport(string imagelocation, string strSourcePath, string strTargetPath, Report objReport, string practiceid, System.Data.DataTable dtsort)
		{
			try
			{


				if (IsFileInUse(strSourcePath))
				{
					return "File already in use, please close and retry.";
				}
				//Only run the reports that are not ran earlier.
				// List<int> reportToBeGenerated = db.Source_InputData.Where(r => r.RowId > lastReportGenerated && r.IDformat == "Y").Select(r => r.RowId).ToList();

				List<int> reportToBeGenerated = db.Source_InputData.Where(r => r.RowId > lastReportGenerated && r.IDformat == "Y").Select(r => r.RowId).ToList();

				if (reportToBeGenerated.Count() < NoOfDocumentsLimit)
				{
					NoOfDocumentsLimit = 1; // reportToBeGenerated.Count();
				}

				//for (int i = 0; i < NoOfDocumentsLimit; i++)
				//{


				for (int i = reportToBeGenerated.Count - 1; i <= reportToBeGenerated.Count - 1; i++)
				{
					DateTime CurrentDateTime = DateTime.Now;
					string filenamestr = "Executive" + objReport.lstInput[0].colQ73 + "_" + objReport.lstInput[0].colIDname + "_" + CurrentDateTime.ToString("MMddyyyy-hhmmss");




					System.Web.HttpContext.Current.Session["varNameexecutive"] = filenamestr;

					/*Muntajib-Remove tempPath For both word & pdf*/
					string tempPath = AppDomain.CurrentDomain.BaseDirectory + filenamestr + ".doc";
					filenamestr = strTargetPath + "\\" + filenamestr + ".doc";

					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					FileStream fs = new FileStream(tempPath, FileMode.Create, FileAccess.ReadWrite);
					fs.Close();
					//  Just to kill WINWORD.EXE if it is running
					//  killprocess("winword");     
					//  copy letter format to temp.doc

					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					File.Copy(strSourcePath, tempPath, true);

					//  create missing object
					object missing = Missing.Value;
					//  create Word application object
					Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
					//  create Word document object
					Microsoft.Office.Interop.Word.Document aDoc = null;
					//  create & define filename object with temp.doc
					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					object filename = tempPath;
					if (File.Exists((string)filename))
					{
						/*//--Shahbaz.Need to enable when we go with Word to Pdf Report.*/
						object SaveChanges = false;

						object readOnly = false;
						object isVisible = false;
						//  make visible Word application
						wordApp.Visible = false;
						//  open Word document named temp.doc
						aDoc = wordApp.Documents.Open(ref filename, ref missing,
					   ref readOnly, ref missing, ref missing, ref missing,
					   ref missing, ref missing, ref missing, ref missing,
					   ref missing, ref isVisible, ref missing, ref missing,
					   ref missing, ref missing);
						aDoc.Activate();

						/*//--Shweta.Need to enable this code,when we go with Word to Pdf Report.
                       
                        //--Shweta.If we want to put it as web app,we just need to install "SaveAsPDFandXPS.exe" on the server.*/

						object outputFileName = filenamestr.Replace(".doc", ".pdf");
						object fileFormat = WdSaveFormat.wdFormatPDF;

						//  Call FindAndReplace()function for each change

						#region Replace Word Documnet Tempalte's content.


						Microsoft.Office.Interop.Word.Find fnd = wordApp.ActiveWindow.Selection.Find;

						fnd.ClearFormatting();
						fnd.Replacement.ClearFormatting();
						fnd.Forward = true;
						fnd.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;

						fnd.Text = "<<name>>";
						fnd.Replacement.Text = objReport.lstInput[0].colQ74;

						fnd.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);


						//if (objReport.lstInput[i].colQ74 != null)
						//{
						//    this.FindAndReplace(wordApp, "<<name>>", objReport.lstInput[i].colQ74);
						//}

						//else
						//{
						//    this.FindAndReplace(wordApp, "<<name>>", "");
						//}





						string pictureName = string.Empty;  // @"C:\Visual Studio 2010\Projects\SurveyApp\SurveyApp\ImageChart\Chart.png";
						int count = aDoc.Bookmarks.Count;
						for (int i1 = 1; i1 < count + 1; i1++)
						{
							object oRange = aDoc.Bookmarks[i1].Range;
							object saveWithDocument = true;
							object missing1 = Type.Missing;
							if (i1 == 1)
							{

								pictureName = imagelocation + @"\Chart.png";  // @"C:\Visual Studio 2010\Projects\SurveyApp\SurveyApp\ImageChart\Chart.png";

								var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
								shape.Width = 650;
								shape.Height = 220;


							}
							if (i1 == 2)
							{
								pictureName = imagelocation + @"\chart1.png";   // @"C:\Visual Studio 2010\Projects\SurveyApp\SurveyApp\ImageChart\chart1.png";

								var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
								shape.Width = 500;
								shape.Height = 240;

							}
							if (i1 == 3)
							{
								pictureName = imagelocation + @"\chart2.png";  // @"C:\Visual Studio 2010\Projects\SurveyApp\SurveyApp\ImageChart\chart2.png";
								var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
								shape.Width = 500;
								shape.Height = 297;

							}
							if (i1 == 4)
							{
								pictureName = imagelocation + @"\chart3.png";  // @"C:\Visual Studio 2010\Projects\SurveyApp\SurveyApp\ImageChart\chart3.png";
								var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
								shape.Width = 500;
								shape.Height = 351;

							}
							if (i1 == 5)
							{
								pictureName = imagelocation + @"\chart4.png"; // @"C:\Visual Studio 2010\Projects\SurveyApp\SurveyApp\ImageChart\chart4.png";

								var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
								shape.Width = 500;
								shape.Height = 135;

							}
							if (i1 == 6)
							{
								pictureName = imagelocation + @"\chart5.png";  // @"C:\Visual Studio 2010\Projects\SurveyApp\SurveyApp\ImageChart\chart5.png";
								var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
								shape.Width = 500;
								shape.Height = 189;

							}
							if (i1 == 7)
							{
								pictureName = imagelocation + @"\chart6.png"; // @"C:\Visual Studio 2010\Projects\SurveyApp\SurveyApp\ImageChart\chart6.png";
								var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
								shape.Width = 500;
								shape.Height = 378;

							}
							if (i1 == 8)
							{
								pictureName = imagelocation + @"\chart7.png";  // @"C:\Visual Studio 2010\Projects\SurveyApp\SurveyApp\ImageChart\chart7.png";
								var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
								shape.Width = 500;
								shape.Height = 135;

							}

						}

						#endregion Replace Word Documnet Tempalte's content.




						string picturegreen = imagelocation + @"\greenboxe.png";
						string expression = "SortedData >= '90' ";
						string sortOrder = "SortedData DESC";
						// DataRow[] foundRows;
						System.Data.DataTable dt2 = new System.Data.DataTable();
						dt2.Columns.Add("SortedData", typeof(long));
						dt2.Columns.Add("DisplayText");
						dt2.Columns.Add("DisplayValue");
						dt2.Columns.Add("Pageno");
						dt2.Columns.Add("75percentile");
						dt2.Columns.Add("75percentilesorted", typeof(long));
						dtsort.Select(expression, sortOrder).CopyToDataTable(dt2, LoadOption.OverwriteChanges);

						int length = dt2.Rows.Count;


						if (length > 0)
						{
							int n = 0;

							foreach (DataRow dr in dt2.Rows)
							{
								int x = n + 1;
								string nam = "G" + x;
								string H = "H" + x;
								string picturename = imagelocation + @"\ime" + x + ".png";
								// this.FindAndReplace(wordApp, "<<val" + x + "hel>>", "Read more about this metric on page no." + dtsort.Rows[n]["Pageno"] + " in the detailed report");
								var shape = aDoc.Bookmarks[nam].Range.InlineShapes.AddPicture(picturename, false, true);


								shape.Width = 60;
								shape.Height = 60;
								Bitmap bmp = new Bitmap(imagelocation + @"\greenboxe.png", false);
								Graphics graphicsobj = Graphics.FromImage(bmp);
								Brush brush = new SolidBrush(Color.White);
								System.Drawing.Point postionWaterMark = new System.Drawing.Point((bmp.Width), (bmp.Height * 9 / 10));
								RectangleF rf = new RectangleF(2, 2, 109, 42);
								graphicsobj.DrawString("Read more about this metric on page no. " + dtsort.Rows[n]["Pageno"] + " in the detailed report", new System.Drawing.Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Pixel), brush, rf);
								string filepath = imagelocation + @"\picturegreene1.png";
								bmp.Save(filepath);
								var shape1 = aDoc.Bookmarks[H].Range.InlineShapes.AddPicture(filepath, false, true);
								this.FindAndReplace(wordApp, "<G" + x + ">", dtsort.Rows[n]["DisplayText"]);
								if (dtsort.Rows[n]["DisplayValue"].ToString().Contains('.'))
								{
									string text = dtsort.Rows[n]["DisplayValue"].ToString().Split('.')[0];
									this.FindAndReplace(wordApp, "<<val" + x + ">>", text);
								}

								else
								{
									this.FindAndReplace(wordApp, "<<val" + x + ">>", dtsort.Rows[n]["DisplayValue"]);
								}

								n++;

								if (n > 5)
								{
									break;
								}
							}

							if (n < 6)
							{
								for (int j = n + 1; j <= 6; j++)
								{


									this.FindAndReplace(wordApp, "<<val" + j + ">>", "");
									this.FindAndReplace(wordApp, "<G" + j + ">", "");
								}
							}

						}

						else
						{
							if (length == 0)
							{
								for (int h = 1; h <= 6; h++)
								{
									this.FindAndReplace(wordApp, "<<val" + h + ">>", "");
									this.FindAndReplace(wordApp, "<G" + h + ">", "");
								}
							}
							else
							{

								for (int j = length + 1; j <= 6; j++)
								{


									this.FindAndReplace(wordApp, "<<val" + j + ">>", "");
									this.FindAndReplace(wordApp, "<G" + j + ">", "");
								}
							}
						}

						string expression1 = "75percentilesorted >= '20000' ";
						string sortOrder1 = "75percentilesorted DESC";
						System.Data.DataTable dt75 = new System.Data.DataTable();
						dt75.Columns.Add("SortedData", typeof(long));
						dt75.Columns.Add("DisplayText");
						dt75.Columns.Add("DisplayValue");
						dt75.Columns.Add("Pageno");
						dt75.Columns.Add("75percentile");
						dt75.Columns.Add("75percentilesorted", typeof(long));

						DataView _dtv1 = dtsort.DefaultView;
						_dtv1.Sort = "75percentilesorted DESC";

						dtsort = _dtv1.ToTable();


						// dtsort.Select(expression1, sortOrder1).CopyToDataTable(dt75, LoadOption.OverwriteChanges);

						if (dtsort.Rows.Count > 0)
						{
							int n = 0;
							int x = 0;
							foreach (DataRow dr in dtsort.Rows)
							{
								if (dr["75percentilesorted"] != null && !string.IsNullOrEmpty(dr["75percentilesorted"].ToString()))
								{
									if (Convert.ToInt64(dr["75percentilesorted"]) >= 20000)
									{
										x = n + 1;
										if (x > 6)
										{
											break;
										}
										string nam = "M" + x;
										string H = "V" + x;
										string picturename = imagelocation + @"\imge" + x + ".png";
										//  this.FindAndReplace(wordApp, "<<val" + x + "help>>", "Read more about this metric on page no." + dtsort.Rows[n]["Pageno"] + " in the detailed report");
										var shape = aDoc.Bookmarks[nam].Range.InlineShapes.AddPicture(picturename, false, true);
										//  var shape1 = aDoc.Bookmarks[H].Range.InlineShapes.AddPicture(picturegreen, false, true);
										shape.Width = 60;
										shape.Height = 60;
										Bitmap bmp = new Bitmap(imagelocation + @"\greenboxe.png", false);
										Graphics graphicsobj = Graphics.FromImage(bmp);
										Brush brush = new SolidBrush(Color.White);
										System.Drawing.Point postionWaterMark = new System.Drawing.Point((bmp.Width), (bmp.Height * 9 / 10));
										RectangleF rf = new RectangleF(2, 2, 109, 42);
										graphicsobj.DrawString("Read more about this metric on page no. " + dtsort.Rows[n]["Pageno"] + " in the detailed report", new System.Drawing.Font("Arial", 9, FontStyle.Regular, GraphicsUnit.Pixel), brush, rf);
										string filepath = imagelocation + @"\picturegreene1.png";
										bmp.Save(filepath);
										var shape1 = aDoc.Bookmarks[H].Range.InlineShapes.AddPicture(filepath, false, true);
										this.FindAndReplace(wordApp, "<M" + x + ">", dtsort.Rows[n]["DisplayText"]);
										this.FindAndReplace(wordApp, "<<val" + x + "desc>>", ReturnAvg(dtsort.Rows[n]["75percentile"].ToString()));

										n++;
									}
								}

							}
							if (x < 6)
							{
								for (int y = x + 1; y <= 6; y++)
								{
									this.FindAndReplace(wordApp, "<M" + y + ">", "");

									this.FindAndReplace(wordApp, "<<val" + y + "desc>>", "");
								}

							}
						}



						aDoc.Save();



						/*--Shahbaz.Need to enable when we go with Word to Pdf Report.*/
						//  Save document into PDF Format
						aDoc.SaveAs(ref outputFileName,
						ref fileFormat, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing);
						System.Threading.Thread.Sleep(5000);
						((_Application)wordApp).Quit(SaveChanges, ref missing, ref missing);

						// wordApp.Quit(ref missing, ref missing, ref missing);

						/*Muntajib-Remove to below if block to For both word & pdf*/

						//systesession["filepath"] = outputFileName;

						//((_Application)wordApp).Quit(SaveChanges, ref missing, ref missing);

						//((_Application)wordApp).Quit(ref missing, ref missing, ref missing);

						/*Muntajib-Remove to below if block to For both word & pdf*/
						System.Threading.Thread.Sleep(5000);
						if (File.Exists(tempPath))
						{
							File.Delete(tempPath);
						}

					}
					else
						return "File does not exist, please check and retry.";
				}
				string updatedStatus = UpdateReportGenerateStatus(objReport.lstOutput);
				if (updatedStatus == "success")
				{
					//var milliseconds = stopwatch.ElapsedMilliseconds;
					return "success";
				}
				else
				{
					return updatedStatus;
				}

			}
			catch (Exception ex)
			{

				// string filePath = @"C:\Error.txt";
				string filePath = null;
				if (ConfigurationManager.AppSettings["ErrorFilePath"] != null)
				{
					filePath = ConfigurationManager.AppSettings["ErrorFilePath"].ToString();// @"C:\Error.txt";
				}
				if (filePath != null)
				{

					using (StreamWriter writer = new StreamWriter(filePath, true))
					{
						writer.WriteLine("Messageexecutive :" + ex.Message);
						writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
					}
				}
				return ex.Message;
			}
		}

		public string GenerateKeyMetricsWordReport(string imagelocation, string strSourcePath, string strTargetPath, Report objReport, string practiceid, System.Data.DataTable dtsort)
		{
			try
			{
				if (IsFileInUse(strSourcePath))
				{
					return "File already in use, please close and retry.";
				}
				//Only run the reports that are not ran earlier.
				// List<int> reportToBeGenerated = db.Source_InputData.Where(r => r.RowId > lastReportGenerated && r.IDformat == "Y").Select(r => r.RowId).ToList();

				List<int> reportToBeGenerated = db.Source_InputDataBenchMarkSource.Where(r => r.RowId > lastReportGenerated && r.IDformat == "Y").Select(r => r.RowId).ToList();

				if (reportToBeGenerated.Count() < NoOfDocumentsLimit)
				{
					NoOfDocumentsLimit = 1; // reportToBeGenerated.Count();
				}

				for (int i = reportToBeGenerated.Count - 1; i <= reportToBeGenerated.Count - 1; i++)
				{
					DateTime CurrentDateTime = DateTime.Now;
					string filenamestr = "KeyMetrics" + objReport.lstInput[0].colQ73 + "_" + objReport.lstInput[0].colIDname + "_" + CurrentDateTime.ToString("MMddyyyy-hhmmss");

					System.Web.HttpContext.Current.Session["varNameKeyMetrics"] = filenamestr;

					/*Muntajib-Remove tempPath For both word & pdf*/
					string tempPath = AppDomain.CurrentDomain.BaseDirectory + filenamestr + ".doc";
					filenamestr = strTargetPath + "\\" + filenamestr + ".doc";

					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					FileStream fs = new FileStream(tempPath, FileMode.Create, FileAccess.ReadWrite);
					fs.Close();
					//  Just to kill WINWORD.EXE if it is running
					//  killprocess("winword");     
					//  copy letter format to temp.doc

					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					File.Copy(strSourcePath, tempPath, true);

					//  create missing object
					object missing = Missing.Value;
					//  create Word application object
					Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
					//  create Word document object
					Microsoft.Office.Interop.Word.Document aDoc = null;
					//  create & define filename object with temp.doc
					/*Muntajib-Change 'tempPath to 'filenamestr' to For both word & pdf*/
					object filename = tempPath;
					if (File.Exists((string)filename))
					{
						/*//--Shahbaz.Need to enable when we go with Word to Pdf Report.*/
						object SaveChanges = false;

						object readOnly = false;
						object isVisible = false;
						//  make visible Word application
						wordApp.Visible = false;
						//  open Word document named temp.doc
						aDoc = wordApp.Documents.Open(ref filename, ref missing,
					   ref readOnly, ref missing, ref missing, ref missing,
					   ref missing, ref missing, ref missing, ref missing,
					   ref missing, ref isVisible, ref missing, ref missing,
					   ref missing, ref missing);
						aDoc.Activate();

						/*//--Shweta.Need to enable this code,when we go with Word to Pdf Report.
                       
                        //--Shweta.If we want to put it as web app,we just need to install "SaveAsPDFandXPS.exe" on the server.*/

						object outputFileName = filenamestr.Replace(".doc", ".pdf");
						object fileFormat = WdSaveFormat.wdFormatPDF;

						//  Call FindAndReplace()function for each change

						#region Replace Word Documnet Tempalte's content.


						Microsoft.Office.Interop.Word.Find fnd = wordApp.ActiveWindow.Selection.Find;

						fnd.ClearFormatting();
						fnd.Replacement.ClearFormatting();
						fnd.Forward = true;
						fnd.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;

						fnd.Text = "<<Customer_Name>>";
						fnd.Replacement.Text = objReport.lstInput[0].colQ74;
						this.FindAndReplace(wordApp, "<<Customer_Name>>", objReport.lstInput[0].colQ74);
						fnd.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);

						foreach (Section aSection in wordApp.ActiveDocument.Sections)
						{
							//It contains multiple headers in blank template.
							foreach (HeaderFooter aHeader in aSection.Headers)
							{
								//Only Replace the header contains "Prepared exclusively for «Q73», «Q74»" in it.
								if (aHeader.Range.Text.Contains("Prepared exclusively for <<Customer_Name>>"))
								{
									aHeader.Range.Text = "Prepared exclusively for " + objReport.lstInput[0].colQ74;
								}

							}
						}

						DateTime dt = DateTime.Now;
						string day = dt.Day.ToString();
						string month = dt.ToString("MMMM", CultureInfo.InvariantCulture);
						string year = dt.Year.ToString();
						this.FindAndReplace(wordApp, "<<DD>>", day);
						this.FindAndReplace(wordApp, "<<MM>>", month);
						this.FindAndReplace(wordApp, "<<YYYY>>", year);

						int currentYearValue = Convert.ToInt32(System.Web.HttpContext.Current.Session["YearName"]);
						string previousYear = Convert.ToString(currentYearValue - 1);
						string currentYear = Convert.ToString(currentYearValue);
						decimal medAllSal = 0;
						decimal annFrmsTrnovrMed = 0;
						decimal presBioPerc = 0;
						decimal singleVisPerc = 0;
						decimal healthVisionPlansPerc = 0;
						decimal directPateintsPerc = 0;
						decimal medicarePerc = 0;
						decimal twoWeekVal = 0;
						decimal monthlyVal = 0;
						decimal singleVsnLensAvg = 0;
						decimal progVsnLensAvg = 0;
						//--- For Keymetrics Report Graphs msinghai --//
						#region Spectacle Report Graphs
						string pictureName = string.Empty;  // @"~\SurveyApp\SurveyApp\ImageChart\Spectacle_Graph_1.png";
						int count = aDoc.Bookmarks.Count;
						for (int i1 = 1; i1 < count + 1; i1++)
						{
							object oRange = aDoc.Bookmarks[i1].Range;
							object saveWithDocument = true;
							object missing1 = Type.Missing;
							decimal grossRev = Convert.ToDecimal(objReport.lstInput[0].colQ24 == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ24)).ToString("#,0")); // Its used most often
							if (i1 == 10 || i1 == 82)
							{
								pictureName = imagelocation + @"\Spectacle_Graph_1.png";  // @"~\SurveyApp\SurveyApp\ImageChart\Spectacle_Graph_1.png";

								//percentages for pie chart
								//decimal totalGrossRevenue = grossRev;
								//decimal presEyewarePerc = totalGrossRevenue == 0 ? 0 : ((Convert.ToDecimal(objReport.lstInput[0].colQ26i == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26i)).ToString("#,0"))) / totalGrossRevenue) * 100;
								//decimal eyeExamPerc = totalGrossRevenue == 0 ? 0 : ((Convert.ToDecimal(objReport.lstInput[0].colQ26 == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26)).ToString("#,0"))) / totalGrossRevenue) * 100;
								//decimal medEyeCarePerc = totalGrossRevenue == 0 ? 0 : ((Convert.ToDecimal(objReport.lstInput[0].colQ26b == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26b)).ToString("#,0"))) / totalGrossRevenue) * 100;
								//decimal othersPerc = totalGrossRevenue == 0 ? 0 : (((Convert.ToDecimal(objReport.lstInput[0].colQ26c == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26c)).ToString("#,0"))) + (Convert.ToDecimal(objReport.lstInput[0].colQ26h == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26h)).ToString("#,0")))) / totalGrossRevenue) * 100;
								//decimal contactlensesPerc = totalGrossRevenue == 0 ? 0 : (((Convert.ToDecimal(objReport.lstInput[0].colQ26a == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26a)).ToString("#,0"))) + (Convert.ToDecimal(objReport.lstInput[0].colQ26g == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26g)).ToString("#,0")))) / totalGrossRevenue) * 100;


								decimal totalGrossRevenue = Math.Round(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 0).Select(x => x.Q24 ?? 0).Average()); ;
								decimal presEyewarePerc = totalGrossRevenue == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q26i > 0).Select(x => x.Q26i ?? 0).Average() / totalGrossRevenue) * 100;

								decimal eyeExamPerc = totalGrossRevenue == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q26 > 0).Select(x => x.Q26 ?? 0).Average() / totalGrossRevenue) * 100;

								decimal medEyeCarePerc = totalGrossRevenue == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q26b > 0).Select(x => x.Q26b ?? 0).Average() / totalGrossRevenue) * 100;

								decimal othersPerc = totalGrossRevenue == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q26c > 0).Select(x => x.Q26c ?? 0).Average() / totalGrossRevenue +
									(db.Source_InputDataBenchMarkSource.Where(x => x.Q26h > 0).Select(x => x.Q26h ?? 0).Average()) / totalGrossRevenue) * 100;

								decimal contactlensesPerc = totalGrossRevenue == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q26a > 0).Select(x => x.Q26a ?? 0).Average() / totalGrossRevenue +
									(db.Source_InputDataBenchMarkSource.Where(x => x.Q26g > 0).Select(x => x.Q26g ?? 0).Average()) / totalGrossRevenue) * 100;



								if (presEyewarePerc + eyeExamPerc + medEyeCarePerc + othersPerc + contactlensesPerc < 100)
								{
									othersPerc = 100 - (presEyewarePerc + eyeExamPerc + medEyeCarePerc + othersPerc + contactlensesPerc);
								}

								List<decimal> percentsList = new List<decimal>();
								percentsList.Add(Math.Round(presEyewarePerc));
								percentsList.Add(Math.Round(eyeExamPerc));
								percentsList.Add(Math.Round(medEyeCarePerc));
								percentsList.Add(Math.Round(othersPerc));
								percentsList.Add(Math.Round(contactlensesPerc));

								int maxIndex = 0;
								if (presEyewarePerc + eyeExamPerc + medEyeCarePerc + othersPerc + contactlensesPerc > 100)
								{
									decimal maxValue = percentsList.Max();
									maxIndex = percentsList.IndexOf(maxValue);
									percentsList[maxIndex] = Math.Round((percentsList.Max() - ((presEyewarePerc + eyeExamPerc + medEyeCarePerc + othersPerc + contactlensesPerc) - 100)));
								}




								//list of colors
								List<Color> pieColors = new List<Color>();
								pieColors.Add(Color.FromArgb(255, 51, 102, 153));
								pieColors.Add(Color.FromArgb(220, 51, 102, 153));
								pieColors.Add(Color.FromArgb(195, 51, 102, 153));
								pieColors.Add(Color.FromArgb(170, 51, 102, 153));
								pieColors.Add(Color.FromArgb(145, 51, 102, 153));

								//list of description of the slices

								List<string> descriptions = new List<string>();
								//descriptions.Add("Prescription Eyewear " + presEyewarePerc + "%");
								//descriptions.Add("Eye Exam " + eyeExamPerc + "%");
								//descriptions.Add("Medical Eye Care " + medEyeCarePerc + "%");
								//descriptions.Add("Others " + othersPerc + "%");
								//descriptions.Add("Contact Lenses " + contactlensesPerc + "%");

								descriptions.Add("Prescription Eyewear");
								descriptions.Add("Eye Exam");
								descriptions.Add("Medical Eye Care");
								descriptions.Add("Others");
								descriptions.Add("Contact Lenses");

								string graphName = i1 == 82 ? "Optometric Practice Sources of Revenue" : "Sources of Revenue";
								string status = CreatePieChart(100, graphName, percentsList, pieColors, descriptions, pictureName);

								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 400;
									shape.Height = 400;
								}
							}
							if (i1 == 11)
							{
								pictureName = imagelocation + @"\KeyMetrics_Independent_RevGrowth_1.png";
								string graphHeader = "Independent OD 2012 Revenue Growth by Source (Average % change versus prior year)";
								List<string> yAxisData = new List<string>() { "Total Revenue",
																			   "Total Professional Fees ", "Eye Exams",
																			   "Medical Eye Care", "Product Sales",
																			   "EyeWear", "Contact Lenses"
																			};
								//----Gross Revenue Percentage Change-----
								string lookuptbl = "Lookup.AvgGrossRevenue_J";
								char type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal avgRevPresentYear = GetAllLookUpValues("Lookup.AvgGrossRevenue_J_" + currentYear).Where(x => x > 0).Average();
								decimal avgRevPreviousYear = GetAllLookUpValues("Lookup.AvgGrossRevenue_J_" + previousYear).Where(x => x > 0).Average();

								decimal percChangeInGrossRev = 0;
								if (avgRevPresentYear >= avgRevPreviousYear)
									percChangeInGrossRev = Math.Round(Convert.ToDecimal(((avgRevPresentYear - avgRevPreviousYear) / avgRevPreviousYear) * 100));
								else
									percChangeInGrossRev = Math.Round(Convert.ToDecimal(((avgRevPreviousYear - avgRevPresentYear) / avgRevPreviousYear) * 100));

								// ----Total Prof Fees Percentage Change---- -
								lookuptbl = "Lookup.TotalProfFees_J";
								type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal totalProfFeesPresentYear = GetAllLookUpValues("Lookup.TotalProfFees_J_" + currentYear).Where(x => x > 0).Average();
								decimal totalProfFeesPreviousYear = GetAllLookUpValues("Lookup.TotalProfFees_J_" + previousYear).Where(x => x > 0).Average();

								decimal percChangeInProfFees = 0;
								if (totalProfFeesPresentYear >= totalProfFeesPreviousYear)
									percChangeInProfFees = Math.Round(Convert.ToDecimal(((totalProfFeesPresentYear - totalProfFeesPreviousYear) / totalProfFeesPreviousYear) * 100));
								else
									percChangeInProfFees = Math.Round(Convert.ToDecimal(((totalProfFeesPreviousYear - totalProfFeesPresentYear) / totalProfFeesPreviousYear) * 100));

								// ----Eye Exams Percentage Change---- -
								lookuptbl = "Lookup.EyeExams_J";
								type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal totalEyeExamsPresentYear = GetAllLookUpValues("Lookup.EyeExams_J_" + currentYear).Where(x => x > 0).Average();
								decimal totalEyeExamsPreviousYear = GetAllLookUpValues("Lookup.EyeExams_J_" + previousYear).Where(x => x > 0).Average();

								decimal percChangeIneyeExams = 0;
								if (totalEyeExamsPresentYear >= totalEyeExamsPreviousYear)
									percChangeIneyeExams = Math.Round(Convert.ToDecimal(((totalEyeExamsPresentYear - totalEyeExamsPreviousYear) / totalEyeExamsPreviousYear) * 100));
								else
									percChangeIneyeExams = Math.Round(Convert.ToDecimal(((totalEyeExamsPreviousYear - totalEyeExamsPresentYear) / totalEyeExamsPreviousYear) * 100));

								// ----Medical Eye Care Percentage Change---- -
								lookuptbl = "Lookup.MedicalEyeCare_J";
								type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal totalMedEyeExamsPresentYear = GetAllLookUpValues("Lookup.MedicalEyeCare_J_" + currentYear).Where(x => x > 0).Average();
								decimal totalMedEyeExamsPreviousYear = GetAllLookUpValues("Lookup.MedicalEyeCare_J_" + previousYear).Where(x => x > 0).Average();

								decimal percChangeInmedEyeCare = 0;
								if (totalMedEyeExamsPresentYear >= totalMedEyeExamsPreviousYear)
									percChangeInmedEyeCare = Math.Round(Convert.ToDecimal(((totalMedEyeExamsPresentYear - totalMedEyeExamsPreviousYear) / totalMedEyeExamsPreviousYear) * 100));
								else
									percChangeInmedEyeCare = Math.Round(Convert.ToDecimal(((totalMedEyeExamsPreviousYear - totalMedEyeExamsPresentYear) / totalMedEyeExamsPreviousYear) * 100));

								// ----Product Sales Percentage Change---- -
								lookuptbl = "Lookup.ProductSales_J";
								type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal totalProductSalesPresentYear = GetAllLookUpValues("Lookup.ProductSales_J_" + currentYear).Where(x => x > 0).Average();
								decimal totalProductSalesPreviousYear = GetAllLookUpValues("Lookup.ProductSales_J_" + previousYear).Where(x => x > 0).Average();

								decimal percChangeInProdSales = 0;
								if (totalProductSalesPresentYear >= totalProductSalesPreviousYear)
									percChangeInProdSales = Math.Round(Convert.ToDecimal(((totalProductSalesPresentYear - totalProductSalesPreviousYear) / totalProductSalesPreviousYear) * 100));
								else
									percChangeInProdSales = Math.Round(Convert.ToDecimal(((totalProductSalesPreviousYear - totalProductSalesPresentYear) / totalProductSalesPreviousYear) * 100));

								// ----Eye Glasses Percentage Change---- -
								lookuptbl = "Lookup.EyeGlassesSales_J";
								type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal totalEyeGlassesPresentYear = GetAllLookUpValues("Lookup.EyeGlassesSales_J_" + currentYear).Where(x => x > 0).Average();
								decimal totalEyeGlassesPreviousYear = GetAllLookUpValues("Lookup.EyeGlassesSales_J_" + previousYear).Where(x => x > 0).Average();

								decimal percChangeInEyeWearSales = 0;
								if (totalEyeGlassesPresentYear >= totalEyeGlassesPreviousYear)
									percChangeInEyeWearSales = Math.Round(Convert.ToDecimal(((totalEyeGlassesPresentYear - totalEyeGlassesPreviousYear) / totalEyeGlassesPreviousYear) * 100));
								else
									percChangeInEyeWearSales = Math.Round(Convert.ToDecimal(((totalEyeGlassesPreviousYear - totalEyeGlassesPresentYear) / totalEyeGlassesPreviousYear) * 100));

								// ----Contact Lens Sales Percentage Change---- -
								lookuptbl = "Lookup.ContactLensSales_J";
								type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal totalContactLensPresentYear = GetAllLookUpValues("Lookup.ContactLensSales_J_" + currentYear).Where(x => x > 0).Average();
								decimal totalContactLensPreviousYear = GetAllLookUpValues("Lookup.ContactLensSales_J_" + previousYear).Where(x => x > 0).Average();

								decimal percChangeInContLens = 0;
								if (totalContactLensPresentYear >= totalContactLensPreviousYear)
									percChangeInContLens = Math.Round(Convert.ToDecimal(((totalContactLensPresentYear - totalContactLensPreviousYear) / totalContactLensPreviousYear) * 100));
								else
									percChangeInContLens = Math.Round(Convert.ToDecimal(((totalContactLensPreviousYear - totalContactLensPresentYear) / totalContactLensPreviousYear) * 100));

								List<decimal> graphData = new List<decimal>() { percChangeInGrossRev, percChangeInProfFees, percChangeIneyeExams,
																			percChangeInmedEyeCare, percChangeInProdSales, percChangeInEyeWearSales, percChangeInContLens };

								string status = CreateIndRevGrowthBarDataGraph(pictureName, graphHeader, yAxisData, graphData);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 350;
								}
							}
							if (i1 == 12)
							{
								pictureName = imagelocation + @"\KeyMetric_EyeExam_DataGraph_1.png";
								string status = CreateEyeExamDataGraph(pictureName);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 300;
								}
							}
							if (i1 == 14)
							{
								pictureName = imagelocation + @"\KeyMetrics_AnnualMedVisit_1.png";
								List<string> eyeCareTypes = new List<string>(){"","Dry Eye","Ocular Infection","Ocular allergy","Glaucoma","Cataract co-management",
																				"Refractive surgery co-management","Foreign body removal","Total"};

								List<string> medianList = new List<string>() { "Median" };

								//Calculation for average
								List<string> averageList = new List<string>() { "Average" };

								decimal avgdryEye = db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20a > 0).Select(x => ((x.Q20a / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000))).Average().Value;
								decimal avgOcularInfc = db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20b > 0).Select(x => ((x.Q20b / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000))).Average().Value;
								decimal avgOcularAllrgy = db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20c > 0).Select(x => ((x.Q20c / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000))).Average().Value;
								decimal avgGlaucoma = db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20d > 0).Select(x => ((x.Q20d / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000))).Average().Value;
								decimal avgCatrctMgmt = db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20e > 0).Select(x => ((x.Q20e / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000))).Average().Value;
								decimal avgRefSrg = db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20f > 0).Select(x => ((x.Q20f / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000))).Average().Value;
								decimal avgFrgnBodyRmvl = db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20g > 0).Select(x => ((x.Q20g / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000))).Average().Value;
								decimal avgTotal = avgdryEye + avgOcularInfc + avgOcularAllrgy + avgGlaucoma + avgCatrctMgmt + avgRefSrg + avgFrgnBodyRmvl;
								averageList.Add(Convert.ToString(Math.Round(avgdryEye)));
								averageList.Add(Convert.ToString(Math.Round(avgOcularInfc)));
								averageList.Add(Convert.ToString(Math.Round(avgOcularAllrgy)));
								averageList.Add(Convert.ToString(Math.Round(avgGlaucoma)));
								averageList.Add(Convert.ToString(Math.Round(avgCatrctMgmt)));
								averageList.Add(Convert.ToString(Math.Round(avgRefSrg)));
								averageList.Add(Convert.ToString(Math.Round(avgFrgnBodyRmvl)));
								averageList.Add(Convert.ToString(Math.Round(avgTotal)));

								//calculations for Median

								decimal meddryEye = GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20a > 0).Select(x => ((x.Q20a / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000)) ?? 0).ToList());
								decimal medOcularInfc = GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20b > 0).Select(x => ((x.Q20b / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000)) ?? 0).ToList());
								decimal medOcularAllrgy = GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20c > 0).Select(x => ((x.Q20c / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000)) ?? 0).ToList());
								decimal medGlaucoma = GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20d > 0).Select(x => ((x.Q20d / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000)) ?? 0).ToList());
								decimal medCatrctMgmt = GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20e > 0).Select(x => ((x.Q20e / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000)) ?? 0).ToList());
								decimal medRefSrg = GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20f > 0).Select(x => ((x.Q20f / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000)) ?? 0).ToList());
								decimal medFrgnBodyRmvl = GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q12 > 0 && x.Q20g > 0).Select(x => ((x.Q20g / (x.Q20a + x.Q20b + x.Q20c + x.Q20d + x.Q20e + x.Q20f + x.Q20g)) / (x.Q12 / 1000)) ?? 0).ToList());
								decimal medTotal = meddryEye + medOcularInfc + medOcularAllrgy + medGlaucoma + medCatrctMgmt + medRefSrg + medFrgnBodyRmvl;
								medianList.Add(Convert.ToString(Math.Round(meddryEye)));
								medianList.Add(Convert.ToString(Math.Round(medOcularInfc)));
								medianList.Add(Convert.ToString(Math.Round(medOcularAllrgy)));
								medianList.Add(Convert.ToString(Math.Round(medGlaucoma)));
								medianList.Add(Convert.ToString(Math.Round(medCatrctMgmt)));
								medianList.Add(Convert.ToString(Math.Round(medRefSrg)));
								medianList.Add(Convert.ToString(Math.Round(medFrgnBodyRmvl)));
								medianList.Add(Convert.ToString(Math.Round(medTotal)));

								string status = CreateAnnualMedicalVisitsDataGraph(pictureName, eyeCareTypes, medianList, averageList);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 350;
								}
							}
							if (i1 == 15)
							{
								pictureName = imagelocation + @"\KeyMetrics_MedEyeCareVisits_1.png";
								string graphHeader = "Medical Eye Care Visits by Type \n (% of Total Medical Eye Care Visits)";
								List<string> xAxisData = new List<string>() {  "Glaucoma","Dry Eye","Ocular allergy",
																				 "Ocular Infection",
																				 "Cataract \n co-management",
																				 "Refractive \n surgery \n co-management",
																				 "Foreign body removal",
																				 };

								List<decimal> graphData = new List<decimal>();
								decimal totalMedicalEyeCareVisits = objReport.lstInput[0].colQ20a == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20a))
																	+ objReport.lstInput[0].colQ20b == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20b))
																	+ objReport.lstInput[0].colQ20c == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20c))
																	+ objReport.lstInput[0].colQ20d == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20d))
																	+ objReport.lstInput[0].colQ20e == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20e))
																	+ objReport.lstInput[0].colQ20f == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20f))
																	+ objReport.lstInput[0].colQ20f == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20g));
								//Add data to graph data
								graphData.Add(Math.Round(objReport.lstInput[0].colQ20d == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20d)) / totalMedicalEyeCareVisits));
								graphData.Add(Math.Round(objReport.lstInput[0].colQ20a == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20a)) / totalMedicalEyeCareVisits));
								graphData.Add(Math.Round(objReport.lstInput[0].colQ20b == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20b)) / totalMedicalEyeCareVisits));
								graphData.Add(Math.Round(objReport.lstInput[0].colQ20e == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20e)) / totalMedicalEyeCareVisits));
								graphData.Add(Math.Round(objReport.lstInput[0].colQ20c == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20c)) / totalMedicalEyeCareVisits));
								graphData.Add(Math.Round(objReport.lstInput[0].colQ20f == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20f)) / totalMedicalEyeCareVisits));
								graphData.Add(Math.Round(objReport.lstInput[0].colQ20g == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ20g)) / totalMedicalEyeCareVisits));

								string status = CreateVerticalGraph(120, pictureName, graphHeader, xAxisData, graphData, true);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 350;
								}
							}
							if (i1 == 17)
							{
								pictureName = imagelocation + @"\KeyMetrics_PieChart_Graph_1.png";

								//percentages for pie chart
								decimal total = Math.Round(db.Source_InputDataBenchMarkSource.Where(x => x.Q27e > 0).Select(x => x.Q27e ?? 0).Average());
								healthVisionPlansPerc = total == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q27a > 0).Select(x => x.Q27a ?? 0).Average() / total) * 100;
								directPateintsPerc = total == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q27d > 0).Select(x => x.Q27d ?? 0).Average() / total) * 100;
								medicarePerc = total == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q27 > 0).Select(x => x.Q27 ?? 0).Average() / total) * 100;

								List<decimal> percentsList = new List<decimal>();
								percentsList.Add(Math.Round(healthVisionPlansPerc));
								percentsList.Add(Math.Round(directPateintsPerc));
								percentsList.Add(Math.Round(medicarePerc));

								//list of colors
								List<Color> pieColors = new List<Color>();
								pieColors.Add(Color.FromArgb(255, 51, 102, 153));
								pieColors.Add(Color.FromArgb(220, 51, 102, 153));
								pieColors.Add(Color.FromArgb(195, 51, 102, 153));


								//list of description of the slices

								List<string> descriptions = new List<string>();
								//descriptions.Add("Health Vision Plans " + healthVisionPlansPerc + "%");
								//descriptions.Add("Direct From Pateints " + directPateintsPerc + "%");
								//descriptions.Add("Medicare " + medicarePerc + "%");
								descriptions.Add("Health Vision Plans");
								descriptions.Add("Direct From Pateints");
								descriptions.Add("Medicare");

								string status = CreatePieChart(100, "MBA Participant Source of Payments", percentsList, pieColors, descriptions, pictureName);

								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 300;
									shape.Height = 300;
								}

							}
							if (i1 == 35)
							{
								pictureName = imagelocation + @"\KeyMetrics_PieChart_Graph_2.png";

								//percentages for 1st pie chart
								presBioPerc = Math.Round(db.Source_InputDataBenchMarkSource.Where(x => x.Q31b > 0).Select(x => x.Q31b ?? 0).Average());
								singleVisPerc = Math.Round(db.Source_InputDataBenchMarkSource.Where(x => x.Q31a > 0).Select(x => x.Q31a ?? 0).Average());

								List<decimal> percentsList = new List<decimal>();
								percentsList.Add(Math.Round(presBioPerc));
								percentsList.Add(Math.Round(singleVisPerc));

								//list of colors
								List<Color> pieColors = new List<Color>();
								pieColors.Add(Color.FromArgb(255, 51, 102, 153));
								pieColors.Add(Color.FromArgb(240, 51, 102, 153));

								//list of description of the slices

								List<string> descriptions = new List<string>();
								descriptions.Add("PresByopic " + presBioPerc + "%");
								descriptions.Add("Single Vision " + singleVisPerc + "%");

								//percentages for 2nd pie chart
								decimal bitrifocal = db.Source_InputDataBenchMarkSource.Where(x => x.Q32b > 0).Select(x => x.Q31b ?? 0).Average();
								decimal progressive = db.Source_InputDataBenchMarkSource.Where(x => x.Q32c > 0).Select(x => x.Q32c ?? 0).Average();
								decimal other = db.Source_InputDataBenchMarkSource.Where(x => x.Q32d > 0).Select(x => x.Q32d ?? 0).Average();

								List<decimal> percentsList1 = new List<decimal>();
								percentsList1.Add(Math.Round(bitrifocal));
								percentsList1.Add(Math.Round(progressive));
								percentsList1.Add(Math.Round(other));

								//list of colors
								List<Color> pieColors1 = new List<Color>();
								pieColors1.Add(Color.FromArgb(255, 51, 102, 153));
								pieColors1.Add(Color.FromArgb(240, 51, 102, 153));
								pieColors1.Add(Color.FromArgb(225, 51, 102, 153));


								//list of description of the slices

								List<string> descriptions1 = new List<string>();
								descriptions1.Add("Bifocal/Trifocal " + bitrifocal + "%");
								descriptions1.Add("Progressive " + progressive + "%");
								descriptions1.Add("Other " + progressive + "%");

								string status = CreateTwoPieCharts("Spectacle Lens Rxes (% of total eyewear Rxes)", percentsList, pieColors, descriptions, percentsList1, pieColors1, descriptions1, pictureName);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 300;
								}

							}
							if (i1 == 42)
							{
								pictureName = imagelocation + @"\KeyMetrics_SpectacleLensMarkups_1.png";
								string graphHeader = "Spectacle Lens Mark-Ups";
								List<string> yAxisData = new List<string>() { "Polycarbonate",
																			   "Polycarbonate, anti-reflective", "High index, anti-reflective",
																			   "Polycarbonate", "Polycarbonate, anti-reflective",
																			   "High index, anti-reflective", "Polycarbonate, photochromic"
																			};

								decimal totalCostOfGoods = db.Source_InputDataBenchMarkSource.Where(x => x.Q52j > 0).Select(x => x.Q52j ?? 0).Average();

								decimal svPoly = totalCostOfGoods == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q92a > 0).Select(x => x.Q92a ?? 0).Average() / totalCostOfGoods;
								decimal svPolyanti = totalCostOfGoods == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q92b > 0).Select(x => x.Q92a ?? 0).Average() / totalCostOfGoods;
								decimal svHighInd = totalCostOfGoods == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q92c > 0).Select(x => x.Q92a ?? 0).Average() / totalCostOfGoods;
								decimal progPoly = totalCostOfGoods == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q92d > 0).Select(x => x.Q92a ?? 0).Average() / totalCostOfGoods;
								decimal progPolyAnti = totalCostOfGoods == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q92e > 0).Select(x => x.Q92a ?? 0).Average() / totalCostOfGoods;
								decimal progHighInd = totalCostOfGoods == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q92f > 0).Select(x => x.Q92a ?? 0).Average() / totalCostOfGoods;
								decimal progPolyPhoto = totalCostOfGoods == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q92g > 0).Select(x => x.Q92a ?? 0).Average() / totalCostOfGoods;

								singleVsnLensAvg = new List<decimal>() { svPoly, svPolyanti, svHighInd }.Average();
								progVsnLensAvg = new List<decimal>() { progPoly, progPolyAnti, progHighInd, progPolyPhoto }.Average();

								List<decimal> graphData = new List<decimal>();
								graphData.Add(Math.Round(svPoly));
								graphData.Add(Math.Round(svPolyanti));
								graphData.Add(Math.Round(svHighInd));
								graphData.Add(Math.Round(progPoly));
								graphData.Add(Math.Round(progPolyAnti));
								graphData.Add(Math.Round(progHighInd));
								graphData.Add(Math.Round(progPolyPhoto));



								List<decimal> xAxisData = new List<decimal>();
								decimal xEnd = 0;
								decimal graphMaxData = graphData.Max();
								if (graphMaxData == 0)
								{
									xEnd = Convert.ToDecimal(graphMaxData % 0.5m == 0 ? graphMaxData + 0.5m : (Math.Ceiling(graphMaxData) - graphMaxData > 0.5m ? Math.Ceiling(graphMaxData) - 0.5m : Math.Ceiling(graphMaxData)));
								}
								else
								{
									xEnd = 2; // if no value is present in graph data then by default draw xAxis data upto value '2'

								}
								decimal xData = 0;
								while (xData <= xEnd)
								{
									xAxisData.Add(xData);
									xData += 0.5m;
								}
								//-----------------------------------------------------------------------------------

								string status = CreateSpectacleBarDataGraph(pictureName, graphHeader, yAxisData, xAxisData, graphData);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 300;
								}

							}
							if (i1 == 43)
							{
								pictureName = imagelocation + @"\KeyMetrics_FramesInventoryTurnover_1.png";
								string graphHeader = "Frames Inventory and Turnover";
								List<decimal> framesInInv = new List<decimal>();
								List<decimal> annComSpecRx = new List<decimal>();
								List<decimal> annFramesTurnover = new List<decimal>();
								List<decimal> valueOfFramesInv = new List<decimal>();
								List<string> medians = new List<string>();
								medians.Add("Total MBA Practices");

								//Get Frames In Inventory
								decimal data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 2133000).Select(x => x.Q34 + x.Q35).Average());
								framesInInv.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 2132999 && x.Q24 >= 1695000).Select(x => x.Q34 + x.Q35).Average());
								framesInInv.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1694999 && x.Q24 >= 1432000).Select(x => x.Q34 + x.Q35).Average());
								framesInInv.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1431999 && x.Q24 >= 1200000).Select(x => x.Q34 + x.Q35).Average());
								framesInInv.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1199999 && x.Q24 >= 1026000).Select(x => x.Q34 + x.Q35).Average());
								framesInInv.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1025999 && x.Q24 >= 883000).Select(x => x.Q34 + x.Q35).Average());
								framesInInv.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 882999 && x.Q24 >= 767000).Select(x => x.Q34 + x.Q35).Average());
								framesInInv.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 766999 && x.Q24 >= 642000).Select(x => x.Q34 + x.Q35).Average());
								framesInInv.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 641999 && x.Q24 >= 493000).Select(x => x.Q34 + x.Q35).Average());
								framesInInv.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 492999).Select(x => x.Q34 + x.Q35).Average());
								framesInInv.Add(Math.Round(data));

								//add median of frames in Inventory
								medians.Add(Convert.ToString(Math.Round(db.Source_InputDataBenchMarkSource.Select(x => x.Q34 + x.Q35).Average().Value)));


								//Get Annual Com Spec Rxes

								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 2133000).Select(x => x.Q28 + x.Q29).Average());
								annComSpecRx.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 2132999 && x.Q24 >= 1695000).Select(x => x.Q28 + x.Q29).Average());
								annComSpecRx.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1694999 && x.Q24 >= 1432000).Select(x => x.Q28 + x.Q29).Average());
								annComSpecRx.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1431999 && x.Q24 >= 1200000).Select(x => x.Q28 + x.Q29).Average());
								annComSpecRx.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1199999 && x.Q24 >= 1026000).Select(x => x.Q28 + x.Q29).Average());
								annComSpecRx.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1025999 && x.Q24 >= 883000).Select(x => x.Q28 + x.Q29).Average());
								annComSpecRx.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 882999 && x.Q24 >= 767000).Select(x => x.Q28 + x.Q29).Average());
								annComSpecRx.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 766999 && x.Q24 >= 642000).Select(x => x.Q28 + x.Q29).Average());
								annComSpecRx.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 641999 && x.Q24 >= 493000).Select(x => x.Q28 + x.Q29).Average());
								annComSpecRx.Add(Math.Round(data));
								data = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 492999).Select(x => x.Q28 + x.Q29).Average());
								annComSpecRx.Add(Math.Round(data));

								//add median of Annual Com Spec Rxex
								medians.Add(Convert.ToString(Math.Round(GetListMedian(db.Source_InputDataBenchMarkSource.Select(x => (x.Q28 + x.Q29) ?? 0).ToList()))));

								//Annual Frames Turnover
								decimal annFramesTrn = 0;
								annFramesTrn = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 2133000 && (x.Q34 > 0 || x.Q35 > 0)).Select(x => x.Q28 / (x.Q34 + x.Q35)).Average());
								annFramesTurnover.Add(Math.Round(annFramesTrn));
								annFramesTrn = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 2132999 && x.Q24 >= 1695000 && (x.Q34 > 0 || x.Q35 > 0)).Select(x => x.Q28 / (x.Q34 + x.Q35)).Average());
								annFramesTurnover.Add(Math.Round(annFramesTrn));
								annFramesTrn = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1694999 && x.Q24 >= 1432000 && (x.Q34 > 0 || x.Q35 > 0)).Select(x => x.Q28 / (x.Q34 + x.Q35)).Average());
								annFramesTurnover.Add(Math.Round(annFramesTrn));
								annFramesTrn = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1431999 && x.Q24 >= 1200000 && (x.Q34 > 0 || x.Q35 > 0)).Select(x => x.Q28 / (x.Q34 + x.Q35)).Average());
								annFramesTurnover.Add(Math.Round(annFramesTrn));
								annFramesTrn = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1199999 && x.Q24 >= 1026000 && (x.Q34 > 0 || x.Q35 > 0)).Select(x => x.Q28 / (x.Q34 + x.Q35)).Average());
								annFramesTurnover.Add(Math.Round(annFramesTrn));
								annFramesTrn = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1025999 && x.Q24 >= 883000 && (x.Q34 > 0 || x.Q35 > 0)).Select(x => x.Q28 / (x.Q34 + x.Q35)).Average());
								annFramesTurnover.Add(Math.Round(annFramesTrn));
								annFramesTrn = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 882999 && x.Q24 >= 767000 && (x.Q34 > 0 || x.Q35 > 0)).Select(x => x.Q28 / (x.Q34 + x.Q35)).Average());
								annFramesTurnover.Add(Math.Round(annFramesTrn));
								annFramesTrn = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 766999 && x.Q24 >= 642000 && (x.Q34 > 0 || x.Q35 > 0)).Select(x => x.Q28 / (x.Q34 + x.Q35)).Average());
								annFramesTurnover.Add(Math.Round(annFramesTrn));
								annFramesTrn = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 641999 && x.Q24 >= 493000 && (x.Q34 > 0 || x.Q35 > 0)).Select(x => x.Q28 / (x.Q34 + x.Q35)).Average());
								annFramesTurnover.Add(Math.Round(annFramesTrn));
								annFramesTrn = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 492999 && (x.Q34 > 0 || x.Q35 > 0)).Select(x => x.Q28 / (x.Q34 + x.Q35)).Average());
								annFramesTurnover.Add(Math.Round(annFramesTrn));

								//add median of Annual Frames Turnover
								annFrmsTrnovrMed = Math.Round(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q34 > 0 || x.Q35 > 0).Select(x => (x.Q28 / (x.Q34 + x.Q35)) ?? 0).ToList()));
								medians.Add(Convert.ToString(annFrmsTrnovrMed));


								//Value Frames Inventory
								decimal valueFramesInv = 0;
								valueFramesInv = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 2133000).Select(x => (x.Q34 + x.Q35) * x.Q36).Average());
								valueOfFramesInv.Add(Math.Round(valueFramesInv));
								valueFramesInv = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 2132999 && x.Q24 >= 1695000).Select(x => (x.Q34 + x.Q35) * x.Q36).Average());
								valueOfFramesInv.Add(Math.Round(valueFramesInv));
								valueFramesInv = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1694999 && x.Q24 >= 1432000).Select(x => (x.Q34 + x.Q35) * x.Q36).Average());
								valueOfFramesInv.Add(Math.Round(valueFramesInv));
								valueFramesInv = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1431999 && x.Q24 >= 1200000).Select(x => (x.Q34 + x.Q35) * x.Q36).Average());
								valueOfFramesInv.Add(Math.Round(valueFramesInv));
								valueFramesInv = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1199999 && x.Q24 >= 1026000).Select(x => (x.Q34 + x.Q35) * x.Q36).Average());
								valueOfFramesInv.Add(Math.Round(valueFramesInv));
								valueFramesInv = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 1025999 && x.Q24 >= 883000).Select(x => (x.Q34 + x.Q35) * x.Q36).Average());
								valueOfFramesInv.Add(Math.Round(valueFramesInv));
								valueFramesInv = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 882999 && x.Q24 >= 767000).Select(x => (x.Q34 + x.Q35) * x.Q36).Average());
								valueOfFramesInv.Add(Math.Round(valueFramesInv));
								valueFramesInv = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 766999 && x.Q24 >= 642000).Select(x => (x.Q34 + x.Q35) * x.Q36).Average());
								valueOfFramesInv.Add(Math.Round(valueFramesInv));
								valueFramesInv = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 641999 && x.Q24 >= 493000).Select(x => (x.Q34 + x.Q35) * x.Q36).Average());
								valueOfFramesInv.Add(Math.Round(valueFramesInv));
								valueFramesInv = Convert.ToInt64(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 492999).Select(x => (x.Q34 + x.Q35) * x.Q36).Sum());
								valueOfFramesInv.Add(Math.Round(valueFramesInv));

								//add median of Annual Frames Turnover
								medians.Add(Convert.ToString(Math.Round(GetListMedian(db.Source_InputDataBenchMarkSource.Select(x => ((x.Q34 + x.Q35) * x.Q36) ?? 0).ToList()))));

								string status = CreateFramesInventoryTurnoverDataGraph(pictureName, graphHeader, framesInInv, annComSpecRx, annFramesTurnover, valueOfFramesInv, medians);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 400;
								}
							}
							if (i1 == 44)
							{
								pictureName = imagelocation + @"\KeyMetrics_FramesInventoryGuideline_1.png";
								string graphHeader = "Frames Inventory Guideline";
								List<string> medianAnnFrames = new List<string>();
								medianAnnFrames = GetExternalTableValues("ES_Frames_Inventory_Guidelines").Values.ToList();

								decimal targetTurnOver = Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ15a == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ15a)).ToString("#,0"))
														/ 3);
								//Ideal Frames turnover
								List<decimal> idealFramesInv = new List<decimal>();
								decimal data = targetTurnOver == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 500000 && x.Q28 > 0).Select(x => x.Q28 ?? 0 / targetTurnOver).Average();
								idealFramesInv.Add(Math.Round(data));
								data = targetTurnOver == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 500000 && x.Q24 <= 800000 && x.Q28 > 0).Select(x => x.Q28 ?? 0 / targetTurnOver).Average();
								idealFramesInv.Add(Math.Round(data));
								data = targetTurnOver == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 800000 && x.Q24 <= 1100000 && x.Q28 > 0).Select(x => x.Q28 ?? 0 / targetTurnOver).Average();
								idealFramesInv.Add(Math.Round(data));
								data = targetTurnOver == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 1100000 && x.Q24 <= 1400000 && x.Q28 > 0).Select(x => x.Q28 ?? 0 / targetTurnOver).Average();
								idealFramesInv.Add(Math.Round(data));
								data = targetTurnOver == 0 ? 0 : db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 1400000 && x.Q24 <= 2000000 && x.Q28 > 0).Select(x => x.Q28 ?? 0 / targetTurnOver).Average();
								idealFramesInv.Add(Math.Round(data));

								List<decimal> excessInv = new List<decimal>();
								foreach (decimal ideal in idealFramesInv)
								{
									excessInv.Add(ideal + (ideal * 0.2m));
								}

								List<decimal> insuffInv = new List<decimal>();
								foreach (decimal ideal in idealFramesInv)
								{
									insuffInv.Add(ideal - (ideal * 0.2m));
								}

								string status = CreateFramesInventoryGuidelineDataGraph(pictureName, graphHeader, medianAnnFrames, idealFramesInv, excessInv, insuffInv);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 300;
								}
							}
							if (i1 == 45)
							{
								pictureName = imagelocation + @"\KeyMetrics_FramesUnitMix_1.png";
								string graphHeader = "Frames Unit Sales Mix by Price Point";
								List<string> xAxisData = new List<string>() {  "$99 or less",
																			   "$100-149",
																			   "$150-199",
																			   "$200-299",
																			   "$300-399",
																			   "$400 or more",
																			};

								List<decimal> graphData = new List<decimal>();
								int framesSold = db.Source_InputDataBenchMarkSource.Where(x => (x.Q36 != null || (x.Q36 != 0))).Count();
								decimal data = framesSold == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q36 <= 99 && (x.Q36 != null || (x.Q36 != 0))).Count() / framesSold) * 100;
								graphData.Add(Math.Round(data));
								data = framesSold == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => (x.Q36 != null || (x.Q36 != 0)) && x.Q36 >= 100 && x.Q36 <= 149).Count() / framesSold) * 100;
								graphData.Add(Math.Round(data));
								data = framesSold == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => (x.Q36 != null || (x.Q36 != 0)) && x.Q36 >= 150 && x.Q36 <= 199).Count() / framesSold) * 100;
								graphData.Add(Math.Round(data));
								data = framesSold == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => (x.Q36 != null || (x.Q36 != 0)) && x.Q36 >= 200 && x.Q36 <= 299).Count() / framesSold) * 100;
								graphData.Add(Math.Round(data));
								data = framesSold == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => (x.Q36 != null || (x.Q36 != 0)) && x.Q36 >= 300 && x.Q36 <= 399).Count() / framesSold) * 100;
								graphData.Add(Math.Round(data));
								data = framesSold == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => (x.Q36 != null || (x.Q36 != 0)) && x.Q36 >= 400).Count() / framesSold) * 100;
								graphData.Add(Math.Round(data));


								string status = CreateVerticalGraph(100, pictureName, graphHeader, xAxisData, graphData, true);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 350;
								}
							}
							if (i1 == 61)
							{
								pictureName = imagelocation + @"\KeyMetrics_AnnualSupPurchase_1.png";
								string graphHeader = "Annual Supply Purchase by Soft Lens Modality";
								List<string> yAxisData = new List<string>() { "Two-Week lenses", "Monthly lenses" };
								twoWeekVal = db.Source_InputDataBenchMarkSource.Where(x => x.Q39b > 0).Select(x => x.Q39b ?? 0).Average();
								monthlyVal = db.Source_InputDataBenchMarkSource.Where(x => x.Q39c > 0).Select(x => x.Q39c ?? 0).Average(); ;
								List<decimal> graphData = new List<decimal>() { Math.Round(twoWeekVal), Math.Round(monthlyVal) };

								bool isXAxisDrawn = true;

								string status = CreateHorizontalBarGraph(100, pictureName, graphHeader, yAxisData, graphData, isXAxisDrawn);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 250;
								}

							}
							if (i1 == 64)
							{
								pictureName = imagelocation + @"\KeyMetrics_SoftLensReq_1.png";
								List<string> practicelAnnGorssRev = new List<string>()
								{
									"Practice Annual Gross Revenue",
									"$350,000",
									"$500,000",
									"$650,000",
									"$800,000",
									"$1,000,000",
									"$1,200,000",
								};
								List<string> medianList = new List<string>() { "Median Monthly Contact Lens Exams" };
								List<string> softLensInv = new List<string>() { "Soft Lens Inventory Requirement (boxes)" };
								List<string> externalData = GetExternalTableValues("ES_Soft_Lens_Inventory_Requirements").Values.ToList();
								for (int j = 0; j < externalData.Count(); j++)
								{
									if (j % 2 != 0)
										medianList.Add(externalData[j]);
									else
										softLensInv.Add(externalData[j]);
								}

								string status = CreateSoftLensReqDataGraph(pictureName, practicelAnnGorssRev, medianList, softLensInv);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 300;
								}
							}
							if (i1 == 66) //Full Time Office Managers
							{
								pictureName = imagelocation + @"\KeyMetrics_FullTimeOfcMgr_1.png";
								string graphHeader = "Full-Time Office Manager by Practice Size";
								List<string> xAxisData = new List<string>() {  "Small $510,000",
																			   "Medium Small $796,000",
																			   "Medium $1.1M",
																			   "Medium Large $1.5M",
																			   "Large $2.2M",
																			};

								List<decimal> graphData = new List<decimal>();
								decimal data = db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 510000).Count() == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 510000 && x.Q97 == true).Count() / db.Source_InputDataBenchMarkSource.Where(x => x.Q24 <= 510000).Count()) * 100;
								graphData.Add(Math.Round(data));
								data = db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 510000 && x.Q24 <= 796000).Count() == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 510000 && x.Q24 <= 796000 && x.Q97 == true).Count() / db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 510000 && x.Q24 <= 796000).Count()) * 100;
								graphData.Add(Math.Round(data));
								data = db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 796000 && x.Q24 <= 1100000).Count() == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 796000 && x.Q24 <= 1100000 && x.Q97 == true).Count() / db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 796000 && x.Q24 <= 1100000).Count()) * 100;
								graphData.Add(Math.Round(data));
								data = db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 1100000 && x.Q24 <= 1500000).Count() == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 1100000 && x.Q24 <= 1500000 && x.Q97 == true).Count() / db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 1100000 && x.Q24 <= 1500000).Count()) * 100;
								graphData.Add(Math.Round(data));
								data = db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 1500000 && x.Q24 <= 2200000).Count() == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 1500000 && x.Q24 <= 2200000 && x.Q97 == true).Count() / db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 1500000 && x.Q24 <= 2200000).Count()) * 100;
								graphData.Add(Math.Round(data));

								string status = CreateVerticalGraph(100, pictureName, graphHeader, xAxisData, graphData, true);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 350;
								}

							}
							if (i1 == 68)
							{
								pictureName = imagelocation + @"\KeyMetrics_StaffHourlyandAnnualSalaries_1.png";
								string graphHeader = "Staff Hourly and Annual Salaries by Position";
								List<decimal> avgHourlySal = new List<decimal>();

								avgHourlySal.Add(22.33m);
								avgHourlySal.Add(13.21m);
								avgHourlySal.Add(14.73m);
								avgHourlySal.Add(16.80m);
								avgHourlySal.Add(17.73m);
								avgHourlySal.Add(13.13m);
								avgHourlySal.Add(17.19m);
								avgHourlySal.Add(15.15m);


								List<decimal> medianHourlySal = new List<decimal>();
								medianHourlySal.Add(20.48m);
								medianHourlySal.Add(12.71m);
								medianHourlySal.Add(14.40m);
								medianHourlySal.Add(16.29m);
								medianHourlySal.Add(17.25m);
								medianHourlySal.Add(12.64m);
								medianHourlySal.Add(16.12m);
								medianHourlySal.Add(15.00m);

								medAllSal = GetListMedian(medianHourlySal);
								List<decimal> medianAnnuallySal = new List<decimal>();
								medianAnnuallySal.Add(42598m);
								medianAnnuallySal.Add(26437m);
								medianAnnuallySal.Add(29952m);
								medianAnnuallySal.Add(33883m);
								medianAnnuallySal.Add(35880m);
								medianAnnuallySal.Add(26291m);
								medianAnnuallySal.Add(33350m);
								medianAnnuallySal.Add(31200m);

								string status = CreateStaffHourlySalariesByPositionDataGraph(pictureName, graphHeader, avgHourlySal, medianHourlySal, medianAnnuallySal);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 350;
								}
							}
							if (i1 == 69)
							{
								pictureName = imagelocation + @"\KeyMetric_ExpenseCategory_DataGraph_1.png";
								string status = CreateExpenseCategoryDataGraph(pictureName);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 450;
								}
							}
							if (i1 == 70)
							{
								pictureName = imagelocation + @"\KeyMetrics_HorizontalPlotGraph_1.png";
								string graphHeader = "Range of Expense Ratios by Category";
								List<string> yAxisData = new List<string>() {
																			 "Cost of Goods",
																			 "Staff Salaries and Benefits",
																			 "Occupancy",
																			 "Equipment",
																			 "Marketing & Promotion",
																			 "General Office Overhead",
																			 "Interest",
																			 "Insurance"
																			 };
								List<decimal> graphDataMin = new List<decimal>();
								List<decimal> graphDataMax = new List<decimal>();

								//Fill Min Max graph data for each type
								List<string> tableList = new List<string>()
								{   "Lookup.ExpenseRatioPercentageByCostOfGoods_J",
									"Lookup.ExpenseRatioPercentageByStaffSalaries_J",
									"Lookup.ExpenseRatioPercentageByOccupancy_J",
									"Lookup.ExpenseRatioPercentageByEquipment_J",
									"Lookup.ExpenseRatioPercentageByMarketing_J",
									"Lookup.ExpenseRatioPercentageByGenOverhead_J",
									"Lookup.ExpenseRatioPercentageByInterest_J",
									"Lookup.ExpenseRatioPercentageByRepairMaintenance_J",
									"Lookup.ExpenseRatioPercentageByInsurance_J",
								};
								List<decimal> minmax = new List<decimal>();
								foreach (var item in tableList)
								{
									minmax = GetMinAndMaxLookUpValue(item, 60); // getting min and max of 60th Percentile
									graphDataMin.Add(minmax.FirstOrDefault());
									graphDataMax.Add(minmax.LastOrDefault());
								}

								string status = CreateHorizontalPlotGraph(pictureName, graphHeader, yAxisData, graphDataMin, graphDataMax);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 300;
								}

							}
							if (i1 == 77) //additionalgraphs 1 ->Average % change in total cost of goods by change in Frame wholesale cost \n (Average % change versus prior year)
							{
								pictureName = imagelocation + @"\KeyMetrics_AvgPerCngInCostOfGoods_1.png";
								string graphHeader = "Average % change in total cost of goods by change in Frame wholesale cost \n (Average % change versus prior year)";

								List<string> yAxisData = new List<string>() { "% Change in Total Cost of Goods", "% Change in Wholesale Cost of Frames" };

								List<decimal> graphData = new List<decimal>();

								//get % change in total cost of goods
								//To get the % change between present year and prior year, we need to generate an annual lookup table on the fly

								string lookuptbl = "Lookup.AvgCostOfGoods_J";
								char type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal costOfGoodsPresentYear = GetAllLookUpValues("Lookup.AvgCostOfGoods_J_" + currentYear).Where(x => x > 0).Average();
								decimal costOfGoodsPreviousYear = GetAllLookUpValues("Lookup.AvgCostOfGoods_J_" + previousYear).Where(x => x > 0).Average();
								decimal PerChangeInCostOfGoods = 0;
								if (costOfGoodsPresentYear >= costOfGoodsPreviousYear)
									PerChangeInCostOfGoods = costOfGoodsPreviousYear == 0 ? 0 : ((costOfGoodsPresentYear - costOfGoodsPreviousYear) / costOfGoodsPreviousYear) * 100;
								else
									PerChangeInCostOfGoods = costOfGoodsPreviousYear == 0 ? 0 : ((costOfGoodsPreviousYear - costOfGoodsPresentYear) / costOfGoodsPreviousYear) * 100;

								graphData.Add(Math.Round(PerChangeInCostOfGoods));

								//Repeat same for % change in frames cost                                

								lookuptbl = "Lookup.AvgFramesWholesaleCost_J";
								type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal costOfFramesPresentYear = GetAllLookUpValues("Lookup.AvgFramesWholesaleCost_J_" + currentYear).Where(x => x > 0).Average();
								decimal costOfFramesPreviousYear = GetAllLookUpValues("Lookup.AvgFramesWholesaleCost_J_" + previousYear).Where(x => x > 0).Average();

								decimal PerChangeInCostOfFrames = 0;
								if (costOfFramesPresentYear >= costOfFramesPreviousYear)
									PerChangeInCostOfFrames = ((costOfFramesPresentYear - costOfFramesPreviousYear) / costOfFramesPreviousYear) * 100;
								else
									PerChangeInCostOfFrames = ((costOfFramesPreviousYear - costOfFramesPresentYear) / costOfFramesPreviousYear) * 100;

								graphData.Add(Math.Round(PerChangeInCostOfFrames));

								string status = CreateHorizontalBarGraph(140, pictureName, graphHeader, yAxisData, graphData);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 250;
								}
							}
							if (i1 == 79) //additionalgraphs 2 ->3.	Average % change in gross revenue by change in total professional fees (Average % change versus prior year)
							{
								pictureName = imagelocation + @"\KeyMetrics_AvgPerCngInGrossRev_1.png";
								string graphHeader = "Average % change in gross revenue by change in total professional fees (Average % change versus prior year)";

								List<string> yAxisData = new List<string>() { "% Change in Gross Revenue", "% Change in Total professional fees" };

								List<decimal> graphData = new List<decimal>();

								string lookuptbl = "Lookup.AvgGrossRevenue_J";
								char type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal avgRevPresentYear = GetAllLookUpValues("Lookup.AvgGrossRevenue_J_" + currentYear).Where(x => x > 0).Average();
								decimal avgRevPreviousYear = GetAllLookUpValues("Lookup.AvgGrossRevenue_J_" + previousYear).Where(x => x > 0).Average();

								decimal perChangeInGrossRev = 0;
								if (avgRevPresentYear >= avgRevPreviousYear)
									perChangeInGrossRev = ((avgRevPresentYear - avgRevPreviousYear) / avgRevPreviousYear) * 100;
								else
									perChangeInGrossRev = ((avgRevPreviousYear - avgRevPresentYear) / avgRevPreviousYear) * 100;


								graphData.Add(Math.Round(perChangeInGrossRev));

								lookuptbl = "Lookup.TotalProfFees_J";
								type = 'A';
								PopulateBenchMarks(lookuptbl, currentYear, type);
								db.SaveChanges();

								PopulateBenchMarks(lookuptbl, previousYear, type);
								db.SaveChanges();

								decimal totalProfFeesPresentYear = GetAllLookUpValues("Lookup.TotalProfFees_J_" + currentYear).Where(x => x > 0).Average();
								decimal totalProfFeesPreviousYear = GetAllLookUpValues("Lookup.TotalProfFees_J_" + previousYear).Where(x => x > 0).Average();

								decimal PerChangeInTotalProfFees = 0;
								if (totalProfFeesPresentYear >= totalProfFeesPreviousYear)
									PerChangeInTotalProfFees = totalProfFeesPreviousYear == 0 ? 0 : ((totalProfFeesPresentYear - totalProfFeesPreviousYear) / totalProfFeesPreviousYear) * 100;
								else
									PerChangeInTotalProfFees = totalProfFeesPreviousYear == 0 ? 0 : ((totalProfFeesPreviousYear - totalProfFeesPresentYear) / totalProfFeesPreviousYear) * 100;

								graphData.Add(Math.Round(PerChangeInTotalProfFees));

								string status = CreateHorizontalBarGraph(130, pictureName, graphHeader, yAxisData, graphData);
								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 250;
								}
							}
							if (i1 == 80) //additionalgraphs 2 ->4.	Percentage of gross revenue from complete eye exams and from total product sales
							{
								pictureName = imagelocation + @"\KeyMetrics_PieChart_Graph_7.png";

								//percentages for pie chart
								//decimal completeEyeExamsPerc = grossRev == 0 ? 0 : ((Convert.ToDecimal(objReport.lstInput[0].colQ14 == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ14)).ToString("#,0"))) / grossRev) * 100;
								//decimal totalProdSalePerc = grossRev == 0 ? 0 : (((Convert.ToDecimal(objReport.lstInput[0].colQ26e == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26e)).ToString("#,0")))
								//						   + (Convert.ToDecimal(objReport.lstInput[0].colQ26f == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26f)).ToString("#,0")))
								//						   + (Convert.ToDecimal(objReport.lstInput[0].colQ26g == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26g)).ToString("#,0"))))
								//						   / grossRev) * 100;
								//decimal otherPerc = grossRev == 0 ? 0 : ((grossRev - (completeEyeExamsPerc + totalProdSalePerc)) / grossRev) * 100;

								decimal totalGrossRevenue = Math.Round(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 > 0).Select(x => x.Q24 ?? 0).Average());
								decimal completeEyeExamsPerc = totalGrossRevenue == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q26c > 0).Select(x => x.Q26c ?? 0).Average() / totalGrossRevenue) * 100;
								decimal totalProdSalePerc = totalGrossRevenue == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q26f > 0).Select(x => x.Q26f ?? 0).Average() / totalGrossRevenue +
									(db.Source_InputDataBenchMarkSource.Where(x => x.Q26g > 0).Select(x => x.Q26g ?? 0).Average()) / totalGrossRevenue + db.Source_InputDataBenchMarkSource.Where(x => x.Q26h > 0).Select(x => x.Q26h ?? 0).Average() / totalGrossRevenue) * 100;

								decimal otherPerc = 100 - (completeEyeExamsPerc + totalProdSalePerc);




								List<decimal> percentsList = new List<decimal>();
								percentsList.Add(Math.Round(completeEyeExamsPerc));
								percentsList.Add(Math.Round(totalProdSalePerc));
								percentsList.Add(Math.Round(otherPerc));

								//list of colors
								List<Color> pieColors = new List<Color>();
								pieColors.Add(Color.FromArgb(255, 51, 102, 153));
								pieColors.Add(Color.FromArgb(220, 51, 102, 153));
								pieColors.Add(Color.FromArgb(195, 51, 102, 153));


								//list of description of the slices

								List<string> descriptions = new List<string>();
								//descriptions.Add("Complete Eye Exams " + completeEyeExamsPerc + "%");
								//descriptions.Add("Total Product sales " + totalProdSalePerc + "%");
								//descriptions.Add("Other " + otherPerc + "%");
								descriptions.Add("Complete Eye Exams");
								descriptions.Add("Total Product sales");
								descriptions.Add("Other");

								string status = CreatePieChart(120, "Percent of gross revenue from \n complete eye exams and total product sales", percentsList, pieColors, descriptions, pictureName);

								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 400;
									shape.Height = 400;
								}

							}
							if (i1 == 81) //Instrument Penetration
							{
								pictureName = imagelocation + @"\KeyMetrics_InstrumentPene_Graph_1.png";
								string graphHeader = "Instrument Penetration (% of practices with one or more instruments)";
								List<decimal> graphData = new List<decimal>();
								List<string> yAxisData = new List<string>()
								{
									"Corneal tachymeter",
									"Retinal camera",
									"Nerve fiber analyzer",
									"(OCT, RTA, HRT, GDx)",
									"Corneal topographer",
									"Anterior segment camera",
									"Computerized refraction system",
									"Wide field scanning",
									"Laser ophthalmoscope (OPTOS)"
								};

								int totalPractices = db.Source_InputDataBenchMarkSource.Count();
								decimal cornealTachPerc = totalPractices == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q96a == true).Count() / totalPractices) * 100;
								decimal retineCameraPerc = totalPractices == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q96b == true).Count() / totalPractices) * 100;
								decimal nerveFibrePerc = totalPractices == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q96c == true).Count() / totalPractices) * 100;
								decimal octRtaPerc = totalPractices == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q96d == true).Count() / totalPractices) * 100;
								decimal cornealTopPerc = totalPractices == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q96e == true).Count() / totalPractices) * 100;
								decimal anteriorPerc = totalPractices == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q96f == true).Count() / totalPractices) * 100;
								decimal compRefPerc = totalPractices == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q96g == true).Count() / totalPractices) * 100;
								decimal wideFieldPerc = totalPractices == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q96h == true).Count() / totalPractices) * 100;
								decimal laserOpthPerc = totalPractices == 0 ? 0 : (db.Source_InputDataBenchMarkSource.Where(x => x.Q96i == true).Count() / totalPractices) * 100;

								graphData.Add(Math.Round(cornealTachPerc));
								graphData.Add(Math.Round(retineCameraPerc));
								graphData.Add(Math.Round(nerveFibrePerc));
								graphData.Add(Math.Round(octRtaPerc));
								graphData.Add(Math.Round(cornealTopPerc));
								graphData.Add(Math.Round(anteriorPerc));
								graphData.Add(Math.Round(compRefPerc));
								graphData.Add(Math.Round(wideFieldPerc));
								graphData.Add(Math.Round(laserOpthPerc));

								bool isXAxisDrawn = true;
								string status = CreateHorizontalBarGraph1(pictureName, graphHeader, yAxisData, graphData, isXAxisDrawn);

								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 500;
									shape.Height = 350;
								}
							}
							if (i1 == 84)
							{
								pictureName = imagelocation + @"\Spectacle_Graph_3.png";  // @"~\SurveyApp\SurveyApp\ImageChart\Spectacle_Graph_3.png";

								decimal plasticPerc = Convert.ToDecimal(objReport.lstInput[0].colQ89b == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ89b)).ToString("#,0"));
								decimal polycarbonatePerc = Convert.ToDecimal(objReport.lstInput[0].colQ89a == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ89a)).ToString("#,0"));
								decimal glassPerc = Convert.ToDecimal(objReport.lstInput[0].colQ89c == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ89c)).ToString("#,0"));

								List<decimal> percentsList = new List<decimal>();
								percentsList.Add(Math.Round(plasticPerc));
								percentsList.Add(Math.Round(polycarbonatePerc));
								percentsList.Add(Math.Round(glassPerc));

								//list of colors
								List<Color> pieColors = new List<Color>();
								pieColors.Add(Color.FromArgb(255, 51, 102, 153));
								pieColors.Add(Color.FromArgb(220, 51, 102, 153));
								pieColors.Add(Color.FromArgb(195, 51, 102, 153));


								//list of description of the slices

								List<string> descriptions = new List<string>();
								//descriptions.Add("Plastic " + plasticPerc + "%");
								//descriptions.Add("Polycarbonate " + polycarbonatePerc + "%");
								//descriptions.Add("Glass " + glassPerc + "%");
								descriptions.Add("Plastic");
								descriptions.Add("Polycarbonate");
								descriptions.Add("Glass");

								string status = CreatePieChart(100, "Material (% of lens pairs)", percentsList, pieColors, descriptions, pictureName);

								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 400;
									shape.Height = 400;
								}

							}
							if (i1 == 85)
							{
								pictureName = imagelocation + @"\Spectacle_Graph_2.png";  // @"~\SurveyApp\SurveyApp\ImageChart\Spectacle_Graph_1.png";

								//percentages for pie chart
								decimal progressivePerc = Convert.ToDecimal(objReport.lstInput[0].colQ90a == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ90a)).ToString("#,0"));
								decimal bifTrifocalPerc = Convert.ToDecimal(objReport.lstInput[0].colQ90b == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ90b)).ToString("#,0"));


								List<decimal> percentsList = new List<decimal>();
								percentsList.Add(Math.Round(progressivePerc));
								percentsList.Add(Math.Round(bifTrifocalPerc));

								//list of colors
								List<Color> pieColors = new List<Color>();
								pieColors.Add(Color.FromArgb(255, 51, 102, 153));
								pieColors.Add(Color.FromArgb(220, 51, 102, 153));

								//list of description of the slices

								List<string> descriptions = new List<string>();
								//descriptions.Add("Progressive " + progressivePerc + "%");
								//descriptions.Add("Bifocal / Trifocal " + bifTrifocalPerc + "%");
								descriptions.Add("Progressive");
								descriptions.Add("Bifocal / Trifocal");

								string status = CreatePieChart(100, "Design (% of lens pairs)", percentsList, pieColors, descriptions, pictureName);

								if (status == "success")
								{
									//aDoc.Bookmarks[i1].Range.Text = string.Empty;
									var shape = aDoc.Bookmarks[i1].Range.InlineShapes.AddPicture(pictureName, false, true);
									shape.Width = 400;
									shape.Height = 400;
								}

							}

						}

						//-- msinghai --- Find And Replace Placeholders
						this.FindAndReplace(wordApp, "<<DD>>", day);
						this.FindAndReplace(wordApp, "<<MM>>", month);
						this.FindAndReplace(wordApp, "<<YYYY>>", year);
						this.FindAndReplace(wordApp, "<<Customer_Name>>", objReport.lstInput[0].colQ74);
						decimal medGrossRev = GetLookUpValue("Lookup.AvgGrossRevenue_J", 50);
						decimal activePatients = objReport.lstInput[0].colQ12 == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ12));
						this.FindAndReplace(wordApp, "<<Q24_MD>>", medGrossRev);
						this.FindAndReplace(wordApp, "<<Q26d>>", objReport.lstInput[0].colQ26d == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26d)));
						this.FindAndReplace(wordApp, "<<Q26i>>", objReport.lstInput[0].colQ26i == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26i)));
						this.FindAndReplace(wordApp, "<<Q26f>>", objReport.lstInput[0].colQ26f == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26f)));
						this.FindAndReplace(wordApp, "<<Q26g>>", objReport.lstInput[0].colQ26g == null ? 0 : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ26g)));
						this.FindAndReplace(wordApp, "<<Q26f_MD>>", GetLookUpValue("Lookup.EyeGlassesSales_J", 50));
						this.FindAndReplace(wordApp, "<<CLSalesPercentGrossRev _MD>>", GetLookUpValue("Lookup.CLSalesPercentGrossRev", 50));
						this.FindAndReplace(wordApp, "<<CLWearerPercentActivePatients _MD>>", GetLookUpValue("Lookup.CLWearerPercentActivePatients", 50, "%"));
						this.FindAndReplace(wordApp, "<<AnnCLSalesPerCLExam_MD>>", GetLookUpValue("Lookup.AnnCLSalesPerCLExam", 50, "$"));
						this.FindAndReplace(wordApp, "<<CLRefitPercentCLExam _MD>>", GetLookUpValue("Lookup.CLRefitPercentCLExam", 50, "%"));
						this.FindAndReplace(wordApp, "<<CLGrossProfitMargin _MD>>", GetLookUpValue("Lookup.CLGrossProfitMargin", 50));
						this.FindAndReplace(wordApp, "<<PercentPatientsCLExamPurchEyewea _MD>>", GetLookUpValue("Lookup.PercentPatientsCLExamPurchEyewea", 50, "%"));
						this.FindAndReplace(wordApp, "<<ChairCostPerComplExam_MD>>", GetLookUpValue("Lookup.ChairCostPerComplExam", 50, "$"));
						this.FindAndReplace(wordApp, "<<NetIncomePercentGrossRev _MD>>", GetLookUpValue("Lookup.NetIncomePercentGrossRev", 50));
						this.FindAndReplace(wordApp, "<<NetIncomePercentGrossRev _75TH>>", GetLookUpValue("Lookup.NetIncomePercentGrossRev", 75));
						this.FindAndReplace(wordApp, "<<AnnMrktSpendPerComplExam_MD>>", GetLookUpValue("Lookup.AnnMrktSpendPerComplExam", 50, "$"));
						this.FindAndReplace(wordApp, "<<AcctRecDaysOutstanding_MD>>", GetLookUpValue("Lookup.AcctRecDaysOutstanding", 50));
						this.FindAndReplace(wordApp, "<<EyewearRxPer100ComplExam_MD>>", GetLookUpValue("Lookup.EyewearRxPer100ComplExam", 50));
						this.FindAndReplace(wordApp, "<<EyewearRxPer100ComplExam_TD>>", GetLookUpValue("Lookup.EyewearRxPer100ComplExam", 95));
						this.FindAndReplace(wordApp, "<<EyewearRxPer100ComplExam_BD>>", GetLookUpValue("Lookup.EyewearRxPer100ComplExam", 5));
						this.FindAndReplace(wordApp, "<<PrescriptionSunwearPercentofEyeWearRxes_J_MD>>", GetLookUpValue("Lookup.PrescriptionSunwearPercentofEyeWearRxes_J", 50));
						this.FindAndReplace(wordApp, "<<MedicalEyeCareVisitPercentTotal_MD>>", GetLookUpValue("Lookup.MedicalEyeCareVisitPercentTotal", 50));
						this.FindAndReplace(wordApp, "<<AnnMedEyeCareVisitPer1000_MD>>", GetLookUpValue("Lookup.AnnMedEyeCareVisitPer1000", 50));
						this.FindAndReplace(wordApp, "<<AnnMedEyeCareVisitPer1000_75th>>", GetLookUpValue("Lookup.AnnMedEyeCareVisitPer1000", 75));
						this.FindAndReplace(wordApp, "<<AnnMedEyeCareVisitPer1000_25th>>", GetLookUpValue("Lookup.AnnMedEyeCareVisitPer1000", 25));
						this.FindAndReplace(wordApp, "<<MultipleEyewearPurchasePercent_MD>>", GetLookUpValue("Lookup.MultipleEyewearPurchasePercent", 50, "%"));
						this.FindAndReplace(wordApp, "<<PercentExamsProvideWMangCareDis_MD>>", GetLookUpValue("Lookup.PercentExamsProvideWMangCareDis", 50, "%"));
						this.FindAndReplace(wordApp, "<<ExamFeeNonCL_MD>>", GetLookUpValue("Lookup.ExamFeeNonCL", 50, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_60P>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 60, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_40P>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 40, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_59P>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 59, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_39P>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 39, "$"));
						this.FindAndReplace(wordApp, "<<CompleteExamsPerODHour_60>>", GetLookUpValue("Lookup.CompleteExamsPerODHour", 60));
						this.FindAndReplace(wordApp, "<<CompleteExamsPerODHour_40>>", GetLookUpValue("Lookup.CompleteExamsPerODHour", 40));
						this.FindAndReplace(wordApp, "<<CompleteExamsPerODHour_59>>", GetLookUpValue("Lookup.CompleteExamsPerODHour", 59));
						this.FindAndReplace(wordApp, "<<CompleteExamsPerODHour_39>>", GetLookUpValue("Lookup.CompleteExamsPerODHour", 39));
						this.FindAndReplace(wordApp, "<<GrossRevPerNonODStaffHr_MD>>", GetLookUpValue("Lookup.GrossRevPerNonODStaffHr", 50, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevPerNonODStaffHr_70>>", GetLookUpValue("Lookup.GrossRevPerNonODStaffHr", 70, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevPerNonODStaffHr_30>>", GetLookUpValue("Lookup.GrossRevPerNonODStaffHr", 30, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevPerNonODStaffHr_69>>", GetLookUpValue("Lookup.GrossRevPerNonODStaffHr", 69, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevPerNonODStaffHr_29>>", GetLookUpValue("Lookup.GrossRevPerNonODStaffHr", 29, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerODHour_70>>", GetLookUpValue("Lookup.GrossRevenuePerODHour", 70, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerODHour_30>>", GetLookUpValue("Lookup.GrossRevenuePerODHour", 30, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerODHour_69>>", GetLookUpValue("Lookup.GrossRevenuePerODHour", 69, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerODHour_29>>", GetLookUpValue("Lookup.GrossRevenuePerODHour", 29, "$"));
						this.FindAndReplace(wordApp, "<<CompleteExamsPer100Active_70>>", GetLookUpValue("Lookup.CompleteExamsPer100Active", 70));
						this.FindAndReplace(wordApp, "<<CompleteExamsPer100Active_30>>", GetLookUpValue("Lookup.CompleteExamsPer100Active", 30));
						this.FindAndReplace(wordApp, "<<CompleteExamsPer100Active_69>>", GetLookUpValue("Lookup.CompleteExamsPer100Active", 69));
						this.FindAndReplace(wordApp, "<<CompleteExamsPer100Active_29>>", GetLookUpValue("Lookup.CompleteExamsPer100Active", 29));
						this.FindAndReplace(wordApp, "<<PercentofCompleteEyeExamsByHealthyEyeExams_J_MD>>", GetLookUpValue("Lookup.PercentofCompleteEyeExamsByHealthyEyeExams_J", 50));
						this.FindAndReplace(wordApp, "<<PercentofCompleteEyeExamsByHealthyEyeExams_J_TD>>", GetLookUpValue("Lookup.PercentofCompleteEyeExamsByHealthyEyeExams_J", 95));
						this.FindAndReplace(wordApp, "<<OpticalDispensaryPercentOfTotalOfficeSpace_J_MD>>", GetLookUpValue("Lookup.OpticalDispensaryPercentOfTotalOfficeSpace_J", 50));
						this.FindAndReplace(wordApp, "<<OpticalDispensaryPercentOfTotalOfficeSpace_J_75th>>", GetLookUpValue("Lookup.OpticalDispensaryPercentOfTotalOfficeSpace_J", 75));
						this.FindAndReplace(wordApp, "<<GrossRevPerSqFt_MD>>", GetLookUpValue("Lookup.GrossRevPerSqFt", 50, "$"));
						this.FindAndReplace(wordApp, "<<FramesAvgWholesaleCostPerFrame_J_MD>>", GetLookUpValue("Lookup.FramesAvgWholesaleCostPerFrame_J", 50));
						this.FindAndReplace(wordApp, "<<FramesAvgWholesaleCostPerFrame_J_25TH>>", GetLookUpValue("Lookup.FramesAvgWholesaleCostPerFrame_J", 25));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_MD>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 50, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_TD>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 95, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_BD>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 5, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_39>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 39, "$"));
						this.FindAndReplace(wordApp, "<<CompleteExamsPerODHour_MD>>", GetLookUpValue("Lookup.CompleteExamsPerODHour", 50));
						this.FindAndReplace(wordApp, "<<CompleteExamsPerODHour_TD>>", GetLookUpValue("Lookup.CompleteExamsPerODHour", 95));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerODHour_MD>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 50, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerODHour_TD>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 95, "$"));
						this.FindAndReplace(wordApp, "<<GrossRevenuePerODHour_BD>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 5, "$"));
						this.FindAndReplace(wordApp, "<<CompleteExamsPer100Active_MD>>", GetLookUpValue("Lookup.CompleteExamsPer100Active", 50));
						this.FindAndReplace(wordApp, "<<GrossRevPerActivePatient_MD>>", GetLookUpValue("Lookup.GrossRevPerActivePatient", 50, "$"));
						this.FindAndReplace(wordApp, "<<PercentofCompleteEyeExamsByHealthyEyeExams_J_MD>>", GetLookUpValue("Lookup.PercentofCompleteEyeExamsByHealthyEyeExams_J", 95));
						this.FindAndReplace(wordApp, "<<AnnMedEyeCareVisitPer1000_MD_Perct>>", activePatients == 0 ? 0 : GetLookUpValue("Lookup.AnnMedEyeCareVisitPer1000", 50) / activePatients);
						this.FindAndReplace(wordApp, "<<AnnMedEyeCareVisitPer1000_75th_Perct>>", activePatients == 0 ? 0 : GetLookUpValue("Lookup.AnnMedEyeCareVisitPer1000", 75) / activePatients);
						this.FindAndReplace(wordApp, "<<AnnMedEyeCareVisitPer1000_25th_Perct>>", activePatients == 0 ? 0 : GetLookUpValue("Lookup.AnnMedEyeCareVisitPer1000", 25) / activePatients);
						this.FindAndReplace(wordApp, "<<EyewearSalePercentageOfGrossRev_MD>>", GetLookUpValue("Lookup.EyewearSalePercentageOfGrossRev", 50));
						this.FindAndReplace(wordApp, "<<CLNewFitsPer100CLExam_MD>>", GetLookUpValue("Lookup.CLNewFitsPer100CLExam", 50));

						decimal eyewearSaleMed = GetLookUpValue("Lookup.GrossRevPerEyewearRx", 50, "$");
						this.FindAndReplace(wordApp, "<<GrossRevPerEyewearRx_MD>>", Math.Round(eyewearSaleMed, 2));

						decimal eyewearSaleTD = GetLookUpValue("Lookup.GrossRevPerEyewearRx", 95, "$");
						this.FindAndReplace(wordApp, "<<GrossRevPerEyewearRx_TD>>", eyewearSaleTD);
						this.FindAndReplace(wordApp, "<<GrossRevPerEyewearRx_BD>>", GetLookUpValue("Lookup.GrossRevPerEyewearRx", 95, "$"));

						this.FindAndReplace(wordApp, "<<GrossRevPerEyewearRx_MD%>>", Math.Round((eyewearSaleMed == 0 ? 100 : (eyewearSaleTD - eyewearSaleMed) / eyewearSaleMed)));
						this.FindAndReplace(wordApp, "<<EyewearGrossProfitMargin_MD>>", GetLookUpValue("Lookup.EyewearGrossProfitMargin", 50));
						decimal totalFramesInInv = db.Source_InputDataBenchMarkSource.Where(x => (x.Q34 > 0 || x.Q35 > 0)).Select(x => (x.Q34 + x.Q35) ?? 0).Average();
						this.FindAndReplace(wordApp, "<<Q34_And_Q35_MD>>", Math.Round(totalFramesInInv, 2));
						this.FindAndReplace(wordApp, "<<Q34_And_Q35_Times_Q36_All_MD>>", Math.Round(totalFramesInInv * db.Source_InputDataBenchMarkSource.Where(x => x.Q36 > 0).Select(x => x.Q36 ?? 0).Average(), 2));
						this.FindAndReplace(wordApp, "<<FramesUnitSalesMixbyPricePoint_Retail300Above>>", GetLookUpValue("Lookup.EyewearSalePercentageOfGrossRev", 50));
						this.FindAndReplace(wordApp, "<<AverageFramesMarkUp_MD>>", GetLookUpValue("lookup.averageframesmarkup_j", 50));
						this.FindAndReplace(wordApp, "<<StaffHourlyandAnnualSalariesbyPosition:2009_MedianHourlySalary>>", medAllSal);
						this.FindAndReplace(wordApp, "<<AnnualOccupancyCostperSquareFoot_MD>>", GetLookUpValue("Lookup.AnnualOccupancyCostperSquareFoot_J", 50));
						this.FindAndReplace(wordApp, "<<AnnualOccupancyCostperSquareFoot_20TH>>", GetLookUpValue("Lookup.AnnualOccupancyCostperSquareFoot_J", 20));
						this.FindAndReplace(wordApp, "<<AnnualOccupancyCostperSquareFoot_80TH>>", GetLookUpValue("Lookup.AnnualOccupancyCostperSquareFoot_J", 80));

						this.FindAndReplace(wordApp, "<<Q2_All_MD>>", GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q2 > 0).Select(x => x.Q2 ?? 0).ToList()));
						this.FindAndReplace(wordApp, "<<FramesTurnover_Median>>", Math.Round(annFrmsTrnovrMed, 2));
						this.FindAndReplace(wordApp, "<<PlanoSunglassInInv_Q38_Avg>>", Math.Round(db.Source_InputDataBenchMarkSource.Where(x => x.Q38 > 0).Select(x => x.Q38 ?? 0).Average()));
						this.FindAndReplace(wordApp, "<<ES_Gross_Revenue_per_Square_Foot_by_Practice_Size_for_refrection_MD>>", GetListMedian(GetExternalTableValues("ES_Gross_Revenue_per_Square_Foot_by_Practice_Size_for_refrection").Select(x => Convert.ToDecimal(x.Value)).ToList()));

						decimal softLensPerc = (db.Source_InputDataBenchMarkSource.Where(x => x.Q46a > 0).ToList().Count / db.Source_InputDataBenchMarkSource.Count()) * 100;
						this.FindAndReplace(wordApp, "<<SoftLens_Q46a_TotalInv%>>", softLensPerc);

						decimal softLensPerc1MilGR = (db.Source_InputDataBenchMarkSource.Where(x => x.Q46a > 0 && x.Q24 >= 1000000).ToList().Count / db.Source_InputDataBenchMarkSource.Where(x => x.Q24 >= 1000000).Count()) * 100;
						this.FindAndReplace(wordApp, "<<SoftLens_Q46a_GrossRev%>>", softLensPerc1MilGR);
						this.FindAndReplace(wordApp, "<<SpectacleLensRxes_Q31a>>", singleVisPerc);
						this.FindAndReplace(wordApp, "<<SpectacleLensRxes_Q31b>>", presBioPerc);

						this.FindAndReplace(wordApp, "<<ManagedCare_HealthVisionPlans>>", Math.Round(healthVisionPlansPerc));
						this.FindAndReplace(wordApp, "<<ManagedCare_Q27>>", Math.Round(medicarePerc));
						this.FindAndReplace(wordApp, "<<ManagedCare_Q27d>>", Math.Round(directPateintsPerc));
						this.FindAndReplace(wordApp, "<<AnnualSupplyPurchasebySoftLensModality_Q39b>>", Math.Round(twoWeekVal, 2));
						this.FindAndReplace(wordApp, "<<PercentPatientsCLExamPurchEyewea_MD>>", GetLookUpValue("Lookup.PercentPatientsCLExamPurchEyewea", 50, "%"));
						this.FindAndReplace(wordApp, "<<NetIncome%_GrossRev_Avg>>", Math.Round(GetAllLookUpValues("Lookup.NetIncomePercentGrossRev").Where(x => x > 0).Average()));
						this.FindAndReplace(wordApp, "<<FramesSold_Q28_MD>>", Math.Round(db.Source_InputDataBenchMarkSource.Where(x => (x.Q28 > 0 || x.Q29 > 0)).Select(x => (x.Q28 + x.Q29) ?? 0).Average(), 2));
						this.FindAndReplace(wordApp, "<<EyeWearGrossProfit>>", Math.Round(db.Source_InputDataBenchMarkSource.Where(x => x.Q26f > 0).Select(x => (x.Q26f - (x.Q52a + x.Q52b + x.Q52c + x.Q52d + x.Q52e)) ?? 0).Average(), 2));

						this.FindAndReplace(wordApp, "<<NoGlareLensPercentSpecLensRx_20TH>>", GetLookUpValue("Lookup.NoGlareLensPercentSpecLensRx", 20, "%"));
						this.FindAndReplace(wordApp, "<<NoGlareLensPercentSpecLensRx_80TH>>", GetLookUpValue("Lookup.NoGlareLensPercentSpecLensRx", 80, "%"));
						this.FindAndReplace(wordApp, "<<PrescriptionSunwearPercentofEyeWearRxes_J_75>>", GetLookUpValue("Lookup.PrescriptionSunwearPercentofEyeWearRxes_J", 75));
						this.FindAndReplace(wordApp, "<<MultipleEyewearPurchasePercent_80TH>>", GetLookUpValue("Lookup.PrescriptionSunwearPercentofEyeWearRxes_J", 80));
						this.FindAndReplace(wordApp, "<<CLGrossProfitMargin_20TH>>", GetLookUpValue("Lookup.CLGrossProfitMargin", 20));
						this.FindAndReplace(wordApp, "<<CLGrossProfitMargin_80TH>>", GetLookUpValue("Lookup.CLGrossProfitMargin", 80));
						//this.FindAndReplace(wordApp, "<<TotalCostOfGoods>>", GetLookUpValue("Lookup.GrossRevenuePerCompleteExam", 50));

						string vsnStr = "";
						if (singleVsnLensAvg > progVsnLensAvg)
							vsnStr = "more than";
						else if (singleVsnLensAvg < progVsnLensAvg)
							vsnStr = "less than";
						else
							vsnStr = "equal to";
						this.FindAndReplace(wordApp, "<<SpectacleLensMark-Ups_moreorless>>", vsnStr);

						#endregion

						#endregion Replace Word Documnet Tempalte's content.

						GenerateBarGraph(imagelocation, wordApp, aDoc, "GrossRevenuePerCompleteExam", "01", "Gross Revenue per Complete Exam Performance Deciles");
						GenerateBarGraph(imagelocation, wordApp, aDoc, "CompleteExamsPerODHour", "02", "Complete Exams per OD Hour Performance Deciles");
						GenerateBarGraph(imagelocation, wordApp, aDoc, "GrossRevPerNonODStaffHr", "03", "Gross Revenue per Staff Hour");
						GenerateBarGraph(imagelocation, wordApp, aDoc, "GrossRevenuePerODHour", "05", "Gross Revenue per OD Hour Performance Deciles");
						GenerateBarGraph(imagelocation, wordApp, aDoc, "CompleteExamsPer100Active", "06", "Complete Exams per 100 Active patients");
						GenerateBarGraph(imagelocation, wordApp, aDoc, "GrossRevPerActivePatient", "07", "Annual Gross Revenue per Active Patient");
						GenerateBarGraph(imagelocation, wordApp, aDoc, "EyewearRxPer100ComplExam", "29", "Eyewear Rxes per 100 Complete Exams Performance Deciles");
						GenerateBarGraph(imagelocation, wordApp, aDoc, "EyewearSalePercentageOfGrossRev", "30", "Eyewear % of Gross Revenue Performance Deciles");
						GenerateBarGraph(imagelocation, wordApp, aDoc, "GrossRevPerEyewearRx", "31", "Eyewear Revenue per Rx Performance Deciles");
						GenerateBarGraph(imagelocation, wordApp, aDoc, "EyewearGrossProfitMargin", "32", "Eyewear Rx Gross Profit Performance Deciles");
						GenerateBarGraph(imagelocation, wordApp, aDoc, "CLSalesPercentGrossRev", "50", "Contact Lens % of Gross Revenue Performance Deciles");

						//---------------------------------
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "MedicalEyeCareVisitPercentTotal", "13", "Medical Eye Care Visits % of Total Patient Visits");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "AnnMedEyeCareVisitPer1000", "14", "Annual Medical Eye Care Visits per 1,000 Active Patients");
						//<<PercentofExamsProvidedwithManagedCareDiscount_Graph>>
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "PercentExamsProvideWMangCareDis", "18", "Percent of Exams Provided with Managed Care Discount");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "PercentOfGrossRevenueDirectPatientPayments_J", "19", "Percent of Gross Revenue from Direct Patient Payments");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "PercentOfGrossRevenueAllHealthVisionPlans_J", "20", "Percent of Gross Revenue from All Health/Vision Plans");
						GenerateLinearGraph2(120, imagelocation, wordApp, aDoc, "PercentOfGrossRevenueVSPPayments_J", "21", "Percent of Gross Revenue from VSP Payments (included in total above)");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "PercentOfGrossRevenueFromMedicarePayments_J", "22", "Percent of Gross Revenue from Medicare Payments");
						//<<AverageCollectedExamRevenueperCompleteExam(direct-payandmanagedcare)_Graph>>
						GenerateLinearGraph2(120, imagelocation, wordApp, aDoc, "AvgCollectFeeRevPerCompl", "23", "Average Collected Exam Revenue per Complete Exam \n (direct-pay and managed care)");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "ExamFeeNonCL", "24", "Non-Contact Lens Exam (direct-pay)");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "ExamFeeSoftNewFitSPHERE", "25", "Contact Lens New Fit Exam – Sphere (direct-pay)");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "ExamFeeSoftNewFitTORIC", "26", "Contact Lens New Fit Exam – Soft Toric (direct-pay)");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "ExamFeeSoftNewFitMULTIFO", "27", "Contact Lens New Fit Exam – Soft Multifocal (direct-pay)");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "ExamFeeSoftLensNOREFITT", "28", "Contact Lens Exam – No Refitting (direct-pay)");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "OpticalDispensaryPercentOfTotalOfficeSpace_J", "33", "Optical Dispensary % of Total Office Space");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "MultipleEyewearPurchasePercent", "34", "Eyewear Multiple Pair Sales Ratio");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "ProgressiveLensAndPresbyopRx", "36", "Progressive Lenses (% of presbyopic Rxes)");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "NoGlareLensPercentSpecLensRx", "37", "No-Glare (anti-reflective) Lens (% of eyewear Rxes)");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "HighIndexLensPercentSpecLensRx", "38", "High Index Lenses (% of eyewear Rxes)");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "PhotochrLensPercentofSpecLensRx", "39", "Photochromic Lenses (% of eyewear Rxes)");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "PrescriptionSunwearPercentofEyeWearRxes_J", "40", "Prescription Sunwear (% of eyewear Rxes)");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "ComputerLensesPercentofEyeWearRxes_J", "41", "Computer Lenses (% of eyewear Rxes)");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "FramesAvgWholesaleCostPerFrame_J", "48", "Frames Average Wholesale Cost per Pair");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "CLWearerPercentActivePatients", "51", "Percent of Active Patients Wearing Contact Lenses");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "AnnCLSalesPerCLExam", "52", "Annual Contact Lens Sales per Contact Lens Eye Exam");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "SiliconeHydroLensWearPercentSoft", "53", "Silicone Hydrogel Wearer % of Soft Lens Wearers");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "DailyDisposableLensPercentSoft", "54", "Daily Disposable Wearer % of Contact Lens Wearers");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "SoftToricPercentSoftLens", "55", "Soft Toric Lens Wearer % of Contact Lens Wearers");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "SoftMultiFocPercentSoftLens", "56", "Soft Multifocal Lens Wearer % of Contact Lens Wearers");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "RGPLensWearerPercentOfCLWeares_J", "57", "RGP Lens Wearer % of Contact Lens Wearers");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "CLRefitPercentCLExam", "58", "Soft Lens Patient Refit Ratio");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "CLNewFitsPer100CLExam", "59", "Soft Lens New Fits per 100 Contact Lens Exams");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "CLGrossProfitMargin", "60", "Soft Lens Gross Profit Margin %");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "PercentPatientsCLExamPurchEyewea", "65", "Percent of Contact Lens Patients Purchasing Eyeglasses");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "ChairCostPerComplExam", "71", "Chair Cost per Complete Exam");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "AnnualOccupancyCostperSquareFoot_j", "72", "Annual Occupancy Cost per Square Foot");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "NetIncomePercentGrossRev", "73", "Net Income % of Gross Revenue");
						GenerateLinearGraph2(100, imagelocation, wordApp, aDoc, "AnnMrktSpendPerComplExam", "75", "Annual Marketing Spending per Complete Exam");
						GenerateLinearGraph2(120, imagelocation, wordApp, aDoc, "AcctRecDaysOutstanding", "76", "Accounts Receivables Aging % 60 days or more (By Performance Decile) Accounts Receivables");
						GenerateLinearGraph(100, imagelocation, wordApp, aDoc, "AnnTotalEyeExamsPer1000Patient", "78", "Annual total eye exams per 1,000 active Patients");
						GenerateLinearGraph3(100, imagelocation, wordApp, aDoc, "ES_Average_Frames_Mark_Up", "47", "Average Frames Mark-Up");


						GenerateSpectacleLinearGraph(100, imagelocation, wordApp, aDoc, "EyewearRxPer100ComplExam", "83", "Eyewear Rxes per 100 Complete Exams Performance Deciles");
						GenerateSpectacleLinearGraph(100, imagelocation, wordApp, aDoc, "GrossRevPerEyewearRx", "86", "Eyewear Revenue per Rx Performance Deciles");
						GenerateSpectacleLinearGraph(100, imagelocation, wordApp, aDoc, "ProgressiveLensAndPresbyopRx", "87", "Progressive Lenses (% of presbyopic Rxes)");
						GenerateSpectacleLinearGraph(100, imagelocation, wordApp, aDoc, "NoGlareLensPercentSpecLensRx", "88", "No-Glare (anti-reflective) Lens (% of eyewear Rxes)");
						GenerateSpectacleLinearGraph(100, imagelocation, wordApp, aDoc, "HighIndexLensPercentSpecLensRx", "89", "High Index Lenses (% of eyewear Rxes)");
						GenerateSpectacleLinearGraph(100, imagelocation, wordApp, aDoc, "PhotochrLensPercentofSpecLensRx", "90", "Photochromic Lenses (% of eyewear Rxes)");
						GenerateSpectacleLinearGraph(100, imagelocation, wordApp, aDoc, "PrescriptionSunwearPercentofEyeWearRxes_J", "92", "Prescription Sunwear (% of eyewear Rxes)");
						GenerateSpectacleLinearGraph(100, imagelocation, wordApp, aDoc, "ComputerLensesPercentofEyeWearRxes_J", "91", "Computer Lenses (% of eyewear Rxes)");
						GenerateSpectacleLinearGraph(100, imagelocation, wordApp, aDoc, "MultipleEyewearPurchasePercent", "93", "Eyewear Multiple Pair Sales % Eyewear Buyers Performance Deciles");


						GenerateLineGraphGrossRevenueperODHourbyPracticeSize(imagelocation, wordApp, aDoc, "GrossRevenuePerODHour", "04", "Gross Revenue per OD Hour by Practice Size");
						GenerateLineGraphPercentofPatientsDispensedfromInventorybyInventorySize(imagelocation, wordApp, aDoc, "ES_Percent_of_Patients_Dispensed_from_Inventory_by_Inventory_Size", "63", "Percent of Patients Dispensed from Inventory by Inventory Size (median)");
						GenerateLineGraphSoftLensInventorybyPracticeSize(imagelocation, wordApp, aDoc, "SoftLensInventorybyPracticeSize*(averageboxes)_Graph", "62", "Soft Lens Inventory by Practice Size* (average boxes)");
						GenerateLineGraphPlanoSunglassInventory(imagelocation, wordApp, aDoc, "PlanoSunglassInventory*byPracticeSize_Graph", "49", "Plano Sunglass Inventory* by Practice Size");
						GenerateLineGraphFramesMarkUpbyRetailPrice(imagelocation, wordApp, aDoc, "FramesMark-UpbyRetailPrice _Graph", "46", "Frames Mark-Up by Retail Price");
						//---------------------
						GenerateLineGraphGrossRevenueperSquareFootbyPracticeSize(imagelocation, wordApp, aDoc, "ES_Gross_Revenue_per_Square_Foot_by_Practice_Size_for_refrection", "Graph1", "08", "Gross Revenue per Square Foot by Practice Size");
						GenerateLineGraphGrossRevenueperSquareFootbyPracticeSizeforrefrection(imagelocation, wordApp, aDoc, "ES_Gross_Revenue_per_Square_Foot_by_Practice_Size_for_refrection", "Graph2", "08", "Gross Revenue per Square Foot by Practice Size");
						//---------------------

						GenerateLineGraphRangeofSquareFootagebyPracticeSizeLargestThird(imagelocation, wordApp, aDoc, "ES_Gross_Revenue_per_Square_Foot_by_Practice_Size_for_refrection", "Graph3", "09", "Range of Square Footage by Practice Size Largest Third");
						GenerateLineGraphRangeofSquareFootagebyPracticeSizeMediumThird(imagelocation, wordApp, aDoc, "ES_Gross_Revenue_per_Square_Foot_by_Practice_Size_for_refrection", "Graph2", "09", "Range of Square Footage by Practice Size Medium Third");
						GenerateLineGraphRangeofSquareFootagebyPracticeSizeSmallestThird(imagelocation, wordApp, aDoc, "ES_Gross_Revenue_per_Square_Foot_by_Practice_Size_for_refrection", "Graph1", "09", "Range of Square Footage by Practice Size Smallest Third");

						//---------------------
						GenerateLineGraphStaffingLevelsbyPracticeSizeOD(imagelocation, wordApp, aDoc, "", "Graph1", "66", "Staffing Levels by Practice Size");
						GenerateLineGraphStaffingLevelsbyPracticeSizenonOD(imagelocation, wordApp, aDoc, "", "Graph2", "66", "Staffing Levels by Practice Size");
						GenerateLineGraphNetIncomePercentGrossRev(imagelocation, wordApp, aDoc, "NetIncomePercentGrossRev", "74", "Net Income % of Gross Revenue by Practice Size");


						aDoc.Save();

						/*--Shahbaz.Need to enable when we go with Word to Pdf Report.*/
						//  Save document into PDF Format
						aDoc.SaveAs(ref outputFileName,
						ref fileFormat, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing,
						ref missing, ref missing, ref missing, ref missing);
						System.Threading.Thread.Sleep(5000);
						((_Application)wordApp).Quit(SaveChanges, ref missing, ref missing);

						// wordApp.Quit(ref missing, ref missing, ref missing);

						/*Muntajib-Remove to below if block to For both word & pdf*/

						//systesession["filepath"] = outputFileName;

						//((_Application)wordApp).Quit(SaveChanges, ref missing, ref missing);

						//((_Application)wordApp).Quit(ref missing, ref missing, ref missing);

						/*Muntajib-Remove to below if block to For both word & pdf*/
						System.Threading.Thread.Sleep(5000);
						if (File.Exists(tempPath))
						{
							File.Delete(tempPath);
						}

					}
					else
						return "File does not exist, please check and retry.";
				}
				string updatedStatus = UpdateReportGenerateStatus(objReport.lstOutput);
				if (updatedStatus == "success")
				{
					//var milliseconds = stopwatch.ElapsedMilliseconds;
					return "success";
				}
				else
				{
					return updatedStatus;
				}

			}
			catch (Exception ex)
			{

				// string filePath = @"C:\Error.txt";
				string filePath = null;
				if (ConfigurationManager.AppSettings["ErrorFilePath"] != null)
				{
					filePath = ConfigurationManager.AppSettings["ErrorFilePath"].ToString();// @"C:\Error.txt";
				}
				if (filePath != null)
				{

					using (StreamWriter writer = new StreamWriter(filePath, true))
					{
						writer.WriteLine("Messageexecutive :" + ex.Message);
						writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
					}
				}

				return ex.Message;
			}
		}

		public void GenerateBarGraph(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string lookUpTable, string bookMarkNum, string title)
		{
			System.Data.DataTable dt = new System.Data.DataTable();
			DataSet ds = new DataSet();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
			SqlConnection con = new SqlConnection(connStr);
			con.Open();
			var cmd = new SqlCommand();
			cmd.Connection = con;

			string tableName = "[dbo].[Lookup." + lookUpTable + "]";
			String strQuery = "select CONVERT(varchar(10), RowId-5)+'th' +CONVERT(varchar(1), '-')+" +
				" CONVERT(varchar(10), RowId + 4) + 'th' + ' percentile' as Heading, * from " + tableName +
	"where RowId like ('%5') or RowId like ('50') ";
			cmd.CommandText = strQuery;
			cmd.CommandType = CommandType.Text;
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			adp.Fill(ds);
			dt = ds.Tables[0];

			int maxRectSize = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1].ItemArray[2]);
			string resize = string.Empty;
			int bitmapsizeXaxis = 1020 + 550;
			int bitmapsizeYaxis = 830;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Bitmap barBitmap1 = new Bitmap(1020 + 550, 100);
			Graphics objGraphic1 = Graphics.FromImage(barBitmap1);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.FromArgb(255, 102, 204, 51));
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			string sngHighestValueNewString = string.Empty;
			Single sngHeight1New = new Single();
			int gapValue = 50;

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial",40, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 20);
			objGraphic.FillRectangle(yellowBrush, 0, 450, (bitmapsizeXaxis - 50), 50);

			objGraphic.DrawString("Lowest", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 20, (bitmapsizeYaxis - 70));

			//if (title == "GrossRevenuePerCompleteExam")
			//{
			//	this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_MD>>", Convert.ToDouble(dt.Rows[5].ItemArray[2].ToString()));
			//	this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_TD>>", Convert.ToDouble(dt.Rows[0].ItemArray[2].ToString()));
			//	this.FindAndReplace(wordApp, "<<GrossRevenuePerCompleteExam_BD>>", Convert.ToDouble(dt.Rows[10].ItemArray[2].ToString()));
			//}
			//if (title == "GrossRevenuePerCompleteExam")
			//{
			//	this.FindAndReplace(wordApp, "<<CompleteExamsPerODHour_MD>>", Convert.ToDouble(dt.Rows[5].ItemArray[2].ToString()));
			//	this.FindAndReplace(wordApp, "<<CompleteExamsPerODHour_TD>>", Convert.ToDouble(dt.Rows[0].ItemArray[2].ToString()));
			//}
			//if (title == "GrossRevenuePerCompleteExam")
			//{
			//	this.FindAndReplace(wordApp, "<<GrossRevenuePerODHour_MD>>", Convert.ToDouble(dt.Rows[5].ItemArray[2].ToString()));
			//	this.FindAndReplace(wordApp, "<<GrossRevenuePerODHour_TD>>", Convert.ToDouble(dt.Rows[0].ItemArray[2].ToString()));
			//	this.FindAndReplace(wordApp, "<<GrossRevenuePerODHour_BD>>", Convert.ToDouble(dt.Rows[0].ItemArray[2].ToString()));
			//}
			//if (title == "GrossRevenuePerCompleteExam")
			//{
			//	this.FindAndReplace(wordApp, "<<CompleteExamsPer100Active_MD>>", Convert.ToDouble(dt.Rows[5].ItemArray[2].ToString()));
			//}
			//if (title == "GrossRevenuePerCompleteExam")
			//{
			//	this.FindAndReplace(wordApp, "<<GrossRevPerActivePatient_MD>>", Convert.ToDouble(dt.Rows[5].ItemArray[2].ToString()));
			//}
			//if (title == "EyewearRxPer100ComplExam")
			//{
			//	this.FindAndReplace(wordApp, "<<EyewearRxPer100ComplExam_MD>>", Convert.ToDouble(dt.Rows[5].ItemArray[2].ToString()));
			//	this.FindAndReplace(wordApp, "<<EyewearRxPer100ComplExam_TD>>", Convert.ToDouble(dt.Rows[0].ItemArray[2].ToString()));
			//	this.FindAndReplace(wordApp, "<<EyewearRxPer100ComplExam_BD>>", Convert.ToDouble(dt.Rows[0].ItemArray[2].ToString()));
			//}
			//if (title == "EyewearSalePercentageOfGrossRev")
			//{
			//	this.FindAndReplace(wordApp, "<<EyewearSalePercentageOfGrossRev_MD>>", Convert.ToDouble(dt.Rows[5].ItemArray[2].ToString()));
			//}
			//if (title == "GrossRevPerEyewearRx")
			//{
			//	this.FindAndReplace(wordApp, "<<GrossRevPerEyewearRx_MD>>", Convert.ToDouble(dt.Rows[5].ItemArray[2].ToString()));
			//	this.FindAndReplace(wordApp, "<<GrossRevPerEyewearRx_TD>>", Convert.ToDouble(dt.Rows[0].ItemArray[2].ToString()));
			//	this.FindAndReplace(wordApp, "<<GrossRevPerEyewearRx_MD%>>", Convert.ToDouble(dt.Rows[5].ItemArray[2].ToString()));
			//}
			//if (title == "EyewearGrossProfitMargin")
			//{
			//	this.FindAndReplace(wordApp, "<<TotalCostOfGoods>>", Convert.ToDouble(dt.Rows[5].ItemArray[2].ToString()));
			//	//this.FindAndReplace(wordApp, "<<EyewearGrossProfitMargin_MD%>>", Convert.ToDouble(dt.Rows[0].ItemArray[2].ToString()));
			//}



			foreach (DataRow row in dt.Rows)
			{
				DataColumnCollection columns = dt.Columns;
				if (columns.Contains("LookupValue$"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue$"]);
					sngHighestValueNewString = "$" + sngHighestValueNew.ToString();

					if (sngHighestValueNew <= 0)
						sngHeight1New = 0;
					else
						sngHeight1New = Convert.ToInt32((Convert.ToSingle(row["LookupValue$"]) / (Convert.ToInt32(dt.Rows[dt.Rows.Count - 1].ItemArray[2])) * 800));

				}

				if (columns.Contains("LookupValue"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue"]);
					sngHighestValueNewString = sngHighestValueNew.ToString();

					if (lookUpTable == "EyewearSalePercentageOfGrossRev")
					{
						sngHighestValueNewString = (sngHighestValueNew).ToString() + "%";
					}
					else if (lookUpTable == "EyewearGrossProfitMargin")
					{
						sngHighestValueNewString = (sngHighestValueNew).ToString() + "%";
					}
					else if (lookUpTable == "CLSalesPercentGrossRev")
					{
						sngHighestValueNewString = (sngHighestValueNew).ToString() + "%";
					}


					if (sngHighestValueNew <= 0)
						sngHeight1New = 0;
					else
						sngHeight1New = Convert.ToInt32((Convert.ToSingle(row["LookupValue"]) / (Convert.ToSingle(dt.Rows[dt.Rows.Count - 1].ItemArray[2])) * 800));
				}

				if (row["Heading"].ToString() == "45th-54th percentile")
				{
					objGraphic.DrawString("Median", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 20, (bitmapsizeYaxis - 70) - gapValue);
				}

				else
					if (row["Heading"].ToString() != "0th-9th percentile")
					objGraphic.DrawString(row["Heading"].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 30, (bitmapsizeYaxis - 70) - gapValue);

				if (row["Heading"].ToString() == "0th-9th percentile")
				{
					objGraphic.DrawString("1st-9th percentile", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 30, (bitmapsizeYaxis - 70) - gapValue);
				}

				objGraphic.FillRectangle(lightblueBrush, 320, ((bitmapsizeYaxis - 70) - gapValue), sngHeight1New, 30);
				objGraphic.DrawString(sngHighestValueNewString, new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, sngHeight1New + 10 + 200 + 120, (bitmapsizeYaxis - 70) - gapValue);

				if (lookUpTable != "EyewearSalePercentageOfGrossRev")
					objGraphic.DrawString(Convert.ToInt32(((Convert.ToDouble(row.ItemArray[2].ToString()) / Convert.ToDouble(dt.Rows[5].ItemArray[2])) * Convert.ToDouble(100))).ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, bitmapsizeXaxis - 150 + 20, (bitmapsizeYaxis - 70) - gapValue);

				gapValue = gapValue + 50;
			}

			objGraphic.DrawString("Highest", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 20, (bitmapsizeYaxis - 70) - gapValue);
			if (lookUpTable != "EyewearSalePercentageOfGrossRev")
				objGraphic.DrawString("Index vs. Median", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, bitmapsizeXaxis - 250, (bitmapsizeYaxis - 70) - gapValue);

			objGraphic.DrawLine(blackPen, 320, (bitmapsizeYaxis - 70), (bitmapsizeXaxis - 50), (bitmapsizeYaxis - 70));
			objGraphic.DrawLine(blackPen, 320, 130, 320, (bitmapsizeYaxis - 70));

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));


			string filepath = imagelocation + @"\bargraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 300;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();

		}

		public void GenerateLinearGraph(int headerHeight, string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string lookUpTable, string bookMarkNum, string title)
		{
			System.Data.DataTable dt = new System.Data.DataTable();
			DataSet ds = new DataSet();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
			SqlConnection con = new SqlConnection(connStr);
			con.Open();
			var cmd = new SqlCommand();
			cmd.Connection = con;

			string tableName = "[dbo].[Lookup." + lookUpTable + "]";
			String strQuery = "select * from " + tableName +
	"where LookupLable like ('%5th') or LookupLable like ('%50th') ";
			cmd.CommandText = strQuery;
			cmd.CommandType = CommandType.Text;
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			adp.Fill(ds);
			dt = ds.Tables[0];

			int bitmapsizeXaxis = 1250;
			int bitmapsizeYaxis = 500;
			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush lightGrayBrush = new SolidBrush(Color.LightGray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, headerHeight));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();
			int gapValue = 0;

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = headerHeight, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 20);

			string sngHighestValueNewString = string.Empty;
			int i = 0;
			int average = 0;
			string type = string.Empty;
			foreach (DataRow row in dt.Rows)
			{
				DataColumnCollection columns = dt.Columns;
				if (columns.Contains("LookupValue$"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue$"]);
					sngHighestValueNewString = "$" + Math.Round(sngHighestValueNew, 1).ToString();
					type = "$";
				}
				if (columns.Contains("LookupValue"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue"]);
					sngHighestValueNewString = Math.Round(sngHighestValueNew, 1).ToString();
					type = string.Empty;
				}

				if (columns.Contains("LookupValue%"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue%"]);
					if (sngHighestValueNew > 100)
					{
						sngHighestValueNew = 100;
					}
					sngHighestValueNewString = Math.Round(sngHighestValueNew, 1).ToString() + "%";
					type = "%";
				}

				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;

				sngHeight1New = sngHighestValueNew;

				int RowId = 0;
				if (row["LookupLable"].ToString().Split('-').Count() == 1)
					RowId = Convert.ToInt32(row["LookupLable"].ToString().Split('-')[0].Replace("th", ""));
				if (row["LookupLable"].ToString().Split('-').Count() == 2)
					RowId = Convert.ToInt32(row["LookupLable"].ToString().Split('-')[1].Replace("th", ""));

				if (RowId == 5)
				{
					objGraphic.DrawString("Improvement", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 68, 140);
					objGraphic.DrawString("Oppurtunity", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 68, 165);
				}

				if (RowId == 50)
				{
					objGraphic.DrawString("Median", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 575, 165);
				}
				if (RowId == 85)
				{
					objGraphic.DrawString("High", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 1020, 140);
					objGraphic.DrawString("Performance", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 1020, 165);
				}

				Brush myBrush = Brushes.LightBlue;
				switch (i)
				{
					case 0:
						myBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
						break;
					case 1:
						myBrush = new SolidBrush(Color.FromArgb(240, 51, 102, 153));
						break;
					case 2:
						myBrush = new SolidBrush(Color.FromArgb(225, 51, 102, 153));
						break;
					case 3:
						myBrush = new SolidBrush(Color.FromArgb(210, 51, 102, 153));
						break;
					case 4:
						myBrush = new SolidBrush(Color.FromArgb(195, 51, 102, 153));
						break;
					case 5:
						myBrush = new SolidBrush(Color.FromArgb(255, 102, 204, 51));
						break;
					case 6:
						myBrush = new SolidBrush(Color.FromArgb(180, 51, 102, 153));
						break;
					case 7:
						myBrush = new SolidBrush(Color.FromArgb(165, 51, 102, 153));
						break;
					case 8:
						myBrush = new SolidBrush(Color.FromArgb(150, 51, 102, 153));
						break;
					case 9:
						myBrush = new SolidBrush(Color.FromArgb(135, 51, 102, 153));
						break;
					case 10:
						myBrush = new SolidBrush(Color.FromArgb(120, 51, 102, 153));
						break;
				}
				objGraphic.FillRectangle(myBrush, 80 + (gapValue - 10), 200, 98, 98);
				objGraphic.FillRectangle(lightblueBrush, 70 + gapValue, 200 + 100, 100, 100);
				objGraphic.DrawString(sngHighestValueNewString, new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 100 + gapValue + 10, 250, format);
				objGraphic.DrawString(RowId.ToString() + "th", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 70 + gapValue + 20, 300);

				//if (RowId == 50)
				//{
				//	objGraphic.DrawString("Percentile", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
				//	blackBrush, 570, 330);
				//	objGraphic.DrawString("Ranking", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
				//blackBrush, 570, 350);

				//}
				gapValue = gapValue + 100;
				i++;
				average = average + Convert.ToInt32(sngHighestValueNew);
			}
			objGraphic.DrawString("Percentile Ranking", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
			blackBrush, 530, 345);
			average = average / 11;
			if (type == "$")
				objGraphic.DrawString("AVERAGE=" + type + average.ToString(), new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 530, 370);
			else if (type == "%")
				objGraphic.DrawString("AVERAGE=" + average.ToString() + type, new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 530, 370);
			else
				objGraphic.DrawString("AVERAGE=" + average.ToString(), new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 530, 370);

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\lineargraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 500;
			shape1.Height = 200;
			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();
		}

		public void GenerateLinearGraph2(int headerHeight, string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string lookUpTable, string bookMarkNum, string title)
		{
			System.Data.DataTable dt = new System.Data.DataTable();
			DataSet ds = new DataSet();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
			SqlConnection con = new SqlConnection(connStr);
			con.Open();
			var cmd = new SqlCommand();
			cmd.Connection = con;

			string tableName = "[dbo].[Lookup." + lookUpTable + "]";
			String strQuery = "select * from " + tableName +
	"where LookupLable like ('%5th') or LookupLable like ('%50th') ";
			cmd.CommandText = strQuery;
			cmd.CommandType = CommandType.Text;
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			adp.Fill(ds);
			dt = ds.Tables[0];

			int bitmapsizeXaxis = 1250;
			int bitmapsizeYaxis = 500;
			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush lightGrayBrush = new SolidBrush(Color.LightGray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, headerHeight));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();
			int gapValue = 0;

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = headerHeight, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 20);

			string sngHighestValueNewString = string.Empty;
			int i = 0;
			int average = 0;
			string type = string.Empty;
			foreach (DataRow row in dt.Rows)
			{
				DataColumnCollection columns = dt.Columns;
				if (columns.Contains("LookupValue$"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue$"]);
					sngHighestValueNewString = "$" + Math.Round(sngHighestValueNew, 1).ToString();
					type = "$";
				}
				if (columns.Contains("LookupValue"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue"]);
					sngHighestValueNewString = Math.Round(sngHighestValueNew, 1).ToString();
					type = string.Empty;
				}

				if (columns.Contains("LookupValue%"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue%"]);
					if (sngHighestValueNew > 100)
					{
						sngHighestValueNew = 100;
					}
					sngHighestValueNewString = Math.Round(sngHighestValueNew, 1).ToString() + "%";
					type = "%";
				}


				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;

				int RowId = 0;
				if (row["LookupLable"].ToString().Split('-').Count() == 1)
					RowId = Convert.ToInt32(row["LookupLable"].ToString().Split('-')[0].Replace("th", ""));
				if (row["LookupLable"].ToString().Split('-').Count() == 2)
					RowId = Convert.ToInt32(row["LookupLable"].ToString().Split('-')[1].Replace("th", ""));

				sngHeight1New = sngHighestValueNew;
				if (RowId == 5)
				{
					objGraphic.DrawString("Low", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 68, 165);
				}

				if (RowId == 50)
				{
					objGraphic.DrawString("Median", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 575, 165);
				}
				if (RowId == 85)
				{
					objGraphic.DrawString("High", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 1110, 165);
				}

				Brush myBrush = Brushes.LightBlue;
				switch (i)
				{
					case 0:
						myBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
						break;
					case 1:
						myBrush = new SolidBrush(Color.FromArgb(240, 51, 102, 153));
						break;
					case 2:
						myBrush = new SolidBrush(Color.FromArgb(225, 51, 102, 153));
						break;
					case 3:
						myBrush = new SolidBrush(Color.FromArgb(210, 51, 102, 153));
						break;
					case 4:
						myBrush = new SolidBrush(Color.FromArgb(195, 51, 102, 153));
						break;
					case 5:
						myBrush = new SolidBrush(Color.FromArgb(255, 102, 204, 51));
						break;
					case 6:
						myBrush = new SolidBrush(Color.FromArgb(180, 51, 102, 153));
						break;
					case 7:
						myBrush = new SolidBrush(Color.FromArgb(165, 51, 102, 153));
						break;
					case 8:
						myBrush = new SolidBrush(Color.FromArgb(150, 51, 102, 153));
						break;
					case 9:
						myBrush = new SolidBrush(Color.FromArgb(135, 51, 102, 153));
						break;
					case 10:
						myBrush = new SolidBrush(Color.FromArgb(120, 51, 102, 153));
						break;
				}
				objGraphic.FillRectangle(myBrush, 80 + (gapValue - 10), 200, 98, 98);
				objGraphic.FillRectangle(lightblueBrush, 70 + gapValue, 200 + 100, 100, 100);
				objGraphic.DrawString(sngHighestValueNewString, new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 100 + gapValue + 10, 250, format);
				objGraphic.DrawString(RowId.ToString() + "th", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 70 + gapValue + 20, 300);

				//if (RowId == 50)
				//{
				//	objGraphic.DrawString("Percentile", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
				//	blackBrush, 570, 330);
				//	objGraphic.DrawString("Ranking", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
				//blackBrush, 570, 350);

				//}
				gapValue = gapValue + 100;
				i++;
				average = average + Convert.ToInt32(sngHighestValueNew);
			}

			objGraphic.DrawString("Percentile Ranking", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
			blackBrush, 530, 345);
			average = average / 11;

			if (type == "$")
				objGraphic.DrawString(type + "AVERAGE=" + average.ToString(), new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 530, 370);
			else if (type == "%")
				objGraphic.DrawString("AVERAGE=" + average.ToString() + type, new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 530, 370);
			else
				objGraphic.DrawString("AVERAGE=" + average.ToString(), new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 530, 370);

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\lineargraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 500;
			shape1.Height = 200;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLinearGraph3(int headerHeight, string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string bookMarkNum, string title)
		{
			System.Data.DataTable dt = new System.Data.DataTable();
			DataSet ds = new DataSet();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];

			SqlConnection con = new SqlConnection(connStr);

			con.Open();

			var cmd = new SqlCommand();
			cmd.Connection = con;
			string tableName = "[PPASurvey_DBProd].[dbo].[" + eSTable + "]";
			String strQuery = "select  * from " + tableName;
			cmd.CommandText = strQuery;
			cmd.CommandType = CommandType.Text;
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			adp.Fill(ds);
			dt = ds.Tables[0];
			int bitmapsizeXaxis = 1250;
			int bitmapsizeYaxis = 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush lightGrayBrush = new SolidBrush(Color.LightGray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, headerHeight));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();
			int gapValue = 0;

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = headerHeight, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 20);

			string sngHighestValueNewString = string.Empty;
			int i = 0;
			foreach (DataRow row in dt.Rows)
			{
				foreach (DataColumn column in dt.Columns)
				{
					DataColumnCollection columns = dt.Columns;
					sngHighestValueNew = Convert.ToSingle(row[column].ToString().Replace("x", ""));
					sngHighestValueNewString = row[column].ToString();

					if (sngHighestValueNew == 0)
						sngHighestValueNew = 1;

					int RowId = 0;
					RowId = Convert.ToInt32(column.ToString().Split('-')[0].Replace("th", ""));

					sngHeight1New = sngHighestValueNew;

					if (RowId == 5)
					{
						objGraphic.DrawString("Low", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 68, 165);
					}

					if (RowId == 50)
					{
						objGraphic.DrawString("Median", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 575, 165);
					}
					if (RowId == 85)
					{
						objGraphic.DrawString("High", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 1110, 165);
					}

					Brush myBrush = Brushes.LightBlue;
					switch (i)
					{
						case 0:
							myBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
							break;
						case 1:
							myBrush = new SolidBrush(Color.FromArgb(240, 51, 102, 153));
							break;
						case 2:
							myBrush = new SolidBrush(Color.FromArgb(225, 51, 102, 153));
							break;
						case 3:
							myBrush = new SolidBrush(Color.FromArgb(210, 51, 102, 153));
							break;
						case 4:
							myBrush = new SolidBrush(Color.FromArgb(195, 51, 102, 153));
							break;
						case 5:
							myBrush = new SolidBrush(Color.FromArgb(255, 102, 204, 51));
							break;
						case 6:
							myBrush = new SolidBrush(Color.FromArgb(180, 51, 102, 153));
							break;
						case 7:
							myBrush = new SolidBrush(Color.FromArgb(165, 51, 102, 153));
							break;
						case 8:
							myBrush = new SolidBrush(Color.FromArgb(150, 51, 102, 153));
							break;
						case 9:
							myBrush = new SolidBrush(Color.FromArgb(135, 51, 102, 153));
							break;
						case 10:
							myBrush = new SolidBrush(Color.FromArgb(120, 51, 102, 153));
							break;
					}
					objGraphic.FillRectangle(myBrush, 80 + (gapValue - 10), 200, 98, 98);
					objGraphic.FillRectangle(lightblueBrush, 70 + gapValue, 200 + 100, 100, 100);
					objGraphic.DrawString(sngHighestValueNewString, new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 100 + gapValue + 10, 250, format);
					objGraphic.DrawString(RowId.ToString() + "th", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 70 + gapValue + 20, 300);

					//if (RowId == 50)
					//{
					//	objGraphic.DrawString("Percentile", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
					//	blackBrush, 570, 330);
					//	objGraphic.DrawString("Ranking", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
					//blackBrush, 570, 350);

					//}
					gapValue = gapValue + 100;
					i++;
				}
			}

			objGraphic.DrawString("Percentile Ranking", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
			blackBrush, 530, 345);
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\lineargraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 500;
			shape1.Height = 200;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateSpectacleLinearGraph(int headerHeight, string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string lookUpTable, string bookMarkNum, string title)
		{
			System.Data.DataTable dt = new System.Data.DataTable();
			DataSet ds = new DataSet();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
			SqlConnection con = new SqlConnection(connStr);
			con.Open();
			var cmd = new SqlCommand();
			cmd.Connection = con;

			string tableName = "[dbo].[Lookup." + lookUpTable + "]";
			String strQuery = "select * from " + tableName +
	"where LookupLable like ('%5th') or LookupLable like ('%50th') ";
			cmd.CommandText = strQuery;
			cmd.CommandType = CommandType.Text;
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			adp.Fill(ds);
			dt = ds.Tables[0];

			int bitmapsizeXaxis = 1250;
			int bitmapsizeYaxis = 500;
			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush lightGrayBrush = new SolidBrush(Color.LightGray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, headerHeight));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();
			int gapValue = 0;

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = headerHeight, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 20);

			string sngHighestValueNewString = string.Empty;
			int i = 0;
			int average = 0;
			string type = string.Empty;
			foreach (DataRow row in dt.Rows)
			{
				DataColumnCollection columns = dt.Columns;
				if (columns.Contains("LookupValue$"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue$"]);
					sngHighestValueNewString = "$" + Math.Round(sngHighestValueNew, 1).ToString();
					type = "$";
				}
				if (columns.Contains("LookupValue"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue"]);
					sngHighestValueNewString = Math.Round(sngHighestValueNew, 1).ToString();
					type = string.Empty;
				}

				if (columns.Contains("LookupValue%"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue%"]);
					if (sngHighestValueNew > 100)
					{
						sngHighestValueNew = 100;
					}
					sngHighestValueNewString = Math.Round(sngHighestValueNew, 1).ToString() + "%";
					type = "%";
				}

				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;

				sngHeight1New = sngHighestValueNew;

				int RowId = 0;
				if (row["LookupLable"].ToString().Split('-').Count() == 1)
					RowId = Convert.ToInt32(row["LookupLable"].ToString().Split('-')[0].Replace("th", ""));
				if (row["LookupLable"].ToString().Split('-').Count() == 2)
					RowId = Convert.ToInt32(row["LookupLable"].ToString().Split('-')[1].Replace("th", ""));

				if (RowId == 5)
				{
					objGraphic.DrawString("Improvement", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 68, 140);
					objGraphic.DrawString("Oppurtunity", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 68, 165);
				}

				if (RowId == 50)
				{
					objGraphic.DrawString("Median", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 575, 165);
				}
				if (RowId == 85)
				{
					objGraphic.DrawString("High", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 1020, 140);
					objGraphic.DrawString("Performance", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 1020, 165);
				}

				Brush myBrush = Brushes.LightBlue;
				switch (i)
				{
					case 0:
						myBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
						break;
					case 1:
						myBrush = new SolidBrush(Color.FromArgb(240, 51, 102, 153));
						break;
					case 2:
						myBrush = new SolidBrush(Color.FromArgb(225, 51, 102, 153));
						break;
					case 3:
						myBrush = new SolidBrush(Color.FromArgb(210, 51, 102, 153));
						break;
					case 4:
						myBrush = new SolidBrush(Color.FromArgb(195, 51, 102, 153));
						break;
					case 5:
						myBrush = new SolidBrush(Color.FromArgb(255, 102, 204, 51));
						break;
					case 6:
						myBrush = new SolidBrush(Color.FromArgb(180, 51, 102, 153));
						break;
					case 7:
						myBrush = new SolidBrush(Color.FromArgb(165, 51, 102, 153));
						break;
					case 8:
						myBrush = new SolidBrush(Color.FromArgb(150, 51, 102, 153));
						break;
					case 9:
						myBrush = new SolidBrush(Color.FromArgb(135, 51, 102, 153));
						break;
					case 10:
						myBrush = new SolidBrush(Color.FromArgb(120, 51, 102, 153));
						break;
				}
				objGraphic.FillRectangle(myBrush, 80 + (gapValue - 10), 200, 98, 98);
				objGraphic.FillRectangle(lightblueBrush, 70 + gapValue, 200 + 100, 100, 100);
				objGraphic.DrawString(sngHighestValueNewString, new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 100 + gapValue + 10, 250, format);
				objGraphic.DrawString(RowId.ToString() + "th", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 70 + gapValue + 20, 300);

				//if (RowId == 50)
				//{
				//	objGraphic.DrawString("Percentile", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
				//	blackBrush, 570, 330);
				//	objGraphic.DrawString("Ranking", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
				//blackBrush, 570, 350);

				//}
				gapValue = gapValue + 100;
				i++;
			}

			objGraphic.DrawString("Percentile Ranking", new System.Drawing.Font("Arial", 22, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 530, 345);
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\Spectaclelineargraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);

			shape1.Width = 500;
			shape1.Height = 200;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}




		public void GenerateLineGraphGrossRevenueperODHourbyPracticeSize(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string lookUpTable, string bookMarkNum, string title)
		{
			System.Data.DataTable dt = new System.Data.DataTable();
			DataSet ds = new DataSet();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
			SqlConnection con = new SqlConnection(connStr);
			con.Open();
			var cmd = new SqlCommand();
			cmd.Connection = con;

			string tableName = "[dbo].[Lookup." + lookUpTable + "]";
			String strQuery = "select CONVERT(varchar(10), RowId-5)+'th' +CONVERT(varchar(1), '-')+" +
				" CONVERT(varchar(10), RowId + 4) + 'th' + ' percentile' as Heading, * from " + tableName +
	"where RowId like ('%5') ";

			cmd.CommandText = strQuery;
			cmd.CommandType = CommandType.Text;
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			adp.Fill(ds);
			dt = ds.Tables[0];

			int length = Convert.ToInt32(Math.Round((Convert.ToInt32(dt.Rows[dt.Rows.Count - 1][2].ToString().Split('.')[0])) / 6 / 100d, 0) * 100);

			int bitmapsizeXaxis = 1400;
			int bitmapsizeYaxis = 600 + 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();


			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);
			int i = 1;
			int j = 0;
			int x = 300;

			objGraphic.DrawString("Gross Revenue per OD Hour (median)", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 300, 230);
			objGraphic.DrawString("Total MBA Practices: $330", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 850, 750);

			string value1 = string.Empty;
			string value2 = string.Empty;
			int rect = 0; int Xaxis = 0;
			foreach (DataRow row in dt.Rows)
			{
				Single nextPoint = 0;
				DataColumnCollection columns = dt.Columns;
				if (columns.Contains("LookupValue$"))
				{
					sngHighestValueNew = Convert.ToSingle(row["LookupValue$"]) / length * 100;
					if (i < dt.Rows.Count)
						nextPoint = Convert.ToSingle(dt.Rows[i][2]) / length * 100;
				}
				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;

				sngHeight1New = sngHighestValueNew;

				switch (j)
				{
					case 0:
						value1 = "< $493";
						value2 = "";
						break;
					case 1:
						value1 = "$493 - ";
						value2 = "$642";
						break;
					case 2:
						value1 = "$642 - ";
						value2 = "$767";
						break;
					case 3:
						value1 = "$767 - ";
						value2 = "$883";
						break;
					case 4:
						value1 = "$883 - ";
						value2 = "$1026";
						break;
					case 5:
						value1 = "$1026 - ";
						value2 = "$1200";
						break;
					case 6:
						value1 = "$1200 - ";
						value2 = "$1432";
						break;
					case 7:
						value1 = "$1432 - ";
						value2 = "$1695";
						break;
					case 8:
						value1 = "$1695 - ";
						value2 = "$2133";
						break;
					case 9:
						value1 = "$2133+";
						value2 = "";
						break;
				}
				if (rect <= Convert.ToInt32(dt.Rows[dt.Rows.Count - 1][2]))
				{
					Xaxis = Xaxis + length;
					objGraphic.DrawLine(blackPen, 200, rect + 200, 1300, rect + 200);
					objGraphic.DrawString("$" + (Xaxis).ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 100, bitmapsizeYaxis - 220 - rect);
					rect = rect + 100;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, x, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew), x + 100, (bitmapsizeYaxis - 200) - nextPoint);

				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 1, 1);

				objGraphic.DrawLine(blackPen, i * (100) + 200, (bitmapsizeYaxis - 200) - 15, i * (100) + 200, (bitmapsizeYaxis - 200) + 15);

				objGraphic.DrawString("$" + sngHighestValueNew.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, x - 10, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) + 10);

				objGraphic.DrawString(value2, new System.Drawing.Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Pixel),
					blackBrush, (100 * i) + 120, (bitmapsizeYaxis - 140));

				objGraphic.DrawString(value1, new System.Drawing.Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Pixel),
					blackBrush, (100 * i) + 120, (bitmapsizeYaxis - 165));


				i++;
				j++;
				x = x + 100;
			}
			objGraphic.DrawString("Annual Gross Revenue ($000)", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 550, (bitmapsizeYaxis - 200) + 100);
			objGraphic.DrawLine(blackPen, 200, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 200, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 200, 200, 200, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));


			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\linegraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphPercentofPatientsDispensedfromInventorybyInventorySize(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string bookMarkNum, string title)
		{
			System.Data.DataTable dt = new System.Data.DataTable();
			DataSet ds = new DataSet();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
			SqlConnection con = new SqlConnection(connStr);
			con.Open();
			var cmd = new SqlCommand();
			cmd.Connection = con;
			String strQuery = "select  * from " + eSTable;

			cmd.CommandText = strQuery;
			cmd.CommandType = CommandType.Text;
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			adp.Fill(ds);
			dt = ds.Tables[0];

			int bitmapsizeXaxis = 1500;
			int bitmapsizeYaxis = Convert.ToInt32(Convert.ToDecimal(dt.Rows[dt.Rows.Count - 1][dt.Columns.Count - 1].ToString().Replace("%", "")) * 10) + 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 140));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 140, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			objGraphic.DrawString("Percent Dispensed From Inventory", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 50, 120);
			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);

			int i = 1;
			int x = 400;
			//objGraphic.RotateTransform(180);
			//objGraphic.TranslateTransform(0, -bitmapsizeYaxis);
			//objGraphic.ScaleTransform(-1, 1);
			//Graphics objGraphicText = Graphics.FromImage(barBitmap);
			//objGraphic.DrawString("Weekly hours", new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel), blackBrush, 300, 330);
			//objGraphic.DrawString("Total MBA Practices: 46", new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel), blackBrush, 850, 750);

			int rect = 100;
			foreach (DataRow row in dt.Rows)
			{
				foreach (DataColumn column in dt.Columns)
				{
					Single nextPoint = 0;
					DataColumnCollection columns = dt.Columns;

					sngHighestValueNew = Convert.ToSingle(Convert.ToDecimal(row[i - 1].ToString().Replace("%", ""))) * 10;

					if (i < dt.Columns.Count)
						nextPoint = Convert.ToSingle(Convert.ToDecimal(row[i].ToString().Replace("%", ""))) * 10;

					if (sngHighestValueNew == 0)
						sngHighestValueNew = 1;

					sngHeight1New = sngHighestValueNew;
					//if (rect - 100 <= Convert.ToInt32(Convert.ToDecimal(row[dt.Columns.Count - 1].ToString().Replace("%", "")) * 10))
					{
						objGraphic.DrawLine(blackPen, 200, rect + 200, (bitmapsizeXaxis - 100), rect + 200);
						objGraphic.DrawString(rect.ToString().Substring(0, rect.ToString().Length - 1) + "%", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 130, bitmapsizeYaxis - 220 - rect);
						rect = rect + 100;
					}
					if (nextPoint != 0)
						objGraphic.DrawLine(redPen, x, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew), x + 200, (bitmapsizeYaxis - 200) - nextPoint);

					objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 6, 6);
					objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 3, 3);
					objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 1, 1);

					objGraphic.DrawLine(blackPen, i * (200) + 200, (bitmapsizeYaxis - 200) - 15, i * (200) + 200, (bitmapsizeYaxis - 200) + 15);

					objGraphic.DrawString(row[i - 1].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, x - 10, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) + 10);
					objGraphic.DrawString((column.ToString()).Replace('_', ' '), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (200 * i) + 150, ((bitmapsizeYaxis - 200) + 30));
					i++;
					x = x + 200;
				}

			}

			objGraphic.DrawString("Inventory Size (boxes)", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 550, (bitmapsizeYaxis - 200) + 100);
			objGraphic.DrawLine(blackPen, 200, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 200, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 200, 200, 200, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\linegraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphSoftLensInventorybyPracticeSize(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string bookMarkNum, string title)
		{
			//float one = 175;// (Convert.ToSingle(objReport.lstInput[0].colQ92a == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92a)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float two = 240;// (Convert.ToSingle(objReport.lstInput[0].colQ92d == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92d)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float three = 340;//    (Convert.ToSingle(objReport.lstInput[0].colQ92e == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92e)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float four = 400; // (Convert.ToSingle(objReport.lstInput[0].colQ92f == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92f)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float five = 500;// (Convert.ToSingle(objReport.lstInput[0].colQ92g == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92g)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));

			decimal one = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 500000).Select(c => c.Q46a ?? 0).Average(), 2);
			decimal two = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 500000 && c.Q24 <= 750000).Select(c => c.Q46a ?? 0).Average(), 2);
			decimal three = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 750000 && c.Q24 <= 1000000).Select(c => c.Q46a ?? 0).Average(), 2);
			decimal four = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1000000 && c.Q24 <= 1250000).Select(c => c.Q46a ?? 0).Average(), 2);
			decimal five = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1250000 && c.Q24 <= 1500000).Select(c => c.Q46a ?? 0).Average(), 2);

			List<decimal> graphData = new List<decimal>() { one, two, three, four, five };

			int length = Convert.ToInt32(Math.Round((Convert.ToInt32(graphData.Max()) / 5) / 100d, 0) * 100);

			System.Data.DataTable dtValue = new System.Data.DataTable();
			dtValue.Clear();
			dtValue.Columns.Add("Xaxis");
			dtValue.Columns.Add("Value");
			DataRow drValue1 = dtValue.NewRow();
			drValue1["Xaxis"] = "$500,000";
			drValue1["Value"] = one;

			DataRow drValue2 = dtValue.NewRow();
			drValue2["Xaxis"] = "$750,000";
			drValue2["Value"] = two;

			DataRow drValue3 = dtValue.NewRow();
			drValue3["Xaxis"] = "$1M";
			drValue3["Value"] = three;

			DataRow drValue4 = dtValue.NewRow();
			drValue4["Xaxis"] = "$1.25M";
			drValue4["Value"] = four;

			DataRow drValue5 = dtValue.NewRow();
			drValue5["Xaxis"] = "$1.5M";
			drValue5["Value"] = five;

			dtValue.Rows.Add(drValue1);
			dtValue.Rows.Add(drValue2);
			dtValue.Rows.Add(drValue3);
			dtValue.Rows.Add(drValue4);
			dtValue.Rows.Add(drValue5);

			int bitmapsizeXaxis = 1500;
			int bitmapsizeYaxis = 1000;
			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			Brush cadetBlueBrush = new SolidBrush(Color.FromArgb(255, 0, 184, 237));

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);



			objGraphic.DrawString("Nubmber of Boxes", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 50, 120);
			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);

			int i = 1;
			int x = 400;
			int rect = 100;
			int Xaxis = 0;
			foreach (DataRow row in dtValue.Rows)
			{
				Single nextPoint = 0;
				sngHighestValueNew = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i - 1][1].ToString())) / length * 100;
				if (i < dtValue.Rows.Count)
					nextPoint = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i][1].ToString())) / length * 100;
				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;

				sngHeight1New = sngHighestValueNew;
				//if (rect - 100 <= Convert.ToInt32(Convert.ToDecimal(graphData.Max())))
				{
					Xaxis = Xaxis + length;
					objGraphic.DrawLine(blackPen, 200, rect + 200, (bitmapsizeXaxis - 100), rect + 200);
					objGraphic.DrawString(Xaxis.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 130, bitmapsizeYaxis - 220 - rect);
					rect = rect + 100;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, x, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew), x + 200, (bitmapsizeYaxis - 200) - nextPoint);


				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 1, 1);
				objGraphic.DrawLine(blackPen, i * (200) + 200, (bitmapsizeYaxis - 200) - 15, i * (200) + 200, (bitmapsizeYaxis - 200) + 15);

				objGraphic.DrawString(dtValue.Rows[i - 1][1].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, x - 10, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) + 10);
				objGraphic.DrawString(dtValue.Rows[i - 1][0].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (200 * i) + 150, ((bitmapsizeYaxis - 200) + 30));
				i++;
				x = x + 200;
			}

			objGraphic.DrawString("Annual Gross Revenue", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 550, (bitmapsizeYaxis - 200) + 100);

			objGraphic.DrawLine(blackPen, 200, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 200, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 200, 200, 200, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));


			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
							new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));


			string filepath = imagelocation + @"\linegraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphPlanoSunglassInventory(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string bookMarkNum, string title)
		{
			//float one = 76;// (Convert.ToSingle(objReport.lstInput[0].colQ92a == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92a)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float two = 89;// (Convert.ToSingle(objReport.lstInput[0].colQ92d == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92d)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float three = 123;//    (Convert.ToSingle(objReport.lstInput[0].colQ92e == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92e)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float four = 124; // (Convert.ToSingle(objReport.lstInput[0].colQ92f == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92f)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float five = 175;// (Convert.ToSingle(objReport.lstInput[0].colQ92g == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92g)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));


			decimal one = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 509000).Select(c => c.Q38 ?? 0).Average(), 2);
			decimal two = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 509000 && c.Q24 <= 796000).Select(c => c.Q38 ?? 0).Average(), 2);
			decimal three = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 796000 && c.Q24 <= 1100000).Select(c => c.Q38 ?? 0).Average(), 2);
			decimal four = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1100000 && c.Q24 <= 1500000).Select(c => c.Q38 ?? 0).Average(), 2);
			decimal five = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1500000 && c.Q24 <= 2200000).Select(c => c.Q38 ?? 0).Average(), 2);

			List<decimal> graphData = new List<decimal>() { one, two, three, four, five };

			System.Data.DataTable dtValue = new System.Data.DataTable();
			dtValue.Clear();
			dtValue.Columns.Add("Xaxis1");
			dtValue.Columns.Add("Xaxis2");
			dtValue.Columns.Add("Value");
			DataRow drValue1 = dtValue.NewRow();
			drValue1["Xaxis1"] = "Small";
			drValue1["Xaxis2"] = "$509,000";
			drValue1["Value"] = one;

			DataRow drValue2 = dtValue.NewRow();
			drValue2["Xaxis1"] = "Medium Sm-";
			drValue2["Xaxis2"] = "all $796,000";
			drValue2["Value"] = two;

			DataRow drValue3 = dtValue.NewRow();
			drValue3["Xaxis1"] = "Medium";
			drValue3["Xaxis2"] = "$1.1M";
			drValue3["Value"] = three;

			DataRow drValue4 = dtValue.NewRow();
			drValue4["Xaxis1"] = "Medium";
			drValue4["Xaxis2"] = "Large $1.5M";
			drValue4["Value"] = four;

			DataRow drValue5 = dtValue.NewRow();
			drValue5["Xaxis1"] = "Large $2.2M";
			drValue5["Xaxis2"] = "";
			drValue5["Value"] = five;

			dtValue.Rows.Add(drValue1);
			dtValue.Rows.Add(drValue2);
			dtValue.Rows.Add(drValue3);
			dtValue.Rows.Add(drValue4);
			dtValue.Rows.Add(drValue5);

			int bitmapsizeXaxis = 1500;
			int bitmapsizeYaxis = 500 + 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			objGraphic.DrawString("Numbers in inventory", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 50, 120);
			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, 30, 28);

			int i = 1;
			int x = 400;
			int yaxisvalue = 75;

			objGraphic.DrawString("50", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 130, bitmapsizeYaxis - 220);
			int rect = 100;
			foreach (DataRow row in dtValue.Rows)
			{

				Single nextPoint = 0;
				sngHighestValueNew = 100 * i;
				if (i < dtValue.Rows.Count)
					nextPoint = 100 * (i + 1);
				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;
				sngHeight1New = sngHighestValueNew;

				//if (rect - 100 <= Convert.ToInt32(Convert.ToDecimal(graphData.Max()) * 10))
				{
					objGraphic.DrawLine(blackPen, 200, rect + 200, 1400, rect + 200);
					objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 130, bitmapsizeYaxis - 220 - rect);
					rect = rect + 100;
					yaxisvalue = yaxisvalue + 25;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, x, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew), x + 200, (bitmapsizeYaxis - 200) - nextPoint);

				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 1, 1);

				objGraphic.DrawLine(blackPen, i * (200) + 200, (bitmapsizeYaxis - 200) - 15, i * (200) + 200, (bitmapsizeYaxis - 200) + 15);

				objGraphic.DrawString(dtValue.Rows[i - 1][2].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, x - 10, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) + 10);
				objGraphic.DrawString(dtValue.Rows[i - 1][0].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (200 * i) + 150, ((bitmapsizeYaxis - 200) + 30));
				objGraphic.DrawString(dtValue.Rows[i - 1][1].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (200 * i) + 150, ((bitmapsizeYaxis - 200) + 60));
				i++;
				x = x + 200;
			}

			objGraphic.DrawString("Total MBA Practices", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 550, (bitmapsizeYaxis - 200) + 120);
			objGraphic.DrawString("*Among practices with any plano sunglass inventory", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 200, (bitmapsizeYaxis - 200) + 160);
			objGraphic.DrawLine(blackPen, 200, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 200, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 200, 200, 200, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));



			string filepath = imagelocation + @"\linegraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphFramesMarkUpbyRetailPrice(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string bookMarkNum, string title)
		{
			//float one = 3.09f;// (Convert.ToSingle(objReport.lstInput[0].colQ92a == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92a)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float two = 2.84f;// (Convert.ToSingle(objReport.lstInput[0].colQ92d == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92d)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float three = 2.76f;//    (Convert.ToSingle(objReport.lstInput[0].colQ92e == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92e)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float four = 2.64f; // (Convert.ToSingle(objReport.lstInput[0].colQ92f == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92f)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float five = 2.59f;// (Convert.ToSingle(objReport.lstInput[0].colQ92g == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92g)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));

			decimal one = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q36 > 0 && c.Q36 <= 99).Select(c => (((c.Q34 + c.Q35) * c.Q36) / c.Q52a) ?? 0).Average(), 2);
			decimal two = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q36 >= 100 && c.Q36 <= 199).Select(c => (((c.Q34 + c.Q35) * c.Q36) / c.Q52a) ?? 0).Average(), 2);
			decimal three = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q36 >= 200 && c.Q36 <= 299).Select(c => (((c.Q34 + c.Q35) * c.Q36) / c.Q52a) ?? 0).Average(), 2);
			decimal four = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q36 >= 300 && c.Q36 <= 399).Select(c => (((c.Q34 + c.Q35) * c.Q36) / c.Q52a) ?? 0).Average(), 2);
			decimal five = Math.Round(db.Source_InputDataBenchMarkSource.Where(c => c.Q36 >= 400 && c.Q36 <= 100000).Select(c => (((c.Q34 + c.Q35) * c.Q36) / c.Q52a) ?? 0).Average(), 2);


			List<decimal> graphData = new List<decimal>() { one, two, three, four, five };

			int length = Convert.ToInt32(Math.Round((Convert.ToInt32(graphData.Max()) / 5) / 100d, 0) * 100);

			System.Data.DataTable dtValue = new System.Data.DataTable();
			dtValue.Clear();
			dtValue.Columns.Add("Xaxis");
			dtValue.Columns.Add("Value");
			DataRow drValue1 = dtValue.NewRow();
			drValue1["Xaxis"] = "$99 or less";
			drValue1["Value"] = Math.Round(one, 0);

			DataRow drValue2 = dtValue.NewRow();
			drValue2["Xaxis"] = "$100-199";
			drValue2["Value"] = Math.Round(two, 0);

			DataRow drValue3 = dtValue.NewRow();
			drValue3["Xaxis"] = "$200-299";
			drValue3["Value"] = Math.Round(three, 0);

			DataRow drValue4 = dtValue.NewRow();
			drValue4["Xaxis"] = "$300-399";
			drValue4["Value"] = Math.Round(four, 0);

			DataRow drValue5 = dtValue.NewRow();
			drValue5["Xaxis"] = "$400 or more";
			drValue5["Value"] = Math.Round(five, 0);

			dtValue.Rows.Add(drValue1);
			dtValue.Rows.Add(drValue2);
			dtValue.Rows.Add(drValue3);
			dtValue.Rows.Add(drValue4);
			dtValue.Rows.Add(drValue5);

			int bitmapsizeXaxis = 1500;
			int bitmapsizeYaxis = 500 + 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);

			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);
			objGraphic.DrawString("Mark-up", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 50, 120);
			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, 30, 28);
			int i = 1;
			int x = 400;
			int Xaxis = 0;
			int rect = 100;
			foreach (DataRow row in dtValue.Rows)
			{
				Single nextPoint = 0;
				sngHighestValueNew = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i - 1][1].ToString())) / length * 100;
				if (i < dtValue.Rows.Count)
					nextPoint = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i][1].ToString())) / length * 100;
				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;
				sngHeight1New = sngHighestValueNew;
				if (rect - 100 <= 700)
				{
					Xaxis = Xaxis + length;
					objGraphic.DrawLine(blackPen, 200, rect + 200, 1400, rect + 200);
					objGraphic.DrawString(Xaxis.ToString().Substring(0, rect.ToString().Length) + "x", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 100, bitmapsizeYaxis - 220 - rect);
					rect = rect + 100;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, x, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew), x + 200, (bitmapsizeYaxis - 200) - nextPoint);

				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, x - 3, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 1, 1);

				objGraphic.DrawLine(blackPen, i * (200) + 200, (bitmapsizeYaxis - 200) - 15, i * (200) + 200, (bitmapsizeYaxis - 200) + 15);
				objGraphic.DrawString(dtValue.Rows[i - 1][1].ToString() + "x", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, x - 10, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) + 10);
				objGraphic.DrawString(dtValue.Rows[i - 1][0].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (200 * i) + 150, ((bitmapsizeYaxis - 200) + 50));

				i++;
				x = x + 200;
			}

			objGraphic.DrawString("Retail Price", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 550, (bitmapsizeYaxis - 200) + 100);
			objGraphic.DrawString("*Selling price divided by cost-of goods", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 200, (bitmapsizeYaxis - 200) + 150);
			objGraphic.DrawLine(blackPen, 200, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 200, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 200, 200, 200, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
								new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));


			string filepath = imagelocation + @"\linegraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}


		public void GenerateLineGraphGrossRevenueperSquareFootbyPracticeSize(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string Graph1, string bookMarkNum, string title)
		{
			System.Data.DataTable dt = new System.Data.DataTable();
			DataSet ds = new DataSet();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
			SqlConnection con = new SqlConnection(connStr);
			con.Open();
			var cmd = new SqlCommand();
			cmd.Connection = con;



			System.Data.DataTable dtValue = new System.Data.DataTable();
			dtValue.Clear();
			dtValue.Columns.Add("value");

			//DataRow drValue1 = dtValue.NewRow();
			//drValue1["value"] = "1700";
			//DataRow drValue2 = dtValue.NewRow();
			//drValue2["value"] = "2000";
			//DataRow drValue3 = dtValue.NewRow();
			//drValue3["value"] = "2150";
			//DataRow drValue4 = dtValue.NewRow();
			//drValue4["value"] = "2500";
			//DataRow drValue5 = dtValue.NewRow();
			//drValue5["value"] = "3000";
			//DataRow drValue6 = dtValue.NewRow();
			//drValue6["value"] = "3000";
			//DataRow drValue7 = dtValue.NewRow();
			//drValue7["value"] = "3375";
			//DataRow drValue8 = dtValue.NewRow();
			//drValue8["value"] = "3300";
			//DataRow drValue9 = dtValue.NewRow();
			//drValue9["value"] = "4400";
			//DataRow drValue10 = dtValue.NewRow();
			//drValue10["value"] = "5000";

			decimal? value = 0;
			DataRow drValue1 = dtValue.NewRow();
			value = db.Source_InputDataBenchMarkSource.Where(c => (c.Q24 == null ? 0 : c.Q24) <= 356 && (c.Q24 == null ? 0 : c.Q24) > 0).Select(c => (c.Q2 == null ? 0 : c.Q2)).Average();
			if (string.IsNullOrEmpty(Convert.ToString(value)))
			{
				drValue1["value"] = 0;
			}
			else
				drValue1["value"] = value;

			DataRow drValue2 = dtValue.NewRow();
			value = db.Source_InputDataBenchMarkSource.Where(c => (c.Q24 == null ? 0 : c.Q24) > 356 && (c.Q24 == null ? 0 : c.Q24) <= 581).Select(c => (c.Q2 == null ? 0 : c.Q2)).Average();
			if (string.IsNullOrEmpty(Convert.ToString(value)))
			{
				drValue2["value"] = 0;
			}
			else
				drValue2["value"] = value;

			DataRow drValue3 = dtValue.NewRow();
			value = db.Source_InputDataBenchMarkSource.Where(c => (c.Q24 == null ? 0 : c.Q24) > 581 && (c.Q24 == null ? 0 : c.Q24) <= 698).Select(c => (c.Q2 == null ? 0 : c.Q2)).Average();
			if (string.IsNullOrEmpty(Convert.ToString(value)))
			{
				drValue3["value"] = 0;
			}
			else
				drValue3["value"] = value;

			DataRow drValue4 = dtValue.NewRow();
			value = db.Source_InputDataBenchMarkSource.Where(c => (c.Q24 == null ? 0 : c.Q24) > 698 && (c.Q24 == null ? 0 : c.Q24) <= 823).Select(c => (c.Q2 == null ? 0 : c.Q2)).Average();
			if (string.IsNullOrEmpty(Convert.ToString(value)))
			{
				drValue4["value"] = 0;
			}
			else
				drValue4["value"] = value;

			DataRow drValue5 = dtValue.NewRow();
			value = db.Source_InputDataBenchMarkSource.Where(c => (c.Q24 == null ? 0 : c.Q24) > 823 && (c.Q24 == null ? 0 : c.Q24) <= 947).Select(c => (c.Q2 == null ? 0 : c.Q2)).Average();
			if (string.IsNullOrEmpty(Convert.ToString(value)))
			{
				drValue5["value"] = 0;
			}
			else
				drValue5["value"] = value;

			DataRow drValue6 = dtValue.NewRow();
			value = db.Source_InputDataBenchMarkSource.Where(c => (c.Q24 == null ? 0 : c.Q24) > 947 && (c.Q24 == null ? 0 : c.Q24) <= 1106).Select(c => (c.Q2 == null ? 0 : c.Q2)).Average();
			if (string.IsNullOrEmpty(Convert.ToString(value)))
			{
				drValue6["value"] = 0;
			}
			else
				drValue6["value"] = value;

			DataRow drValue7 = dtValue.NewRow();
			value = db.Source_InputDataBenchMarkSource.Where(c => (c.Q24 == null ? 0 : c.Q24) > 1106 && (c.Q24 == null ? 0 : c.Q24) <= 1300).Select(c => (c.Q2 == null ? 0 : c.Q2)).Average();
			if (string.IsNullOrEmpty(Convert.ToString(value)))
			{
				drValue7["value"] = 0;
			}
			else
				drValue7["value"] = value;

			DataRow drValue8 = dtValue.NewRow();
			value = db.Source_InputDataBenchMarkSource.Where(c => (c.Q24 == null ? 0 : c.Q24) > 1300 && (c.Q24 == null ? 0 : c.Q24) <= 1532).Select(c => (c.Q2 == null ? 0 : c.Q2)).Average();
			if (string.IsNullOrEmpty(Convert.ToString(value)))
			{
				drValue8["value"] = 0;
			}
			else
				drValue8["value"] = value;

			DataRow drValue9 = dtValue.NewRow();
			value = db.Source_InputDataBenchMarkSource.Where(c => (c.Q24 == null ? 0 : c.Q24) > 1532 && (c.Q24 == null ? 0 : c.Q24) <= 1852).Select(c => (c.Q2 == null ? 0 : c.Q2)).Average();
			if (string.IsNullOrEmpty(Convert.ToString(value)))
			{
				drValue9["value"] = 0;
			}
			else
				drValue9["value"] = value;

			DataRow drValue10 = dtValue.NewRow();
			value = db.Source_InputDataBenchMarkSource.Where(c => (c.Q24 == null ? 0 : c.Q24) > 1852 && (c.Q24 == null ? 0 : c.Q24) <= 2950).Select(c => (c.Q2 == null ? 0 : c.Q2)).Average();
			if (string.IsNullOrEmpty(Convert.ToString(value)) || value == 0)
			{
				drValue10["value"] = value;
			}
			else
				drValue10["value"] = value;

			dtValue.Rows.Add(drValue1);
			dtValue.Rows.Add(drValue2);
			dtValue.Rows.Add(drValue3);
			dtValue.Rows.Add(drValue4);
			dtValue.Rows.Add(drValue5);
			dtValue.Rows.Add(drValue6);
			dtValue.Rows.Add(drValue7);
			dtValue.Rows.Add(drValue8);
			dtValue.Rows.Add(drValue9);
			dtValue.Rows.Add(drValue10);


			String strQuery = "select  * from " + eSTable;

			//		cmd.CommandText = " select CONVERT(varchar(10), RowId-6)+'th' +CONVERT(varchar(1), '-')+" +
			//			" CONVERT(varchar(10), RowId + 3) + 'th' + ' percentile' as Heading, * from [dbo].[Lookup.GrossRevenuePerCompleteExam] " +
			//"where RowId like ('%6') or RowId like ('50') ";
			cmd.CommandText = strQuery;
			cmd.CommandType = CommandType.Text;
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			adp.Fill(ds);
			dt = ds.Tables[0];

			int bitmapsizeXaxis = 1400;

			int length = Convert.ToInt32(Math.Round((Convert.ToInt32(Convert.ToDecimal(dtValue.Rows[dtValue.Rows.Count - 1][dtValue.Columns.Count - 1].ToString())) / 6) / 100d, 0) * 100);

			if (length == 0)
			{
				length = 1;
			}

			int bitmapsizeYaxis = 700 + 500;
			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);

			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			objGraphic.DrawString("Office Square Feet", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 50, 120);

			objGraphic.DrawString("0", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 100, bitmapsizeYaxis - 220);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arail", 28, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);

			int i = 1;
			int x = 300;
			int Xaxis = 0;
			int rect = 100;

			foreach (DataRow row in dtValue.Rows)
			{
				Single nextPoint = 0;
				if (dtValue.Rows[i - 1][0].ToString() != "0")
					sngHighestValueNew = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i - 1][0].ToString())) / length * 100;
				else
					sngHighestValueNew = 1;

				if (i < dtValue.Rows.Count)
				{
					if (dtValue.Rows[i][0].ToString() != "0")
						nextPoint = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i][0].ToString())) / length * 100;
					else
						nextPoint = 1;
				}

				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;

				sngHeight1New = sngHighestValueNew;

				if ((rect - 100 < 800) && nextPoint != 0)
				{
					if (length == 1)
					{
						length = length + 1;
					}
					Xaxis = Xaxis + length;
					objGraphic.DrawLine(blackPen, 200, rect + 200, (bitmapsizeXaxis - 100), rect + 200);
					objGraphic.DrawString((Xaxis).ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 100, bitmapsizeYaxis - 220 - rect);
					rect = rect + 100;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, x - 50, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew), x + 50, (bitmapsizeYaxis - 200) - nextPoint);

				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 1, 1);

				objGraphic.DrawLine(blackPen, i * (100) + 200, (bitmapsizeYaxis - 200) - 15, i * (100) + 200, (bitmapsizeYaxis - 200) + 15);
				objGraphic.DrawString(Math.Round(Convert.ToDouble(dtValue.Rows[i - 1][0]), 2).ToString(), new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, x - 60, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) + 10);
				i++;
				x = x + 100;
			}

			i = 1;
			x = 300;
			rect = 100;
			foreach (DataRow row in dt.Rows)
			{
				foreach (DataColumn column in dt.Columns)
				{
					Single nextPoint = 0;
					DataColumnCollection columns = dt.Columns;

					sngHighestValueNew = Convert.ToSingle(Convert.ToDecimal(row[i - 1].ToString())) * 100;

					if (i < dt.Columns.Count)
						nextPoint = Convert.ToSingle(Convert.ToDecimal(row[i].ToString())) * 100;

					if (sngHighestValueNew == 0)
						sngHighestValueNew = 1;
					sngHeight1New = sngHighestValueNew;

					objGraphic.DrawString(column.ToString(), new System.Drawing.Font("Arial", 23, FontStyle.Bold, GraphicsUnit.Pixel),
						blackBrush, (100 * i) + 110, (bitmapsizeYaxis - 200) + 45);
					i++;
					x = x + 100;
				}

			}
			objGraphic.DrawString("Annual Gross Revenue ($000)", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 550, (bitmapsizeYaxis - 200) + 100);
			objGraphic.DrawLine(blackPen, 200, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 200, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 200, 200, 200, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));


			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			objGraphic.DrawLine(new Pen(Brushes.White, 2), 200, 1100, 1300, 1100);

			string filepath = imagelocation + @"\linegraph" + Graph1 + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphGrossRevenueperSquareFootbyPracticeSizeforrefrection(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string Graph2, string bookMarkNum, string title)
		{
			System.Data.DataTable dt = new System.Data.DataTable();
			DataSet ds = new DataSet();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
			SqlConnection con = new SqlConnection(connStr);
			con.Open();
			var cmd = new SqlCommand();
			cmd.Connection = con;

			String strQuery = "select  * from " + eSTable;
			cmd.CommandText = strQuery;
			cmd.CommandType = CommandType.Text;
			SqlDataAdapter adp = new SqlDataAdapter(cmd);
			adp.Fill(ds);
			dt = ds.Tables[0];

			int bitmapsizeXaxis = 1400;
			int bitmapsizeYaxis = Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(dt.Rows[dt.Rows.Count - 1][dt.Columns.Count - 1].ToString())) * 100) + 400;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);


			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			objGraphic.DrawString("Refraction Rooms", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 50, 120);
			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);

			int i = 1;
			int x = 300;
			int rect = 100;
			foreach (DataRow row in dt.Rows)
			{
				foreach (DataColumn column in dt.Columns)
				{
					Single nextPoint = 0;
					DataColumnCollection columns = dt.Columns;

					sngHighestValueNew = Convert.ToSingle(Convert.ToDecimal(row[i - 1].ToString())) * 100;

					if (i < dt.Columns.Count)
						nextPoint = Convert.ToSingle(Convert.ToDecimal(row[i].ToString())) * 100;

					if (sngHighestValueNew == 0)
						sngHighestValueNew = 1;

					sngHeight1New = sngHighestValueNew;

					if (rect <= Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(dt.Rows[dt.Rows.Count - 1][dt.Columns.Count - 1].ToString())) * 100))
					{
						objGraphic.DrawLine(blackPen, 200, rect + 200, (bitmapsizeXaxis - 100), rect + 200);
						objGraphic.DrawString(rect.ToString().Substring(0, rect.ToString().Length - 2), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 150, bitmapsizeYaxis - 120 - rect);
						rect = rect + 100;
					}
					if (nextPoint != 0)
						objGraphic.DrawLine(redPen, x - 50, (bitmapsizeYaxis - 100) - Convert.ToInt32(sngHighestValueNew), x + 50, (bitmapsizeYaxis - 100) - nextPoint);

					objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 100) - Convert.ToInt32(sngHighestValueNew) - 3, 6, 6);
					objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 100) - Convert.ToInt32(sngHighestValueNew) - 3, 3, 3);
					objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 100) - Convert.ToInt32(sngHighestValueNew) - 3, 1, 1);

					objGraphic.DrawLine(blackPen, i * (100) + 200, (bitmapsizeYaxis - 200) - 15, i * (100) + 200, (bitmapsizeYaxis - 200) + 15);
					objGraphic.DrawString(row[i - 1].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, x - 30, (bitmapsizeYaxis - 100) - Convert.ToInt32(sngHighestValueNew) + 10);
					objGraphic.DrawString(column.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (100 * i) + 115, (bitmapsizeYaxis - 200) + 20);
					i++;
					x = x + 100;
				}

			}
			objGraphic.DrawString("Annual Gross Revenue ($000)", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 550, (bitmapsizeYaxis - 100) + 10);
			objGraphic.DrawLine(blackPen, 200, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 200, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 200, 200, 200, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));


			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
							new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));


			string filepath = imagelocation + @"\linegraph" + Graph2 + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphRangeofSquareFootagebyPracticeSizeSmallestThird(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string Graph1, string bookMarkNum, string title)
		{
			//float one = 1200;// (Convert.ToSingle(objReport.lstInput[0].colQ92a == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92a)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float two = 1600;// (Convert.ToSingle(objReport.lstInput[0].colQ92d == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92d)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float three = 1950;//    (Convert.ToSingle(objReport.lstInput[0].colQ92e == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92e)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float four = 2150; // (Convert.ToSingle(objReport.lstInput[0].colQ92f == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92f)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float five = 3250;// (Convert.ToSingle(objReport.lstInput[0].colQ92g == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92g)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));

			int count = 0;
			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 509000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal one = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 509000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 509000 && c.Q24 <= 796000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal two = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 509000 && c.Q24 <= 796000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 796000 && c.Q24 <= 1100000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal three = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 796000 && c.Q24 <= 1100000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1100000 && c.Q24 <= 1500000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal four = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1100000 && c.Q24 <= 1500000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1500000 && c.Q24 <= 2200000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal five = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1500000 && c.Q24 <= 2200000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Take(count).Average();


			List<decimal> graphData = new List<decimal>() { one, two, three, four, five };

			System.Data.DataTable dtValue = new System.Data.DataTable();
			dtValue.Clear();
			dtValue.Columns.Add("Yaxis1");
			dtValue.Columns.Add("Yaxis2");
			dtValue.Columns.Add("Value");
			DataRow drValue1 = dtValue.NewRow();
			drValue1["Yaxis1"] = "Small $509,000";
			drValue1["Yaxis2"] = "";
			drValue1["Value"] = one;

			DataRow drValue2 = dtValue.NewRow();
			drValue2["Yaxis1"] = "Medium"; drValue2["Yaxis2"] = "Small $790,000";
			drValue2["Value"] = two;

			DataRow drValue3 = dtValue.NewRow();
			drValue3["Yaxis1"] = "Medium $1.1M";
			drValue3["Yaxis2"] = "";
			drValue3["Value"] = three;

			DataRow drValue4 = dtValue.NewRow();
			drValue4["Yaxis1"] = "Medium";
			drValue4["Yaxis2"] = "Large $1.4M";
			drValue4["Value"] = four;

			DataRow drValue5 = dtValue.NewRow();
			drValue5["Yaxis1"] = "Large $2.2M";
			drValue5["Yaxis2"] = "";
			drValue5["Value"] = five;

			dtValue.Rows.Add(drValue1);
			dtValue.Rows.Add(drValue2);
			dtValue.Rows.Add(drValue3);
			dtValue.Rows.Add(drValue4);
			dtValue.Rows.Add(drValue5);

			int bitmapsizeXaxis = 1600;
			int bitmapsizeYaxis = 500 + 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);

			objGraphic.DrawString("Practice Size Quintiles", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 300, 160);
			int i = 1;
			int x = 300;
			int yaxisvalue = 1000;

			int rect = 100;
			foreach (DataRow row in dtValue.Rows)
			{

				Single nextPoint = 0;
				sngHighestValueNew = Convert.ToSingle(dtValue.Rows[i - 1][2]) / 10;
				if (i < dtValue.Rows.Count)
					nextPoint = Convert.ToSingle(dtValue.Rows[i][2]) / 10;
				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;
				sngHeight1New = sngHighestValueNew;

				//if (rect - 100 <= Convert.ToInt32(Convert.ToDecimal(graphData.Max()) * 10))
				{
					objGraphic.DrawLine(blackPen, 300, rect + 200, 1500, rect + 200);

					objGraphic.DrawString(dtValue.Rows[i - 1][0].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 30, bitmapsizeYaxis - 200 - rect);
					objGraphic.DrawString(dtValue.Rows[i - 1][1].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 30, bitmapsizeYaxis - 230 - rect);
					//objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel),
					//blackBrush, 130, bitmapsizeYaxis - 220 - rect);
					//rect = rect + 100;
					//yaxisvalue = yaxisvalue + 25;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, (300 + sngHighestValueNew), (bitmapsizeYaxis - 200) - rect,
						(300 + nextPoint), (bitmapsizeYaxis - 200) - (rect + 100));

				objGraphic.DrawEllipse(redPen, (300 + sngHighestValueNew) - 3, (bitmapsizeYaxis - 200) - rect - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, (300 + sngHighestValueNew) - 3, (bitmapsizeYaxis - 200) - rect - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, (300 + sngHighestValueNew) - 3, (bitmapsizeYaxis - 200) - rect - 3, 1, 1);
				//objGraphic.DrawLine(blackPen, i * (100) + 200, (bitmapsizeYaxis - 200) - 15, i * (100) + 200, (bitmapsizeYaxis - 200) + 15);

				objGraphic.DrawString(Math.Round(Convert.ToDouble(dtValue.Rows[i - 1][2]), 2).ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (300 + sngHighestValueNew) + 10, (bitmapsizeYaxis - 200) - rect + 10);

				//objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel),
				//		blackBrush, (100 * i) + 150, ((bitmapsizeYaxis - 200) + 30));
				//yaxisvalue = yaxisvalue + 1000;
				i++;
				x = x + 100;
				rect = rect + 100;

			}
			i = 1;
			x = 300;

			int width = 0;

			if (graphData.Max().ToString().Split('.')[0].Length > 4)
			{
				width = Convert.ToInt32(graphData.Max().ToString().Substring(0, 2));
			}
			else
				width = Convert.ToInt32(graphData.Max().ToString().Substring(0, 1));

			for (int z = 0; z <= 10; z++)
			{
				objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (100 * i) + 250, ((bitmapsizeYaxis - 200) + 30));
				objGraphic.DrawLine(blackPen, i * (100) + 300, (bitmapsizeYaxis - 200) - 15, i * (100) + 300, (bitmapsizeYaxis - 200) + 15);
				i++;
				x = x + 100;
				yaxisvalue = yaxisvalue + 1000;
			}
			//objGraphic.DrawString("Total MBA Practices", new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel),
			//			blackBrush, 550, (bitmapsizeYaxis - 200) + 100);
			objGraphic.DrawString("Median Sq. Ft.", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 750, (bitmapsizeYaxis - 200) + 100);

			objGraphic.DrawLine(blackPen, 300, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 300, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 300, 200, 300, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
									new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));


			string filepath = imagelocation + @"\linegraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphRangeofSquareFootagebyPracticeSizeMediumThird(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string Graph2, string bookMarkNum, string title)
		{
			//float one = 1800;// (Convert.ToSingle(objReport.lstInput[0].colQ92a == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92a)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float two = 2400;// (Convert.ToSingle(objReport.lstInput[0].colQ92d == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92d)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float three = 3000;//    (Convert.ToSingle(objReport.lstInput[0].colQ92e == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92e)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float four = 3300; // (Convert.ToSingle(objReport.lstInput[0].colQ92f == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92f)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float five = 5000;// (Convert.ToSingle(objReport.lstInput[0].colQ92g == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92g)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));


			int count = 0;
			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 509000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal one = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 509000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Skip(count).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 509000 && c.Q24 <= 796000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal two = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 509000 && c.Q24 <= 796000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Skip(count).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 796000 && c.Q24 <= 1100000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal three = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 796000 && c.Q24 <= 1100000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Skip(count).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1100000 && c.Q24 <= 1500000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal four = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1100000 && c.Q24 <= 1500000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Skip(count).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1500000 && c.Q24 <= 2200000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal five = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1500000 && c.Q24 <= 2200000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Skip(count).Take(count).Average();


			List<decimal> graphData = new List<decimal>() { one, two, three, four, five };

			System.Data.DataTable dtValue = new System.Data.DataTable();
			dtValue.Clear();
			dtValue.Columns.Add("Yaxis1");
			dtValue.Columns.Add("Yaxis2");
			dtValue.Columns.Add("Value");
			DataRow drValue1 = dtValue.NewRow();
			drValue1["Yaxis1"] = "Small $509,000";
			drValue1["Yaxis2"] = "";
			drValue1["Value"] = one;

			DataRow drValue2 = dtValue.NewRow();
			drValue2["Yaxis1"] = "Medium"; drValue2["Yaxis2"] = "Small $790,000";
			drValue2["Value"] = two;

			DataRow drValue3 = dtValue.NewRow();
			drValue3["Yaxis1"] = "Medium $1.1M";
			drValue3["Yaxis2"] = "";
			drValue3["Value"] = three;

			DataRow drValue4 = dtValue.NewRow();
			drValue4["Yaxis1"] = "Medium";
			drValue4["Yaxis2"] = "Large $1.4M";
			drValue4["Value"] = four;

			DataRow drValue5 = dtValue.NewRow();
			drValue5["Yaxis1"] = "Large $2.2M";
			drValue5["Yaxis2"] = "";
			drValue5["Value"] = five;

			dtValue.Rows.Add(drValue1);
			dtValue.Rows.Add(drValue2);
			dtValue.Rows.Add(drValue3);
			dtValue.Rows.Add(drValue4);
			dtValue.Rows.Add(drValue5);

			int bitmapsizeXaxis = 1600;
			int bitmapsizeYaxis = 500 + 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);

			objGraphic.DrawString("Practice Size Quintiles", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 300, 160);
			int i = 1;
			int x = 300;
			int yaxisvalue = 1000;

			int rect = 100;
			foreach (DataRow row in dtValue.Rows)
			{

				Single nextPoint = 0;
				sngHighestValueNew = Convert.ToSingle(dtValue.Rows[i - 1][2]) / 10;
				if (i < dtValue.Rows.Count)
					nextPoint = Convert.ToSingle(dtValue.Rows[i][2]) / 10;
				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;
				sngHeight1New = sngHighestValueNew;

				//if (rect - 100 <= Convert.ToInt32(Convert.ToDecimal(graphData.Max()) * 10))
				{
					objGraphic.DrawLine(blackPen, 300, rect + 200, 1500, rect + 200);

					objGraphic.DrawString(dtValue.Rows[i - 1][0].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 30, bitmapsizeYaxis - 200 - rect);
					objGraphic.DrawString(dtValue.Rows[i - 1][1].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 30, bitmapsizeYaxis - 230 - rect);
					//objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel),
					//blackBrush, 130, bitmapsizeYaxis - 220 - rect);
					//rect = rect + 100;
					//yaxisvalue = yaxisvalue + 25;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, (300 + sngHighestValueNew), (bitmapsizeYaxis - 200) - rect,
						(300 + nextPoint), (bitmapsizeYaxis - 200) - (rect + 100));

				objGraphic.DrawEllipse(redPen, (300 + sngHighestValueNew) - 3, (bitmapsizeYaxis - 200) - rect - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, (300 + sngHighestValueNew) - 3, (bitmapsizeYaxis - 200) - rect - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, (300 + sngHighestValueNew) - 3, (bitmapsizeYaxis - 200) - rect - 3, 1, 1);
				//objGraphic.DrawLine(blackPen, i * (100) + 200, (bitmapsizeYaxis - 200) - 15, i * (100) + 200, (bitmapsizeYaxis - 200) + 15);

				objGraphic.DrawString(Math.Round(Convert.ToDouble(dtValue.Rows[i - 1][2]), 2).ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (300 + sngHighestValueNew) + 10, (bitmapsizeYaxis - 200) - rect + 10);

				//objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel),
				//		blackBrush, (100 * i) + 150, ((bitmapsizeYaxis - 200) + 30));
				//yaxisvalue = yaxisvalue + 1000;
				i++;
				x = x + 100;
				rect = rect + 100;

			}
			i = 1;
			x = 300;

			int width = 0;

			if (graphData.Max().ToString().Split('.')[0].Length > 4)
			{
				width = Convert.ToInt32(graphData.Max().ToString().Substring(0, 2));
			}
			else
				width = Convert.ToInt32(graphData.Max().ToString().Substring(0, 1));

			for (int z = 0; z <= 10; z++)
			{
				objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (100 * i) + 250, ((bitmapsizeYaxis - 200) + 30));
				objGraphic.DrawLine(blackPen, i * (100) + 300, (bitmapsizeYaxis - 200) - 15, i * (100) + 300, (bitmapsizeYaxis - 200) + 15);
				i++;
				x = x + 100;
				yaxisvalue = yaxisvalue + 1000;
			}
			//objGraphic.DrawString("Total MBA Practices", new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel),
			//			blackBrush, 550, (bitmapsizeYaxis - 200) + 100);
			objGraphic.DrawString("Median Sq. Ft.", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 750, (bitmapsizeYaxis - 200) + 100);

			objGraphic.DrawLine(blackPen, 300, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 300, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 300, 200, 300, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
									new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\linegraph" + Graph2 + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphRangeofSquareFootagebyPracticeSizeLargestThird(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string Graph3, string bookMarkNum, string title)
		{
			//float one = 2600;// (Convert.ToSingle(objReport.lstInput[0].colQ92a == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92a)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float two = 3400;// (Convert.ToSingle(objReport.lstInput[0].colQ92d == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92d)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float three = 4200;//    (Convert.ToSingle(objReport.lstInput[0].colQ92e == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92e)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float four = 5000; // (Convert.ToSingle(objReport.lstInput[0].colQ92f == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92f)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));
			//float five = 7500;// (Convert.ToSingle(objReport.lstInput[0].colQ92g == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ92g)).ToString("#,0"))) / (Convert.ToSingle(objReport.lstInput[0].colQ52j == null ? null : Math.Round(Convert.ToDecimal(objReport.lstInput[0].colQ52j)).ToString("#,0")));

			int count = 0;
			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 509000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal one = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 509000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Skip((count * 2)).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 509000 && c.Q24 <= 796000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal two = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 509000 && c.Q24 <= 796000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Skip((count * 2)).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 796000 && c.Q24 <= 1100000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal three = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 796000 && c.Q24 <= 1100000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Skip((count * 2)).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1100000 && c.Q24 <= 1500000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal four = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1100000 && c.Q24 <= 1500000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Skip((count * 2)).Take(count).Average();

			count = Convert.ToInt32((db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1500000 && c.Q24 <= 2200000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Count()) / 3);
			if (string.IsNullOrEmpty(Convert.ToString(count)))
			{
				count = 0;
			}
			decimal five = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1500000 && c.Q24 <= 2200000).OrderBy(c => c.Q2).Select(c => c.Q2 ?? 0).Skip((count * 2)).Take(count).Average();


			List<decimal> graphData = new List<decimal>() { one, two, three, four, five };

			System.Data.DataTable dtValue = new System.Data.DataTable();
			dtValue.Clear();
			dtValue.Columns.Add("Yaxis1");
			dtValue.Columns.Add("Yaxis2");
			dtValue.Columns.Add("Value");
			DataRow drValue1 = dtValue.NewRow();
			drValue1["Yaxis1"] = "Small $509,000";
			drValue1["Yaxis2"] = "";
			drValue1["Value"] = one;

			DataRow drValue2 = dtValue.NewRow();
			drValue2["Yaxis1"] = "Medium"; drValue2["Yaxis2"] = "Small $790,000";
			drValue2["Value"] = two;

			DataRow drValue3 = dtValue.NewRow();
			drValue3["Yaxis1"] = "Medium $1.1M";
			drValue3["Yaxis2"] = "";
			drValue3["Value"] = three;

			DataRow drValue4 = dtValue.NewRow();
			drValue4["Yaxis1"] = "Medium";
			drValue4["Yaxis2"] = "Large $1.4M";
			drValue4["Value"] = four;

			DataRow drValue5 = dtValue.NewRow();
			drValue5["Yaxis1"] = "Large $2.2M";
			drValue5["Yaxis2"] = "";
			drValue5["Value"] = five;

			dtValue.Rows.Add(drValue1);
			dtValue.Rows.Add(drValue2);
			dtValue.Rows.Add(drValue3);
			dtValue.Rows.Add(drValue4);
			dtValue.Rows.Add(drValue5);

			int bitmapsizeXaxis = 1600;
			int bitmapsizeYaxis = 500 + 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);

			objGraphic.DrawString("Practice Size Quintiles", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 300, 160);
			int i = 1;
			int x = 300;
			int yaxisvalue = 1000;

			int rect = 100;
			foreach (DataRow row in dtValue.Rows)
			{

				Single nextPoint = 0;
				sngHighestValueNew = Convert.ToSingle(dtValue.Rows[i - 1][2]) / 10;
				if (i < dtValue.Rows.Count)
					nextPoint = Convert.ToSingle(dtValue.Rows[i][2]) / 10;
				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;
				sngHeight1New = sngHighestValueNew;

				//if (rect - 100 <= Convert.ToInt32(Convert.ToDecimal(graphData.Max()) * 10))
				{
					objGraphic.DrawLine(blackPen, 300, rect + 200, 1500, rect + 200);

					objGraphic.DrawString(dtValue.Rows[i - 1][0].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 30, bitmapsizeYaxis - 200 - rect);
					objGraphic.DrawString(dtValue.Rows[i - 1][1].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 30, bitmapsizeYaxis - 230 - rect);
					//objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel),
					//blackBrush, 130, bitmapsizeYaxis - 220 - rect);
					//rect = rect + 100;
					//yaxisvalue = yaxisvalue + 25;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, (300 + sngHighestValueNew), (bitmapsizeYaxis - 200) - rect,
						(300 + nextPoint), (bitmapsizeYaxis - 200) - (rect + 100));

				objGraphic.DrawEllipse(redPen, (300 + sngHighestValueNew) - 3, (bitmapsizeYaxis - 200) - rect - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, (300 + sngHighestValueNew) - 3, (bitmapsizeYaxis - 200) - rect - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, (300 + sngHighestValueNew) - 3, (bitmapsizeYaxis - 200) - rect - 3, 1, 1);
				//objGraphic.DrawLine(blackPen, i * (100) + 200, (bitmapsizeYaxis - 200) - 15, i * (100) + 200, (bitmapsizeYaxis - 200) + 15);

				objGraphic.DrawString(Math.Round(Convert.ToDouble(dtValue.Rows[i - 1][2]), 2).ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (300 + sngHighestValueNew) + 10, (bitmapsizeYaxis - 200) - rect + 10);

				//objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel),
				//		blackBrush, (100 * i) + 150, ((bitmapsizeYaxis - 200) + 30));
				//yaxisvalue = yaxisvalue + 1000;
				i++;
				x = x + 100;
				rect = rect + 100;

			}
			i = 1;
			x = 300;

			int width = 0;

			if (graphData.Max().ToString().Split('.')[0].Length > 4)
			{
				width = Convert.ToInt32(graphData.Max().ToString().Substring(0, 2));
			}
			else
				width = Convert.ToInt32(graphData.Max().ToString().Substring(0, 1));

			for (int z = 0; z <= 10; z++)
			{
				objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, (100 * i) + 250, ((bitmapsizeYaxis - 200) + 30));
				objGraphic.DrawLine(blackPen, i * (100) + 300, (bitmapsizeYaxis - 200) - 15, i * (100) + 300, (bitmapsizeYaxis - 200) + 15);
				i++;
				x = x + 100;
				yaxisvalue = yaxisvalue + 1000;
			}
			//objGraphic.DrawString("Total MBA Practices", new System.Drawing.Font("Arial", 30, FontStyle.Bold, GraphicsUnit.Pixel),
			//			blackBrush, 550, (bitmapsizeYaxis - 200) + 100);
			objGraphic.DrawString("Median Sq. Ft.", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 750, (bitmapsizeYaxis - 200) + 100);

			objGraphic.DrawLine(blackPen, 300, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 300, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 300, 200, 300, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
									new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\linegraph" + Graph3 + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphStaffingLevelsbyPracticeSizeOD(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string Graph1, string bookMarkNum, string title)
		{
			System.Data.DataTable dtValue = new System.Data.DataTable();
			dtValue.Clear();
			dtValue.Columns.Add("value");

			DataRow drValue1 = dtValue.NewRow();
			drValue1["value"] = "1.1";
			DataRow drValue2 = dtValue.NewRow();
			drValue2["value"] = "1.1";
			DataRow drValue3 = dtValue.NewRow();
			drValue3["value"] = "1.2";
			DataRow drValue4 = dtValue.NewRow();
			drValue4["value"] = "1.3";
			DataRow drValue5 = dtValue.NewRow();
			drValue5["value"] = "1.3";
			DataRow drValue6 = dtValue.NewRow();
			drValue6["value"] = "1.6";
			DataRow drValue7 = dtValue.NewRow();
			drValue7["value"] = "1.8";
			DataRow drValue8 = dtValue.NewRow();
			drValue8["value"] = "2.1";
			DataRow drValue9 = dtValue.NewRow();
			drValue9["value"] = "2.4";
			DataRow drValue10 = dtValue.NewRow();
			drValue10["value"] = "3.3";


			//DataRow drValue1 = dtValue.NewRow();
			//drValue1["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 493).Select(c => c.Q4 ?? 0).Average();
			//DataRow drValue2 = dtValue.NewRow();
			//drValue2["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 493 && c.Q24 <= 642).Select(c => c.Q4 ?? 0).Average();
			//DataRow drValue3 = dtValue.NewRow();
			//drValue3["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 642 && c.Q24 <= 767).Select(c => c.Q4 ?? 0).Average();
			//DataRow drValue4 = dtValue.NewRow();
			//drValue4["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 767 && c.Q24 <= 493).Select(c => c.Q4 ?? 0).Average();
			//DataRow drValue5 = dtValue.NewRow();
			//drValue5["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 883 && c.Q24 <= 883).Select(c => c.Q4 ?? 0).Average();
			//DataRow drValue6 = dtValue.NewRow();
			//drValue6["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1026 && c.Q24 <= 1200).Select(c => c.Q4 ?? 0).Average();
			//DataRow drValue7 = dtValue.NewRow();
			//drValue7["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1200 && c.Q24 <= 1432).Select(c => c.Q4 ?? 0).Average();
			//DataRow drValue8 = dtValue.NewRow();
			//drValue8["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1432 && c.Q24 <= 1695).Select(c => c.Q4 ?? 0).Average();
			//DataRow drValue9 = dtValue.NewRow();
			//drValue9["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1695 && c.Q24 <= 2133).Select(c => c.Q4 ?? 0).Average();
			//DataRow drValue10 = dtValue.NewRow();
			//drValue10["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 2133).Select(c => c.Q4 ?? 0).Average();


			dtValue.Rows.Add(drValue1);
			dtValue.Rows.Add(drValue2);
			dtValue.Rows.Add(drValue3);
			dtValue.Rows.Add(drValue4);
			dtValue.Rows.Add(drValue5);
			dtValue.Rows.Add(drValue6);
			dtValue.Rows.Add(drValue7);
			dtValue.Rows.Add(drValue8);
			dtValue.Rows.Add(drValue9);
			dtValue.Rows.Add(drValue10);

			int bitmapsizeXaxis = 1400;
			int bitmapsizeYaxis = Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(dtValue.Rows[dtValue.Rows.Count - 1][dtValue.Columns.Count - 1].ToString())) * 100) + 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);

			int i = 1;
			int j = 0;
			int x = 300;
			int rect = 0;
			string value1 = string.Empty;
			string value2 = string.Empty;

			objGraphic.DrawString("Full-time Employed ODs", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 100, 140);

			foreach (DataRow row in dtValue.Rows)
			{
				Single nextPoint = 0;
				DataColumnCollection columns = dtValue.Columns;

				sngHighestValueNew = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i - 1][0].ToString())) * 100;

				if (i < dtValue.Rows.Count)
					nextPoint = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i][0].ToString())) * 100;

				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;

				sngHeight1New = sngHighestValueNew;
				switch (j)
				{
					case 0:
						value1 = "< $493";
						value2 = "";
						break;
					case 1:
						value1 = "$493 - ";
						value2 = "$642";
						break;
					case 2:
						value1 = "$642 - ";
						value2 = "$767";
						break;
					case 3:
						value1 = "$767 - ";
						value2 = "$883";
						break;
					case 4:
						value1 = "$883 - ";
						value2 = "$1026";
						break;
					case 5:
						value1 = "$1026 - ";
						value2 = "$1200";
						break;
					case 6:
						value1 = "$1200 - ";
						value2 = "$1432";
						break;
					case 7:
						value1 = "$1432 - ";
						value2 = "$1695";
						break;
					case 8:
						value1 = "$1695 - ";
						value2 = "$2133";
						break;
					case 9:
						value1 = "$2133+";
						value2 = "";
						break;
				}
				if (rect <= Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(dtValue.Rows[dtValue.Rows.Count - 1][dtValue.Columns.Count - 1].ToString())) * 100))
				{
					objGraphic.DrawLine(blackPen, 200, rect + 200, (bitmapsizeXaxis - 100), rect + 200);
					if (rect != 0)
						objGraphic.DrawString(rect.ToString().Substring(0, rect.ToString().Length - 2), new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 150, bitmapsizeYaxis - 220 - rect);
					else
						objGraphic.DrawString(rect.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 150, bitmapsizeYaxis - 220 - rect);
					rect = rect + 100;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, x - 50, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew), x + 50, (bitmapsizeYaxis - 200) - nextPoint);

				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 1, 1);
				objGraphic.DrawLine(blackPen, i * (100) + 200, (bitmapsizeYaxis - 200) - 15, i * (100) + 200, (bitmapsizeYaxis - 200) + 15);
				objGraphic.DrawString(dtValue.Rows[i - 1][0].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, x - 30, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) + 10);
				objGraphic.DrawString(value1, new System.Drawing.Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Pixel),
				blackBrush, (100 * i) + 120, (bitmapsizeYaxis - 200) + 20);

				objGraphic.DrawString(value2, new System.Drawing.Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Pixel),
					blackBrush, (100 * i) + 120, (bitmapsizeYaxis - 200) + 45);
				i++;
				j++;
				x = x + 100;
			}

			objGraphic.DrawString("Annual Gross Revenue ($000)", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 550, (bitmapsizeYaxis - 150) + 50);

			objGraphic.DrawString("Note: Full time equivalent(FTE) equals 2,080 hours per year", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 150, (bitmapsizeYaxis - 50) + 10);
			objGraphic.DrawLine(blackPen, 200, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 200, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 200, 200, 200, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));


			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\linegraph" + Graph1 + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphStaffingLevelsbyPracticeSizenonOD(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string Graph2, string bookMarkNum, string title)
		{
			System.Data.DataTable dtValue = new System.Data.DataTable();
			dtValue.Clear();
			dtValue.Columns.Add("value");

			DataRow drValue1 = dtValue.NewRow();
			drValue1["value"] = "2.8";
			DataRow drValue2 = dtValue.NewRow();
			drValue2["value"] = "3.5";
			DataRow drValue3 = dtValue.NewRow();
			drValue3["value"] = "4.3";
			DataRow drValue4 = dtValue.NewRow();
			drValue4["value"] = "4.9";
			DataRow drValue5 = dtValue.NewRow();
			drValue5["value"] = "5.7";
			DataRow drValue6 = dtValue.NewRow();
			drValue6["value"] = "6.6";
			DataRow drValue7 = dtValue.NewRow();
			drValue7["value"] = "7.5";
			DataRow drValue8 = dtValue.NewRow();
			drValue8["value"] = "8.8";
			DataRow drValue9 = dtValue.NewRow();
			drValue9["value"] = "10.2";
			DataRow drValue10 = dtValue.NewRow();
			drValue10["value"] = "13.0";


			//DataRow drValue1 = dtValue.NewRow();
			//drValue1["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 493).Select(c => c.Q11 ?? 0).Average();
			//DataRow drValue2 = dtValue.NewRow();
			//drValue2["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 493 && c.Q24 <= 642).Select(c => c.Q11 ?? 0).Average();
			//DataRow drValue3 = dtValue.NewRow();
			//drValue3["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 642 && c.Q24 <= 767).Select(c => c.Q11 ?? 0).Average();
			//DataRow drValue4 = dtValue.NewRow();
			//drValue4["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 767 && c.Q24 <= 493).Select(c => c.Q11 ?? 0).Average();
			//DataRow drValue5 = dtValue.NewRow();
			//drValue5["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 883 && c.Q24 <= 883).Select(c => c.Q11 ?? 0).Average();
			//DataRow drValue6 = dtValue.NewRow();
			//drValue6["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1026 && c.Q24 <= 1200).Select(c => c.Q11 ?? 0).Average();
			//DataRow drValue7 = dtValue.NewRow();
			//drValue7["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1200 && c.Q24 <= 1432).Select(c => c.Q11 ?? 0).Average();
			//DataRow drValue8 = dtValue.NewRow();
			//drValue8["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1432 && c.Q24 <= 1695).Select(c => c.Q11 ?? 0).Average();
			//DataRow drValue9 = dtValue.NewRow();
			//drValue9["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1695 && c.Q24 <= 2133).Select(c => c.Q11 ?? 0).Average();
			//DataRow drValue10 = dtValue.NewRow();
			//drValue10["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 2133).Select(c => c.Q11 ?? 0).Average();


			dtValue.Rows.Add(drValue1);
			dtValue.Rows.Add(drValue2);
			dtValue.Rows.Add(drValue3);
			dtValue.Rows.Add(drValue4);
			dtValue.Rows.Add(drValue5);
			dtValue.Rows.Add(drValue6);
			dtValue.Rows.Add(drValue7);
			dtValue.Rows.Add(drValue8);
			dtValue.Rows.Add(drValue9);
			dtValue.Rows.Add(drValue10);

			int bitmapsizeXaxis = 1400;
			int bitmapsizeYaxis = Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(dtValue.Rows[dtValue.Rows.Count - 1][dtValue.Columns.Count - 1].ToString()) / 2) * 100) + 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);

			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);

			int i = 1;
			int j = 0;
			int x = 300;
			int rect = 0;
			int yaxisvalue = 0;
			string value1 = string.Empty;
			string value2 = string.Empty;

			objGraphic.DrawString("Full-time Employed Staff", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel), blackBrush, 100, 140);

			foreach (DataRow row in dtValue.Rows)
			{
				Single nextPoint = 0;
				DataColumnCollection columns = dtValue.Columns;

				sngHighestValueNew = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i - 1][0].ToString()) / 2) * 100;

				if (i < dtValue.Rows.Count)
					nextPoint = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i][0].ToString()) / 2) * 100;

				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;

				sngHeight1New = sngHighestValueNew;
				switch (j)
				{
					case 0:
						value1 = "< $493";
						value2 = "";
						break;
					case 1:
						value1 = "$493 - ";
						value2 = "$642";
						break;
					case 2:
						value1 = "$642 - ";
						value2 = "$767";
						break;
					case 3:
						value1 = "$767 - ";
						value2 = "$883";
						break;
					case 4:
						value1 = "$883 - ";
						value2 = "$1026";
						break;
					case 5:
						value1 = "$1026 - ";
						value2 = "$1200";
						break;
					case 6:
						value1 = "$1200 - ";
						value2 = "$1432";
						break;
					case 7:
						value1 = "$1432 - ";
						value2 = "$1695";
						break;
					case 8:
						value1 = "$1695 - ";
						value2 = "$2133";
						break;
					case 9:
						value1 = "$2133+";
						value2 = "";
						break;
				}
				if (rect <= Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(dtValue.Rows[dtValue.Rows.Count - 1][dtValue.Columns.Count - 1].ToString()) / 2) * 100))
				{
					objGraphic.DrawLine(blackPen, 200, rect + 200, (bitmapsizeXaxis - 100), rect + 200);
					objGraphic.DrawString(yaxisvalue.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 150, bitmapsizeYaxis - 220 - rect);
					rect = rect + 100;
					yaxisvalue = yaxisvalue + 2;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, x - 50, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew), x + 50, (bitmapsizeYaxis - 200) - nextPoint);

				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 1, 1);
				objGraphic.DrawLine(blackPen, i * (100) + 200, (bitmapsizeYaxis - 200) - 15, i * (100) + 200, (bitmapsizeYaxis - 200) + 15);
				objGraphic.DrawString(dtValue.Rows[i - 1][0].ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, x - 30, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) + 10);
				objGraphic.DrawString(value1, new System.Drawing.Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Pixel),
				blackBrush, (100 * i) + 120, (bitmapsizeYaxis - 200) + 20);

				objGraphic.DrawString(value2, new System.Drawing.Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Pixel),
					blackBrush, (100 * i) + 120, (bitmapsizeYaxis - 200) + 45);
				i++;
				j++;
				x = x + 100;
			}

			objGraphic.DrawString("Annual Gross Revenue ($000)", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 550, (bitmapsizeYaxis - 150) + 50);

			objGraphic.DrawString("Note: Full time equivalent(FTE) equals 2,080 hours per year", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 150, (bitmapsizeYaxis - 50));
			objGraphic.DrawLine(blackPen, 200, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 200, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 200, 200, 200, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\linegraph" + Graph2 + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public void GenerateLineGraphNetIncomePercentGrossRev(string imagelocation, Microsoft.Office.Interop.Word.Application wordApp, Microsoft.Office.Interop.Word.Document aDoc, string eSTable, string bookMarkNum, string title)
		{
			System.Data.DataTable dtValue = new System.Data.DataTable();
			dtValue.Clear();
			dtValue.Columns.Add("value");


			DataRow drValue1 = dtValue.NewRow();
			drValue1["value"] = "26.6";
			DataRow drValue2 = dtValue.NewRow();
			drValue2["value"] = "29.6";
			DataRow drValue3 = dtValue.NewRow();
			drValue3["value"] = "27.6";
			DataRow drValue4 = dtValue.NewRow();
			drValue4["value"] = "30.7";
			DataRow drValue5 = dtValue.NewRow();
			drValue5["value"] = "31.8";
			DataRow drValue6 = dtValue.NewRow();
			drValue6["value"] = "30.3";
			DataRow drValue7 = dtValue.NewRow();
			drValue7["value"] = "32.2";
			DataRow drValue8 = dtValue.NewRow();
			drValue8["value"] = "32.0";
			DataRow drValue9 = dtValue.NewRow();
			drValue9["value"] = "33.2";
			DataRow drValue10 = dtValue.NewRow();
			drValue10["value"] = "35.0";

			//DataRow drValue1 = dtValue.NewRow();
			//drValue1["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 0 && c.Q24 <= 493).Select(c => ((c.Q24 - ((c.Q52j ?? 0 + c.Q53 ?? 0 + c.Q54 ?? 0 + c.Q55 ?? 0 + c.Q56 ?? 0 + c.Q57 ?? 0 + c.Q58 ?? 0 + c.Q58 ?? 0 + c.Q59 ?? 0 + c.Q60 ?? 0))) / c.Q24) ?? 0 * 100).Average();
			//DataRow drValue2 = dtValue.NewRow();
			//drValue2["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 493 && c.Q24 <= 642).Select(c => ((c.Q24 - ((c.Q52j ?? 0 + c.Q53 ?? 0 + c.Q54 ?? 0 + c.Q55 ?? 0 + c.Q56 ?? 0 + c.Q57 ?? 0 + c.Q58 ?? 0 + c.Q58 ?? 0 + c.Q59 ?? 0 + c.Q60 ?? 0))) / c.Q24) ?? 0 * 100).Average();
			//DataRow drValue3 = dtValue.NewRow();
			//drValue3["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 642 && c.Q24 <= 767).Select(c => ((c.Q24 - ((c.Q52j ?? 0 + c.Q53 ?? 0 + c.Q54 ?? 0 + c.Q55 ?? 0 + c.Q56 ?? 0 + c.Q57 ?? 0 + c.Q58 ?? 0 + c.Q58 ?? 0 + c.Q59 ?? 0 + c.Q60 ?? 0))) / c.Q24) ?? 0 * 100).Average();
			//DataRow drValue4 = dtValue.NewRow();
			//drValue4["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 767 && c.Q24 <= 493).Select(c => ((c.Q24 - ((c.Q52j ?? 0 + c.Q53 ?? 0 + c.Q54 ?? 0 + c.Q55 ?? 0 + c.Q56 ?? 0 + c.Q57 ?? 0 + c.Q58 ?? 0 + c.Q58 ?? 0 + c.Q59 ?? 0 + c.Q60 ?? 0))) / c.Q24) ?? 0 * 100).Average();
			//DataRow drValue5 = dtValue.NewRow();
			//drValue5["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 883 && c.Q24 <= 883).Select(c => ((c.Q24 - ((c.Q52j ?? 0 + c.Q53 ?? 0 + c.Q54 ?? 0 + c.Q55 ?? 0 + c.Q56 ?? 0 + c.Q57 ?? 0 + c.Q58 ?? 0 + c.Q58 ?? 0 + c.Q59 ?? 0 + c.Q60 ?? 0))) / c.Q24) ?? 0 * 100).Average();
			//DataRow drValue6 = dtValue.NewRow();
			//drValue6["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1026 && c.Q24 <= 1200).Select(c => ((c.Q24 - ((c.Q52j ?? 0 + c.Q53 ?? 0 + c.Q54 ?? 0 + c.Q55 ?? 0 + c.Q56 ?? 0 + c.Q57 ?? 0 + c.Q58 ?? 0 + c.Q58 ?? 0 + c.Q59 ?? 0 + c.Q60 ?? 0))) / c.Q24) ?? 0 * 100).Average();
			//DataRow drValue7 = dtValue.NewRow();
			//drValue7["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1200 && c.Q24 <= 1432).Select(c => ((c.Q24 - ((c.Q52j ?? 0 + c.Q53 ?? 0 + c.Q54 ?? 0 + c.Q55 ?? 0 + c.Q56 ?? 0 + c.Q57 ?? 0 + c.Q58 ?? 0 + c.Q58 ?? 0 + c.Q59 ?? 0 + c.Q60 ?? 0))) / c.Q24) ?? 0 * 100).Average();
			//DataRow drValue8 = dtValue.NewRow();
			//drValue8["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1432 && c.Q24 <= 1695).Select(c => ((c.Q24 - ((c.Q52j ?? 0 + c.Q53 ?? 0 + c.Q54 ?? 0 + c.Q55 ?? 0 + c.Q56 ?? 0 + c.Q57 ?? 0 + c.Q58 ?? 0 + c.Q58 ?? 0 + c.Q59 ?? 0 + c.Q60 ?? 0))) / c.Q24) ?? 0 * 100).Average();
			//DataRow drValue9 = dtValue.NewRow();
			//drValue9["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 1695 && c.Q24 <= 2133).Select(c => ((c.Q24 - ((c.Q52j ?? 0 + c.Q53 ?? 0 + c.Q54 ?? 0 + c.Q55 ?? 0 + c.Q56 ?? 0 + c.Q57 ?? 0 + c.Q58 ?? 0 + c.Q58 ?? 0 + c.Q59 ?? 0 + c.Q60 ?? 0))) / c.Q24) ?? 0 * 100).Average();
			//DataRow drValue10 = dtValue.NewRow();
			//drValue10["value"] = db.Source_InputDataBenchMarkSource.Where(c => c.Q24 > 2133).Select(c => ((c.Q24 - ((c.Q52j ?? 0 + c.Q53 ?? 0 + c.Q54 ?? 0 + c.Q55 ?? 0 + c.Q56 ?? 0 + c.Q57 ?? 0 + c.Q58 ?? 0 + c.Q58 ?? 0 + c.Q59 ?? 0 + c.Q60 ?? 0))) / c.Q24) ?? 0 * 100).Average();


			dtValue.Rows.Add(drValue1);
			dtValue.Rows.Add(drValue2);
			dtValue.Rows.Add(drValue3);
			dtValue.Rows.Add(drValue4);
			dtValue.Rows.Add(drValue5);
			dtValue.Rows.Add(drValue6);
			dtValue.Rows.Add(drValue7);
			dtValue.Rows.Add(drValue8);
			dtValue.Rows.Add(drValue9);
			dtValue.Rows.Add(drValue10);

			int bitmapsizeXaxis = 1400;
			int bitmapsizeYaxis = Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(dtValue.Rows[dtValue.Rows.Count - 1][dtValue.Columns.Count - 1].ToString())) * 10) + 50 + 500;

			Bitmap barBitmap = new Bitmap(bitmapsizeXaxis, bitmapsizeYaxis);
			Graphics objGraphic = Graphics.FromImage(barBitmap);

			Brush lightblueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
			Brush blueBrush = new SolidBrush(Color.Blue);
			Brush greenBrush = new SolidBrush(Color.Green);
			Brush whiteBrush = new SolidBrush(Color.White);
			Brush rectBrush = new SolidBrush(Color.FromArgb(255, 231, 231, 233));
			Brush blackBrush = new SolidBrush(Color.Black);
			Brush grayBrush = new SolidBrush(Color.Gray);
			Brush yellowBrush = new SolidBrush(Color.Yellow);
			Pen grayPen = new Pen(Color.Gray, 2);
			Pen redPen = new Pen(Color.Red, 3);
			Pen blackPen = new Pen(Color.Black, 1);

			objGraphic.FillRectangle(whiteBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.FillRectangle(lightblueBrush, new System.Drawing.Rectangle(0, 0, bitmapsizeXaxis, 100));

			Single sngHighestValueNew = new Single();
			Single sngHeight1New = new Single();

			RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 100, Width = bitmapsizeXaxis } };
			//graphics.DrawRectangle(pen, rect2);
			//To write header text
			StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
			objGraphic.DrawString(title, new System.Drawing.Font("Arial", 40, FontStyle.Italic, GraphicsUnit.Pixel), whiteBrush, header, format);
			objGraphic.DrawString("% of Gross Revenue", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
				blackBrush, 50, 120);
			//objGraphic.DrawString(title, new System.Drawing.Font("Arial", 35, FontStyle.Italic, GraphicsUnit.Pixel), blackBrush, 30, 28);

			int i = 1;
			int j = 0;
			int x = 300;
			int rect = 0;
			string value1 = string.Empty;
			string value2 = string.Empty;

			//objGraphic.DrawString("Full-time Employed ODs", new System.Drawing.Font("Arial", 25, FontStyle.Bold, GraphicsUnit.Pixel), blackBrush, 100, 140);

			foreach (DataRow row in dtValue.Rows)
			{
				Single nextPoint = 0;
				DataColumnCollection columns = dtValue.Columns;

				sngHighestValueNew = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i - 1][0].ToString())) * 10;

				if (i < dtValue.Rows.Count)
					nextPoint = Convert.ToSingle(Convert.ToDecimal(dtValue.Rows[i][0].ToString())) * 10;

				if (sngHighestValueNew == 0)
					sngHighestValueNew = 1;

				sngHeight1New = sngHighestValueNew;
				switch (j)
				{
					case 0:
						value1 = "< $493";
						value2 = "";
						break;
					case 1:
						value1 = "$493 - ";
						value2 = "$642";
						break;
					case 2:
						value1 = "$642 - ";
						value2 = "$767";
						break;
					case 3:
						value1 = "$767 - ";
						value2 = "$883";
						break;
					case 4:
						value1 = "$883 - ";
						value2 = "$1026";
						break;
					case 5:
						value1 = "$1026 - ";
						value2 = "$1200";
						break;
					case 6:
						value1 = "$1200 - ";
						value2 = "$1432";
						break;
					case 7:
						value1 = "$1432 - ";
						value2 = "$1695";
						break;
					case 8:
						value1 = "$1695 - ";
						value2 = "$2133";
						break;
					case 9:
						value1 = "$2133+";
						value2 = "";
						break;
				}
				if (rect - 100 <= Convert.ToInt32(Math.Ceiling(Convert.ToDecimal(dtValue.Rows[dtValue.Rows.Count - 1][dtValue.Columns.Count - 1].ToString())) * 10))
				{
					objGraphic.DrawLine(blackPen, 200, rect + 200, (bitmapsizeXaxis - 100), rect + 200);
					if (rect != 0)
						objGraphic.DrawString(rect.ToString().Substring(0, rect.ToString().Length - 1) + "%", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 120, bitmapsizeYaxis - 220 - rect);
					else
						objGraphic.DrawString(rect.ToString(), new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, 150, bitmapsizeYaxis - 220 - rect);
					rect = rect + 100;
				}
				if (nextPoint != 0)
					objGraphic.DrawLine(redPen, x - 50, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew), x + 50, (bitmapsizeYaxis - 200) - nextPoint);

				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 6, 6);
				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 3, 3);
				objGraphic.DrawEllipse(redPen, x - 53, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) - 3, 1, 1);
				objGraphic.DrawLine(blackPen, i * (100) + 200, (bitmapsizeYaxis - 200) - 15, i * (100) + 200, (bitmapsizeYaxis - 200) + 15);
				objGraphic.DrawString(dtValue.Rows[i - 1][0].ToString() + "%", new System.Drawing.Font("Arial", 25, FontStyle.Regular, GraphicsUnit.Pixel),
					blackBrush, x - 10, (bitmapsizeYaxis - 200) - Convert.ToInt32(sngHighestValueNew) + 10);
				objGraphic.DrawString(value1, new System.Drawing.Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Pixel),
				blackBrush, (100 * i) + 120, (bitmapsizeYaxis - 200) + 20);

				objGraphic.DrawString(value2, new System.Drawing.Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Pixel),
					blackBrush, (100 * i) + 120, (bitmapsizeYaxis - 200) + 45);
				i++;
				j++;
				x = x + 100;
			}

			objGraphic.DrawString("Annual Gross Revenue ($000)", new System.Drawing.Font("Arial", 30, FontStyle.Regular, GraphicsUnit.Pixel),
						blackBrush, 550, (bitmapsizeYaxis - 150) + 40);
			objGraphic.DrawLine(blackPen, 200, (bitmapsizeYaxis - 200), (bitmapsizeXaxis - 100), (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, 200, 200, bitmapsizeXaxis - 100, 200);
			objGraphic.DrawLine(blackPen, 200, 200, 200, (bitmapsizeYaxis - 200));
			objGraphic.DrawLine(blackPen, bitmapsizeXaxis - 100, 200, bitmapsizeXaxis - 100, (bitmapsizeYaxis - 200));

			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(bitmapsizeXaxis, 0));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(bitmapsizeXaxis, 0),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 6), new System.Drawing.Point(0, bitmapsizeYaxis),
				new System.Drawing.Point(bitmapsizeXaxis, bitmapsizeYaxis));
			objGraphic.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
				new System.Drawing.Point(0, bitmapsizeYaxis));

			string filepath = imagelocation + @"\linegraph" + bookMarkNum + ".png";
			barBitmap.Save(filepath, ImageFormat.Png);
			string bookMark = "KeyMetrics_Graph_" + bookMarkNum;
			//aDoc.Bookmarks[Convert.ToInt32(bookMarkNum)].Range.Text = string.Empty;
			var shape1 = aDoc.Bookmarks[bookMark].Range.InlineShapes.AddPicture(filepath, false, true);
			shape1.Width = 550;
			shape1.Height = 415;

			//Dispose off the Graphics and Bitmap objects
			objGraphic.Dispose();
			barBitmap.Dispose();


		}

		public string UpdateReportGenerateStatus(List<Output> lstOutput)
		{
			try
			{
				foreach (Output op in lstOutput)
				{
					Target_OutputData objdbOutput = new Target_OutputData();
					objdbOutput = db.Target_OutputData.Where(r => r.SourceDataRefId == op.SourceDataRefId).Select(r => r).ToList().FirstOrDefault();
					objdbOutput.IsReportGenerated = true;
					//In below line we are getting current windows logon UserName.
					string CurrentUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
					objdbOutput.ReportGeneratedBy = CurrentUserName;
					//objdbOutput.ReportGeneratedBy = "Administrator";
					objdbOutput.ReportGeneratedDate = DateTime.Now;

				}
				db.SaveChanges();

			}
			catch (Exception ex)
			{
				return ex.Message;
			}
			return "success";
		}

		public string GetOrdinal(int num)
		{
			try
			{
				switch (num % 100)
				{
					case 11:
					case 12:
					case 13:
						return num.ToString() + "th";
				}

				switch (num % 10)
				{
					case 1:
						return num.ToString() + "st";
					case 2:
						return num.ToString() + "nd";
					case 3:
						return num.ToString() + "rd";
					default:
						return num.ToString() + "th";
				}
			}
			catch (Exception ex)
			{
				return ex.Message;
			}


		}

		public decimal ReturnValue(string arr)
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


			return Convert.ToDecimal(arr);
		}

		static decimal GetListMedian(List<decimal> dataList)
		{
			decimal median = 0;

			if (dataList.Count > 0)
				median = dataList.OrderBy(x => x).Skip(dataList.Count() / 2).First();
			return median;
		}

		void PopulateBenchMarks(string lookuptableName, string year, char type)
		{
			try
			{
				//string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
				//SqlConnection con = new SqlConnection(connStr);
				//using (SqlCommand cmd = new SqlCommand("SP_PopulateAllBenchmarks", con))
				//{
				//    cmd.CommandType = CommandType.StoredProcedure;



				//    cmd.Parameters.Add("@year", SqlDbType.VarChar).Value = year;
				//    cmd.Parameters.Add("@type", SqlDbType.Char).Value = type;
				//    con.Open();
				//    cmd.ExecuteNonQuery();
				//    con.Close();
				//}
				string StrSPQuery = "";
				if (string.IsNullOrEmpty(lookuptableName))
				{
					StrSPQuery = "EXEC [dbo].[SP_PopulateAllBenchmarks] @LookupName = NULL, @year = N'" + year + "', @type = N'" + type + "'";
				}
				else
					StrSPQuery = "EXEC [dbo].[SP_PopulateAllBenchmarks] @LookupName = [" + lookuptableName + "], @year = N'" + year + "', @type = N'" + type + "'";
				db.ExecuteStoreCommand(StrSPQuery);
				db.SaveChanges();
			}
			catch (Exception ex)
			{
				throw ex;
			}

		}

		public IEnumerable<GetQuestionData_Result> GetQuestionData(int practiceid, int yearid)
		{
			System.Data.DataTable retVal = new System.Data.DataTable();
			ObjectParameter[] PracticeIdParameter = new ObjectParameter[2];
			ObjectParameter PracticeIdParameter1 = new ObjectParameter("PracticeId", practiceid);
			PracticeIdParameter[0] = PracticeIdParameter1;
			PracticeIdParameter1 = new ObjectParameter("YearId", yearid);
			PracticeIdParameter[1] = PracticeIdParameter1;
			var result = db.ExecuteFunction<GetQuestionData_Result>("GetQuestionData", MergeOption.OverwriteChanges,
										  PracticeIdParameter);

			//var result = db.GetQuestionData(practiceid, Convert.ToInt16(yearid)).AsEnumerable<GetQuestionData_Result>();

			IEnumerable<GetQuestionData_Result> query = from order in result.AsEnumerable()
														select order;


			return query;
		}

		public IEnumerable<GetSurveyData_Result> GetSurveyData(string practiceid)
		{

			ObjectParameter PracticeParameter;
			PracticeParameter = new ObjectParameter("PracticeId", practiceid);
			var result = db.ExecuteFunction<GetSurveyData_Result>("GetSurveyData", MergeOption.OverwriteChanges, PracticeParameter);
			// var result = db.GetSurveyData(Convert.ToInt32(practiceid));

			IEnumerable<GetSurveyData_Result> query = from order in result.AsEnumerable()
													  select order;

			return query;

		}

		public System.Data.DataTable GetSurveyDataSave(string practiceid)
		{

			System.Data.DataTable dt = new System.Data.DataTable();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
			SqlConnection con = new SqlConnection(connStr);

			con.Open();

			var cmd = new SqlCommand();
			cmd.Connection = con;
			cmd.CommandText = "GetSurveyDataForSave";
			cmd.CommandType = CommandType.StoredProcedure;

			var numParam = new SqlParameter();
			numParam.ParameterName = "@PracticeId";
			numParam.SqlDbType = SqlDbType.Int;
			numParam.Value = Convert.ToInt32(practiceid); //   
			cmd.Parameters.Add(numParam);



			numParam = new SqlParameter();
			numParam.ParameterName = "@Year";
			numParam.SqlDbType = SqlDbType.VarChar;
			numParam.Value = System.Web.HttpContext.Current.Session["YearName"].ToString(); //   
			cmd.Parameters.Add(numParam);


			numParam = new SqlParameter();
			numParam.ParameterName = "@Name";
			numParam.SqlDbType = SqlDbType.VarChar;
			numParam.Value = System.Web.HttpContext.Current.Session["namesave"].ToString(); //   
			cmd.Parameters.Add(numParam);


			numParam = new SqlParameter();
			numParam.ParameterName = "@PracticeName";
			numParam.SqlDbType = SqlDbType.VarChar;
			numParam.Value = System.Web.HttpContext.Current.Session["practicenamesave"].ToString(); //   
			cmd.Parameters.Add(numParam);


			SqlDataAdapter adp = new SqlDataAdapter();
			adp.SelectCommand = cmd;

			adp.Fill(dt);

			con.Close();


			//System.Data.DataTable dt = new System.Data.DataTable();
			//ObjectParameter PracticeParameter;
			//PracticeParameter = new ObjectParameter("PracticeId", practiceid);
			//var result = db.ExecuteFunction<GetSurveyData_Result>("GetSurveyData", MergeOption.OverwriteChanges, PracticeParameter);
			//// var result = db.GetSurveyData(Convert.ToInt32(practiceid));

			//IEnumerable<GetSurveyData_Result> query = from order in result.AsEnumerable()
			//                                          select order;

			return dt;

		}

		public int GetPageSection(int practiceid, int yearid)
		{
			int pageno = 0;
			ObjectParameter[] PracticeIdParameter = new ObjectParameter[2];
			ObjectParameter PracticeIdParameter1 = new ObjectParameter("PracticeId", practiceid);
			PracticeIdParameter[0] = PracticeIdParameter1;
			PracticeIdParameter1 = new ObjectParameter("YearId", yearid);
			PracticeIdParameter[1] = PracticeIdParameter1;

			try
			{
				var result = db.ExecuteFunction<int>("GetPageSection", MergeOption.OverwriteChanges, PracticeIdParameter).FirstOrDefault();
				if (result != null)
				{
					pageno = Convert.ToInt32(result);
				}

				else
				{
					pageno = 0;
				}

			}

			catch
			{

				pageno = 0;
			}
			// var result = db.GetPageSection((Byte)practiceid, (Byte)yearid);



			return pageno;
		}

		//public IEnumerable<GetAdminInfo_Result> GetAdminInfo()
		//{
		//    System.Data.DataTable retVal = new System.Data.DataTable();
		//    ObjectParameter PracticeIdParameter;
		//    //PracticeIdParameter = new ObjectParameter("PracticeId", practiceid);
		//    var result = db.ExecuteFunction<GetAdminInfo_Result>("GetAdminInfo", MergeOption.OverwriteChanges);

		//    IEnumerable<GetAdminInfo_Result> query = from order in result.AsEnumerable()
		//                                                 select order;


		//    return query;
		//}

		public System.Data.DataTable GetSurveyYear()
		{


			//var result = db.ExecuteFunction<SynchNewYearSurveyData_Result>("SynchNewYearSurveyData", MergeOption.OverwriteChanges);
			//// var result = db.SynchNewYearSurveyData();

			//IEnumerable<SynchNewYearSurveyData_Result> query = from order in result.AsEnumerable()
			//                                                   select order;






			//return query;
			DataSet ds = new DataSet();
			System.Data.DataTable dt = new System.Data.DataTable();
			string connStr = ConfigurationSettings.AppSettings["myConnectionString"];
			SqlConnection con = new SqlConnection(connStr);
			var cmd = new SqlCommand();
			cmd.Connection = con;
			cmd.CommandText = "SynchNewYearSurveyData";
			cmd.CommandType = CommandType.StoredProcedure;

			SqlDataAdapter adp = new SqlDataAdapter();
			adp.SelectCommand = cmd;

			adp.Fill(ds);
			if (ds.Tables.Count > 1)
			{
				return ds.Tables[1];
			}

			else
			{

				return ds.Tables[0];
			}

		}

		public IEnumerable<GetSurveyTranscation_Result> GetSurveyTranscation()
		{


			var result = db.ExecuteFunction<GetSurveyTranscation_Result>("GetSurveyTranscation", MergeOption.OverwriteChanges);
			IEnumerable<GetSurveyTranscation_Result> query = from order in result.AsEnumerable()
															 select order;


			return query;
		}

		public string ReturnAvg(string arr)
		{
			string dc = "";
			string arr1 = "";
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
					if (arr.Contains('.'))
					{

						arr = arr.Split('.')[0].ToString();

					}
					dc = arr + "st";
				}
				else if (arr.Contains('t'))
				{

					arr = arr.Split('t')[0];
					if (arr.Contains('.'))
					{

						arr = arr.Split('.')[0].ToString();

					}
					dc = arr + "th";

				}

				else if (arr.Contains('n'))
				{

					arr = arr.Split('n')[0];
					if (arr.Contains('.'))
					{

						arr = arr.Split('.')[0].ToString();

					}
					dc = arr + "nd";
				}

				else if (arr.Contains('r'))
				{

					arr = arr.Split('r')[0];
					if (arr.Contains('.'))
					{

						arr = arr.Split('.')[0].ToString();

					}
					dc = arr + "rd";
				}
				else if (arr == null)
				{

					arr = "0";
					dc = arr;
				}

				else
				{
					if (arr.Contains('.'))
					{

						arr = arr.Split('.')[0].ToString();

					}
					dc = arr;
				}




			}


			return dc;
		}

		//Methods to generate graphs for Key metrics report
		private string CreatePieChart(int headerHeight, string graphTitle, List<decimal> piePercents, List<Color> pieColors, List<string> _description, string imagePath)
		{
			try
			{
				Bitmap bmp = new Bitmap(700, 800);

				Graphics pieGraphics = Graphics.FromImage(bmp);
				pieGraphics.Clear(Color.White);
				Size pieSize = new Size(400, 400);



				//Check if sections add up to 100.
				int sum = 0;
				foreach (int percent in piePercents)
				{
					sum += percent;
				}


				int a = 0, b = 0;

				if (piePercents.Count == pieColors.Count)
				{
					int _radius = 100;
					Random _random = new Random();

					System.Drawing.Rectangle rect = new System.Drawing.Rectangle(new System.Drawing.Point(150, 130), pieSize);

					float startAngle = 0;
					for (int i = 0; i < piePercents.ToArray().Length; i++)
					{
						Color color = pieColors[i];
						float arcAngle = Convert.ToSingle((piePercents[i] * 360) / 100);
						using (SolidBrush brush = new SolidBrush(color))
						{
							pieGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
							pieGraphics.FillPie(brush, rect, startAngle, arcAngle);

							using (Pen pen = new Pen(Color.White, 1))
							{
								pieGraphics.DrawPie(pen, rect, startAngle, arcAngle);
							}

						}
						double centerAngle = (startAngle + arcAngle / 2) / 180 * Math.PI;

						//Calculate out the center point of the string region, also the center of the pie.
						PointF center = new PointF();
						center.X = (float)(rect.X + rect.Width / 2 + _radius / 2 * Math.Cos(centerAngle));
						center.Y = (float)(rect.Y + rect.Height / 2 + _radius / 2 * Math.Sin(centerAngle));

						if (piePercents[i] > 0)
						{
							//Get the region of the string
							SizeF size = pieGraphics.MeasureString(_description[i], new System.Drawing.Font("Arial", 15));
							//Calculate out the string rectangle.
							RectangleF stringRect = new RectangleF(15, 580, 685, 685);
							//Draw the string.
							using (Brush stringBrush = new SolidBrush(Color.Black))
							{
								StringFormat formater = new StringFormat() { Alignment = StringAlignment.Center };
								pieGraphics.FillRectangle(new SolidBrush(pieColors[i]), 30, 590 + b, 20, 20);
								pieGraphics.DrawString(_description[i], new System.Drawing.Font("Arial", 15), stringBrush, 120, 590 + b);
							}
						}

						StringFormat string_format = new StringFormat();
						string_format.Alignment = StringAlignment.Center;
						string_format.LineAlignment = StringAlignment.Center;

						// Find the center of the rectangle.
						float cx = (rect.Left + rect.Right) / 2f;
						float cy = (rect.Top + rect.Bottom) / 2f;

						// Place the label about 2/3 of the way out to the edge.
						float radius = (rect.Width + rect.Height) / 2f * 0.33f;
						double label_angle =
								Math.PI * (startAngle + arcAngle / 2f) / 180f;
						float x = cx + (float)(radius * Math.Cos(label_angle));
						float y = cy + (float)(radius * Math.Sin(label_angle));
						pieGraphics.DrawString(piePercents[i].ToString() + "%",
							new System.Drawing.Font("Arial", 15), new SolidBrush(Color.Black), x, y, string_format);

						startAngle += arcAngle;
						a = a + 20;
						b = b + 40;
					}

					//To create border around the Image
					System.Drawing.Rectangle borderRect = new System.Drawing.Rectangle(System.Drawing.Point.Empty, bmp.Size);
					ControlPaint.DrawBorder(pieGraphics, borderRect, Color.Black, ButtonBorderStyle.Solid);

					//TO draw Pie Chart Header
					RectangleF header = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = headerHeight, Width = 700 } };
					SolidBrush blueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
					pieGraphics.FillRectangle(blueBrush, header);

					//To write header text
					StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
					pieGraphics.DrawString(graphTitle, new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, header, format);
					pieGraphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(800, 0));
					pieGraphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(700, 0),
						new System.Drawing.Point(700, 800));
					pieGraphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 800),
						new System.Drawing.Point(700, 800));
					pieGraphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(0, 800));

					bmp.Save(imagePath);
					return "success";
				}
				else
					return "The total percentage is not 100 or percentage and colors count does not match";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateTwoPieCharts(string graphTitle, List<decimal> piePercents, List<Color> pieColors, List<string> _description, List<decimal> piePercents1, List<Color> pieColors1, List<string> _description1, string imagePath)
		{

			try
			{
				Bitmap bmp = new Bitmap(1000, 600);

				Graphics pieGraphics = Graphics.FromImage(bmp);
				pieGraphics.Clear(Color.White);
				Size pieSize = new Size(300, 300);


				//Check if sections add up to 100.
				int sum = 0;
				foreach (int percent in piePercents)
				{
					sum += percent;
				}

				if (piePercents.Count == pieColors.Count)
				{
					int _radius = 100;
					Random _random = new Random();

					System.Drawing.Rectangle rect = new System.Drawing.Rectangle(new System.Drawing.Point(100, 130), pieSize);

					float startAngle = 270;
					for (int i = 0; i < piePercents.ToArray().Length; i++)
					{
						Color color = pieColors[i];
						float arcAngle = Convert.ToSingle((piePercents[i] * 360) / 100);
						using (SolidBrush brush = new SolidBrush(color))
						{
							pieGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
							pieGraphics.FillPie(brush, rect, startAngle, arcAngle);

							using (Pen pen = new Pen(Color.White, 1))
							{
								pieGraphics.DrawPie(pen, rect, startAngle, arcAngle);
							}

						}
						double centerAngle = (startAngle + arcAngle / 2) / 180 * Math.PI;

						//Calculate out the center point of the string region, also the center of the pie.
						PointF center = new PointF();
						center.X = (float)(rect.X + rect.Width / 2 + _radius / 2 * Math.Cos(centerAngle));
						center.Y = (float)(rect.Y + rect.Height / 2 + _radius / 2 * Math.Sin(centerAngle));
						if (piePercents[i] > 0)
						{
							//Get the region of the string
							SizeF size = pieGraphics.MeasureString(_description[i], new System.Drawing.Font("Arial", 15));
							//Calculate out the string rectangle.
							RectangleF stringRect = new RectangleF(center.X - size.Width / 2, center.Y - size.Height / 2,
								size.Width, size.Height);
							//Draw the string.
							using (Brush stringBrush = new SolidBrush(Color.White))
							{
								pieGraphics.DrawString(_description[i], new System.Drawing.Font("Arial", 15), stringBrush, stringRect);
							}
						}
						startAngle += arcAngle;
					}

					System.Drawing.Rectangle rect_1 = new System.Drawing.Rectangle(new System.Drawing.Point(460, 100), pieSize);

					Pen pen1 = new Pen(Color.Black, 2);
					pieGraphics.DrawLine(pen1, new System.Drawing.Point(510, 300), new System.Drawing.Point(650, 300));

					for (int i = 0; i < piePercents1.ToArray().Length; i++)
					{
						Color color = pieColors1[i];
						float arcAngle = Convert.ToSingle(piePercents1[i] * 360);
						using (SolidBrush brush = new SolidBrush(color))
						{
							pieGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
							pieGraphics.FillPie(brush, rect_1, startAngle, arcAngle);

							using (Pen pen = new Pen(Color.White, 1))
							{
								pieGraphics.DrawPie(pen, rect_1, startAngle, arcAngle);
							}
						}

						double centerAngle = (startAngle + arcAngle / 2) / 180 * Math.PI;

						//Calculate out the center point of the string region, also the center of the pie.
						PointF center = new PointF();
						center.X = (float)(rect_1.X + rect_1.Width / 2 + _radius / 2 * Math.Cos(centerAngle));
						center.Y = (float)(rect_1.Y + rect_1.Height / 2 + _radius / 2 * Math.Sin(centerAngle));
						if (piePercents.Count() > 0)
						{
							//Get the region of the string
							SizeF size = pieGraphics.MeasureString(_description1[i], new System.Drawing.Font("Arial", 15));
							//Calculate out the string rectangle.
							RectangleF stringRect = new RectangleF(center.X - size.Width / 2, center.Y - size.Height / 2,
								size.Width, size.Height);
							//Draw the string.
							using (Brush stringBrush = new SolidBrush(Color.White))
							{
								pieGraphics.DrawString(_description1[i], new System.Drawing.Font("Arial", 15), stringBrush, stringRect);
							}
						}
						startAngle += arcAngle;
					}

					//To create border around the Image
					System.Drawing.Rectangle borderRect = new System.Drawing.Rectangle(System.Drawing.Point.Empty, bmp.Size);
					ControlPaint.DrawBorder(pieGraphics, borderRect, Color.Black, ButtonBorderStyle.Solid);

					//TO draw Pie Chart Header
					RectangleF header = new RectangleF() { Location = new PointF() { X = 5, Y = 5 }, Size = new SizeF() { Height = 100, Width = 1000 } };
					SolidBrush blueBrush = new SolidBrush(Color.FromArgb(255, 51, 102, 153));
					pieGraphics.FillRectangle(blueBrush, header);

					//To write header text
					StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
					pieGraphics.DrawString(graphTitle, new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, header, format);


					pieGraphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
							new System.Drawing.Point(1000, 0));
					pieGraphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1000, 0),
						new System.Drawing.Point(1000, 600));
					pieGraphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 600),
						new System.Drawing.Point(1000, 600));
					pieGraphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(0, 600));

					bmp.Save(imagePath);
					return "success";
				}
				else
					return "The total percentage is not 100 or percentage and colors count does not match";
			}
			catch (Exception ex)
			{

				return ex.Message;
			}
		}

		private string CreateEyeExamDataGraph(string imagePath)
		{
			try
			{
				//Calculate Data
				string[] headerBlockStringList1 = new string[5] { "Well Below Average", "Below Average", "Average", "Above Average", "Well Above Average" };
				string[] headerBlockStringList2 = new string[5] { "1st-19th Percentile", "20th-39th Percentile", "40th-59th Percentile", "60th-79th Percentile", "80th-99th Percentile" };

				List<decimal> eyeGlassesExams = new List<decimal>();
				eyeGlassesExams = GetQuintileLookUpValues("Lookup.PercentofCompleteEyeExamsByEyeGlass_J");

				List<decimal> contactEyeExams = new List<decimal>();
				contactEyeExams = GetQuintileLookUpValues("Lookup.PercentofCompleteEyeExamsByCLExams_J");

				List<decimal> healthyEyeExams = new List<decimal>();
				healthyEyeExams = GetQuintileLookUpValues("Lookup.PercentofCompleteEyeExamsByHealthyEyeExams_J");

				Bitmap bmp = new Bitmap(1050, 700);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);
				Pen pen = new Pen(Color.Black, 2);


				//Draw Header Rectangle 

				Size size1 = new Size(1050, 180);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				RectangleF headerRowStringRect = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 80, Width = 1100 } };
				StringFormat rowStringFormat = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				using (Brush stringBrush = new SolidBrush(Color.White))
				{
					graphics.DrawString("Percent of Complete Eye Exams by Type", new System.Drawing.Font("Arial", 25, FontStyle.Italic), stringBrush, headerRowStringRect, rowStringFormat);
				}

				int x1 = 0;
				for (int j = 0; j < 5; j++)
				{
					headerRowStringRect = new RectangleF() { Location = new PointF() { X = x1, Y = 82 }, Size = new SizeF() { Height = 60, Width = 180 } };
					rowStringFormat = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
					using (Brush stringBrush = new SolidBrush(Color.White))
					{
						graphics.DrawString(headerBlockStringList1[j], new System.Drawing.Font("Arial", 20, FontStyle.Italic), stringBrush, headerRowStringRect, rowStringFormat);
					}

					headerRowStringRect = new RectangleF() { Location = new PointF() { X = x1, Y = 142 }, Size = new SizeF() { Height = 30, Width = 180 } };
					using (Brush stringBrush = new SolidBrush(Color.White))
					{
						graphics.DrawString(headerBlockStringList2[j], new System.Drawing.Font("Arial", 14, FontStyle.Italic), stringBrush, headerRowStringRect, rowStringFormat);
					}

					x1 += 210;
				}


				Size size2 = new Size(1050, 700);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 180), size2);
				graphics.DrawRectangle(pen, rect2);


				using (Brush stringBrush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(stringBrush, rect2);
				}


				//draw data rects
				Size dataRectSize1 = new Size(220, 40);
				int xBlock1;
				int yBlock1 = 210;
				string[] dataStringList1 = new string[3] { "EyeGlass Exams", "Contact Lens Exams", "Healthy Eye Exams" };

				Size dataRectSize2 = new Size(1015, 80);
				int xBlock2;
				int yBlock2 = 240;
				for (int i = 0; i < 3; i++)
				{
					xBlock1 = 15;
					System.Drawing.Point point = new System.Drawing.Point(xBlock1, yBlock1);
					System.Drawing.Rectangle dataRect = new System.Drawing.Rectangle(point, dataRectSize1);
					StringFormat format = new StringFormat() { Alignment = StringAlignment.Near };
					using (SolidBrush stringBrush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(dataStringList1[i], new System.Drawing.Font("Arial", 15, FontStyle.Regular), stringBrush, dataRect, format);
					}

					yBlock1 += 160;

					//Fill data rectangles 
					xBlock2 = 15;
					point = new System.Drawing.Point(xBlock2, yBlock2);
					dataRect = new System.Drawing.Rectangle(point, dataRectSize2);
					using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 208, 208, 208)))
					{
						graphics.FillRectangle(brush, dataRect);
					}

					format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
					for (int j = 1; j <= 5; j++)
					{
						RectangleF rectBlock = new RectangleF() { Location = new PointF() { X = xBlock2, Y = yBlock2 }, Size = new SizeF() { Height = 80, Width = 190 } };
						if (j == 3)
						{
							using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 102, 204, 51)))
							{
								graphics.FillRectangle(brush, rectBlock);
							}
						}

						switch (i)
						{
							case 0:
								graphics.DrawString(eyeGlassesExams[j - 1] + "%", new System.Drawing.Font("Arial", 15), Brushes.Black, rectBlock, format);
								break;

							case 1:
								graphics.DrawString(contactEyeExams[j - 1] + "%", new System.Drawing.Font("Arial", 15), Brushes.Black, rectBlock, format);
								break;

							case 2:
								graphics.DrawString(healthyEyeExams[j - 1] + "%", new System.Drawing.Font("Arial", 15), Brushes.Black, rectBlock, format);
								break;
						}
						xBlock2 += 210;
					}
					yBlock2 += 160;
				}

				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1050, 700)));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(1050, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1050, 0),
					new System.Drawing.Point(1050, 700));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 700),
					new System.Drawing.Point(1050, 700));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 700));
				bmp.Save(imagePath);
				return "success";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}

		}

		private string CreateExpenseCategoryDataGraph(string imagePath)
		{
			try
			{
				//Calculate Data
				string[] headerBlockStringList1 = new string[6] { "$537,000", "$803,000", "$1.1Million", "$1.5Million", "$2.1Million", "Total" };
				string[] headerBlockStringList2 = new string[6] { "1st-19th Percentile", "20th-39th Percentile", "40th-59th Percentile", "60th-79th Percentile", "80th-99th Percentile", "(median)" };

				string[] dataStringList1 = new string[7] { "Cost-of Goods (% of Gross Revenue)",
														   "Non-OD Staff (% of Gross Revenue)",
														   "General Overhead (% of Gross Revenue)",
														   "Occupancy (% of Gross Revenue)",
														   "Equipment (% of Gross Revenue)",
														   "Marketing (% of Gross Revenue)",
														   "Interest (% of Gross Revenue)" };

				List<decimal> costOfGoodsPec = new List<decimal>();
				costOfGoodsPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0 && x.Q24 <= 537000).Select(x => ((x.Q52j / x.Q24) * 100) ?? 0).ToList()));
				costOfGoodsPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 537000 && x.Q24 <= 803000).Select(x => ((x.Q52j / x.Q24) * 100) ?? 0).ToList()));
				costOfGoodsPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 803000 && x.Q24 <= 1100000).Select(x => ((x.Q52j / x.Q24) * 100) ?? 0).ToList()));
				costOfGoodsPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1100000 && x.Q24 <= 1500000).Select(x => ((x.Q52j / x.Q24) * 100) ?? 0).ToList()));
				costOfGoodsPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1500000 && x.Q24 <= 2100000).Select(x => ((x.Q52j / x.Q24) * 100) ?? 0).ToList()));
				costOfGoodsPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0).Select(x => ((x.Q52j / x.Q24) * 100) ?? 0).ToList()));

				List<decimal> nonODStaffPec = new List<decimal>();
				nonODStaffPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0 && x.Q24 <= 537000).Select(x => ((x.Q53 / x.Q24) * 100) ?? 0).ToList()));
				nonODStaffPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 537000 && x.Q24 <= 803000).Select(x => ((x.Q53 / x.Q24) * 100) ?? 0).ToList()));
				nonODStaffPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 803000 && x.Q24 <= 1100000).Select(x => ((x.Q53 / x.Q24) * 100) ?? 0).ToList()));
				nonODStaffPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1100000 && x.Q24 <= 1500000).Select(x => ((x.Q53 / x.Q24) * 100) ?? 0).ToList()));
				nonODStaffPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1500000 && x.Q24 <= 2100000).Select(x => ((x.Q53 / x.Q24) * 100) ?? 0).ToList()));
				nonODStaffPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0).Select(x => ((x.Q53 / x.Q24) * 100) ?? 0).ToList()));

				List<decimal> generalOvrHdPec = new List<decimal>();
				generalOvrHdPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0 && x.Q24 <= 537000).Select(x => ((x.Q57 / x.Q24) * 100) ?? 0).ToList()));
				generalOvrHdPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 537000 && x.Q24 <= 803000).Select(x => ((x.Q57 / x.Q24) * 100) ?? 0).ToList()));
				generalOvrHdPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 803000 && x.Q24 <= 1100000).Select(x => ((x.Q57 / x.Q24) * 100) ?? 0).ToList()));
				generalOvrHdPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1100000 && x.Q24 <= 1500000).Select(x => ((x.Q57 / x.Q24) * 100) ?? 0).ToList()));
				generalOvrHdPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1500000 && x.Q24 <= 2100000).Select(x => ((x.Q57 / x.Q24) * 100) ?? 0).ToList()));
				generalOvrHdPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0).Select(x => ((x.Q57 / x.Q24) * 100) ?? 0).ToList()));

				List<decimal> occupancyPec = new List<decimal>();
				occupancyPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0 && x.Q24 <= 537000).Select(x => ((x.Q54 / x.Q24) * 100) ?? 0).ToList()));
				occupancyPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 537000 && x.Q24 <= 803000).Select(x => ((x.Q54 / x.Q24) * 100) ?? 0).ToList()));
				occupancyPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 803000 && x.Q24 <= 1100000).Select(x => ((x.Q54 / x.Q24) * 100) ?? 0).ToList()));
				occupancyPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1100000 && x.Q24 <= 1500000).Select(x => ((x.Q54 / x.Q24) * 100) ?? 0).ToList()));
				occupancyPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1500000 && x.Q24 <= 2100000).Select(x => ((x.Q54 / x.Q24) * 100) ?? 0).ToList()));
				occupancyPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0).Select(x => ((x.Q54 / x.Q24) * 100) ?? 0).ToList()));

				List<decimal> equipmentPec = new List<decimal>();
				equipmentPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0 && x.Q24 <= 537000).Select(x => ((x.Q55 / x.Q24) * 100) ?? 0).ToList()));
				equipmentPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 537000 && x.Q24 <= 803000).Select(x => ((x.Q55 / x.Q24) * 100) ?? 0).ToList()));
				equipmentPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 803000 && x.Q24 <= 1100000).Select(x => ((x.Q55 / x.Q24) * 100) ?? 0).ToList()));
				equipmentPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1100000 && x.Q24 <= 1500000).Select(x => ((x.Q55 / x.Q24) * 100) ?? 0).ToList()));
				equipmentPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1500000 && x.Q24 <= 2100000).Select(x => ((x.Q55 / x.Q24) * 100) ?? 0).ToList()));
				equipmentPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0).Select(x => ((x.Q55 / x.Q24) * 100) ?? 0).ToList()));

				List<decimal> marketingPec = new List<decimal>();
				marketingPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0 && x.Q24 <= 537000).Select(x => ((x.Q56 / x.Q24) * 100) ?? 0).ToList()));
				marketingPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 537000 && x.Q24 <= 803000).Select(x => ((x.Q56 / x.Q24) * 100) ?? 0).ToList()));
				marketingPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 803000 && x.Q24 <= 1100000).Select(x => ((x.Q56 / x.Q24) * 100) ?? 0).ToList()));
				marketingPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1100000 && x.Q24 <= 1500000).Select(x => ((x.Q56 / x.Q24) * 100) ?? 0).ToList()));
				marketingPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1500000 && x.Q24 <= 2100000).Select(x => ((x.Q56 / x.Q24) * 100) ?? 0).ToList()));
				marketingPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0).Select(x => ((x.Q56 / x.Q24) * 100) ?? 0).ToList()));

				List<decimal> interestPec = new List<decimal>();
				interestPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0 && x.Q24 <= 537000).Select(x => ((x.Q58 / x.Q24) * 100) ?? 0).ToList()));
				interestPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 537000 && x.Q24 <= 803000).Select(x => ((x.Q58 / x.Q24) * 100) ?? 0).ToList()));
				interestPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 803000 && x.Q24 <= 1100000).Select(x => ((x.Q58 / x.Q24) * 100) ?? 0).ToList()));
				interestPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1100000 && x.Q24 <= 1500000).Select(x => ((x.Q58 / x.Q24) * 100) ?? 0).ToList()));
				interestPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 1500000 && x.Q24 <= 2100000).Select(x => ((x.Q58 / x.Q24) * 100) ?? 0).ToList()));
				interestPec.Add(GetListMedian(db.Source_InputDataBenchMarkSource.Where(x => x.Q24 != null && x.Q24 > 0).Select(x => ((x.Q58 / x.Q24) * 100) ?? 0).ToList()));

				Bitmap bmp = new Bitmap(1300, 1000);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);
				Pen pen = new Pen(Color.Black, 2);


				//Draw Header Rectangle 

				Size size1 = new Size(1300, 180);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				RectangleF headerRowStringRect = new RectangleF() { Location = new PointF() { X = 0, Y = 0 }, Size = new SizeF() { Height = 60, Width = 1300 } };
				StringFormat rowStringFormat = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				using (Brush stringBrush = new SolidBrush(Color.White))
				{
					graphics.DrawString("Expense Category % of Gross Revenue by Practice Size", new System.Drawing.Font("Arial", 25, FontStyle.Italic), stringBrush, headerRowStringRect, rowStringFormat);
					//graphics.DrawLine(new Pen(stringBrush), new System.Drawing.Point(30, 35), new System.Drawing.Point(670, 35));
				}

				int x1 = 20;
				for (int j = 0; j < 6; j++)
				{
					headerRowStringRect = new RectangleF() { Location = new PointF() { X = x1, Y = 61 }, Size = new SizeF() { Height = 50, Width = 220 } };
					rowStringFormat = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
					using (Brush stringBrush = new SolidBrush(Color.White))
					{
						graphics.DrawString(headerBlockStringList1[j], new System.Drawing.Font("Arial", 20, FontStyle.Italic), stringBrush, headerRowStringRect, rowStringFormat);
					}

					headerRowStringRect = new RectangleF() { Location = new PointF() { X = x1, Y = 111 }, Size = new SizeF() { Height = 30, Width = 220 } };
					using (Brush stringBrush = new SolidBrush(Color.White))
					{
						graphics.DrawString(headerBlockStringList2[j], new System.Drawing.Font("Arial", 15, FontStyle.Italic), stringBrush, headerRowStringRect, rowStringFormat);
					}

					x1 += 200;
				}

				//----Draw Header Rectangle End-----

				Size size2 = new Size(1300, 900);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 150), size2);
				graphics.DrawRectangle(pen, rect2);


				using (Brush stringBrush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(stringBrush, rect2);
				}


				//draw data rects
				Size dataRectSize1 = new Size(1300, 30);
				int xBlock1;
				int yBlock1 = 180;

				Size dataRectSize2 = new Size(1220, 40);
				int xBlock2;
				int yBlock2 = 220;
				for (int i = 0; i < 7; i++)
				{
					xBlock1 = 30;
					System.Drawing.Point point = new System.Drawing.Point(xBlock1, yBlock1);
					System.Drawing.Rectangle dataRect = new System.Drawing.Rectangle(point, dataRectSize1);
					StringFormat format = new StringFormat() { Alignment = StringAlignment.Near };
					using (SolidBrush stringBrush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(dataStringList1[i], new System.Drawing.Font("Arial", 18, FontStyle.Regular), stringBrush, dataRect, format);
					}

					yBlock1 += 120;

					//Fill data rectangles 
					xBlock2 = 30;
					point = new System.Drawing.Point(xBlock2, yBlock2);
					dataRect = new System.Drawing.Rectangle(point, dataRectSize2);
					using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 208, 208, 208)))
					{
						graphics.FillRectangle(brush, dataRect);
					}

					format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
					for (int j = 0; j < 6; j++)
					{
						RectangleF rectBlock = new RectangleF() { Location = new PointF() { X = xBlock2, Y = yBlock2 }, Size = new SizeF() { Height = 40, Width = 180 } };
						if (j == 5)
						{
							using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 102, 204, 51)))
							{
								graphics.FillRectangle(brush, rectBlock);
							}
						}
						switch (i)
						{
							case 0:
								graphics.DrawString(Math.Round(costOfGoodsPec[j]) + "%", new System.Drawing.Font("Arial", 18), Brushes.Black, rectBlock, format);
								break;
							case 1:
								graphics.DrawString(Math.Round(nonODStaffPec[j]) + "%", new System.Drawing.Font("Arial", 18), Brushes.Black, rectBlock, format);
								break;
							case 2:
								graphics.DrawString(Math.Round(generalOvrHdPec[j]) + "%", new System.Drawing.Font("Arial", 18), Brushes.Black, rectBlock, format);
								break;
							case 3:
								graphics.DrawString(Math.Round(occupancyPec[j]) + "%", new System.Drawing.Font("Arial", 18), Brushes.Black, rectBlock, format);
								break;
							case 4:
								graphics.DrawString(Math.Round(equipmentPec[j]) + "%", new System.Drawing.Font("Arial", 18), Brushes.Black, rectBlock, format);
								break;
							case 5:
								graphics.DrawString(Math.Round(marketingPec[j]) + "%", new System.Drawing.Font("Arial", 20), Brushes.Black, rectBlock, format);
								break;
							case 6:
								graphics.DrawString(Math.Round(interestPec[j]) + "%", new System.Drawing.Font("Arial", 20), Brushes.Black, rectBlock, format);
								break;
						}

						xBlock2 += 210;
					}
					yBlock2 += 120;
				}

				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1300, 1000)));

				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(1300, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1300, 0),
					new System.Drawing.Point(1300, 1000));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 1000),
					new System.Drawing.Point(1300, 1000));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 1000));
				bmp.Save(imagePath);
				return "success";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}

		}

		private string CreateAnnualMedicalVisitsDataGraph(string imagePath, List<string> eyeCareTypes, List<string> medianList, List<string> averageList)
		{
			try
			{
				Bitmap bmp = new Bitmap(1000, 600);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);


				//Draw Header Rectangle 
				Size size1 = new Size(1000, 120);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1000, 480);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 100), size2);
				//graphics.DrawRectangle(pen, rect2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString("Annual Medical Eye Care Visits by Type per 1,000 Active Patients", new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, rect1, format);


				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				Size size3 = new Size(900, 400);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(50, 150), size3);
				Pen pen = new Pen(Color.Black, 2);
				graphics.DrawRectangle(pen, rect3);
				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}

				int y = 200;

				for (int i = 0; i < eyeCareTypes.Count; i++)
				{
					int x = 70;
					format = new StringFormat() { Alignment = StringAlignment.Near };
					System.Drawing.Point pt;
					System.Drawing.Font font;
					Size sizeDtCol;
					for (int j = 0; j < 3; j++)
					{
						int colheight = 0;
						if (i == eyeCareTypes.Count - 1 || i == 0)
						{
							pt = new System.Drawing.Point(x, i != 0 ? y + 20 : y - 2);
							font = new System.Drawing.Font("Arial", 16, FontStyle.Regular);
							colheight = 50;
						}
						else
						{
							pt = new System.Drawing.Point(x, y + 10);
							font = new System.Drawing.Font("Arial", 16, FontStyle.Regular);
							colheight = 30;

						}
						System.Drawing.Rectangle dataRow;
						switch (j)
						{
							case 0:
								sizeDtCol = new Size(450, colheight);
								dataRow = new System.Drawing.Rectangle(pt, sizeDtCol);
								graphics.DrawString(eyeCareTypes[i], font, Brushes.Black, dataRow, format);
								x += 450;
								break;

							case 1:
								sizeDtCol = new Size(200, colheight);
								dataRow = new System.Drawing.Rectangle(pt, sizeDtCol);
								graphics.DrawString(medianList[i], font, Brushes.Black, dataRow, format);
								x += 200;
								break;

							case 2:
								sizeDtCol = new Size(200, colheight);
								dataRow = new System.Drawing.Rectangle(pt, sizeDtCol);
								graphics.DrawString(averageList[i], font, Brushes.Black, dataRow, format);
								x += 200;
								break;
						}
					}
					y += 30;
				}
				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1000, 600)));

				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(1000, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1000, 0),
					new System.Drawing.Point(1000, 600));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 600),
					new System.Drawing.Point(1000, 600));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 600));
				bmp.Save(imagePath);
				return "success";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateSoftLensReqDataGraph(string imagePath, List<string> practicalAnnGrossRev, List<string> medianList, List<string> softLenInv)
		{
			try
			{
				Bitmap bmp = new Bitmap(1300, 540);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);



				//Draw Header Rectangle 
				Size size1 = new Size(1300, 100);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1300, 500);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 100), size2);
				//graphics.DrawRectangle(pen, rect2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString("Soft Lens Inventory Requirements", new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, rect1, format);


				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				Size size3 = new Size(1200, 350);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(50, 140), size3);
				Pen pen = new Pen(Color.Black, 2);
				graphics.DrawRectangle(pen, rect3);
				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}

				int y = 160;

				for (int i = 0; i < practicalAnnGrossRev.Count; i++)
				{
					int x = 80;
					format = new StringFormat() { Alignment = StringAlignment.Near };
					System.Drawing.Point pt;
					System.Drawing.Font font;
					Size sizeDtCol;
					for (int j = 0; j < 3; j++)
					{
						int colheight = 0;
						if (i == practicalAnnGrossRev.Count - 1 || i == 0)
						{
							pt = new System.Drawing.Point(x, i != 0 ? y + 20 : y);
							font = new System.Drawing.Font("Arial", 18, FontStyle.Regular);
							colheight = 60;
						}
						else
						{
							pt = new System.Drawing.Point(x, y + 10);
							font = new System.Drawing.Font("Arial", 18);
							colheight = 60;

						}
						System.Drawing.Rectangle dataRow;
						switch (j)
						{
							case 0:
								sizeDtCol = new Size(425, colheight);
								dataRow = new System.Drawing.Rectangle(pt, sizeDtCol);
								graphics.DrawString(practicalAnnGrossRev[i], font, Brushes.Black, dataRow, format);
								x += 425;
								break;

							case 1:
								sizeDtCol = new Size(375, colheight);
								dataRow = new System.Drawing.Rectangle(pt, sizeDtCol);
								graphics.DrawString(medianList[i], font, Brushes.Black, dataRow, format);
								x += 375;
								break;

							case 2:
								sizeDtCol = new Size(375, colheight);
								dataRow = new System.Drawing.Rectangle(pt, sizeDtCol);
								graphics.DrawString(softLenInv[i], font, Brushes.Black, dataRow, format);
								x += 375;
								break;
						}
					}
					if (i == 0)
						y += 60;
					else
						y += 40;
				}
				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1300, 540)));

				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(1300, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1300, 0),
					new System.Drawing.Point(1300, 540));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 540),
					new System.Drawing.Point(1300, 540));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 540));
				bmp.Save(imagePath);
				return "success";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateAnnPurSoftLensModDataGraph(string imagePath, string graphHeader, List<string> yAxisData, List<string> graphData, List<string> xAxisData = null)
		{
			try
			{
				Bitmap bmp = new Bitmap(550, 250);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);



				//Draw Header Rectangle 
				Size size1 = new Size(550, 50);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(550, 200);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 52), size2);
				//graphics.DrawRectangle(pen, rect2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 15, FontStyle.Italic), Brushes.White, rect1, format);


				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				Size size3 = new Size(350, 150);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(150, 85), size3);
				Pen pen = new Pen(Color.Black, 1);
				//graphics.DrawRectangle(pen, rect3);
				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}

				graphics.DrawLine(pen, new System.Drawing.Point(150, 200), new System.Drawing.Point(500, 200));

				int distanceOfXAxisPts = 350 / (xAxisData.Count + 1);
				int x1 = 150;

				for (int i = 1; i <= xAxisData.Count; i++)
				{
					x1 += distanceOfXAxisPts;
					System.Drawing.Point p1 = new System.Drawing.Point() { X = x1, Y = 198 };
					System.Drawing.Point p2 = new System.Drawing.Point() { X = x1, Y = 202 };
					graphics.DrawLine(pen, p1, p2);
					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(xAxisData[i - 1], new System.Drawing.Font("Arial", 7, FontStyle.Bold), brush, x1 - 7, 205);
					}

				}

				Size barSize;
				int width = 130;
				int yBar = 120;
				int xBar = 150;
				for (int i = 1; i <= yAxisData.Count; i++)
				{
					barSize = new Size(width, 10);
					System.Drawing.Rectangle bar = new System.Drawing.Rectangle(new System.Drawing.Point(xBar, yBar), barSize);
					using (SolidBrush brush = new SolidBrush(Color.SkyBlue))
					{
						graphics.FillRectangle(brush, bar);
					}
					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(yAxisData[i - 1], new System.Drawing.Font("Arial", 7, FontStyle.Bold), brush, 60, yBar);
						graphics.DrawString(graphData[i - 1], new System.Drawing.Font("Arial", 7, FontStyle.Bold), brush, (xBar + width + 5), yBar);
					}
					yBar += 30;
					width += 170;
				}
				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(550, 250)));

				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(550, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(550, 0),
					new System.Drawing.Point(550, 250));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 250),
					new System.Drawing.Point(550, 250));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 250));
				bmp.Save(imagePath);
				return "success";

			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateSpectacleBarDataGraph(string imagePath, string graphHeader, List<string> yAxisData, List<decimal> xAxisData, List<decimal> graphData)
		{
			try
			{
				Bitmap bmp = new Bitmap(1200, 800);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);


				//Draw Header Rectangle 
				Size size1 = new Size(1200, 100);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1200, 700);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 100), size2);
				//graphics.DrawRectangle(pen, rect2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 30, FontStyle.Italic), Brushes.White, rect1, format);


				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				Size size3 = new Size(800, 480);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(400, 170), size3);
				Pen pen = new Pen(Color.Black, 1);
				//graphics.DrawRectangle(pen, rect3);
				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}
				graphics.DrawLine(pen, new System.Drawing.Point(400, 600), new System.Drawing.Point(400, 170));
				graphics.DrawLine(pen, new System.Drawing.Point(400, 600), new System.Drawing.Point(1120, 600));

				int distanceOfXAxisPts = 700 / (xAxisData.Count + 1);
				int x1 = 400;
				List<int> xAxisPtsX = new List<int>();
				for (int i = 1; i <= xAxisData.Count; i++)
				{
					x1 += distanceOfXAxisPts;
					System.Drawing.Point p1 = new System.Drawing.Point() { X = x1, Y = 596 };
					System.Drawing.Point p2 = new System.Drawing.Point() { X = x1, Y = 604 };
					graphics.DrawLine(pen, p1, p2);
					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(xAxisData[i - 1] + "x", new System.Drawing.Font("Arial", 20, FontStyle.Regular), brush, x1 - 7, 610);
					}
					xAxisPtsX.Add(x1);
				}

				Size barSize;
				int width = 0;
				int yBar = 240;
				int xBar = 400;
				for (int i = 1; i <= yAxisData.Count; i++)
				{
					width = Convert.ToInt32((graphData[i - 1] / (xAxisData.Max() + 0.5m)) * 700);
					barSize = new Size(width, 30);
					System.Drawing.Rectangle bar = new System.Drawing.Rectangle(new System.Drawing.Point(xBar, yBar), barSize);
					using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
					{
						graphics.FillRectangle(brush, bar);
					}
					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						if (i == 1)
						{
							graphics.DrawString("Single Vision Lenses", new System.Drawing.Font("Arial", 20, FontStyle.Regular), brush, 20, yBar - 25);
						}
						if (i == 4)
						{
							graphics.DrawString("Progressive Lenses", new System.Drawing.Font("Arial", 20, FontStyle.Regular), brush, 20, yBar - 25);
						}
						graphics.DrawString(yAxisData[i - 1], new System.Drawing.Font("Arial", 20, FontStyle.Regular), brush, 20, yBar);
						graphics.DrawString(graphData[i - 1] + "x", new System.Drawing.Font("Arial", 20, FontStyle.Regular), brush, (xBar + width + 5), yBar);
					}
					if (i != 3)
						yBar += 40;
					else
						yBar += 80;
				}

				using (SolidBrush brush = new SolidBrush(Color.Black))
				{
					graphics.DrawString("Average Mark-Up*", new System.Drawing.Font("Arial", 20, FontStyle.Regular), brush, new System.Drawing.Point(800, 120));
					graphics.DrawString("*Selling Price divided by cost-of Goods", new System.Drawing.Font("Arial", 20, FontStyle.Regular), brush, new System.Drawing.Point(400, 660));
				}

				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1200, 800)));



				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(1200, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1200, 0),
					new System.Drawing.Point(1200, 800));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 800),
					new System.Drawing.Point(1200, 800));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 800));
				bmp.Save(imagePath);
				return "success";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateIndRevGrowthBarDataGraph(string imagePath, string graphHeader, List<string> yAxisData, List<decimal> graphData)
		{
			try
			{
				Bitmap bmp = new Bitmap(1100, 600);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);

				//Draw Header Rectangle 
				Size size1 = new Size(1100, 140);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1100, 600);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 120), size2);
				//graphics.DrawRectangle(pen, rect2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, rect1, format);


				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				Size size3 = new Size(700, 340);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(310, 190), size3);
				Pen pen = new Pen(Color.Black, 1);
				//graphics.DrawRectangle(pen, rect3);
				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}

				format.Alignment = StringAlignment.Far;
				Size barSize;
				int width = 0;
				int yBar = 160;
				int xBar = 310;
				for (int i = 1; i <= yAxisData.Count; i++)
				{
					width = Convert.ToInt32((graphData[i - 1] / 700) * 1000);
					barSize = new Size(width, 30);
					System.Drawing.Rectangle bar = new System.Drawing.Rectangle(new System.Drawing.Point(xBar, yBar), barSize);
					System.Drawing.Rectangle yDataRect = new System.Drawing.Rectangle(new System.Drawing.Point(25, yBar), new Size(250, 40));
					using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
					{
						graphics.FillRectangle(brush, bar);
					}
					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(yAxisData[i - 1], new System.Drawing.Font("Arial", 16, FontStyle.Regular), brush, yDataRect, format);
						if (!(graphData[i - 1] <= 0))
							graphics.DrawString("+" + graphData[i - 1] + "%", new System.Drawing.Font("Arial", 16, FontStyle.Regular), brush, (xBar + width + 5), yBar);
					}
					yBar += 50;
				}

				using (SolidBrush brush = new SolidBrush(Color.Black))
				{
					graphics.DrawString("Source : PPA Estimations", new System.Drawing.Font("Arial", 16, FontStyle.Regular), brush, new System.Drawing.Point(60, 555));
				}

				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1100, 600)));

				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(1100, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1100, 0),
					new System.Drawing.Point(1100, 600));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 600),
					new System.Drawing.Point(1100, 600));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 600));

				bmp.Save(imagePath);
				return "success";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateVerticalGraph(int headerHeight, string imagePath, string graphHeader, List<string> xAxisData, List<decimal> graphData, bool isYAxisDataShown = false)
		{
			try
			{
				Bitmap bmp = new Bitmap(1100, 760);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);

				//Draw Header Rectangle 
				Size size1 = new Size(1100, headerHeight);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1100, 556);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 120), size2);
				//graphics.DrawRectangle(pen, rect2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, rect1, format);


				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				Size size3 = new Size(900, 500);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(65, 120), size3);
				Pen pen = new Pen(Color.Black, 1);
				//graphics.DrawRectangle(pen, rect3);
				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}
				graphics.DrawLine(pen, new System.Drawing.Point(60, 600), new System.Drawing.Point(60, 150));
				graphics.DrawLine(pen, new System.Drawing.Point(60, 600), new System.Drawing.Point(1030, 600));
				graphics.DrawString("% with Full-Time Office Manager", new System.Drawing.Font("Arial", 16, FontStyle.Regular), Brushes.Black, 20, 130, format);
				graphics.DrawString("Adds to 100%", new System.Drawing.Font("Arial", 16, FontStyle.Regular), Brushes.Black, 925, 130, format);
				if (graphHeader.Contains("Unit Sales Mix"))
				{
					graphics.DrawString("Frames Retail Price", new System.Drawing.Font("Arial", 16, FontStyle.Regular), Brushes.Black, 485, 250, format);

				}
				else if (graphHeader.Contains("Full Time Office"))
				{
					graphics.DrawString("Annual Gross Revenue", new System.Drawing.Font("Arial", 16, FontStyle.Regular), Brushes.Black, 485, 250, format);

				}
				//draw y axis points

				decimal yEnd = graphData.Max() % 10 == 0 ? graphData.Max() + 10 : Math.Ceiling(graphData.Max() / 10) * 10;

				if (isYAxisDataShown)
				{
					List<int> yAxisData = new List<int>();
					yAxisData = Enumerable.Range(10, Convert.ToInt32(yEnd)).Where(x => x % 10 == 0).Distinct().OrderBy(x => x).ToList();

					int distanceOfYAxisPts = 410 / (yAxisData.Count + 1);
					int y1 = 180;

					for (int i = yAxisData.Count; i >= 1; i--)
					{
						y1 += distanceOfYAxisPts;
						System.Drawing.Point p1 = new System.Drawing.Point() { X = 60, Y = y1 };
						System.Drawing.Point p2 = new System.Drawing.Point() { X = 80, Y = y1 };
						graphics.DrawLine(pen, p1, p2);
						using (SolidBrush brush = new SolidBrush(Color.Black))
						{
							graphics.DrawString(yAxisData[i - 1] + "%", new System.Drawing.Font("Arial", 16, FontStyle.Regular), brush, 15, y1 - 5);
						}
					}
				}

				Size barSize;
				int height = 0;
				int yBar = 0;
				int xBar = 120;
				int distanceOfXAxisPts = 960 / (xAxisData.Count);
				for (int i = 1; i <= xAxisData.Count; i++)
				{
					height = Convert.ToInt32((graphData[i - 1] / 410) * 1000);
					yBar = 600 - height;
					barSize = new Size(40, height);
					System.Drawing.Rectangle bar = new System.Drawing.Rectangle(new System.Drawing.Point(xBar, yBar), barSize);

					using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
					{
						graphics.FillRectangle(brush, bar);
					}
					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(xAxisData[i - 1], new System.Drawing.Font("Arial", 16, FontStyle.Regular), brush, new RectangleF(xBar - 30, 620, 140, 130), new System.Drawing.StringFormat() { Alignment = StringAlignment.Center, });
						if (!(graphData[i - 1] <= 0))
							graphics.DrawString(graphData[i - 1] + "%", new System.Drawing.Font("Arial", 16, FontStyle.Regular), brush, xBar, yBar - 30);
					}
					xBar += distanceOfXAxisPts;
				}

				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1100, 760)));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(1100, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1100, 0),
					new System.Drawing.Point(1100, 760));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 760),
					new System.Drawing.Point(1100, 760));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 760));
				bmp.Save(imagePath);
				return "success";

			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateHorizontalPlotGraph(string imagePath, string graphHeader, List<string> yAxisData, List<decimal> graphDataMin, List<decimal> graphDataMax)
		{
			try
			{
				Bitmap bmp = new Bitmap(1550, 800);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);


				//Draw Header Rectangle 
				Size size1 = new Size(1550, 100);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1550, 692);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 100), size2);
				//graphics.DrawRectangle(pen, rect2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 30, FontStyle.Italic), Brushes.White, rect1, format);

				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				graphics.DrawString("Expense Category", new System.Drawing.Font("Arial", 20, FontStyle.Regular), Brushes.Black, 360, 140);
				format.Alignment = StringAlignment.Far;
				graphics.DrawString("Range for middle 60% of MBA Practices", new System.Drawing.Font("Arial", 20, FontStyle.Regular), Brushes.Black, 1160, 140, format);

				Size size3 = new Size(1100, 560);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(360, 180), size3);

				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}

				Pen pen = new Pen(Color.Black, 1);
				int y = 180;
				int x = 70;
				int x1, x2, y1, y2 = 0;
				//Draw yAxis Data Rectangle 
				Size sizeYAxis;
				System.Drawing.Rectangle rectY;
				//get end data point of x axis
				decimal xEnd = Math.Ceiling(graphDataMax.Max() / 10) * 10;

				for (int i = 1; i <= yAxisData.Count; i++)
				{
					//Drawing Y Axis Data
					y += 605 / (yAxisData.Count + 1);
					sizeYAxis = new Size(360, 30);
					rectY = new System.Drawing.Rectangle(new System.Drawing.Point(0, y), sizeYAxis);
					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(yAxisData[i - 1], new System.Drawing.Font("Arial", 20, FontStyle.Regular), brush, rectY, new StringFormat() { Alignment = StringAlignment.Far });
					}

					//Drawing Graph Plots
					decimal temp = (graphDataMin[i - 1] / xEnd) * 1000;
					x1 = 450 + Convert.ToInt32(temp);

					decimal temp2 = (graphDataMax[i - 1] / xEnd) * 1000;
					x2 = 450 + Convert.ToInt32(temp2);
					y1 = y;

					int width = x2 - x1;
					sizeYAxis = new Size(width, 30);
					rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(x1, y1), sizeYAxis);
					using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
					{
						graphics.FillRectangle(brush, rect1);
					}
					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						//Display Minimum value only when its value is greater than 1(else will be displayed overlapped with Maximum values)
						if (graphDataMin[i - 1] > 1)
							graphics.DrawString(graphDataMin[i - 1] + "%", new System.Drawing.Font("Arial", 20, FontStyle.Regular), brush, x1, y1 + 11, new StringFormat() { Alignment = StringAlignment.Far });
						graphics.DrawString(graphDataMax[i - 1] + "%", new System.Drawing.Font("Arial", 20, FontStyle.Regular), brush, x2, y1 + 11, new StringFormat() { Alignment = StringAlignment.Near });
					}

				}

				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1550, 800)));

				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(1550, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1550, 0),
					new System.Drawing.Point(1550, 800));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 800),
					new System.Drawing.Point(1550, 800));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 800));


				bmp.Save(imagePath);
				return "success";

			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateHorizontalBarGraph(int headerHeight, string imagePath, string graphHeader, List<string> yAxisData, List<decimal> graphData, bool isXAxisDrawn = false)
		{
			try
			{
				Bitmap bmp = new Bitmap(1150, 500);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);

				//Draw Header Rectangle 
				Size size1 = new Size(1150, headerHeight);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1150, 400);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, headerHeight), size2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };

				if (graphHeader.Contains("(Average % change versus prior year)"))
				{
					Size size11 = new Size(1150, 140);
					System.Drawing.Rectangle rect11 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size11);
					graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, rect11, format);
				}
				else
					graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, rect1, format);

				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				if (graphHeader.Contains("Lens Modality"))
				{
					rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(600, 110), new Size(500, 60));
					graphics.DrawString("Median % of Patients Purchasing Annual Supply", new System.Drawing.Font("Arial", 16, FontStyle.Regular), Brushes.Black, rect2);
				}

				Size size3 = new Size(800, 200);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(310, 170), size3);
				Pen pen = new Pen(Color.Black, 1);
				//graphics.DrawRectangle(pen, rect3);
				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}

				graphics.DrawLine(pen, new System.Drawing.Point(310, 400), new System.Drawing.Point(1010, 400));
				graphics.DrawLine(pen, new System.Drawing.Point(310, 170), new System.Drawing.Point(310, 400));

				int length = Convert.ToInt32(Math.Round((Convert.ToInt32(graphData.Max()) / 4d), 0));
				//decimal xEnd = graphData.Max() % 10 == 0 ? graphData.Max() + 10 : Math.Ceiling(graphData.Max() / 10) * 10;
				if (isXAxisDrawn)
				{
					int x = 310;
					List<int> xAxisData = new List<int>();

					//xAxisData = Enumerable.Range(10, Convert.ToInt32(xEnd)).Where(x => x % 10 == 0).Distinct().OrderBy(x => x).ToList();

					//graphics.DrawLine(pen, new System.Drawing.Point(310, 400), new System.Drawing.Point(1010, 400));
					//graphics.DrawLine(pen, new System.Drawing.Point(310, 170), new System.Drawing.Point(310, 400));

					int distanceOfXAxisPts = 700 / 4;
					int x1 = 300;

					for (int i = 1; i <= 4; i++)
					{
						x = x + length;
						x1 += distanceOfXAxisPts;
						System.Drawing.Point p1 = new System.Drawing.Point() { X = x1, Y = 396 };
						System.Drawing.Point p2 = new System.Drawing.Point() { X = x1, Y = 404 };
						graphics.DrawLine(pen, p1, p2);
						using (SolidBrush brush = new SolidBrush(Color.Black))
						{
							graphics.DrawString(length + "%", new System.Drawing.Font("Arial", 16, FontStyle.Regular), brush, x1 - 7, 410);
						}
						length = length + Convert.ToInt32(Math.Round((Convert.ToInt32(graphData.Max()) / 4d), 0));
					}
				}

				int width = 0;
				int yBar = 170;
				int xBar = 310;
				length = Convert.ToInt32(Math.Round((Convert.ToInt32(graphData.Max()) / 4d), 0));
				System.Drawing.Rectangle yDataRect;
				System.Drawing.Rectangle barRect;
				for (int i = 1; i <= yAxisData.Count; i++)
				{
					yBar += 200 / (yAxisData.Count + 1);
					yDataRect = new System.Drawing.Rectangle(new System.Drawing.Point(1, yBar - 5), new Size(300, 100));
					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(yAxisData[i - 1], new System.Drawing.Font("Arial", 16, FontStyle.Regular), brush, yDataRect, new StringFormat() { Alignment = StringAlignment.Far });
					}

					width = Convert.ToInt32((graphData[i - 1] / length) * 175);
					barRect = new System.Drawing.Rectangle(new System.Drawing.Point(310, yBar), new Size(width, 40));
					using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
					{
						graphics.FillRectangle(brush, barRect);
					}

					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(graphData[i - 1] + "%", new System.Drawing.Font("Arial", 16, FontStyle.Regular), brush, (xBar + width + 5), yBar);
					}
				}

				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1150, 500)));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(1150, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1150, 0),
					new System.Drawing.Point(1150, 500));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 500),
					new System.Drawing.Point(1150, 500));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 500));

				bmp.Save(imagePath);
				return "success";

			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateHorizontalBarGraph1(string imagePath, string graphHeader, List<string> yAxisData, List<decimal> graphData, bool isXAxisDrawn = false)
		{
			try
			{
				Bitmap bmp = new Bitmap(1150, 720);

				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);


				//Draw Header Rectangle 
				Size size1 = new Size(1150, 100);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1150, 650);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 100), size2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, rect1, format);

				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				Size size3 = new Size(800, 200);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(310, 170), size3);
				Pen pen = new Pen(Color.Black, 1);
				//graphics.DrawRectangle(pen, rect3);
				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}

				decimal xEnd = graphData.Max() % 10 == 0 ? graphData.Max() + 10 : Math.Ceiling(graphData.Max() / 10) * 10;
				if (isXAxisDrawn)
				{
					List<int> xAxisData = new List<int>();

					xAxisData = Enumerable.Range(10, Convert.ToInt32(xEnd)).Where(x => x % 10 == 0).Distinct().OrderBy(x => x).ToList();

					graphics.DrawLine(pen, new System.Drawing.Point(310, 650), new System.Drawing.Point(1010, 650));
					graphics.DrawLine(pen, new System.Drawing.Point(310, 170), new System.Drawing.Point(310, 650));

					int distanceOfXAxisPts = 800 / (xAxisData.Count + 1);
					int x1 = 300;

					for (int i = 1; i <= xAxisData.Count; i++)
					{
						x1 += distanceOfXAxisPts;
						System.Drawing.Point p1 = new System.Drawing.Point() { X = x1, Y = 646 };
						System.Drawing.Point p2 = new System.Drawing.Point() { X = x1, Y = 654 };
						graphics.DrawLine(pen, p1, p2);
						using (SolidBrush brush = new SolidBrush(Color.Black))
						{
							graphics.DrawString(xAxisData[i - 1] + "%", new System.Drawing.Font("Arial", 15, FontStyle.Regular), brush, x1 - 7, 670);
						}

					}
				}

				int width = 0;
				int yBar = 170;
				int xBar = 310;

				System.Drawing.Rectangle yDataRect;
				System.Drawing.Rectangle barRect;
				for (int i = 1; i <= yAxisData.Count; i++)
				{
					yBar += 480 / (yAxisData.Count + 1);
					yDataRect = new System.Drawing.Rectangle(new System.Drawing.Point(1, yBar - 5), new Size(300, 100));
					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(yAxisData[i - 1], new System.Drawing.Font("Arial", 15, FontStyle.Regular), brush, yDataRect, new StringFormat() { Alignment = StringAlignment.Far });
					}

					width = Convert.ToInt32((graphData[i - 1] / xEnd) * 700);
					barRect = new System.Drawing.Rectangle(new System.Drawing.Point(310, yBar), new Size(width, 30));
					using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
					{
						graphics.FillRectangle(brush, barRect);
					}

					using (SolidBrush brush = new SolidBrush(Color.Black))
					{
						graphics.DrawString(graphData[i - 1] + "%", new System.Drawing.Font("Arial", 15, FontStyle.Regular), brush, (xBar + width + 5), yBar);
					}
				}

				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1150, 720)));


				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(1150, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1150, 0),
					new System.Drawing.Point(1150, 720));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 720),
					new System.Drawing.Point(1150, 720));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 720));

				bmp.Save(imagePath);
				return "success";

			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}


		private string CreateFramesInventoryTurnoverDataGraph(string imagePath, string graphHeader, List<decimal> framesInInv, List<decimal> annComSpecRx, List<decimal> annFramesTurnover, List<decimal> valueOfFramesInv, List<string> medians)
		{
			try
			{
				//-----------Data Fetch / Calculations----------------------------------
				List<string> headerList = new List<string>()
								{ "Price Size Decile", "Frames in Inventory", "Annual Complete Spec. Rxes", "Annual Frames Turnover","Value of Frames Inventory" };
				List<string> priceSizeDeciles = new List<string>() {  "$2,133,000 or more",
																					  "$1,695,000-$2,132,999",
																					  "$1,432,000-$1,694,999",
																					  "$1,200,000-$1,431,999",
																					  "$1,026,000-$1,199,999",
																					  "$883,000-$1,025,999",
																					  "$767,000-$882,999",
																					  "$642,000-$2,766,999",
																					  "$493,000-$641,999",
																					  "$492,999 or less",
																					};


				Bitmap bmp = new Bitmap(1130, 840);
				Graphics graphics = Graphics.FromImage(bmp);


				//Draw Header Rectangle 
				Size size1 = new Size(1130, 100);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1130, 740);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 100), size2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, rect1, format);


				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				Size size3 = new Size(1050, 700);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(40, 120), size3);
				Pen pen = new Pen(Color.Black, 1);

				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}

				format = new StringFormat() { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center };
				System.Drawing.Rectangle dataRect = new System.Drawing.Rectangle();
				graphics.DrawRectangle(pen, rect3);
				int widthOfDataCol = rect3.Width / 6;
				Size dataSizeCol1 = new Size(widthOfDataCol * 2, 100); ;
				Size dataSizeRestCols = new Size(widthOfDataCol, 100); ;
				int x1 = 50;
				int y1 = 130;
				for (int i = 1; i <= headerList.Count(); i++)
				{
					if (i == 1)
						dataRect = new System.Drawing.Rectangle(new System.Drawing.Point(x1 + 10, 130), dataSizeCol1);

					else
						dataRect = new System.Drawing.Rectangle(new System.Drawing.Point(x1 + 10, 130), dataSizeRestCols);

					graphics.DrawString(headerList[i - 1], new System.Drawing.Font("Arial", 15, FontStyle.Regular), Brushes.Black, dataRect, format);
					if (i == 1)
						x1 += widthOfDataCol * 2;
					else
						x1 += widthOfDataCol;
				}

				//Draw Price Size Decile
				x1 = 50;
				y1 = 210;
				foreach (var itemPrice in priceSizeDeciles)
				{
					dataRect = new System.Drawing.Rectangle(new System.Drawing.Point(x1 + 10, y1), dataSizeCol1);
					graphics.DrawString(itemPrice, new System.Drawing.Font("Arial", 15), Brushes.Black, dataRect, format);
					y1 += 50;

				}

				//Draw Frames In Inv
				x1 += dataRect.Width;
				y1 = 210;
				foreach (var itemFrames in framesInInv)
				{
					dataRect = new System.Drawing.Rectangle(new System.Drawing.Point(x1 + 10, y1), dataSizeRestCols);
					graphics.DrawString(Convert.ToString(itemFrames), new System.Drawing.Font("Arial", 15), Brushes.Black, dataRect, format);
					y1 += 50;
				}

				//Draw Annual Complete Spec Rxes
				x1 += dataRect.Width;
				y1 = 210;
				foreach (var itemCom in annComSpecRx)
				{
					dataRect = new System.Drawing.Rectangle(new System.Drawing.Point(x1 + 10, y1), dataSizeRestCols);
					graphics.DrawString(Convert.ToString(itemCom), new System.Drawing.Font("Arial", 15), Brushes.Black, dataRect, format);
					y1 += 50;
				}

				//Draw Annual Turnover
				x1 += dataRect.Width;
				y1 = 210;
				foreach (var itemTrn in annFramesTurnover)
				{
					dataRect = new System.Drawing.Rectangle(new System.Drawing.Point(x1 + 10, y1), dataSizeRestCols);
					graphics.DrawString(Convert.ToString(itemTrn), new System.Drawing.Font("Arial", 15), Brushes.Black, dataRect, format);
					y1 += 50;
				}

				//Draw Value of Frames Inventory
				x1 += dataRect.Width;
				y1 = 210;
				foreach (var itemValue in valueOfFramesInv)
				{
					dataRect = new System.Drawing.Rectangle(new System.Drawing.Point(x1 + 10, y1), dataSizeRestCols);
					graphics.DrawString(Convert.ToString(itemValue), new System.Drawing.Font("Arial", 15), Brushes.Black, dataRect, format);
					y1 += 50;
				}

				x1 = 50;
				for (int i = 1; i <= medians.Count(); i++)
				{
					if (i == 1)
						dataRect = new System.Drawing.Rectangle(new System.Drawing.Point(x1 + 10, y1 + 5), dataSizeCol1);

					else
						dataRect = new System.Drawing.Rectangle(new System.Drawing.Point(x1 + 10, y1 + 5), dataSizeRestCols);

					graphics.DrawString(medians[i - 1], new System.Drawing.Font("Arial", 15, FontStyle.Regular), Brushes.Black, dataRect, format);
					if (i == 1)
						x1 += widthOfDataCol * 2;
					else
						x1 += widthOfDataCol;

				}
				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1130, 840)));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
						new System.Drawing.Point(1130, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1130, 0),
					new System.Drawing.Point(1130, 840));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 840),
					new System.Drawing.Point(1130, 840));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 840));
				bmp.Save(imagePath);
				return "success";

			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateFramesInventoryGuidelineDataGraph(string imagePath, string graphHeader, List<string> medianAnnFrames, List<decimal> idealFramesInv, List<decimal> excessInv, List<decimal> insuffInv)
		{
			try
			{
				List<string> annGrossRev = new List<string>()
				{   "$500,000",
					"$800,000",
					"$1.1million",
					"$1.4million",
					"$2millions+"
				};

				List<string> guideList = new List<string>()
				{   "Median annual frames inventory turnover",
					"Ideal Frames Inventory ",
					"Excessive Inventory ",
					"Insufficient Inventory"
				};

				Bitmap bmp = new Bitmap(1250, 600);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);

				//Draw Header Rectangle 
				Size size1 = new Size(1250, 100);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1250, 540);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 64), size2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, rect1, format);


				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				Size size3 = new Size(1150, 460);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(60, 100), size3);
				Pen pen = new Pen(Color.Black, 1);
				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}
				graphics.DrawRectangle(pen, rect3);

				int widthOfDataCol = rect3.Width / 7;
				Size dataSizeCol1 = new Size(widthOfDataCol * 2, 60);
				Size dataSizeRestCols = new Size(widthOfDataCol, 40);
				int x1 = 70;
				int y1 = 110;
				System.Drawing.Point pt1 = new System.Drawing.Point(x1 + (widthOfDataCol * 2), y1);
				System.Drawing.Rectangle dataRect = new System.Drawing.Rectangle(pt1, new System.Drawing.Size(widthOfDataCol * 5, 30));

				y1 += 50;
				graphics.DrawString("Annual Gross Revenue", new System.Drawing.Font("Arial", 18, FontStyle.Regular), Brushes.Black, dataRect, format);
				graphics.DrawLine(pen, x1 + (widthOfDataCol * 2), y1, x1 + (widthOfDataCol * 7) - 20, y1);

				format.Alignment = StringAlignment.Near;
				y1 += 50;
				for (int j = 1; j <= 5; j++)
				{
					x1 = 70 + (widthOfDataCol * 2);
					for (int i = 0; i < annGrossRev.Count(); i++)
					{
						//if (i == 0)
						//    x1 += (widthOfDataCol * 2);
						pt1 = new System.Drawing.Point(x1, y1);
						dataRect = new System.Drawing.Rectangle(pt1, dataSizeRestCols);
						switch (j)
						{
							case 1:
								graphics.DrawString(annGrossRev[i], new System.Drawing.Font("Arial", 18, FontStyle.Regular | FontStyle.Underline), Brushes.Black, dataRect, format);
								break;
							case 2:
								graphics.DrawString(medianAnnFrames[i], new System.Drawing.Font("Arial", 18, FontStyle.Regular), Brushes.Black, dataRect, format);
								break;
							case 3:
								graphics.DrawString(Convert.ToString(idealFramesInv[i]), new System.Drawing.Font("Arial", 15, FontStyle.Regular), Brushes.Black, dataRect, format);
								break;
							case 4:
								graphics.DrawString(excessInv[i] + "+", new System.Drawing.Font("Arial", 18, FontStyle.Regular), Brushes.Black, dataRect, format);
								break;
							case 5:
								graphics.DrawString("<" + insuffInv[i], new System.Drawing.Font("Arial", 15, FontStyle.Regular), Brushes.Black, dataRect, format);
								break;
						}
						x1 += widthOfDataCol;
					}
					y1 += 60;
				}

				y1 = 260;
				x1 = 70;
				for (int i = 0; i < guideList.Count(); i++)
				{
					if (i == 0)
					{
						pt1 = new System.Drawing.Point(x1, y1);
						dataRect = new System.Drawing.Rectangle(pt1, dataSizeCol1);
					}
					else
					{
						pt1 = new System.Drawing.Point(x1, y1);
						dataRect = new System.Drawing.Rectangle(pt1, dataSizeCol1);
					}

					graphics.DrawString(guideList[i], new System.Drawing.Font("Arial", 15), Brushes.Black, dataRect, format);
					y1 += 60;
				}

				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1250, 600)));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(1250, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1250, 0),
					new System.Drawing.Point(1250, 600));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 600),
					new System.Drawing.Point(1250, 600));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 600));

				bmp.Save(imagePath);
				return "success";

			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		private string CreateStaffHourlySalariesByPositionDataGraph(string imagePath, string graphHeader, List<decimal> avgHourlySal, List<decimal> medianHourlySal, List<decimal> medianAnnualSal)
		{
			try
			{
				List<string> headerList = new List<string>()
				{   "Position",
					"Average Hourly Salary*",
					"Median Hourly Salary*",
					"Median Annual Salary*"
				};

				List<string> positionList = new List<string>()
				{   "Office Manager",
					"Optometric Assistant",
					"Contact Lens Technician",
					"Optician/Frames Stylist",
					"Lab Manager/Technician",
					"Receptionist",
					"Bookkeeper",
					"Insurance Clerk"
				};

				Bitmap bmp = new Bitmap(1120, 650);
				Graphics graphics = Graphics.FromImage(bmp);
				graphics.Clear(Color.White);

				//Draw Header Rectangle 
				Size size1 = new Size(1120, 100);
				System.Drawing.Rectangle rect1 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), size1);
				//graphics.DrawRectangle(pen, rect1);
				using (SolidBrush brush = new SolidBrush(Color.FromArgb(255, 51, 102, 153)))
				{
					graphics.FillRectangle(brush, rect1);
				}

				Size size2 = new Size(1120, 500);
				System.Drawing.Rectangle rect2 = new System.Drawing.Rectangle(new System.Drawing.Point(0, 100), size2);
				//To write header text
				StringFormat format = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
				graphics.DrawString(graphHeader, new System.Drawing.Font("Arial", 25, FontStyle.Italic), Brushes.White, rect1, format);


				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect2);
				}

				Size size3 = new Size(1000, 470);
				System.Drawing.Rectangle rect3 = new System.Drawing.Rectangle(new System.Drawing.Point(60, 120), size3);
				Pen pen = new Pen(Color.Black, 1);
				using (SolidBrush brush = new SolidBrush(Color.White))
				{
					graphics.FillRectangle(brush, rect3);
				}
				graphics.DrawRectangle(pen, rect3);

				int widthOfDataCol = rect3.Width / 5;
				Size dataSizeCol1 = new Size(widthOfDataCol * 2, 60);
				Size dataSizeRestCols = new Size(widthOfDataCol, 40);
				int x1 = 100;
				int y1 = 90;
				System.Drawing.Point pt1 = new System.Drawing.Point(x1 + (widthOfDataCol * 2), y1);
				System.Drawing.Rectangle dataRect = new System.Drawing.Rectangle(pt1, new System.Drawing.Size(widthOfDataCol * 5, 20));

				y1 += 120;
				////graphics.DrawString("Annual Gross Revenue", new System.Drawing.Font("Arial", 9, FontStyle.Bold), Brushes.Black, dataRect, format);
				graphics.DrawLine(pen, x1, y1, rect3.Width, y1);

				format.Alignment = StringAlignment.Near;

				for (int j = 1; j <= headerList.Count(); j++)
				{
					y1 = 140;
					pt1 = new System.Drawing.Point(x1, y1);
					if (j == 1)
						graphics.DrawString(headerList[j - 1], new System.Drawing.Font("Arial", 15, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(pt1, dataSizeCol1), format);
					else
					{
						Size tempSize = dataSizeRestCols;
						tempSize.Height = 60;
						graphics.DrawString(headerList[j - 1], new System.Drawing.Font("Arial", 15, FontStyle.Regular), Brushes.Black, new System.Drawing.Rectangle(pt1, tempSize), format);
					}

					y1 = 210;
					for (int i = 0; i < positionList.Count(); i++)
					{
						pt1 = new System.Drawing.Point(x1, y1);
						dataRect = new System.Drawing.Rectangle(pt1, dataSizeRestCols);
						switch (j)
						{
							case 1:
								dataRect = new System.Drawing.Rectangle(pt1, dataSizeCol1);
								graphics.DrawString(positionList[i], new System.Drawing.Font("Arial", 15, FontStyle.Regular), Brushes.Black, dataRect, format);
								break;
							case 2:
								graphics.DrawString("$" + avgHourlySal[i], new System.Drawing.Font("Arial", 15, FontStyle.Regular), Brushes.Black, dataRect, format);
								break;
							case 3:
								graphics.DrawString("$" + medianHourlySal[i], new System.Drawing.Font("Arial", 15, FontStyle.Regular), Brushes.Black, dataRect, format);
								break;
							case 4:
								graphics.DrawString("$" + medianAnnualSal[i], new System.Drawing.Font("Arial", 8, FontStyle.Bold), Brushes.Black, dataRect, format);
								break;
						}
						y1 += 45;
					}
					x1 += dataRect.Width - 10;
				}

				graphics.DrawRectangle(pen, new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new Size(1120, 650)));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(1120, 0));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(1100, 0),
					new System.Drawing.Point(1120, 650));
				graphics.DrawLine(new Pen(Brushes.Gray, 5), new System.Drawing.Point(0, 650),
					new System.Drawing.Point(1120, 650));
				graphics.DrawLine(new Pen(Brushes.Gray, 3), new System.Drawing.Point(0, 0),
					new System.Drawing.Point(0, 650));


				bmp.Save(imagePath);
				return "success";
			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}
	}
}