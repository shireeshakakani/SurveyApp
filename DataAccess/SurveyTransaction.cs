//------------------------------------------------------------------------------
// <auto-generated>
//    This code was generated from a template.
//
//    Manual changes to this file may cause unexpected behavior in your application.
//    Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace PracticePerformanceAssessmentDataAccess
{
    using System;
    using System.Collections.Generic;
    
    public partial class SurveyTransaction
    {
        public int Id { get; set; }
        public int SurveyId { get; set; }
        public int PracticeId { get; set; }
        public string UserName { get; set; }
        public System.DateTime Entrydate { get; set; }
        public string DetailedPath { get; set; }
        public string InfographicPath { get; set; }
        public string ExecutivePath { get; set; }
        public string CSVPath { get; set; }
        public Nullable<bool> Isactive { get; set; }
        public int YearId { get; set; }
    }
}
