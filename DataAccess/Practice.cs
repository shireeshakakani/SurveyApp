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
    
    public partial class Practice
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public Nullable<System.DateTime> CreatedOn { get; set; }
        public Nullable<bool> IsActive { get; set; }
        public int SurveyId { get; set; }
    
        public virtual SurveyTbl SurveyTbl { get; set; }
    }
}
