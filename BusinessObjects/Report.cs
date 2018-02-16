using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.DataAnnotations;
using System.Runtime.Serialization;

namespace BusinessObjects
{
    [DataContract]
    public class Report
    {
        [DataMember]
        public List<Input> lstInput { get; set; }
        
        [DataMember]
        public List<Output> lstOutput { get; set; }
        
        public Report()
        {
            lstInput = new List<Input>();
            lstOutput = new List<Output>();
        }
    }
}
