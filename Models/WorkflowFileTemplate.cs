using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharepointAPI.Models
{
    public class WorkflowFileTemplate
    {
        public int FileFormatId { get; set; }
        public string ExcelSheetName { get; set; }
        public int? HeaderStart { get; set; }
    }
}