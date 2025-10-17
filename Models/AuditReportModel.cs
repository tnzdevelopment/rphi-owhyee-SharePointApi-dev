using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharepointAPI.Models
{
    public class AuditReportModel
    {
        public string ReportingPeriod { get; set; }
        public string WorkFlowName { get; set; }
        public string FileFolder { get; set; }
        public string FileName { get; set; }
        public string FileExtension { get; set; }
        public string RecordsCount { get; set; }
        public string Error { get; set; }

    }
}