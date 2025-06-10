using BoldReports.Web.ReportViewer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Mvc;
using BoldReports.Web.ReportViewer;
using BoldReports.Web;
using System.Reflection;
using System.Web;

namespace PPTRDLConverter.Controllers
{

    public class ReportViewerController : ApiController, IReportController,  IReportLogger
    {
        public void LogError(string message, Exception exception, MethodBase methodType, ErrorType errorType)
        {
            WriteLogs(string.Format("Error Message: {0} \n Stack Trace: {1}", message, exception.StackTrace));
        }

        public void LogError(string errorCode, string message, Exception exception, string errorDetail, string methodName, string className)
        {
            WriteLogs(string.Format("Class Name: {0} \n Method Name: {1} \n Error Message: {2} \n Stack Trace: {3}", className, methodName, errorDetail, exception.StackTrace));
        }

        internal void WriteLogs(string errorMessage)
        {
            string filePath = HttpContext.Current.Server.MapPath("/App_Data/Errordetails.txt");
            using (StreamWriter writer = new StreamWriter(filePath, true))
            {
                writer.AutoFlush = true;
                writer.WriteLine(errorMessage);
            }
        }
        // Post action for processing the RDL/RDLC report
        public object PostReportAction(Dictionary<string, object> jsonResult)
        {
            return ReportHelper.ProcessReport(jsonResult, this);
        }

        // Get action for getting resources from the report
        [System.Web.Http.ActionName("GetResource")]
        [System.Web.Http.AcceptVerbs("GET")]
        public object GetResource(string key, string resourcetype, bool isPrint)
        {
            return ReportHelper.GetResource(key, resourcetype, isPrint);
        }

        // Method that will be called when initialize the report options before start processing the report
        [System.Web.Http.NonAction]
        public void OnInitReportOptions(ReportViewerOptions reportOption)
        {
            // You can update report options here
        }

        // Method that will be called when reported is loaded
        [System.Web.Http.NonAction]
        public void OnReportLoaded(ReportViewerOptions reportOption)
        {
            //You can update report options here
        }
    }
}