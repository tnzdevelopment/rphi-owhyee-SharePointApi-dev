using OfficeOpenXml;
using SharepointAPI.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Security;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Http;
using SHP = Microsoft.SharePoint.Client;
using SIO = System.IO;

namespace SharepointAPI.Controllers
{
    public class HomeController : ApiController
    {
        private const string siteUrl = "https://340bplus.sharepoint.com/sites/Home";
        private const string username = "TaimurB@pillrhealth.com";
        private const string password = "kind=act3693!";

        private static void SendEmail(string mailbody, string mailsubject)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(ConfigurationManager.AppSettings["MailFrom"]);
            
            string mailToConfig = ConfigurationManager.AppSettings["Mailto"];
            if (!string.IsNullOrWhiteSpace(mailToConfig))
            {
                var addresses = mailToConfig.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var address in addresses)
                    mail.To.Add(address.Trim());
            }

            mail.Subject = mailsubject;
            mail.Body = mailbody;

            SmtpClient smtpClient = new SmtpClient(ConfigurationManager.AppSettings["MailSMTP"]); // Replace with your SMTP host
            smtpClient.Port = int.Parse(ConfigurationManager.AppSettings["MailPort"]);
            smtpClient.EnableSsl = bool.Parse(ConfigurationManager.AppSettings["MailSSL"]); // Enable SSL/TLS if required by your server
            smtpClient.Credentials = new System.Net.NetworkCredential("imran.haq@tnzinternational.com", "Helloduniya#"); // Your SMTP server credentials

            try
            {
                smtpClient.Send(mail);
                // Email sent successfully
            }
            catch (SmtpException ex)
            {
                // Handle SMTP-specific errors
            }
            catch (Exception ex)
            {
                // Handle other general errors
            }

        }

        private static void GetSubFoldersFiles(SHP.ClientContext clientContext, SHP.Folder folder, string reportperiod, ref List<AuditReportModel> auditReportModels)
        {
            
            string connectionString = ConfigurationManager.ConnectionStrings["MyDatabaseConnection"].ConnectionString;

            clientContext.Load(folder, f => f.Name);
            clientContext.ExecuteQuery();


            string foldername = folder.Name;

            clientContext.Load(folder.Folders);
            clientContext.ExecuteQuery();

            foreach (SHP.Folder subFolder in folder.Folders)
            {
                List<int> descId = new List<int>();
                clientContext.Load(subFolder, f => f.Name);
                clientContext.ExecuteQuery();
                string workflowname = subFolder.Name;

                string fullpath= "/"+ foldername+"/"+workflowname;
                fullpath = fullpath.Replace(" ","%20");

                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();

                string selectQuery = "SELECT XWorkflowDescriptionID FROM [WorkFlow].[WorkflowFilePath] WHERE SharePointWorkflowFolder like @spflder";
                SqlCommand command = new SqlCommand(selectQuery, connection);
                command.Parameters.AddWithValue("@spflder", fullpath);

                SqlDataReader reader2 = command.ExecuteReader();
                while (reader2.Read())
                {
                    int productId = reader2.GetInt32(0); // By index
                    descId.Add(productId);

                    Console.WriteLine($"Product ID: {productId}");
                }

                reader2.Close();
                connection.Close();

                clientContext.Load(subFolder.Files);
                clientContext.ExecuteQuery();

                foreach (int descriptionid in descId)
                {
                    List<WorkflowFileTemplate> lstfiletemplate = new List<WorkflowFileTemplate>();

                    SqlConnection connectionfile = new SqlConnection(connectionString);
                    connectionfile.Open();

                    string selectFileQuery = "SELECT [FileFormatID],[ExcelSheetName],[HeaderRowStart] FROM [WorkFlow].[WorkflowFileTemplate] WHERE XWorkflowDescriptionID = @workflowid";
                    SqlCommand commandfile = new SqlCommand(selectFileQuery, connectionfile);
                    commandfile.Parameters.AddWithValue("@workflowid", descriptionid);

                    SqlDataReader filereader = commandfile.ExecuteReader();
                    while (filereader.Read())
                    {
                        WorkflowFileTemplate fltmp = new WorkflowFileTemplate();
                        fltmp.FileFormatId = filereader.GetInt32(0);
                        fltmp.ExcelSheetName = "";
                        if(!filereader.IsDBNull(1)) fltmp.ExcelSheetName=filereader.GetString(1);
                        fltmp.HeaderStart = null;
                        if(!filereader.IsDBNull(2)) fltmp.HeaderStart = filereader.GetInt32(2);
                        
                        lstfiletemplate.Add(fltmp);
                    }

                    filereader.Close();
                    connectionfile.Close();

                    foreach (SHP.File file in subFolder.Files)
                    {
                        string filename = file.Name;
                        //clientContext.Load(file);
                        SHP.ClientResult<SIO.Stream> stream = file.OpenBinaryStream();
                        clientContext.ExecuteQuery();

                        int rowCount = 0;
                        string fileextension = "";
                        string error = null;
                        using (StreamReader reader = new StreamReader(stream.Value))
                        {
                            if (filename.Contains("..")) error = "Bad file name.";

                            if (filename.Contains(".xlsx"))
                            {
                                ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization");
                                using (ExcelPackage package = new ExcelPackage(stream.Value))
                                {
                                    foreach (WorkflowFileTemplate template in lstfiletemplate)
                                    {
                                        if (template.FileFormatId == 1)
                                        {
                                            ExcelWorksheet worksheet =  package.Workbook.Worksheets[template.ExcelSheetName];
                                            if (worksheet != null)
                                            {
                                                //ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming first worksheet
                                                rowCount = worksheet.Dimension.End.Row; // Get the last row with data
                                                int hdstart = 0;
                                                if (template.HeaderStart != null) hdstart = int.Parse(template.HeaderStart.ToString());
                                                rowCount = rowCount - hdstart;
                                                fileextension = "xlsx";
                                            }
                                        }
                                    }
                                }
                            }
                            else if (filename.Contains(".csv"))
                            {
                                fileextension = filename.Substring(filename.Length - 3);
                                while (reader.ReadLine() != null)
                                {
                                    rowCount++;
                                }
                            }
                            else
                            {
                                fileextension = filename.Substring(filename.Length - 3);
                                while (reader.ReadLine() != null)
                                {
                                    rowCount++;
                                }
                            }
                            // rowCount now holds the number of lines
                        }

                        int checkcnt = rowCount;


                        AuditReportModel armodel = new AuditReportModel();
                        armodel.ReportingPeriod = reportperiod;
                        armodel.WorkFlowName = workflowname;
                        armodel.FileFolder = foldername;
                        armodel.FileName = filename;
                        armodel.FileExtension = fileextension;
                        armodel.RecordsCount = checkcnt.ToString();
                        armodel.Error = error;
                        auditReportModels.Add(armodel);
                    }
                }
            }
        }

        private static void UpdateWorkflowStatus(string query, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {

                    try
                    {
                        connection.Open();
                        int rowsAffected = command.ExecuteNonQuery();
                        connection.Close();
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }

        }

        [HttpPost]
        [Route("api/Home/PostHome")]
        public IHttpActionResult PostHome(PostModel data)
        {           
            string ReportPeriod = data?.ReportingPeriod;
            int WorkFlowId = data.StepId;
            DateTime dt = DateTime.Parse(ReportPeriod);
            int month = dt.Month;

            DateTimeFormatInfo dtf = CultureInfo.CurrentCulture.DateTimeFormat;
            string fullMonthName = dtf.GetMonthName(month);
            string year = dt.Year.ToString();

            string connectionString = ConfigurationManager.ConnectionStrings["MyDatabaseConnection"].ConnectionString;
            string returnMessage = "";
            List<AuditReportModel> auditReportModels = new List<AuditReportModel>();

            try
            {
                using (SHP.ClientContext clientContext = SharePointHelper.GetContext(siteUrl, username, password))
                {
                    // Perform operations on the SharePoint site
                    SHP.Web web = clientContext.Web;
                    clientContext.Load(web);
                    clientContext.ExecuteQuery();

                    SHP.List specificList = web.Lists.GetByTitle("Data Secure");
                    clientContext.Load(specificList);
                    clientContext.ExecuteQuery();

                    // Get the specific folder
                    //Folder parentFolder = web.GetFolderByServerRelativeUrl(specificList.RootFolder.ServerRelativeUrl + "/Invoices");
                    Guid folderid = Guid.Parse("d95aea5a-9598-2584-76e8-4091bd1e1f6b");
                    SHP.Folder parentFolder = web.GetFolderById(folderid);

                    clientContext.Load(parentFolder, f => f.ServerRelativeUrl);
                    clientContext.ExecuteQuery();

                    string folderServerRelativePath = parentFolder.ServerRelativeUrl;
                    string folderFullUrl = clientContext.Web.Url + parentFolder.ServerRelativeUrl;

                    string pathname = "/" + year + " Invoices/Invoices " + year + " " + fullMonthName;
                    SHP.Folder testFolder = web.GetFolderByServerRelativeUrl(folderServerRelativePath + pathname);
                    clientContext.Load(testFolder, f => f.ServerRelativeUrl);
                    clientContext.ExecuteQuery();

                    clientContext.Load(testFolder, f => f.Folders);
                    clientContext.ExecuteQuery();

                    // Iterate through subfolders
                    foreach (SHP.Folder subFolder in testFolder.Folders)
                    {
                        //                    Console.WriteLine($"  Subfolder: {subFolder.Name}");
                        GetSubFoldersFiles(clientContext, subFolder, ReportPeriod, ref auditReportModels);

                    }

                    List<AuditReportModel> armerrorlist = auditReportModels.Where(a => a.Error != null).ToList();

                    string endDate = DateTime.Now.ToString();


                    if (armerrorlist.Count == 0)
                    {
                        string mailbody = "Files successfully processed.";
                        string mailsubject = "Files successfully processed.";
                        SendEmail(mailbody, mailsubject);

                        string query = $"INSERT INTO [OwyheeWorkflow].[Audit].[AuditReportFiles] (ReportingPeriod, WorkFlowName, FileFolder, FileName,FileExtension, RecordsCount) VALUES (@ReportingPeriod, @WorkFlowName, @FileFolder, @FileName, @FileExtension, @RecordsCount)";

                        foreach (AuditReportModel arm in auditReportModels)
                        {
                            using (SqlConnection connection = new SqlConnection(connectionString))
                            {
                                using (SqlCommand command = new SqlCommand(query, connection))
                                {
                                    command.Parameters.AddWithValue("@ReportingPeriod", arm.ReportingPeriod);
                                    command.Parameters.AddWithValue("@WorkFlowName", arm.WorkFlowName);
                                    command.Parameters.AddWithValue("@FileFolder", arm.FileFolder);
                                    command.Parameters.AddWithValue("@FileName", arm.FileName);
                                    command.Parameters.AddWithValue("@FileExtension", arm.FileExtension);
                                    command.Parameters.AddWithValue("@RecordsCount", arm.RecordsCount);

                                    try
                                    {
                                        connection.Open();
                                        int rowsAffected = command.ExecuteNonQuery();
                                        connection.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error inserting data: {ex.Message}");
                                    }
                                }
                            }
                        }

                        query = "UPDATE [OwyheeWorkflow].[Log].[WorkFlowStepStatus] SET Status = 'Completed', StepEndDateTime = '" + endDate + "' WHERE WorkflowStepStatusId = '" + WorkFlowId + "'";

                        UpdateWorkflowStatus(query, connectionString);
                    }
                    else
                    {
                        string mailbody = "";
                        foreach (AuditReportModel armerror in armerrorlist)
                        {
                            mailbody = mailbody + "Bad file name - " + armerror.FileFolder + " - " + armerror.WorkFlowName + " - " + armerror.FileName + "\r\n";
                        }

                        string mailsubject = "Error in file processing.";
                        SendEmail(mailbody, mailsubject);

                        string query = "UPDATE [OwyheeWorkflow].[Log].[WorkFlowStepStatus] SET Status = 'Error', StepEndDateTime = '" + endDate + "', Error='" + mailbody + "' WHERE WorkflowStepStatusId = '" + WorkFlowId + "'";

                        UpdateWorkflowStatus(query, connectionString);
                    }
                }
            }
            catch (Exception ex)
            {
                string query = "UPDATE [OwyheeWorkflow].[Log].[WorkFlowStepStatus] SET Status = 'Error', StepEndDateTime = '" + DateTime.Now.ToString() + "', Error='" + ex.Message + "' WHERE WorkflowStepStatusId = '" + WorkFlowId + "'";

                UpdateWorkflowStatus(query, connectionString);

                returnMessage = ex.Message;
            }


            return Ok(returnMessage);
        }


        [HttpPost]
        [Route("api/Home/CheckDrugFileExistsOnSharePoint")]
        public async Task<IHttpActionResult> CheckDrugFileExistsOnSharePoint(PostModel data)
        {            
            string ReportPeriod = data?.ReportingPeriod;
            int WorkFlowId = data.StepId;
            DateTime dt = DateTime.Parse(ReportPeriod);
            int month = dt.Month;
            string query = string.Empty;

            DateTimeFormatInfo dtf = CultureInfo.CurrentCulture.DateTimeFormat;
            string fullMonthName = dtf.GetMonthName(month);
            string previousFullMonthName = dtf.GetMonthName(month - 1);
            string year = dt.Year.ToString();

            string pricingFolder = GetFileNameFromSharePointFolder(year, previousFullMonthName, "Invoices " + year + " " + previousFullMonthName + "/PDMI/Pricing");
            string medispan = GetFileNameFromSharePointFolder(year, fullMonthName, "SQL Reference Files/Medispan File/Active");
            string pvp = GetFileNameFromSharePointFolder(year, fullMonthName, "SQL Reference Files/PVP File/Active");

            string[] filenames = {
            pricingFolder,
            medispan,
            pvp
            };

            string result = "No";

            foreach (var filename in filenames)
            {
                var combinedResult = ExtractYearMonth(filename);

                if (combinedResult.HasValue && combinedResult.Value.Year == dt.Year && combinedResult.Value.Month == dt.Month)
                {
                    result = "Yes";
                }
                else
                {
                    result = "No";
                    break;
                }
            }


            string connectionString = ConfigurationManager.ConnectionStrings["MyDatabaseConnection"].ConnectionString;
            string returnMessage = "";
            List<AuditReportModel> auditReportModels = new List<AuditReportModel>();

            if (result == "Yes")
            {
                string mailbody = "Drug file located and process execution initiated. Notification will be sent upon completion!";
                string mailsubject = "Drug file located and process execution initiated.";
                SendEmail(mailbody, mailsubject);

                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = new SqlCommand())
                {
                    command.Connection = connection;
                    command.CommandText = @"INSERT INTO [Log].[WorkFlowStepStatus] 
                                ([StepId], [StepName], [ReportingPeriod], [Status], [StepStartDateTime])
                         VALUES (@StepId, @StepName, @ReportingPeriod, @Status, @StepStartDateTime)
                         SELECT CAST(SCOPE_IDENTITY() AS INT);";

                    // Add parameters
                    command.Parameters.AddWithValue("@StepId", 6);
                    command.Parameters.AddWithValue("@StepName", "DrugFileStep");
                    command.Parameters.AddWithValue("@ReportingPeriod", dt.Year + "-" + dt.Month + "-" + dt.Day);
                    command.Parameters.AddWithValue("@Status", "InProgress");
                    command.Parameters.AddWithValue("@StepStartDateTime", DateTime.Now);

                    connection.Open();
                    int WorkFlowStepStatusId = Convert.ToInt32(command.ExecuteScalar());
                    connection.Close();

                    command.Parameters.Clear();

                    string jobName = "3.0 Owhyee Load Drug Files";
                    try
                    {
                        ExecuteSQLAgentJob(connectionString, jobName);
                        returnMessage = "Yes";

                        command.CommandText = @"
                                 UPDATE [Log].[WorkFlowStepStatus] 
                                 SET 
                                    [Status] = @Status,                                    
                                    [StepEndDateTime] = @StepEndDateTime
                                 WHERE 
                                    [WorkflowStepStatusId] = @WorkflowStepStatusId";

                        // Add parameters
                        command.Parameters.AddWithValue("@Status", "Completed");
                        command.Parameters.AddWithValue("@StepEndDateTime", DateTime.Now);
                        command.Parameters.AddWithValue("@WorkflowStepStatusId", WorkFlowStepStatusId);

                        connection.Open();
                        command.ExecuteScalar();
                        connection.Close();

                        mailbody = "Drug file processing completed successfully.";
                        mailsubject = "Drug file processing completed successfully.";
                        SendEmail(mailbody, mailsubject);
                    }
                    catch (Exception ex)
                    {
                        returnMessage = $"Error: {ex.Message}";

                        command.CommandText = @"
                                 UPDATE [Log].[WorkFlowStepStatus] 
                                 SET 
                                    [Status] = @Status,
                                    [Error] = @Error,
                                    [StepEndDateTime] = @StepEndDateTime
                                 WHERE 
                                    [WorkflowStepStatusId] = @WorkflowStepStatusId";

                        // Add parameters
                        command.Parameters.AddWithValue("@Status", "ERROR");
                        command.Parameters.AddWithValue("@StepEndDateTime", DateTime.Now);
                        command.Parameters.AddWithValue("@WorkflowStepStatusId", WorkFlowStepStatusId);
                        command.Parameters.AddWithValue("@Error", ex.Message);

                        connection.Open();
                        command.ExecuteScalar();
                        connection.Close();
                        
                        mailbody = "Drug file Error Details are given below.\n \n" + ex.Message;
                        mailsubject = "Got error while processing Drug file.";
                        SendEmail(mailbody, mailsubject);
                    }
                }
            }

            else
            {
                string mailbody = "Drug files not found at the specified location.";
                string mailsubject = "Drug files not found!";
                SendEmail(mailbody, mailsubject);
            }

            returnMessage = result;
            return Ok(returnMessage);
        }

        private static string GetFileNameFromSharePointFolder(string year, string fullMonthName, string subPath)
        {
            using (SHP.ClientContext clientContext = SharePointHelper.GetContext(siteUrl, username, password))
            {
                SHP.Web web = clientContext.Web;
                clientContext.Load(web);
                clientContext.ExecuteQuery();

                SHP.List specificList = web.Lists.GetByTitle("Data Secure");
                clientContext.Load(specificList);
                clientContext.Load(specificList.RootFolder);
                clientContext.ExecuteQuery();

                string folderPath = specificList.RootFolder.ServerRelativeUrl +
                    "/Invoices/" + year + " Invoices/" + subPath;

                try
                {
                    SHP.Folder pricingFolder = web.GetFolderByServerRelativeUrl(folderPath);
                    clientContext.Load(pricingFolder);
                    clientContext.Load(pricingFolder.Files);
                    clientContext.ExecuteQuery();

                    if (pricingFolder.Files.Count > 0)
                    {
                        return pricingFolder.Files[0].Name;
                    }

                    return null;
                }
                catch (SHP.ServerException)
                {
                    // Folder does not exist
                    return null;
                }
            }
        }

        public static class SharePointHelper
        {
            public static SHP.ClientContext GetContext(string siteUrl, string username, string password)
            {
                SHP.ClientContext clientContext = new SHP.ClientContext(siteUrl);
                SecureString securePassword = new SecureString();
                foreach (char c in password)
                {
                    securePassword.AppendChar(c);
                }
                clientContext.Credentials = new SHP.SharePointOnlineCredentials(username, securePassword);
                return clientContext;
            }
        }

        static void ExecuteSQLAgentJob(string connectionString, string jobName)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("msdb.dbo.sp_start_job", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@job_name", jobName);
                    command.ExecuteNonQuery();
                }

                connection.Close();
            }
        }

        static (int Year, int Month)? ExtractYearMonth(string filename)
        {
            // Remove extension
            string nameWithoutExt = Path.GetFileNameWithoutExtension(filename);

            // Pattern 1: MM_DD_YYYY format
            var match1 = Regex.Match(nameWithoutExt, @"(\d{2})_(\d{2})_(\d{4})$");
            if (match1.Success)
            {
                int month = int.Parse(match1.Groups[1].Value);
                int year = int.Parse(match1.Groups[3].Value);
                return (year, month);
            }

            // Pattern 2: YYYYMMDD format
            var match2 = Regex.Match(nameWithoutExt, @"(\d{4})(\d{2})(\d{2})$");
            if (match2.Success)
            {
                int year = int.Parse(match2.Groups[1].Value);
                int month = int.Parse(match2.Groups[2].Value);
                return (year, month);
            }

            return null;
        }
    }
}
