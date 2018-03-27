param(
    [string]$address,
    [string]$timeDelay, 
    [string]$doneStatusAs, 
    [string]$report, 
    [string]$schids,
    [string]$schedules
)


function Extract-Nupkg($nupkg, $out)
{
    Add-Type -AssemblyName 'System.IO.Compression.FileSystem' # PowerShell lacks native support for zip
    
    $zipFile = [IO.Compression.ZipFile]
    $zipFile::ExtractToDirectory($nupkg, $out)
}

function Add-NewtonsoftJsonReference
{
    $url = 'https://www.nuget.org/api/v2/package/Newtonsoft.Json'

    $directory = $PSScriptRoot, 'bin', 'Newtonsoft.Json' -Join '\'
    $nupkg = Join-Path $directory "Newtonsoft.Json.nupkg"
    $assemblyPath = $directory, 'lib', 'net45',  "Newtonsoft.Json.dll" -Join '\'
    
    if (Test-Path $assemblyPath)
    {
        # Already downloaded it from a previous script run/function call
        Add-Type -Path $assemblyPath
        return
    }
    
    ri -Recurse -Force $directory 2>&1 | Out-Null
    mkdir -f $directory | Out-Null # prevent this from being interpreted as a return value
    iwr $url -OutFile $nupkg
    Extract-Nupkg $nupkg -Out $directory
    Add-Type -Path $assemblyPath
}

function Add-NetHttpReference 
{
    $url = 'https://www.nuget.org/api/v2/package/System.Net.Http'

    $directory = $PSScriptRoot, 'bin', 'System.Net.Http' -Join '\'
    $nupkg = Join-Path $directory "System.Net.Http.nupkg"
    $assemblyPath = $directory, 'lib', 'net46',  "System.Net.Http.dll" -Join '\'
    
    if (Test-Path $assemblyPath)
    {
        # Already downloaded it from a previous script run/function call
        Add-Type -Path $assemblyPath
        return
    }
    
    ri -Recurse -Force $directory 2>&1 | Out-Null
    mkdir -f $directory | Out-Null # prevent this from being interpreted as a return value
    iwr $url -OutFile $nupkg
    Extract-Nupkg $nupkg -Out $directory
    Add-Type -Path $assemblyPath
}

function Add-XMLSerializationReference
{
	$url = 'https://www.nuget.org/api/v2/package/System.Xml.XmlSerializer'

    $directory = $PSScriptRoot, 'bin', 'System.Xml.XmlSerializer' -Join '\'
    $nupkg = Join-Path $directory "System.Xml.XmlSerializer"
    $assemblyPath = $directory, 'lib', 'netstandard1.3',  "System.Xml.XmlSerializer.dll" -Join '\'
    
    if (Test-Path $assemblyPath)
    {
        # Already downloaded it from a previous script run/function call
        Add-Type -Path $assemblyPath
        return
    }
    
    ri -Recurse -Force $directory 2>&1 | Out-Null
    mkdir -f $directory | Out-Null # prevent this from being interpreted as a return value
    iwr $url -OutFile $nupkg
    Extract-Nupkg $nupkg -Out $directory
    Add-Type -Path $assemblyPath
}

function Get-NewtonsoftJsonAssembly
{   
    return $PSScriptRoot, 'bin', "Newtonsoft.Json", 'lib','net45','Newtonsoft.Json.dll' -Join '\' 
}

function Get-NetHttpAssembly
{
    return $PSScriptRoot, 'bin', "System.Net.Http", 'lib','net46','System.Net.Http.dll' -Join '\'  
}

function Get-XMLSerializationAssembly
{
	    return $PSScriptRoot, 'bin', "System.Xml.XmlSerializer", 'lib','netstandard1.3','System.Xml.XmlSerializer.dll' -Join '\'  
}

Add-NetHttpReference
Add-NewtonsoftJsonReference
#Add-XMLSerializationReference



$assemblies = @()
$assemblies += Get-NewtonsoftJsonAssembly 
$assemblies += Get-NetHttpAssembly
#$assemblies += Get-XMLSerializationAssembly
$assemblies += "System.Runtime, Version=4.0.20.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
$assemblies += "System.IO, Version=4.0.10.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
$assemblies += "System.Xml.ReaderWriter, Version=4.0.10.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
$assemblies += "System.Xml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"

$sourceCode = @"
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;
using System.Threading;
using System.Threading.Tasks;


namespace TFSIntegrationConsole
{
    public class SimpleLogger
    {
        private string DatetimeFormat;
        private string Filename;

        /// <summary>
        /// Initialize a new instance of SimpleLogger class.
        /// Log file will be created automatically if not yet exists, else it can be either a fresh new file or append to the existing file.
        /// Default is create a fresh new log file.
        /// </summary>
        public SimpleLogger(string filePath)
        {
            DatetimeFormat = "yyyy-MM-dd HH:mm:ss.fff";
            Filename = filePath;

            // Log file header line
            string logHeader = Filename + " is created.";
            if (!File.Exists(Filename))
            {
                WriteLine(String.Format("{0} {1}",DateTime.Now.ToString(DatetimeFormat), logHeader), false);
            }

        }

        public void Debug(string text)
        {
            Console.Out.WriteLine(text);
            WriteFormattedLog(LogLevel.DEBUG, text);
        }

        public void Error(string text)
        {
            Console.Error.WriteLine(text);
            WriteFormattedLog(LogLevel.ERROR, text);
        }

        public void Fatal(string text)
        {
            Console.Error.WriteLine(text);
            WriteFormattedLog(LogLevel.FATAL, text);
        }

        public void Info(string text)
        {
            Console.Out.WriteLine(text);
            WriteFormattedLog(LogLevel.INFO, text);
        }

        public void Trace(string text)
        {
            Console.Out.WriteLine(text);
            WriteFormattedLog(LogLevel.TRACE, text);
        }

        public void Warning(string text)
        {
            Console.Out.WriteLine(text);
            WriteFormattedLog(LogLevel.WARNING, text);
        }

        /// <summary>
        /// Format a log message based on log level
        /// </summary>
        private void WriteFormattedLog(LogLevel level, string text)
        {
            string pretext;
            switch (level)
            {
                case LogLevel.TRACE: pretext = String.Format("{0} {1} ", DateTime.Now.ToString(DatetimeFormat), "[TRACE]"); break;
                case LogLevel.INFO: pretext = String.Format("{0} {1} ", DateTime.Now.ToString(DatetimeFormat), "[INFO] "); break;
                case LogLevel.DEBUG: pretext = String.Format("{0} {1} ", DateTime.Now.ToString(DatetimeFormat), "[DEBUG]"); break;
                case LogLevel.WARNING: pretext = String.Format("{0} {1} ", DateTime.Now.ToString(DatetimeFormat), "[WARNING]"); break;
                case LogLevel.ERROR: pretext = String.Format("{0} {1} ", DateTime.Now.ToString(DatetimeFormat), "[ERROR]"); break;
                case LogLevel.FATAL: pretext = String.Format("{0} {1} ", DateTime.Now.ToString(DatetimeFormat), "[FATAL]"); break;
                default: pretext = String.Empty; break;
            }

            WriteLine(String.Format("{0}{1}  ", pretext, text)); //double space is required for .md file new-line formatting 
        }

        private void WriteLine(string text, bool append = true)
        {
            try
            {
                using (StreamWriter Writer = new StreamWriter(Filename, append, Encoding.UTF8))
                {
                    if (!text.Equals(String.Empty)) Writer.WriteLine(text);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("Problems with logs:");
                Console.Error.WriteLine(e.Message);
            }
            
        }


        [Flags]
        private enum LogLevel
        {
            TRACE,
            INFO,
            DEBUG,
            WARNING,
            ERROR,
            FATAL
        }
    }

    [XmlRoot(ElementName = "testsuites")]
    public class ScheduleCollection
    {
        public ScheduleCollection()
        {
            Schedules = new List<Schedule>();

            TotalTests = 0;
            PassedTests = 0;
            FailedTests = 0;
            Errors = 0;
            Disabled = 0;
            TotalTime = 0;
        }

        [XmlElement(ElementName = "testsuite")]
        public List<Schedule> Schedules { get; set; }

        [XmlAttribute(AttributeName = "total")]
        public uint TotalTests { get; set; }

        [XmlAttribute(AttributeName = "tests")]
        public uint PassedTests { get; set; }

        [XmlAttribute(AttributeName = "failures")]
        public uint FailedTests { get; set; }

        [XmlAttribute(AttributeName = "errors")]
        public uint Errors { get; set; }

        [XmlAttribute(AttributeName = "disabled")]
        public uint Disabled { get; set; }

        [XmlAttribute(AttributeName = "time")]
        public double TotalTime { get; set; }


        public void AddPassedTests(uint passed) { this.PassedTests += passed; }
        public void AddFailedTests(uint failed) { this.FailedTests += failed; }
        public void AddErrors(uint errors) { this.Errors += errors; }
        public void AddTotalTime(double time) { this.TotalTime += time; }
    }

    public class Schedule
    {
        public Schedule() { Cases = new List<Case>(); }

        public Schedule(string title)
        {
            Cases = new List<Case>();
            ScheduleTitle = title;
            Failed = 0;
            Passed = 0;
            Errors = 0;
            Time = 0;
        }

        public Schedule(string schid, string title)
        {
            Cases = new List<Case>();
            ScheduleId = schid;
            ScheduleTitle = title;
            Failed = 0;
            Passed = 0;
            Errors = 0;
        }

        [XmlAttribute(AttributeName = "name")]
        public string ScheduleTitle { get; set; }

        [XmlAttribute(AttributeName = "schId")]
        public string ScheduleId { get; set; }

        [XmlAttribute(AttributeName = "id")]
        public int Id { get; set; }

        [XmlAttribute(AttributeName = "passed")]
        public uint Passed { get; set; }

        [XmlAttribute(AttributeName = "failures")]
        public uint Failed { get; set; }

        [XmlAttribute(AttributeName = "tests")]
        public uint Total { get; set; }

        [XmlAttribute(AttributeName = "errors")]
        public uint Errors { get; set; }

        [XmlAttribute(AttributeName = "status")]
        public string Status { get; set; }

        [XmlAttribute(AttributeName = "error")]
        public string Error { get; set; }

        [XmlAttribute(AttributeName = "time")]
        public double Time { get; set; }

        [XmlElement(ElementName = "testcase")]
        public List<Case> Cases { get; set; }

        public void incErrors()
        {
            this.Errors++;
        }
    }

    public class Case
    {
        public Case() { }

        public Case(string caseTitle, string caseStatus, double elapsed, string schedule)
        {
            CaseName = caseTitle;
            CaseStatus = caseStatus;
            ElapsedTime = elapsed;
            classname = schedule;
            failure = null;
        }

        public Case(string caseTitle, string caseStatus, double elapsed, string stacktrace, string schedule)
        {
            CaseName = caseTitle;
            CaseStatus = caseStatus;
            ElapsedTime = elapsed;
            failure = new Failure(stacktrace);
            classname = schedule;
        }
        [XmlAttribute(AttributeName = "name")]
        public string CaseName { get; set; }

        [XmlAttribute(AttributeName = "status")]
        public string CaseStatus { get; set; }

        [XmlAttribute(AttributeName = "time")]
        public double ElapsedTime { get; set; }

        [XmlAttribute(AttributeName = "classname")]
        public string classname { get; set; }

        public Failure failure { get; set; }
    }

    public class Failure
    {
        public Failure() { }
        public Failure(string stacktrace)
        {
            Message = stacktrace;
            Type = "dummy";
        }

        [XmlAttribute(AttributeName = "message")]
        public string Message { get; set; }

        [XmlAttribute(AttributeName = "type")]
        public string Type { get; set; }
    }

    public class InvalidSchedule
    {
        public String Name
        {
            get; private set;
        }
        public String StackTrace
        {
            get; private set;
        }

        public InvalidSchedule(String name, String stackTrace)
        {
            this.Name = name;
            this.StackTrace = stackTrace;
        }
    }

    public enum RUN_RESULT
    {
        RUN_SUCCESS,
        RUN_FAIL,
        RUN_REPEAT
    }

    public class Messages
    {
        public static readonly String SCHEDULE_FORMAT = "{0}[{1}]";
        public static readonly String SCHEDULE_DETECTED = "Schedule {0}[{1}] successfully detected!";
        public static readonly String SCHEDULE_RUN_SUCCESS = "Schedule {0}[{1}] Launched Successfully!";

        public static readonly String SCHEDULE_TITLE_OR_ID_ARE_NOT_GOT = "Tried to get schedule title or id! Check connection to the controller!";
        public static readonly String SCHEDULE_RUN_FAILURE = "Failed to run {0}[{1}]!";
        public static readonly String SCHEDULE_STATE_FAILURE = "Tried to get {0}[{1}] state!";
        public static readonly String NO_SUCH_SCHEDULE = "No such schedule! This may occur if try to run schedule that controller does not have. It can be deleted. Or you simply have forgotten to select schedules after changing controller address;";
        public static readonly String NO_SUCH_SCHEDULE_WAS_FOUND = "Could not find {0}[{1}] schedule! It was likely deleted!";
        public static readonly String SCHEDULE_HAS_NO_CASES = "Schedule {0}[{1}] has no cases! Add them in your Leaptest studio and try again!";
        public static readonly String SCHEDULE_IS_RUNNING_NOW = "Schedule {0}[{1}] is already running or queued now! Try to run it again when it's finished or you can try stop it, then run it!";

        public static readonly String REPORT_FILE_NOT_FOUND = "Couldn't find report file! Wrong path! Press \"help\" button nearby \"report\" textbox! ";
        public static readonly String REPORT_FILE_CREATION_FAILURE = "Failed to create a report file!";

        public static readonly String CASE_CONSOLE_LOG_SEPARATOR = "----------------------------------------------------------------------------------------";
        public static readonly String SCHEDULE_CONSOLE_LOG_SEPARATOR = "//////////////////////////////////////////////////////////////////////////////////////";

        public static readonly String CASE_INFORMATION = "Case: {0} | Status: {1} | Elapsed: {2}";
        public static readonly String CASE_STACKTRACE_FORMAT = "{0} - {1}";

        public static readonly String GET_ALL_AVAILABLE_SCHEDULES_URI = "{0}/api/v1/runSchedules";
        public static readonly String GET_LEAPTEST_VERSION_AND_API_URI = "{0}/api/v1/misc/version";
        public static readonly String RUN_SCHEDULE_URI = "{0}/api/v1/runSchedules/{1}/runNow";
        public static readonly String STOP_SCHEDULE_URI = "{0}/api/v1/runSchedules/{1}/stop";
        public static readonly String GET_SCHEDULE_STATE_URI = "{0}/api/v1/runSchedules/state/{1}";

        public static readonly String INVALID_SCHEDULES = "INVALID SCHEDULES";
        public static readonly String PLUGIN_NAME = "Leaptest Integration";

        public static readonly String NO_SCHEDULES = "No Schedules to run! All schedules you've selected could be deleted. Or you simply have forgotten to select schedules after changing controller address;";

        public static readonly String PLUGIN_SUCCESSFUL_FINISH = "Leaptest for TFS  plugin  successfully finished!";
        public static readonly String PLUGIN_ERROR_FINISH = "Leaptest for TFS plugin finished with errors!";

        public static readonly String CONTROLLER_RESPONDED_WITH_ERRORS = "Controller responded with errors! Please check controller logs and try again! If does not help, try to restart controller.";
        public static readonly String PLEASE_CONTACT_SUPPORT = "If nothing helps, please contact support https://leaptest.com/support and provide the next information:  \n1.Plugin Logs  \n2.Leaptest and plugin version  \n3.Controller logs from the moment you've run the plugin.  \n4.Assets without videos if possible.  \nYou can find them {Path to Leaptest}/LEAPTEST/Assets  \nThank you";

        public static readonly String ERROR_CODE_MESSAGE = "Code: {0} Status: {1}!";
        public static readonly String COULD_NOT_CONNECT_TO = "Could not connect to {0}! Check it and try again! ";
        public static readonly String COULD_NOT_CONNECT_TO_BUT_WAIT = "Could not connect to {0}! Check connection! The plugin is waiting for connection reestablishment! ";
        public static readonly String CONNECTION_LOST = "Connection to controller is lost: {0}! Check connection! The plugin is waiting for connection reestablishment!";
        public static readonly String INTERRUPTED_EXCEPTION = "Interrupted exception: {0}!";
        public static readonly String EXECUTION_EXCEPTION = "Execution exception: {0}!";
        public static readonly String IO_EXCEPTION = "I/O exception: {0}!";
        public static readonly String EXCEPTION = "Exception: {0}!";
        public static readonly String CACHE_TIMEOUT_EXCEPTION = "Cache time out exception has occurred! Don't worry! This schedule {0}[{1}] will be run again later!";

        public static readonly String LICENSE_EXPIRED = "Your Leaptest license has expired. Please contact support https://leaptest.com/support";

        public static readonly String SCHEDULE_IS_STILL_RUNNING = "Schedule {0}[{1}] is still running!";

        public static readonly String STOPPING_SCHEDULE = "Stopping schedule {0}[{1}]!";
        public static readonly String STOP_SUCCESS = "Schedule {0}[{1}] stopped successfully!";
        public static readonly String STOP_FAIL = "Failed to stop schedule {0}[{1}]!";

        public static readonly String BUILD_SUCCEEDED = "SUCCEEDED";
        public static readonly String BUILD_SUCCEEDED_WITH_ISSUES = "SUCCEEDED_WITH_ISSUES";
        public static readonly String BUILD_FAILED = "FAILED";
        public static readonly String BUILD_ABORTED = "ABORTED";

        public static readonly String SCHEDULE_HAS_NOT_REFRESHED_YET = "Leaptest controller has not refreshed yet after refreshing/restarting";
    }

    public class Program
    {
         //Old versions of MSBuild do not support C# 6 features: Null Condition operators ?. http://bartwullems.blogspot.com.by/2016/03/tfs-build-error-invalid-expression-term.html 
        private static String DefaultTokenStringValueIfNull(String tokenName, JToken parentToken)
        {
            JToken token = parentToken.SelectToken(tokenName);
            if (token != null)
                return token.Value<string>();
            else
                return String.Empty;
        }

        private static String GetJunitReportFilePath(String reportFileName)
        {
            if (!reportFileName.Contains(".xml"))
            {
                reportFileName += ".xml";
            }
            return reportFileName;
        }

        private static List<String> GetRawScheduleList(String rawScheduleIds, String rawScheduleTitles)
        {
            List<String> rawScheduleList = new List<String>();

            Regex regex = new Regex("\n|, |,");

            if (!String.IsNullOrEmpty(rawScheduleIds))
            {
                String[] schidsArray = regex.Split(rawScheduleIds);
                for (int i = 0; i < schidsArray.Length; i++)
                {
                    if (!String.IsNullOrEmpty(schidsArray[i]))
                        rawScheduleList.Add(schidsArray[i]);
                }
            }

            if (!String.IsNullOrEmpty(rawScheduleTitles))
            {
                String[] testsArray = regex.Split(rawScheduleTitles);

                for (int i = 0; i < testsArray.Length; i++)
                {
                    if (!String.IsNullOrEmpty(testsArray[i]))
                        rawScheduleList.Add(testsArray[i]);
                }
            }
            else
            {
                throw new Exception(Messages.NO_SCHEDULES);
            }

            return rawScheduleList;
        }

        private static int GetTimeDelay(String rawTimeDelay)
        {
            int defaultTimeDelay = 3;
            Int32.TryParse(rawTimeDelay, out defaultTimeDelay);
            return defaultTimeDelay;
        }

        private static bool IsScheduleStillRunning (JObject jsonState)
        {
            String status = DefaultTokenStringValueIfNull("Status", jsonState); // String status = jsonState.SelectToken("Status")?.Value<string>();

            if (!String.IsNullOrEmpty(status) && (status.Equals("Running") || status.Equals("Queued")))
                return true;
            else
                return false;
        }

        private static uint CaseStatusCount(String statusName, JObject jsonState)
        {
            uint count = 0;
            uint.TryParse(DefaultTokenStringValueIfNull(String.Format("LastRun.{0}", statusName), jsonState), out count);
            return count;
        }

        private static String DefaultElapsedIfNull(JToken token)
        {
            if (token != null)
            {
                string rawElapsed = token.Value<string>();
                if (rawElapsed != null)
                    return rawElapsed;
                else
                    return "00:00:00.0000000";
            }
            else
                return "00:00:00.0000000";
        }

        private static double ParseExecutionTimeToSeconds(String rawExecutionTime)
        {
            Regex regex = new Regex(":|\\.");
            String[] ExecutionTotalTime = regex.Split(rawExecutionTime);

            return Double.Parse(ExecutionTotalTime[0]) * 60 * 60 +  //hours
                Double.Parse(ExecutionTotalTime[1]) * 60 +        //minutes
                Double.Parse(ExecutionTotalTime[2]) +             //seconds
                Double.Parse("0." + ExecutionTotalTime[3]);     //milliseconds
        }


        private static async Task<Dictionary<String, String>> GetSchedulesIdTitleDictionary(
                String leaptestAddress,
                List<String> rawScheduleList,
                ScheduleCollection buildResult,
                List<InvalidSchedule> invalidSchedules,
                SimpleLogger logger
        )
        {
            Dictionary<String, String> schedulesIdTitleDictionary = new Dictionary<string, string>();

            String scheduleListUri = String.Format(Messages.GET_ALL_AVAILABLE_SCHEDULES_URI, leaptestAddress);

            try
            {
                try
                {
                    using (HttpClient client = new HttpClient())
                    using (HttpResponseMessage response = await client.GetAsync(scheduleListUri)) //get schedule Title and/or Id
                    {
                        int statusCode = (int)response.StatusCode;
                        string status = response.StatusCode.ToString();

                        switch (statusCode)
                        {
                            case 200: //SUCCESS

                                using (HttpContent content = response.Content)
                                {
                                    string responseContent = await content.ReadAsStringAsync();

                                    JArray jsonScheduleList = JArray.Parse(responseContent);

                                    foreach (String rawSchedule in rawScheduleList)
                                    {
                                        bool isSuccessfullyMapped = false;

                                        foreach (JObject jsonSchedule in jsonScheduleList)
                                        {
                                            String Id = DefaultTokenStringValueIfNull("Id", jsonSchedule);
                                            String Title = DefaultTokenStringValueIfNull("Title", jsonSchedule);

                                            if (Id.Equals(rawSchedule)) //Id match
                                            {
                                                if (!schedulesIdTitleDictionary.ContainsValue(Title)) //Avoid repeat
                                                {
                                                    schedulesIdTitleDictionary.Add(rawSchedule, Title);
                                                    buildResult.Schedules.Add(new Schedule(rawSchedule, Title));
                                                    logger.Info(String.Format(Messages.SCHEDULE_DETECTED, Title, rawSchedule));
                                                }
                                                isSuccessfullyMapped = true;
                                            }

                                            if (Title.Equals(rawSchedule)) //Title match 
                                            {
                                                if (!schedulesIdTitleDictionary.ContainsKey(Id)) //Avoid repeat
                                                {
                                                    schedulesIdTitleDictionary.Add(Id, rawSchedule);
                                                    buildResult.Schedules.Add(new Schedule(Id, rawSchedule));
                                                    logger.Info(String.Format(Messages.SCHEDULE_DETECTED, rawSchedule, Id));
                                                }
                                                isSuccessfullyMapped = true;
                                            }
                                        }

                                        if (!isSuccessfullyMapped)
                                            invalidSchedules.Add(new InvalidSchedule(rawSchedule, Messages.NO_SUCH_SCHEDULE));
                                    }
                                }

                                break;

                            case 445://LICENSE EXPIRED
                                String errorMessage445 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage445 += String.Format("\n{0}", Messages.LICENSE_EXPIRED);
                                throw new Exception(errorMessage445);

                            case 500://CONTROLLER ERROR
                                String errorMessage500 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage500 += String.Format("\n{0}", Messages.CONTROLLER_RESPONDED_WITH_ERRORS);
                                throw new Exception(errorMessage500);

                            default:
                                String errorMessage = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                throw new Exception(errorMessage);
                        }
                    }
                }
                catch (HttpRequestException e)
                {
                    String connectionErrorMessage = String.Format(Messages.COULD_NOT_CONNECT_TO, e.Message);
                    throw new Exception(connectionErrorMessage);
                }
            }
            catch (Exception e)
            {
                logger.Error(Messages.SCHEDULE_TITLE_OR_ID_ARE_NOT_GOT);
                throw e;
            }

            return schedulesIdTitleDictionary;
        }


        private static async Task<RUN_RESULT> RunSchedule(
            String leaptestAddress,
            String scheduleId,
            String scheduleTitle,
            int currentScheduleIndex,
            ScheduleCollection buildResult,
            List<InvalidSchedule> invalidSchedules,
            SimpleLogger logger

        )
        {
            RUN_RESULT isSuccessfullyRun = RUN_RESULT.RUN_FAIL;

            String uri = String.Format(Messages.RUN_SCHEDULE_URI, leaptestAddress, scheduleId);

            try
            {
                try
                {
                    using (HttpClient client = new HttpClient())
                    using (HttpResponseMessage response = await client.PutAsync(uri, new StringContent(String.Empty))) //Send PUT request and launch schedule
                    {
                        int statusCode = (int)response.StatusCode;
                        string status = response.StatusCode.ToString();

                        switch (statusCode)
                        {
                            case 204://SUCCESS
                                isSuccessfullyRun = RUN_RESULT.RUN_SUCCESS;
                                string successMessage = String.Format(Messages.SCHEDULE_RUN_SUCCESS, scheduleTitle, scheduleId);
                                buildResult.Schedules[currentScheduleIndex].Id = currentScheduleIndex;
                                logger.Info(Messages.SCHEDULE_CONSOLE_LOG_SEPARATOR);
                                logger.Info(successMessage);
                                break;

                            case 404://SCHEDULE NOT FOUND
                                String errorMessage404 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage404 += String.Format("\n{0}", String.Format(Messages.NO_SUCH_SCHEDULE_WAS_FOUND, scheduleTitle, scheduleId));
                                throw new Exception(errorMessage404);

                            case 444://SCHEDULE HAS NO CASES
                                String errorMessage444 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage444 += String.Format("\n{0}", String.Format(Messages.SCHEDULE_HAS_NO_CASES, scheduleTitle, scheduleId));
                                throw new Exception(errorMessage444);

                            case 445://LICENSE EXPIRED
                                String errorMessage445 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage445 += String.Format("\n{0}", Messages.LICENSE_EXPIRED);
                                throw new Exception(errorMessage445);

                            case 448://CACHE TIMEOUT
                                String errorMessage448 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage448 += String.Format("\n{0}", String.Format(Messages.CACHE_TIMEOUT_EXCEPTION, scheduleTitle, scheduleId));
                                isSuccessfullyRun = RUN_RESULT.RUN_REPEAT;
                                Console.Error.WriteLine(errorMessage448);
                                break;

                            case 500://SCHEDULE IS RUNNING NOW
                                String errorMessage500 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage500 += String.Format("\n{0}", String.Format(Messages.SCHEDULE_IS_RUNNING_NOW, scheduleTitle, scheduleId));
                                throw new Exception(errorMessage500);

                            default:
                                String errorMessage = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                throw new Exception(errorMessage);
                        }
                    }
                }
                catch (HttpRequestException e)
                {
                    String connectionErrorMessage = String.Format(Messages.COULD_NOT_CONNECT_TO_BUT_WAIT, e.Message);
                    logger.Error(connectionErrorMessage);
                    return RUN_RESULT.RUN_REPEAT; //In case of problems with connection to controller the plugin will be waiting for connection reestablishment
                }
            }
            catch (Exception e)
            {
                String errorMessage = String.Format(Messages.SCHEDULE_RUN_FAILURE, scheduleTitle, scheduleId);
                logger.Error(errorMessage);
                logger.Error(e.Message);
                logger.Error(Messages.PLEASE_CONTACT_SUPPORT);
                buildResult.Schedules[currentScheduleIndex].Error = String.Format("{0}\n{1}", errorMessage, e.Message);
                buildResult.Schedules[currentScheduleIndex].incErrors();
                invalidSchedules.Add(new InvalidSchedule(String.Format(Messages.SCHEDULE_FORMAT, scheduleTitle, scheduleId), buildResult.Schedules[currentScheduleIndex].Error));
                return RUN_RESULT.RUN_FAIL;
            }

            return isSuccessfullyRun;

        }

        private static async Task<bool> StopSchedule(String leaptestAddress, String scheduleId, String scheduleTitle, SimpleLogger logger)
        {
            bool isSuccessfullyStopped = false;

            logger.Error(String.Format(Messages.STOPPING_SCHEDULE, scheduleTitle, scheduleId));
            String uri = String.Format(Messages.STOP_SCHEDULE_URI, leaptestAddress, scheduleId);
            try
            {
                using (HttpClient client = new HttpClient())
                using (HttpResponseMessage response = await client.PutAsync(uri, new StringContent(String.Empty))) //Send PUT request and launch schedule
                {

                    if (response.StatusCode == HttpStatusCode.NoContent)     // 204 Response means correct schedule launching 
                    {
                        logger.Error(String.Format(Messages.STOP_SUCCESS, scheduleTitle, scheduleId));
                        isSuccessfullyStopped = true;
                    }
                    else
                    {
                        String errorMessage = String.Format(Messages.ERROR_CODE_MESSAGE, (int)response.StatusCode, response.StatusCode.ToString());
                        throw new WebException(errorMessage);
                    }
                }
            }
            catch (Exception e)
            {
                logger.Error(String.Format(Messages.STOP_FAIL, scheduleTitle, scheduleId));
                logger.Error(e.Message);
                return isSuccessfullyStopped;
            }
            return isSuccessfullyStopped;
        }


        private static async Task<bool> GetScheduleState(
            String leaptestAddress,
            String scheduleId,
            String scheduleTitle,
            int currentScheduleIndex,
            String doneStatusValue,
            ScheduleCollection buildResult,
            List<InvalidSchedule> invalidSchedules,
            SimpleLogger logger
            )
        {
            bool isScheduleStillRunning = true;

            String uri = String.Format(Messages.GET_SCHEDULE_STATE_URI, leaptestAddress, scheduleId);

            try
            {

                try
                {
                    using (HttpClient client = new HttpClient())
                    using (HttpResponseMessage response = await client.GetAsync(uri)) //get schedule Title and/or Id
                    {
                        int statusCode = (int)response.StatusCode;
                        string status = response.StatusCode.ToString();

                        switch (statusCode)
                        {
                            case 200://SUCCESS

                                using (HttpContent content = response.Content)
                                {
                                    string responseContent = await content.ReadAsStringAsync();

                                    JObject jsonState = JObject.Parse(responseContent);

                                    String ScheduleId = DefaultTokenStringValueIfNull("ScheduleId", jsonState);

                                    if (IsScheduleStillRunning(jsonState))
                                        isScheduleStillRunning = true;
                                    else
                                    {
                                        isScheduleStillRunning = false;

                                        /////////Schedule Info

                                        String ScheduleTitle = DefaultTokenStringValueIfNull("LastRun.ScheduleTitle", jsonState);

                                        buildResult.Schedules[currentScheduleIndex].Time = TimeSpan.Parse(jsonState.SelectToken("LastRun.ExecutionTotalTime").Value<string>()).TotalSeconds;

                                        uint passedCount = CaseStatusCount("PassedCount", jsonState);
                                        uint failedCount = CaseStatusCount("FailedCount", jsonState);
                                        uint doneCount = CaseStatusCount("DoneCount", jsonState);

                                        if (doneStatusValue.Equals("Failed"))
                                            failedCount += doneCount;
                                        else
                                            passedCount += doneCount;


                                        ///////////AutomationRunItemsInfo
                                        List<string> AutomationRunId = new List<string>();
                                        foreach (JToken token in jsonState.SelectTokens("LastRun.AutomationRunItems[*].AutomationRunId")) AutomationRunId.Add(token.Value<string>());
                                        List<string> Statuses = new List<string>();
                                        foreach (JToken token in jsonState.SelectTokens("LastRun.AutomationRunItems[*].Status")) Statuses.Add(token.Value<string>());
                                        List<string> Elapsed = new List<string>();
                                        foreach (JToken token in jsonState.SelectTokens("LastRun.AutomationRunItems[*].Elapsed")) Elapsed.Add(DefaultElapsedIfNull(token));
                                        List<string> Environments = new List<string>();
                                        foreach (JToken token in jsonState.SelectTokens("LastRun.AutomationRunItems[*].Environment.Title")) Environments.Add(token.Value<string>());

                                        //CaseInfo
                                        List<string> CaseTitles = new List<string>();
                                        foreach (JToken token in jsonState.SelectTokens("LastRun.AutomationRunItems[*]"))
                                        {
                                            String caseTitle = DefaultTokenStringValueIfNull("Case.Title", token);
                                            if (String.IsNullOrEmpty(caseTitle))
                                                CaseTitles.Add(CaseTitles[CaseTitles.Count - 1]);
                                            else
                                                CaseTitles.Add(caseTitle);
                                        }


                                        for (int i = 0; i < AutomationRunId.Count; i++)
                                        {

                                            //double seconds = jsonArray.getJSONObject(i).getDouble("TotalSeconds");
                                            double seconds = ParseExecutionTimeToSeconds(Elapsed[i]);

                                            logger.Info(Messages.CASE_CONSOLE_LOG_SEPARATOR);

                                            if (Statuses[i].Equals("Failed") || (Statuses[i].Equals("Done") && doneStatusValue.Equals("Failed")) || Statuses[i].Equals("Error") || Statuses[i].Equals("Cancelled"))
                                            {
                                                if (Statuses[i].Equals("Error") || Statuses[i].Equals("Cancelled"))
                                                    failedCount++;

                                                //KeyframeInfo
                                                List<string> KeyFrameTimeStamps = new List<string>();
                                                foreach (JToken token in jsonState.SelectTokens(String.Format("LastRun.AutomationRunItems[{0}].Keyframes[*].Timestamp", i)))
                                                {
                                                    var value = (DateTime)token.ToObject(typeof(DateTime));
                                                    KeyFrameTimeStamps.Add(value.ToString("dd-MM-yyyy hh:mm:ss.fff"));
                                                }

                                                List<string> KeyFrameStatuses = new List<string>();
                                                foreach (JToken token in jsonState.SelectTokens(String.Format("LastRun.AutomationRunItems[{0}].Keyframes[*].Status", i))) KeyFrameStatuses.Add(token.Value<string>()); //Old versions of MSBuild do not support C# 6 features: Interpolated strings ($ operator), String.Format instead
                                                List<string> KeyFrameLogMessages = new List<string>();
                                                foreach (JToken token in jsonState.SelectTokens(String.Format("LastRun.AutomationRunItems[{0}].Keyframes[*].LogMessage", i))) KeyFrameLogMessages.Add(token.Value<string>());

                                                logger.Info(String.Format(Messages.CASE_INFORMATION, CaseTitles[i], Statuses[i], Elapsed[i]));

                                                StringBuilder fullKeyframes =  new StringBuilder();
                                                int currentKeyFrameIndex = 0;

                                                for (int j = 0; j < KeyFrameStatuses.Count; j++)
                                                {
                                                    string level = DefaultTokenStringValueIfNull(String.Format("LastRun.AutomationRunItems[{0}].Keyframes[{1}].Level", i, j), jsonState);
                                                    if (!String.IsNullOrEmpty(level) && !level.Contains("Trace"))
                                                    {
                                                        String keyFrame = String.Format(Messages.CASE_STACKTRACE_FORMAT, KeyFrameTimeStamps[currentKeyFrameIndex], KeyFrameLogMessages[currentKeyFrameIndex]);
                                                        logger.Info(keyFrame);
                                                        fullKeyframes.AppendLine(keyFrame);
                                                    }
                                                    currentKeyFrameIndex++;
                                                }

                                                fullKeyframes.AppendLine("Environment: " + Environments[i]);
                                                logger.Info("Environment: " + Environments[i]);
                                                buildResult.Schedules[currentScheduleIndex].Cases.Add(new Case(CaseTitles[i], Statuses[i], seconds, fullKeyframes.ToString(), ScheduleTitle/* + "[" + ScheduleId + "]"*/));
                                                fullKeyframes = null;
                                            }
                                            else
                                            {
                                                logger.Info(String.Format(Messages.CASE_INFORMATION, CaseTitles[i], Statuses[i], Elapsed[i]));
                                                buildResult.Schedules[currentScheduleIndex].Cases.Add(new Case(CaseTitles[i], Statuses[i], seconds, ScheduleTitle/* + "[" + ScheduleId + "]"*/));
                                            }
                                        }

                                        buildResult.Schedules[currentScheduleIndex].Passed = passedCount;
                                        buildResult.Schedules[currentScheduleIndex].Failed = failedCount;

                                        if (buildResult.Schedules[currentScheduleIndex].Failed > 0)
                                            buildResult.Schedules[currentScheduleIndex].Status = "Failed";
                                        else
                                            buildResult.Schedules[currentScheduleIndex].Status = "Passed";
                                    }
                                }

                                break;

                            case 404://SCHEDULE NOT FOUND, LIKELY WAS DELETED WHILE PREVIOUS SCHEDULES WERE RUNNING
                                String errorMessage404 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage404 += String.Format("\n{0}", String.Format(Messages.NO_SUCH_SCHEDULE_WAS_FOUND, scheduleTitle, scheduleId));
                                throw new Exception(errorMessage404);

                            case 444://SCHEDULE HAS NOT REFRESHED YET AFTER START/RESTART
                                String errorMessage444 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage444 += String.Format("\n{0}", Messages.SCHEDULE_HAS_NOT_REFRESHED_YET);
                                isScheduleStillRunning = true;
                                logger.Error(errorMessage444);
                                break;

                            case 445://LICENSE EXPIRED
                                String errorMessage445 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage445 += String.Format("\n{0}", Messages.LICENSE_EXPIRED);
                                throw new Exception(errorMessage445);

                            case 448://CACHE TIMEOUT
                                String errorMessage448 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage448 += String.Format("\n{0}", String.Format(Messages.CACHE_TIMEOUT_EXCEPTION, scheduleTitle, scheduleId));
                                isScheduleStillRunning = true;
                                logger.Error(errorMessage448);
                                break;

                            case 500://CONTROLLER ERROR
                                String errorMessage500 = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                errorMessage500 += String.Format("\n{0}", String.Format(Messages.CONTROLLER_RESPONDED_WITH_ERRORS, scheduleTitle, scheduleId));
                                throw new Exception(errorMessage500);

                            default:
                                String errorMessage = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                                throw new Exception(errorMessage);
                        }
                    }
                }
                catch (HttpRequestException e)  //In case of problems with connection to controller the plugin will be waiting for connection reestablishment
                {
                    String connectionLostErrorMessage = String.Format(Messages.CONNECTION_LOST, e.Message);
                    logger.Error(connectionLostErrorMessage);
                    return true;
                }
            }
            catch (Exception e)
            {
                String errorMessage = String.Format(Messages.SCHEDULE_STATE_FAILURE, scheduleTitle, scheduleId);
                logger.Error(errorMessage);
                logger.Error(e.Message);
                logger.Error(Messages.PLEASE_CONTACT_SUPPORT);
                buildResult.Schedules[currentScheduleIndex].Error = String.Format("{0}\n{1}", errorMessage, e.Message);
                buildResult.Schedules[currentScheduleIndex].incErrors();
                invalidSchedules.Add(new InvalidSchedule(String.Format(Messages.SCHEDULE_FORMAT, scheduleTitle, scheduleId), buildResult.Schedules[currentScheduleIndex].Error));
                return false;
            }

            return isScheduleStillRunning;
        }

        private static void CreateJunitReport(string reportPath, ScheduleCollection buildResult, SimpleLogger logger)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ScheduleCollection));

                using (FileStream fs = new FileStream(reportPath, FileMode.Create))
                {

                    serializer.Serialize(fs, buildResult);
                }
            }
            catch (FileNotFoundException e)
            {
                logger.Info(Messages.REPORT_FILE_NOT_FOUND);
                throw e;
            }
        }

        public static string Call(string LeaptestControllerURL, string time, string doneStatus, string report, string log, string ids, string titles)
        {

            SimpleLogger logger = new SimpleLogger(log);

            logger.Info(String.Format("Leaptest controller URL: {0}", LeaptestControllerURL));
            logger.Info(String.Format("Time Delay: {0}", time));
            logger.Info(String.Format("Done Status As: {0}", doneStatus));
            logger.Info(String.Format("Raw report file path: {0}", report));
            logger.Info(String.Format("Log file path: {0}", log));
            logger.Info(String.Format("Schedule ids: {0}", ids));
            logger.Info(String.Format("Schedule names: {0}", titles));

            Dictionary<String, String> schedulesIdTitleDictionary = null; // Id-Title
            List<InvalidSchedule> invalidSchedules = new List<InvalidSchedule>();
            ScheduleCollection buildResult = new ScheduleCollection();
            List<String> rawScheduleList = null;

            String junitReportPath = GetJunitReportFilePath(report); //checks if .xml in the path exists
            logger.Info(String.Format("Full Report file path: {0}", junitReportPath));

            String schId = null;
            String schTitle = null;

            rawScheduleList = GetRawScheduleList(ids, titles);

            int timeDelay = GetTimeDelay(time);

            try
            {
                schedulesIdTitleDictionary = GetSchedulesIdTitleDictionary(LeaptestControllerURL, rawScheduleList, buildResult, invalidSchedules, logger).Result;
                rawScheduleList = null;//don't need that anymore

                List<String> schIdsList = new List<String>(schedulesIdTitleDictionary.Keys);
                schIdsList.Reverse();

                int currentScheduleIndex = 0;
                bool needSomeSleep = false;   //this time is required if there are schedules to rerun left, caused by controller cache timeout exception
                while (schIdsList.Count != 0)
                {

                    if (needSomeSleep)
                    {
                        Thread.Sleep(timeDelay * 1000); //Time delay
                        needSomeSleep = false;
                    }

                    for (int i = schIdsList.Count - 1; i >= 0; i--)
                    {
                        schId = schIdsList[i];
                        schTitle = schedulesIdTitleDictionary[schId];
                        RUN_RESULT runResult = RunSchedule(LeaptestControllerURL, schId, schTitle, currentScheduleIndex, buildResult, invalidSchedules, logger).Result;
                        logger.Info("Current schedule index: " + currentScheduleIndex);

                        if (runResult.Equals(RUN_RESULT.RUN_SUCCESS)) // if schedule was successfully run
                        {
                            bool isStillRunning = true;

                            do
                            {
                                Thread.Sleep(timeDelay * 1000); //Time delay
                                isStillRunning = GetScheduleState(LeaptestControllerURL, schId, schTitle, currentScheduleIndex, doneStatus, buildResult, invalidSchedules, logger).Result;
                                if (isStillRunning) logger.Info(String.Format(Messages.SCHEDULE_IS_STILL_RUNNING, schTitle, schId));
                            }
                            while (isStillRunning);

                            schIdsList.RemoveAt(i);

                            currentScheduleIndex++;
                        }
                        else if (runResult.Equals(RUN_RESULT.RUN_REPEAT))
                        {
                            needSomeSleep = true;
                        }
                        else
                        {
                            schIdsList.RemoveAt(i);
                            currentScheduleIndex++;
                        }
                    }
                }

                schIdsList = null;
                schedulesIdTitleDictionary = null;


                if (invalidSchedules.Count > 0)
                {
                    logger.Info(Messages.INVALID_SCHEDULES);
                    buildResult.Schedules.Add(new Schedule(Messages.INVALID_SCHEDULES));

                    foreach (InvalidSchedule invalidSchedule in invalidSchedules)
                    {
                        logger.Info(invalidSchedule.Name);
                        buildResult.Schedules[buildResult.Schedules.Count - 1].Cases.Add(new Case(invalidSchedule.Name, "Failed", 0, invalidSchedule.StackTrace, "INVALID SCHEDULE"));
                    }
                }

                foreach (Schedule schedule in buildResult.Schedules)
                {
                    buildResult.AddFailedTests(schedule.Failed);
                    buildResult.AddPassedTests(schedule.Passed);
                    buildResult.AddErrors(schedule.Errors);
                    schedule.Total = schedule.Passed + schedule.Failed;
                    buildResult.AddTotalTime(schedule.Time);
                }
                buildResult.TotalTests = buildResult.FailedTests + buildResult.PassedTests;

                CreateJunitReport(junitReportPath, buildResult, logger);

                logger.Info(Messages.PLUGIN_SUCCESSFUL_FINISH);

                if (buildResult.Errors > 0 || buildResult.FailedTests > 0 || invalidSchedules.Count > 0)
                {
                    logger.Info(Messages.BUILD_SUCCEEDED_WITH_ISSUES);
                    return Messages.BUILD_SUCCEEDED_WITH_ISSUES;
                }
                else
                {
                    logger.Info(Messages.BUILD_SUCCEEDED);
                    return Messages.BUILD_SUCCEEDED;
                }           
            }

            catch (Exception e)
            {
                if (e is ThreadAbortException || e is ThreadInterruptedException)//In case of interruption the plugin tries to send HTTP Schedule STOP request
                {
                    String interruptedExceptionMessage = String.Format(Messages.INTERRUPTED_EXCEPTION, e.Message);
                    logger.Error(interruptedExceptionMessage);
                    bool isSuccessfullyStopped = StopSchedule(LeaptestControllerURL, schId, schTitle, logger).Result;
                    logger.Error(Messages.BUILD_ABORTED);
                    return Messages.BUILD_ABORTED;
                }
                else
                {
                    logger.Error(Messages.PLUGIN_ERROR_FINISH);
                    logger.Error(e.Message);
                    logger.Error(Messages.PLEASE_CONTACT_SUPPORT);
                    logger.Error(Messages.BUILD_FAILED);
                    return Messages.BUILD_FAILED;
                }
            }
            
        }
    }
}
"@

$logFile = Split-Path -Path $report
$logFile = $logFile, "LeaptestIntegrationAgent$Env:AGENT_ID.Build$Env:BUILD_BUILDID.log" -Join "\"

Add-Type -ReferencedAssemblies $assemblies -TypeDefinition $sourceCode -Language CSharp
$buildResult = [TFSIntegrationConsole.Program]::Call($address,$timeDelay,$doneStatusAs,$report,$logFile,$schids,$schedules)

if($buildResult -eq "SUCCEEDED")
{
	Write-Output "##vso[task.complete result=Succeeded;]DONE"
}
elseif($buildResult -eq "SUCCEEDED_WITH_ISSUES")
{
	Write-Output "##vso[task.complete result=SucceededWithIssues;]DONE"
}
elseif($buildResult -eq "ABORTED")
{
	Write-Output "##vso[task.complete result=Cancelled;]DONE"
}
else
{
	Write-Output "##vso[task.complete result=Failed;]DONE"
}

#Write-Output "##vso[task.addattachment type=Distributedtask.Core.Summary;name=Leaptest-TFS Integration Console Output;]$logFile"

Write-Output "##vso[build.uploadlog]$logFile"



