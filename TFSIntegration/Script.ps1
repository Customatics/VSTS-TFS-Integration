$leapworkHostname = Get-VstsInput -Name leapworkHostname -Require
$leapworkPort = Get-VstsInput -Name leapworkPort -Require
$leapworkAccessKey = Get-VstsInput -Name leapworkAccessKey -Require
$leapworkTimeDelay = Get-VstsInput -Name leapworkTimeDelay
$leapworkDoneStatusAs = Get-VstsInput -Name leapworkDoneStatusAs -Require
$leapworkReport = Get-VstsInput -Name leapworkReport -Require
$leapworkSchids = Get-VstsInput -Name leapworkSchids
$leapworkSchedules = Get-VstsInput -Name leapworkSchedules -Require

function Get-NewtonsoftJsonAssembly
{   
    return $PSScriptRoot, 'ps_modules', "Newtonsoft.Json", 'lib','net45','Newtonsoft.Json.dll' -Join '\' 
}

function Get-NetHttpAssembly
{
    return $PSScriptRoot, 'ps_modules', "System.Net.Http", 'lib','net46','System.Net.Http.dll' -Join '\'  
}

$newtonsoft = Get-NewtonsoftJsonAssembly
$systemNetHttp = Get-NetHttpAssembly

Write-Host "newtonsoft path: $newtonsoft"
Write-Host "systemNetHttp path: $systemNetHttp"

[Reflection.Assembly]::LoadFile($newtonsoft)
[Reflection.Assembly]::LoadFile($systemNetHttp)

$assemblies = @()
$assemblies += $newtonsoft 
$assemblies += $systemNetHttp
$assemblies += "System.Runtime, Version=4.0.20.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
$assemblies += "System.IO, Version=4.0.10.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
$assemblies += "System.Xml.ReaderWriter, Version=4.0.10.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
$assemblies += "System.Xml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"


$sourceCode = @"
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using System.Threading;
using System.Threading.Tasks;


namespace TFSIntegrationConsole
{

    public class SimpleLogger
    {
        private readonly string DatetimeFormat;
        private readonly string Filename;

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
                WriteLine(String.Format("{0} {1}", DateTime.Now.ToString(DatetimeFormat), logHeader), false);
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
    public class RunCollection
    {
        public RunCollection()
        {
            Runs = new List<LeapworkRun>();

            TotalTests = 0;
            PassedTests = 0;
            FailedTests = 0;
            Errors = 0;
            Disabled = 0;
            TotalTime = 0;
        }

        [XmlElement(ElementName = "testsuite")]
        public List<LeapworkRun> Runs { get; set; }

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

        [XmlIgnore] private double totalTime;

        [XmlAttribute(AttributeName = "time")]
        public double TotalTime
        {
            get { return Math.Round(totalTime, 2); }
            set { totalTime = value; }
        }


        public void AddPassedTests(uint passed) { this.PassedTests += passed; }
        public void AddFailedTests(uint failed) { this.FailedTests += failed; }
        public void AddErrors(uint errors) { this.Errors += errors; }
        public void AddTotalTime(double time) { this.TotalTime += time; }
    }

    public class LeapworkRun
    {
        public LeapworkRun() { RunItems = new List<RunItem>(); }

        public LeapworkRun(string title)
        {
            RunItems = new List<RunItem>();
            ScheduleTitle = title;
            Failed = 0;
            Passed = 0;
            Errors = 0;
            Time = 0;
        }

        public LeapworkRun(string runId, string title, string scheduleId)
        {
            RunItems = new List<RunItem>();
            RunId = runId;
            ScheduleTitle = title;
            Failed = 0;
            Passed = 0;
            Errors = 0;
        }

        [XmlAttribute(AttributeName = "name")]
        public string ScheduleTitle { get; set; }

        [XmlAttribute(AttributeName = "schId")]
        public string RunId { get; set; }

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

        [XmlIgnore] private double time;

        [XmlAttribute(AttributeName = "time")]
        public double Time
        {
            get { return Math.Round(time, 2); }
            set { time = value; }
        }

        [XmlElement(ElementName = "testcase")]
        public List<RunItem> RunItems { get; set; }

        public void IncErrors()
        {
            this.Errors++;
            IncTotal();
        }
        public void IncPassed()
        {
            this.Passed++;
            IncTotal();
        }
        public void IncFailed()
        {
            this.Failed++;
            IncTotal();
        }

        private void IncTotal()
        {
            this.Total++;
        }

    }

    public class RunItem
    {
        public RunItem() { }

        public RunItem(string caseTitle, string flowStatus, double elapsed, string schedule)
        {
            FlowTitle = caseTitle;
            FlowStatus = flowStatus;
            ElapsedTime = elapsed;
            classname = schedule;
            failure = null;
        }

        public RunItem(string caseTitle, string flowStatus, double elapsed, string stacktrace, string schedule)
        {
            FlowTitle = caseTitle;
            FlowStatus = flowStatus;
            ElapsedTime = elapsed;
            failure = new Failure(stacktrace);
            classname = schedule;
        }
        [XmlAttribute(AttributeName = "name")]
        public string FlowTitle { get; set; }

        [XmlAttribute(AttributeName = "status")]
        public string FlowStatus { get; set; }

        [XmlIgnore] private double elapsedTime;

        [XmlAttribute(AttributeName = "time")]
        public double ElapsedTime
        {
            get { return Math.Round(elapsedTime, 2); }
            set { elapsedTime = value; }
        }

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

    public class Messages
    {


        public static readonly String SCHEDULE_FORMAT = "{0}[{1}]";
        public static readonly String SCHEDULE_DETECTED = "Schedule {0}[{1}] successfully detected!";
        public static readonly String SCHEDULE_RUN_SUCCESS = "Schedule {0}[{1}] Launched Successfully!";

        public static readonly String SCHEDULE_TITLE_OR_ID_ARE_NOT_GOT = "Tried to get schedule title or id! Check connection to the controller!";
        public static readonly String SCHEDULE_RUN_FAILURE = "Failed to run {0}[{1}]!";
        public static readonly String NO_SUCH_SCHEDULE = "No such schedule! This may occur if try to run schedule that controller does not have. It can be deleted. Or you simply have forgotten to select schedules after changing controller address;";
        public static readonly String NO_SUCH_SCHEDULE_WAS_FOUND = "Could not find {0}[{1}] schedule! It was likely deleted!";

        public static readonly String NO_SUCH_RUN_WAS_FOUND = "Could not find run {0} of {1} schedule!";
        public static readonly String NO_SUCH_RUN_ITEM_WAS_FOUND = "Could not find run item {0} of {1} schedule!";
        public static readonly String NO_SUCH_RUN = "No such run {0} !";

        public static readonly String REPORT_FILE_NOT_FOUND = "Couldn't find report file! Wrong path! Press \"help\" button nearby \"report\" textbox! ";
        public static readonly String REPORT_FILE_CREATION_FAILURE = "Failed to create a report file!";

        public static readonly String CASE_CONSOLE_LOG_SEPARATOR = "----------------------------------------------------------------------------------------";
        public static readonly String SCHEDULE_CONSOLE_LOG_SEPARATOR = "//////////////////////////////////////////////////////////////////////////////////////";

        public static readonly String CASE_INFORMATION = "Run Item: {0} | Status: {1} | Elapsed: {2}";
        public static readonly String CASE_STACKTRACE_FORMAT = "{0} - {1}";

        public static readonly String GET_ALL_AVAILABLE_SCHEDULES_URI = "{0}/api/v3/schedules";
        public static readonly String RUN_SCHEDULE_URI = "{0}/api/v3/schedules/{1}/runNow";
        public static readonly String STOP_RUN_URI = "{0}/api/v3/run/{1}/stop";
        public static readonly String GET_RUN_STATUS_URI = "{0}/api/v3/run/{1}/status";
        public static readonly String GET_RUN_ITEMS_IDS_URI = "{0}/api/v3/run/{1}/runItemIds";
        public static readonly String GET_RUN_ITEM_URI = "{0}/api/v3/runItems/{1}";
        public static readonly String GET_RUN_ITEM_KEYFRAMES = "{0}/api/v3/runItems/{1}/keyframes";

        public static readonly String INVALID_SCHEDULES = "INVALID SCHEDULES";
        public static readonly String PLUGIN_NAME = "Leapwork Integration";

        public static readonly String NO_SCHEDULES = "No Schedules to run! All schedules you've selected could be deleted. Or you simply have forgotten to select schedules after changing controller address;";

        public static readonly String PLUGIN_SUCCESSFUL_FINISH = "Leapwork for TFS  plugin  successfully finished!";
        public static readonly String PLUGIN_ERROR_FINISH = "Leapwork for TFS plugin finished with errors!";

        public static readonly String CONTROLLER_RESPONDED_WITH_ERRORS = "Controller responded with errors! Please check controller logs and try again! If does not help, try to restart controller.";
        public static readonly String PLEASE_CONTACT_SUPPORT = "If nothing helps, please contact support https://leapwork.com/chat and provide the next information:  \n1.Plugin Logs  \n2.Leapwork and plugin version  \n3.Controller logs from the moment you've run the plugin.  \n4.Assets without videos if possible.  \nYou can find them {Path to Leaptest}/LEAPTEST/Assets  \nThank you";

        public static readonly String ERROR_CODE_MESSAGE = "Code: {0} Status: {1}!";
        public static readonly String COULD_NOT_CONNECT_TO = "Could not connect to {0}! Check it and try again! ";
        public static readonly String COULD_NOT_CONNECT_TO_BUT_WAIT = "Could not connect to {0}! Check connection! The plugin is waiting for connection reestablishment! ";
        public static readonly String CONNECTION_LOST = "Connection to controller is lost: {0}! Check connection! The plugin is waiting for connection reestablishment!";
        public static readonly String INTERRUPTED_EXCEPTION = "Interrupted exception: {0}!";


        public static readonly String STOPPING_RUN = "Stopping schedule {0} run {1}!";

        public static readonly String STOP_RUN_FAIL = "Failed to stop schedule {0} run {1}!";

        public static readonly String INVALID_ACCESS_KEY = "Invalid or empty access key!";

        public static readonly String DATABASE_NOT_RESPONDING = "Data base is not responding!";

        public static readonly String INVALID_VARIABLE_KEY_NAME = "Variable name is invalid or variable with such name is already in request!";

        public static readonly String NO_DISK_SPACE = "No enough disk space to start schedule!";

        public static readonly String PORT_NUMBER_IS_INVALID = "Port number is invalid, setting to default {0}";

        public static readonly String TIME_DELAY_NUMBER_IS_INVALID = "Time delay number is invalid, setting to default {0}";


        public static readonly String SCHEDULE_DISABLED = "Schedule {0}[{1}] is disabled!";

        public static readonly String BUILD_SUCCEEDED = "SUCCEEDED";
        public static readonly String BUILD_SUCCEEDED_WITH_ISSUES = "SUCCEEDED_WITH_ISSUES";
        public static readonly String BUILD_FAILED = "FAILED";
        public static readonly String BUILD_ABORTED = "ABORTED";

        public static readonly String INVALID_BOOLEAN_TOKEN = "Failed to extract boolean token {0}, setting default value {1}";

        public static readonly String FAILED_TO_PARSE_STRING_TO_GUID = "Failed to parse token {0} {1} to guid value";
        public static readonly String STRING_TOKEN_NOT_FOUND = "Failed to find string token {0}";

        public static readonly String INPUT_VALUES_MESSAGE = "LEAPWORK Plugin input parameters:";
        public static readonly String INPUT_HOSTNAME_VALUE = "LEAPWORK controller hostname: {0}";
        public static readonly String INPUT_PORT_VALUE = "LEAPWORK controller port: {0}";
        public static readonly String INPUT_ACCESS_KEY_VALUE = "LEAPWORK controller Access Key: {0}";
        public static readonly String INPUT_REPORT_VALUE = "JUnit report file name: {0}";
        public static readonly String INPUT_LOG_FILEPATH_VALUE = "Log file path: {0}";
        public static readonly String INPUT_SCHEDULE_NAMES_VALUE = "Schedule names: {0}";
        public static readonly String INPUT_SCHEDULE_IDS_VALUE = "Schedule ids: {0}";
        public static readonly String INPUT_DELAY_VALUE = "Delay between status checks: {0}";
        public static readonly String INPUT_DONE_VALUE = "Done Status As: {0}";
        public static readonly String INPUT_LEAPWORK_CONTROLLER_URL = "LEAPWORK Controller URL: {0}";

        public static readonly String SCHEDULE_TITLE = "Schedule: {0}";
        public static readonly String CASES_PASSED = "Passed testcases: {0}";
        public static readonly String CASES_FAILED = "Failed testcases: {0}";
        public static readonly String CASES_ERRORED = "Error testcases: {0}";

        public static readonly String TOTAL_SEPARATOR = "|---------------------------------------------------------------";
        public static readonly String TOTAL_CASES_PASSED = "| Total passed testcases: {0}";
        public static readonly String TOTAL_CASES_FAILED = "| Total failed testcases: {0}";
        public static readonly String TOTAL_CASES_ERROR = "| Total error testcases: {0}";

        public static readonly String FAILED_TO_PARSE_RESPONSE_KEYFRAME_JSON_ARRAY = "Failed to parse response keyframe json array";
        public static readonly String ERROR_NOTIFICATION = "There were detected case(s) with status 'Failed', 'Error', 'Inconclusive', 'Timeout' or 'Cancelled'. Please check the report or console output for details. Set the build status to FAILURE as the results of the cases are not deterministic.";
    }

    public class Program
    {

        //Old versions of MSBuild do not support C# 6 features: Null Condition operators ?. http://bartwullems.blogspot.com.by/2016/03/tfs-build-error-invalid-expression-term.html 
        private static Guid DefaultTokenGuidValueIfNull(String tokenName, JToken parentToken, SimpleLogger logger)
        {
            JToken token = parentToken.SelectToken(tokenName);
            if (token != null)
            {
                string strGuid = token.Value<string>();
                Guid resultGuid;
                if (Guid.TryParse(strGuid, out resultGuid))
                    return resultGuid;
                else
                {
                    logger.Warning(string.Format(Messages.FAILED_TO_PARSE_STRING_TO_GUID, tokenName, strGuid));
                    return Guid.Empty;
                }
            }
            else
            {
                logger.Warning(string.Format(Messages.FAILED_TO_PARSE_STRING_TO_GUID, tokenName, ""));
                return Guid.Empty;

            }

        }

        private static string DefaultTokenStringValueIfNull(String tokenName, JToken parentToken, SimpleLogger logger, string defaultValue = "")
        {
            JToken token = parentToken.SelectToken(tokenName);
            if (token != null)
            {
                string str = token.Value<string>();
                return str;
            }
            else
            {
                return defaultValue;
            }

        }

        private static double DefaultTokenDoubleValueIfNull(String tokenName, JToken parentToken, SimpleLogger logger, double defaultValue)
        {

            JToken token = parentToken.SelectToken(tokenName);
            if (token != null)
            {
                try
                {
                    string strDouble = token.Value<string>().Replace(',', '.');
                    return double.Parse(strDouble, CultureInfo.InvariantCulture);
                }
                catch (Exception)
                {
                    return defaultValue;
                }
            }
            else
                return defaultValue;
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

            if (!string.IsNullOrEmpty(rawScheduleIds))
            {
                string[] schidsArray = regex.Split(rawScheduleIds);
                for (int i = 0; i < schidsArray.Length; i++)
                {
                    if (!string.IsNullOrEmpty(schidsArray[i]))
                        rawScheduleList.Add(schidsArray[i]);
                }
            }

            if (!string.IsNullOrEmpty(rawScheduleTitles))
            {
                string[] testsArray = regex.Split(rawScheduleTitles);

                for (int i = 0; i < testsArray.Length; i++)
                {
                    if (!string.IsNullOrEmpty(testsArray[i]))
                        rawScheduleList.Add(testsArray[i]);
                }
            }
            else
            {
                throw new Exception(Messages.NO_SCHEDULES);
            }

            return rawScheduleList;
        }

        private static int GetPortNumber(string rawPortStr, SimpleLogger logger)
        {
            int defaultPortNumber = 9001;

            int portNumber;
            if (int.TryParse(rawPortStr, out portNumber))
                return portNumber;
            else
            {
                logger.Warning(String.Format(Messages.PORT_NUMBER_IS_INVALID, defaultPortNumber));
                return defaultPortNumber;
            }

        }

        private static string GetControllerApiHttpAddress(string hostname, string rawPort, SimpleLogger logger)
        {
            var stringBuilder = new StringBuilder();
            int port = GetPortNumber(rawPort, logger);
            stringBuilder.Append("http://").Append(hostname).Append(":").Append(port);
            return stringBuilder.ToString();
        }

        private static int GetTimeDelay(String rawTimeDelay, SimpleLogger logger)
        {
            int defaultTimeDelay = 5;
            int timeDelay;
            if (Int32.TryParse(rawTimeDelay, out timeDelay))
                return timeDelay;
            else
            {
                logger.Warning(String.Format(Messages.TIME_DELAY_NUMBER_IS_INVALID, defaultTimeDelay));
                return defaultTimeDelay;
            }


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

        private static bool DefaultTokenBooleanIfNull(string tokenTitle, JObject jObject, bool defaultValue, SimpleLogger logger)
        {
            if (jObject != null)
            {
                try
                {
                    JToken boolToken = jObject.SelectToken(tokenTitle);
                    return boolToken.Value<bool>();
                }
                catch (Exception)
                {
                    //logger.Warning(string.Format(Messages.INVALID_BOOLEAN_TOKEN, tokenTitle, defaultValue));
                    return defaultValue;
                }
            }
            else
            {
                //logger.Warning(string.Format(Messages.INVALID_BOOLEAN_TOKEN, tokenTitle, defaultValue));
                return defaultValue;
            }

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


        private static async Task<Dictionary<Guid, string>> GetSchedulesIdTitleDictionary(
                HttpClient client,
                string leapworkControllerApiHttpAddress,
                List<string> rawScheduleList,
                List<InvalidSchedule> invalidSchedules,
                SimpleLogger logger
        )
        {
            Dictionary<Guid, string> schedulesIdTitleDictionary = new Dictionary<Guid, string>();

            String scheduleListUri = String.Format(Messages.GET_ALL_AVAILABLE_SCHEDULES_URI, leapworkControllerApiHttpAddress);

            try
            {
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
                                        Guid Id = DefaultTokenGuidValueIfNull("Id", jsonSchedule, logger);
                                        string Title = DefaultTokenStringValueIfNull("Title", jsonSchedule, logger);
                                        bool isEnabled = DefaultTokenBooleanIfNull("IsEnabled", jsonSchedule, false, logger);

                                        if (Id.ToString().Equals(rawSchedule)) //Id match
                                        {
                                            if (!schedulesIdTitleDictionary.ContainsValue(Title)) //Avoid repeat
                                            {
                                                if (isEnabled)
                                                {
                                                    schedulesIdTitleDictionary.Add(Id, Title);
                                                    logger.Info(String.Format(Messages.SCHEDULE_DETECTED, Title, rawSchedule));
                                                }
                                                else
                                                {
                                                    invalidSchedules.Add(new InvalidSchedule(rawSchedule, String.Format(Messages.SCHEDULE_DISABLED, Title, Id)));
                                                }
                                            }
                                            isSuccessfullyMapped = true;
                                        }

                                        if (Title.Equals(rawSchedule)) //Title match 
                                        {
                                            if (!schedulesIdTitleDictionary.ContainsKey(Id)) //Avoid repeat
                                            {
                                                if (isEnabled)
                                                {
                                                    schedulesIdTitleDictionary.Add(Id, rawSchedule);
                                                    logger.Info(String.Format(Messages.SCHEDULE_DETECTED, rawSchedule, Id));
                                                }
                                                else
                                                {
                                                    invalidSchedules.Add(new InvalidSchedule(rawSchedule, string.Format(Messages.SCHEDULE_DISABLED, Title, Id)));
                                                }
                                            }
                                            isSuccessfullyMapped = true;
                                        }
                                    }

                                    if (!isSuccessfullyMapped)
                                        invalidSchedules.Add(new InvalidSchedule(rawSchedule, Messages.NO_SUCH_SCHEDULE));
                                }
                            }

                            break;

                        case 401://INVALID ACCESS KEY
                            StringBuilder errorMessage401 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage401.AppendLine(Messages.INVALID_ACCESS_KEY);
                            errorMessage401.AppendLine(Messages.SCHEDULE_TITLE_OR_ID_ARE_NOT_GOT);
                            throw new Exception(errorMessage401.ToString());

                        case 500://CONTROLLER ERROR
                            StringBuilder errorMessage500 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage500.AppendLine(Messages.CONTROLLER_RESPONDED_WITH_ERRORS);
                            errorMessage500.AppendLine(Messages.SCHEDULE_TITLE_OR_ID_ARE_NOT_GOT);
                            throw new Exception(errorMessage500.ToString());

                        default:
                            StringBuilder errorMessage = new StringBuilder(String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage.AppendLine(Messages.SCHEDULE_TITLE_OR_ID_ARE_NOT_GOT);
                            throw new Exception(errorMessage.ToString());
                    }
                }
            }
            catch (HttpRequestException e)
            {
                StringBuilder connectionErrorMessage = new StringBuilder(String.Format(Messages.COULD_NOT_CONNECT_TO, scheduleListUri));
                connectionErrorMessage.AppendLine(Messages.SCHEDULE_TITLE_OR_ID_ARE_NOT_GOT);
                connectionErrorMessage.AppendLine(e.Message);
                connectionErrorMessage.AppendLine(e.StackTrace);
                throw new Exception(connectionErrorMessage.ToString());
            }

            return schedulesIdTitleDictionary;
        }


        private static async Task<Guid> RunSchedule(
            HttpClient client,
            String leapworkControllerApiHttpAddress,
            Guid scheduleId,
            String scheduleTitle,
            LeapworkRun leapworkRun,
            SimpleLogger logger

        )
        {

            String uri = String.Format(Messages.RUN_SCHEDULE_URI, leapworkControllerApiHttpAddress, scheduleId);

            try
            {
                using (HttpResponseMessage response = await client.PutAsync(uri, new StringContent(String.Empty))) //Send PUT request and launch schedule
                {
                    int statusCode = (int)response.StatusCode;
                    string status = response.StatusCode.ToString();

                    switch (statusCode)
                    {
                        case 200://SUCCESS

                            using (HttpContent content = response.Content)
                            {
                                string responseContent = await content.ReadAsStringAsync();

                                JObject jRunId = JObject.Parse(responseContent);

                                Guid runId = DefaultTokenGuidValueIfNull("RunId", jRunId, logger);

                                string successMessage = String.Format(Messages.SCHEDULE_RUN_SUCCESS, scheduleTitle, scheduleId);
                                logger.Info(Messages.SCHEDULE_CONSOLE_LOG_SEPARATOR);
                                logger.Info(successMessage);

                                return runId;
                            }

                        case 400:
                            StringBuilder errorMessage400 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage400.AppendLine(Messages.INVALID_VARIABLE_KEY_NAME);
                            return OnScheduleRunFailure(errorMessage400, leapworkRun, scheduleId, logger);

                        case 401:
                            StringBuilder errorMessage401 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage401.AppendLine(string.Format(Messages.INVALID_ACCESS_KEY));
                            return OnScheduleRunFailure(errorMessage401, leapworkRun, scheduleId, logger);

                        case 404:
                            StringBuilder errorMessage404 = new StringBuilder(String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage404.AppendLine(string.Format(Messages.NO_SUCH_SCHEDULE_WAS_FOUND, scheduleTitle, scheduleId));
                            return OnScheduleRunFailure(errorMessage404, leapworkRun, scheduleId, logger);

                        case 446:
                            StringBuilder errorMessage446 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage446.AppendLine(Messages.NO_DISK_SPACE);
                            return OnScheduleRunFailure(errorMessage446, leapworkRun, scheduleId, logger);

                        case 455:
                            StringBuilder errorMessage455 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage455.AppendLine(Messages.DATABASE_NOT_RESPONDING);
                            return OnScheduleRunFailure(errorMessage455, leapworkRun, scheduleId, logger);

                        case 500:
                            StringBuilder errorMessage500 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage500.AppendLine(Messages.CONTROLLER_RESPONDED_WITH_ERRORS);
                            return OnScheduleRunFailure(errorMessage500, leapworkRun, scheduleId, logger);

                        default:
                            StringBuilder errorMessage = new StringBuilder(String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            return OnScheduleRunFailure(errorMessage, leapworkRun, scheduleId, logger);
                    }
                }
            }
            catch (HttpRequestException e)
            {
                String connectionErrorMessage = String.Format(Messages.COULD_NOT_CONNECT_TO_BUT_WAIT, uri);
                logger.Error(connectionErrorMessage);
                PrintException(logger,e);
                return Guid.Empty; //In case of problems with connection to controller the plugin will be waiting for connection reestablishment
            }
        }

        private static Guid OnScheduleRunFailure(StringBuilder errorMessage, LeapworkRun failedRun, Guid scheduleId, SimpleLogger logger)
        {
            errorMessage.AppendLine(String.Format(Messages.SCHEDULE_RUN_FAILURE, failedRun.ScheduleTitle, scheduleId));
            logger.Error(errorMessage.ToString());
            failedRun.Error = errorMessage.ToString();
            failedRun.IncErrors();
            return Guid.Empty;
        }

        private static async Task<List<Guid>> GetRunRunItems(HttpClient client, string controllerApiHttpAddress, Guid runId)
        {
            string uri = string.Format(Messages.GET_RUN_ITEMS_IDS_URI, controllerApiHttpAddress, runId.ToString());

            using (HttpResponseMessage response = await client.GetAsync(uri)) //Send PUT request and launch schedule
            {
                int statusCode = (int)response.StatusCode;
                string status = response.StatusCode.ToString();

                switch (statusCode)
                {
                    case 200:

                        using (HttpContent content = response.Content)
                        {
                            string responseContent = await content.ReadAsStringAsync();

                            JObject jsonRunItemIdListObject = JObject.Parse(responseContent);

                            JArray jsonRunItemIdList = jsonRunItemIdListObject.SelectToken("RunItemIds") as JArray;

                            List<Guid> runItems = new List<Guid>();

                            if (jsonRunItemIdList != null)
                            {
                                foreach (JToken jsonRunItemId in jsonRunItemIdList)
                                {
                                    string strGuid = jsonRunItemId.Value<string>();
                                    Guid guid;
                                    if (Guid.TryParse(strGuid, out guid))
                                    {
                                        runItems.Add(guid);
                                    }
                                }
                            }
                            return runItems;
                        }


                    case 401:
                        StringBuilder errorMessage401 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage401.AppendLine(Messages.INVALID_ACCESS_KEY);
                        throw new Exception(errorMessage401.ToString());

                    case 404:
                        StringBuilder errorMessage404 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage404.AppendLine(Messages.NO_SUCH_RUN);
                        throw new Exception(errorMessage404.ToString());

                    case 446:
                        StringBuilder errorMessage446 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage446.AppendLine(Messages.NO_DISK_SPACE);
                        throw new Exception(errorMessage446.ToString());

                    case 455:
                        StringBuilder errorMessage455 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage455.AppendLine(Messages.DATABASE_NOT_RESPONDING);
                        throw new Exception(errorMessage455.ToString());

                    case 500:
                        StringBuilder errorMessage500 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage500.AppendLine(Messages.CONTROLLER_RESPONDED_WITH_ERRORS);
                        throw new Exception(errorMessage500.ToString());

                    default:
                        String errorMessage = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                        throw new Exception(errorMessage);
                }
            }

        }

        private static async Task<RunItem> GetRunItem(HttpClient client, string controllerApiHttpAddress, Guid runItemId, string scheduleName, bool doneStatusAsSuccess, SimpleLogger logger)
        {
            String uri = string.Format(Messages.GET_RUN_ITEM_URI, controllerApiHttpAddress, runItemId);

            using (HttpResponseMessage response = await client.GetAsync(uri))
            {
                int statusCode = (int)response.StatusCode;
                string status = response.StatusCode.ToString();

                switch (statusCode)
                {
                    case 200:

                        using (HttpContent content = response.Content)
                        {
                            string responseContent = await content.ReadAsStringAsync();

                            JObject jsonRunItem = JObject.Parse(responseContent);

                            //FlowInfo
                            JToken jsonFlowInfo = jsonRunItem.SelectToken("FlowInfo");

                            Guid flowId = DefaultTokenGuidValueIfNull("FlowId", jsonFlowInfo, logger);

                            String flowTitle = DefaultTokenStringValueIfNull("FlowTitle", jsonFlowInfo, logger);


                            String flowStatus =
                                DefaultTokenStringValueIfNull("Status", jsonFlowInfo, logger, "NoStatus");

                            //EnvironmentInfo
                            JToken jsonEnvironmentInfo = jsonRunItem.SelectToken("EnvironmentInfo");

                            Guid environmentId =
                                DefaultTokenGuidValueIfNull("EnvironmentId", jsonEnvironmentInfo, logger);
                            String environmentTitle =
                                DefaultTokenStringValueIfNull("EnvironmentTitle", jsonEnvironmentInfo, logger);
                            String environmentConnectionType = DefaultTokenStringValueIfNull("ConnectionType",
                                jsonEnvironmentInfo, logger, "Not defined");
                            Guid runId = DefaultTokenGuidValueIfNull("RunId", jsonRunItem, logger);

                            String elapsed = DefaultElapsedIfNull(jsonRunItem.SelectToken("Elapsed"));
                            double seconds = DefaultTokenDoubleValueIfNull("ElapsedSeconds", jsonRunItem, logger, 0);

                            RunItem runItem = new RunItem(flowTitle, flowStatus, seconds, scheduleName);

                            if (!flowStatus.Equals("Initializing") &&
                                !flowStatus.Equals("Connecting") &&
                                !flowStatus.Equals("Connected") &&
                                !flowStatus.Equals("Running") &&
                                !flowStatus.Equals("NoStatus") &&
                                !flowStatus.Equals("Passed") &&
                                !(flowStatus.Equals("Done") && doneStatusAsSuccess))
                            {

                                Failure keyFrames = GetRunItemKeyframes(client, controllerApiHttpAddress, runItemId,
                                    runItem, scheduleName, environmentTitle, logger).Result;

                                runItem.failure = keyFrames;
                            }

                            return runItem;
                        }

                    case 401:
                        StringBuilder errorMessage401 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage401.AppendLine(Messages.INVALID_ACCESS_KEY);
                        throw new Exception(errorMessage401.ToString());

                    case 404:
                        StringBuilder errorMessage404 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage404.AppendLine(string.Format(Messages.NO_SUCH_RUN_ITEM_WAS_FOUND, runItemId, scheduleName));
                        throw new Exception(errorMessage404.ToString());

                    case 446:
                        StringBuilder errorMessage446 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage446.AppendLine(Messages.NO_DISK_SPACE);
                        throw new Exception(errorMessage446.ToString());

                    case 455:
                        StringBuilder errorMessage455 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage455.AppendLine(Messages.DATABASE_NOT_RESPONDING);
                        throw new Exception(errorMessage455.ToString());

                    case 500:
                        StringBuilder errorMessage500 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage500.AppendLine(Messages.CONTROLLER_RESPONDED_WITH_ERRORS);
                        throw new Exception(errorMessage500.ToString());

                    default:
                        String errorMessage = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                        throw new Exception(errorMessage);
                }
            }

        }

        private static async Task<Failure> GetRunItemKeyframes(HttpClient client, string controllerApiHttpAddress, Guid runItemId,
            RunItem runItem, string scheduleName, string environmentTitle, SimpleLogger logger)
        {
            String uri = string.Format(Messages.GET_RUN_ITEM_KEYFRAMES, controllerApiHttpAddress, runItemId);

            using (HttpResponseMessage response = await client.GetAsync(uri))
            {
                int statusCode = (int)response.StatusCode;
                string status = response.StatusCode.ToString();

                switch (statusCode)
                {
                    case 200:

                        using (HttpContent content = response.Content)
                        {
                            string responseContent = await content.ReadAsStringAsync();

                            JArray jsonKeyframes = JArray.Parse(responseContent);

                            if (jsonKeyframes != null)
                            {
                                logger.Info(Messages.CASE_CONSOLE_LOG_SEPARATOR);
                                logger.Info(string.Format(Messages.CASE_INFORMATION, runItem.FlowTitle, runItem.FlowStatus, runItem.ElapsedTime));

                                StringBuilder fullKeyframes = new StringBuilder("");

                                foreach (var jsonKeyFrame in jsonKeyframes)
                                {

                                    string level = DefaultTokenStringValueIfNull("Level", jsonKeyFrame, logger, "Trace");
                                    if (!string.IsNullOrEmpty(level) && !level.Contains("Trace"))
                                    {
                                        JToken token = jsonKeyFrame.SelectToken("Timestamp").SelectToken("Value");
                                        var timeStampValue = (DateTime)token.ToObject(typeof(DateTime));
                                        string timestamp = timeStampValue.ToString("dd-MM-yyyy hh:mm:ss.fff");

                                        string logMessage = DefaultTokenStringValueIfNull("LogMessage", jsonKeyFrame, logger);

                                        string keyFrame = string.Format(Messages.CASE_STACKTRACE_FORMAT, timestamp, logMessage);

                                        logger.Info(keyFrame);

                                        fullKeyframes.AppendLine(keyFrame);

                                    }

                                }

                                fullKeyframes.AppendLine("Environment: " + environmentTitle);
                                logger.Info("Environment: " + environmentTitle);
                                fullKeyframes.AppendLine("Schedule: " + scheduleName);
                                logger.Info("Schedule: " + scheduleName);

                                return new Failure(fullKeyframes.ToString());
                            }
                            else
                            {
                                logger.Error(Messages.FAILED_TO_PARSE_RESPONSE_KEYFRAME_JSON_ARRAY);
                                return new Failure(Messages.FAILED_TO_PARSE_RESPONSE_KEYFRAME_JSON_ARRAY);
                            }
                        }

                    case 401:
                        StringBuilder errorMessage401 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage401.AppendLine(Messages.INVALID_ACCESS_KEY);
                        logger.Error(errorMessage401.ToString());
                        break;

                    case 404:
                        StringBuilder errorMessage404 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage404.AppendLine(string.Format(Messages.NO_SUCH_RUN_ITEM_WAS_FOUND, runItemId, scheduleName));
                        logger.Error(errorMessage404.ToString());
                        break;

                    case 446:
                        StringBuilder errorMessage446 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage446.AppendLine(Messages.NO_DISK_SPACE);
                        logger.Error(errorMessage446.ToString());
                        break;

                    case 455:
                        StringBuilder errorMessage455 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage455.AppendLine(Messages.DATABASE_NOT_RESPONDING);
                        logger.Error(errorMessage455.ToString());
                        break;

                    case 500:
                        StringBuilder errorMessage500 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage500.AppendLine(Messages.CONTROLLER_RESPONDED_WITH_ERRORS);
                        logger.Error(errorMessage500.ToString());
                        break;

                    default:
                        String errorMessage = String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                        logger.Error(errorMessage);
                        break;
                }
            }

            return null;
        }

        private static async Task<string> GetRunStatus(HttpClient client, string controllerApiHttpAddress, Guid runId, SimpleLogger logger)
        {
            String uri = string.Format(Messages.GET_RUN_STATUS_URI, controllerApiHttpAddress, runId);

            using (HttpResponseMessage response = await client.GetAsync(uri))
            {
                int statusCode = (int)response.StatusCode;
                string status = response.StatusCode.ToString();

                switch (statusCode)
                {
                    case 200:

                        using (HttpContent content = response.Content)
                        {
                            string responseContent = await content.ReadAsStringAsync();

                            JObject jsonRunStatus = JObject.Parse(responseContent);

                            string runStatus = DefaultTokenStringValueIfNull("Status", jsonRunStatus, logger, "Queued");

                            return runStatus;
                        }

                    case 401:
                        StringBuilder errorMessage401 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage401.AppendLine(Messages.INVALID_ACCESS_KEY);
                        throw new Exception(errorMessage401.ToString());

                    case 404:
                        StringBuilder errorMessage404 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage404.AppendLine(Messages.NO_SUCH_RUN);
                        throw new Exception(errorMessage404.ToString());

                    case 446:
                        StringBuilder errorMessage446 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage446.AppendLine(Messages.NO_DISK_SPACE);
                        throw new Exception(errorMessage446.ToString());

                    case 455:
                        StringBuilder errorMessage455 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage455.AppendLine(Messages.DATABASE_NOT_RESPONDING);
                        throw new Exception(errorMessage455.ToString());

                    case 500:
                        StringBuilder errorMessage500 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage500.AppendLine(Messages.CONTROLLER_RESPONDED_WITH_ERRORS);
                        throw new Exception(errorMessage500.ToString());
                    default:
                        String errorMessage = string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status);
                        throw new Exception(errorMessage);
                }
            }
        }

        private static async Task<bool> StopRun(HttpClient client, string controllerApiHttpAddress, Guid runId, string scheduleName, SimpleLogger logger)
        {
            bool isSuccessfullyStopped = false;

            logger.Error(String.Format(Messages.STOPPING_RUN, scheduleName, runId));
            String uri = String.Format(Messages.STOP_RUN_URI, controllerApiHttpAddress, runId);

            using (HttpResponseMessage response = await client.GetAsync(uri))
            {
                int statusCode = (int)response.StatusCode;
                string status = response.StatusCode.ToString();

                switch (statusCode)
                {
                    case 200:

                        using (HttpContent content = response.Content)
                        {
                            string responseContent = await content.ReadAsStringAsync();

                            JObject jsonOperationResult = JObject.Parse(responseContent);

                            isSuccessfullyStopped = DefaultTokenBooleanIfNull("OperationCompleted", jsonOperationResult, isSuccessfullyStopped, logger);
                        }
                        break;
                    case 401:
                        StringBuilder errorMessage401 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage401.AppendLine(Messages.INVALID_ACCESS_KEY);
                        errorMessage401.AppendLine(String.Format(Messages.STOP_RUN_FAIL, scheduleName, runId));
                        logger.Error(errorMessage401.ToString());
                        break;

                    case 404:
                        StringBuilder errorMessage404 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage404.AppendLine(string.Format(Messages.NO_SUCH_RUN_WAS_FOUND, runId, scheduleName));
                        errorMessage404.AppendLine(String.Format(Messages.STOP_RUN_FAIL, scheduleName, runId));
                        logger.Error(errorMessage404.ToString());
                        break;

                    case 446:
                        StringBuilder errorMessage446 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage446.AppendLine(Messages.NO_DISK_SPACE);
                        errorMessage446.AppendLine(String.Format(Messages.STOP_RUN_FAIL, scheduleName, runId));
                        logger.Error(errorMessage446.ToString());
                        break;

                    case 455:
                        StringBuilder errorMessage455 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage455.AppendLine(Messages.DATABASE_NOT_RESPONDING);
                        errorMessage455.AppendLine(String.Format(Messages.STOP_RUN_FAIL, scheduleName, runId));
                        logger.Error(errorMessage455.ToString());
                        break;

                    case 500:
                        StringBuilder errorMessage500 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage500.AppendLine(Messages.CONTROLLER_RESPONDED_WITH_ERRORS);
                        errorMessage500.AppendLine(String.Format(Messages.STOP_RUN_FAIL, scheduleName, runId));
                        logger.Error(errorMessage500.ToString());
                        break;

                    default:
                        StringBuilder errorMessage = new StringBuilder(String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                        errorMessage.AppendLine(String.Format(Messages.STOP_RUN_FAIL, scheduleName, runId));
                        logger.Error(errorMessage.ToString());
                        break;
                }
            }

            return isSuccessfullyStopped;
        }

        private static void CreateJunitReport(string reportPath, RunCollection buildResult, SimpleLogger logger)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(RunCollection));

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


        public static string Call(string leapworkHostname, string leapworkPort, string leapworkAccessKey, string leapworkTime, string leapworkDoneStatus, string leapworkReport, string leapworkLog, string leapworkIds, string leapworkTitles)
        {

            SimpleLogger logger = new SimpleLogger(leapworkLog);

            logger.Info(string.Format(Messages.INPUT_VALUES_MESSAGE));
            logger.Info(string.Format(Messages.CASE_CONSOLE_LOG_SEPARATOR));
            logger.Info(string.Format(Messages.INPUT_HOSTNAME_VALUE, leapworkHostname));
            logger.Info(string.Format(Messages.INPUT_PORT_VALUE, leapworkPort));
            //logger.Info(string.Format(Messages.INPUT_ACCESS_KEY_VALUE, leapworkAccessKey));
            logger.Info(string.Format(Messages.INPUT_REPORT_VALUE, leapworkReport));
            logger.Info(string.Format(Messages.INPUT_LOG_FILEPATH_VALUE, leapworkLog));
            logger.Info(string.Format(Messages.INPUT_SCHEDULE_NAMES_VALUE, leapworkTitles));
            logger.Info(string.Format(Messages.INPUT_SCHEDULE_IDS_VALUE, leapworkIds));
            logger.Info(string.Format(Messages.INPUT_DELAY_VALUE, leapworkTime));
            logger.Info(string.Format(Messages.INPUT_DONE_VALUE, leapworkDoneStatus));
            string controllerApiHttpAddress = GetControllerApiHttpAddress(leapworkHostname, leapworkPort, logger);
            logger.Info(String.Format(Messages.INPUT_LEAPWORK_CONTROLLER_URL, controllerApiHttpAddress));

            String junitReportPath = GetJunitReportFilePath(leapworkReport); //checks if .xml in the path exists
            logger.Info(String.Format("Full Report file path: {0}", junitReportPath));

            Dictionary<Guid, string> schedulesIdTitleDictionary; // Id-Title
            List<InvalidSchedule> invalidSchedules = new List<InvalidSchedule>();
            List<String> rawScheduleList = GetRawScheduleList(leapworkIds, leapworkTitles);

            int timeDelay = GetTimeDelay(leapworkTime, logger);
            bool isDoneStatusIsSuccess = leapworkDoneStatus.Equals("Success");

            Dictionary<Guid, LeapworkRun> resultsMap = new Dictionary<Guid, LeapworkRun>();

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("AccessKey", leapworkAccessKey);

                    schedulesIdTitleDictionary = GetSchedulesIdTitleDictionary(client, controllerApiHttpAddress, rawScheduleList, invalidSchedules, logger).Result;
                    rawScheduleList.Clear();
                    rawScheduleList = null;//don't need that anymore

                    List<Guid> schIdsList = new List<Guid>(schedulesIdTitleDictionary.Keys);
                    schIdsList.Reverse();

                    for (int i = schIdsList.Count - 1; i >= 0; i--)
                    {
                        Guid schId = schIdsList[i];
                        string schTitle = schedulesIdTitleDictionary[schId];

                        LeapworkRun leapworkRun = new LeapworkRun(schTitle);

                        Guid runId = RunSchedule(client, controllerApiHttpAddress, schId, schTitle, leapworkRun, logger).Result;

                        if (runId != Guid.Empty)
                        {
                            resultsMap.Add(runId, leapworkRun);
                            leapworkRun.RunId = runId.ToString();
                            CollectScheduleRunResults(client, controllerApiHttpAddress, runId, schTitle, timeDelay, isDoneStatusIsSuccess, leapworkRun, logger);
                        }
                        else
                        {
                            resultsMap.Add(Guid.NewGuid(), leapworkRun);
                        }
                    }
                }


                RunCollection buildResult = new RunCollection();


                if (invalidSchedules.Count > 0)
                {
                    logger.Info(Messages.INVALID_SCHEDULES);

                    foreach (InvalidSchedule invalidSchedule in invalidSchedules)
                    {
                        logger.Warning(string.Format("{0}: {1}", invalidSchedule.Name, invalidSchedule.StackTrace));
                        LeapworkRun notFoundLeapworkRun = new LeapworkRun(invalidSchedule.Name);
                        RunItem invalidRunItem = new RunItem("Error", "Error", 0, invalidSchedule.StackTrace, invalidSchedule.Name);
                        notFoundLeapworkRun.RunItems.Add(invalidRunItem);
                        notFoundLeapworkRun.Error = invalidSchedule.StackTrace;
                        buildResult.Runs.Add(notFoundLeapworkRun);
                    }
                }

                List<LeapworkRun> resultRuns = new List<LeapworkRun>(resultsMap.Values);

                logger.Info(Messages.TOTAL_SEPARATOR);
                foreach (LeapworkRun leapworkRun in resultRuns)
                {
                    buildResult.Runs.Add(leapworkRun);
                    buildResult.AddFailedTests(leapworkRun.Failed);
                    buildResult.AddPassedTests(leapworkRun.Passed);
                    buildResult.AddErrors(leapworkRun.Errors);
                    leapworkRun.Total = leapworkRun.Passed + leapworkRun.Failed;
                    buildResult.AddTotalTime(leapworkRun.Time);
                    logger.Info(string.Format(Messages.SCHEDULE_TITLE, leapworkRun.ScheduleTitle));
                    logger.Info(string.Format(Messages.CASES_PASSED, leapworkRun.Passed));
                    logger.Info(string.Format(Messages.CASES_FAILED, leapworkRun.Failed));
                    logger.Info(string.Format(Messages.CASES_ERRORED, leapworkRun.Errors));
                }
                buildResult.TotalTests = buildResult.FailedTests + buildResult.PassedTests;

                logger.Info(Messages.TOTAL_SEPARATOR);
                logger.Info(string.Format(Messages.TOTAL_CASES_PASSED, buildResult.PassedTests));
                logger.Info(string.Format(Messages.TOTAL_CASES_FAILED, buildResult.FailedTests));
                logger.Info(string.Format(Messages.TOTAL_CASES_ERROR, buildResult.Errors));

                CreateJunitReport(junitReportPath, buildResult, logger);

                logger.Info(Messages.PLUGIN_SUCCESSFUL_FINISH);

                if (buildResult.Errors > 0 || buildResult.FailedTests > 0 || invalidSchedules.Count > 0)
                {
                    logger.Warning(Messages.ERROR_NOTIFICATION);
                    logger.Warning(Messages.BUILD_SUCCEEDED_WITH_ISSUES);
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
                if (e is ThreadAbortException || e is ThreadInterruptedException)
                {
                    logger.Error(Messages.BUILD_ABORTED);
                    return Messages.BUILD_ABORTED;
                }
                else
                {
                    logger.Error(Messages.PLUGIN_ERROR_FINISH);
                    PrintException(logger, e);

                    var aggregateException = e as AggregateException; //C# 6 feature are not supported
                    if (aggregateException != null && aggregateException.InnerExceptions != null)
                    {
                        var hasInnerException = aggregateException.InnerException != null;
                        foreach (var innerException in aggregateException.InnerExceptions)
                        {
                            if (hasInnerException == false || hasInnerException && innerException != aggregateException.InnerException)
                            {
                                PrintException(logger, innerException);
                            }
                        }
                    }

                    logger.Error(Messages.PLEASE_CONTACT_SUPPORT);
                    logger.Error(Messages.BUILD_FAILED);
                    return Messages.BUILD_FAILED;
                }
            }

        }

        private static void PrintException(SimpleLogger logger, Exception e)
        {
            logger.Error(e.Message);
            logger.Error(e.StackTrace);
            if (e.InnerException != null)
            {
                PrintException(logger, e.InnerException);
            }
        }

        private static void CollectScheduleRunResults(HttpClient client, string controllerApiHttpAddress, Guid runId, string scheduleName, int timeDelay, bool isDoneStatusAsSuccess, LeapworkRun resultRun, SimpleLogger logger)
        {
            List<Guid> runItemsId = new List<Guid>();

            try
            {
                bool isStillRunning = true;

                do
                {

                    Task.Delay(timeDelay * 1000).Wait();

                    List<Guid> executedRunItems = GetRunRunItems(client, controllerApiHttpAddress, runId).Result;

                    foreach (Guid guid in runItemsId)
                    {
                        executedRunItems.Remove(guid); //left only new
                    }

                    executedRunItems.Reverse();

                    for (int i = executedRunItems.Count - 1; i >= 0; i--)
                    {
                        Guid runItemId = executedRunItems[i];
                        RunItem runItem = GetRunItem(client, controllerApiHttpAddress, runItemId, scheduleName, isDoneStatusAsSuccess, logger).Result;

                        String status = runItem.FlowStatus;


                        resultRun.Time += runItem.ElapsedTime;
                        switch (status)
                        {
                            case "NoStatus":
                            case "Initializing":
                            case "Connecting":
                            case "Connected":
                            case "Running":
                                executedRunItems.RemoveAt(i);
                                break;
                            case "Passed":
                                resultRun.IncPassed();
                                resultRun.RunItems.Add(runItem);
                                break;
                            case "Failed":
                                resultRun.IncFailed();
                                resultRun.RunItems.Add(runItem);
                                break;
                            case "Error":
                            case "Inconclusive":
                            case "Timeout":
                            case "Cancelled":
                                resultRun.IncErrors();
                                resultRun.RunItems.Add(runItem);
                                break;
                            case "Done":
                                resultRun.RunItems.Add(runItem);
                                if (isDoneStatusAsSuccess)
                                    resultRun.IncPassed();
                                else
                                    resultRun.IncFailed();
                                break;

                        }

                    }

                    runItemsId.AddRange(executedRunItems);

                    String runStatus = GetRunStatus(client, controllerApiHttpAddress, runId, logger).Result;
                    if (runStatus.Equals("Finished"))
                    {
                        List<Guid> allExecutedRunItems = GetRunRunItems(client, controllerApiHttpAddress, runId).Result;
                        if (allExecutedRunItems.Count > 0 && allExecutedRunItems.Count <= runItemsId.Count)//todo ==
                            isStillRunning = false;
                    }

                    if (isStillRunning)
                        logger.Info(string.Format("The schedule status is already '{0}' - wait a minute...", runStatus));

                }
                while (isStillRunning);

            }
            catch (Exception e)
            {

                if (e is ThreadAbortException || e is ThreadInterruptedException)//In case of interruption the plugin tries to send HTTP Schedule STOP request
                {
                    String interruptedExceptionMessage = String.Format(Messages.INTERRUPTED_EXCEPTION, e.Message);
                    logger.Error(interruptedExceptionMessage);
                    bool isSuccessfullyStopped = StopRun(client, controllerApiHttpAddress, runId, scheduleName, logger).Result;
                    logger.Error(Messages.BUILD_ABORTED);
                    RunItem invalidItem = new RunItem("Aborted run", "Cancelled", 0, interruptedExceptionMessage, scheduleName);
                    resultRun.IncErrors();
                    resultRun.RunItems.Add(invalidItem);
                    throw;
                }
                else
                {
                    PrintException(logger, e);
                    RunItem invalidItem = new RunItem("Invalid run", "Error", 0, e.StackTrace, scheduleName);
                    resultRun.IncErrors();
                    resultRun.RunItems.Add(invalidItem);
                }

            }
        }
    }
}
"@

$logFile = Split-Path -Path $leapworkReport
$logFile = $logFile, "LeapworkIntegrationAgent$Env:AGENT_ID.Build$Env:BUILD_BUILDID.log" -Join "\"

Add-Type -ReferencedAssemblies $assemblies -TypeDefinition $sourceCode -Language CSharp
$buildResult = [TFSIntegrationConsole.Program]::Call($leapworkHostname,$leapworkPort,$leapworkAccessKey,$leapworkTimeDelay,$leapworkDoneStatusAs,$leapworkReport,$logFile,$leapworkSchids,$leapworkSchedules)

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

Write-Output "##vso[build.uploadlog]$logFile"



