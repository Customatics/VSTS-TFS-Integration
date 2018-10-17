param(
	[string]$leapworkHostname,
	[string]$leapworkPort,
	[string]$leapworkAccessKey,
    [string]$leapworkTimeDelay, 
    [string]$leapworkDoneStatusAs, 
    [string]$leapworkReport, 
    [string]$leapworkSchids,
    [string]$leapworkSchedules
)


function Extract-Nupkg($nupkg, $out)
{
    Add-Type -AssemblyName 'System.IO.Compression.FileSystem' # PowerShell lacks native support for zip
    
    $zipFile = [IO.Compression.ZipFile]
    $zipFile::ExtractToDirectory($nupkg, $out)
}

function Add-NewtonsoftJsonReference
{
    $url = 'https://www.nuget.org/api/v2/package/Newtonsoft.Json/11.0.2'

    $directory = $PSScriptRoot, 'bin', 'Newtonsoft.Json' -Join '\'
    $nupkg = Join-Path $directory "newtonsoft.json.11.0.2.nupkg"
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
    $url = 'https://www.nuget.org/api/v2/package/System.Net.Http/4.3.3'

    $directory = $PSScriptRoot, 'bin', 'System.Net.Http' -Join '\'
    $nupkg = Join-Path $directory "system.net.http.4.3.3.nupkg"
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

        [XmlAttribute(AttributeName = "time")]
        public double TotalTime { get; set; }


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

        public LeapworkRun(string runId, string title)
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

        [XmlAttribute(AttributeName = "time")]
        public double Time { get; set; }

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
    }

    public class Program
    {
        private static void MoveDirectory(string source, string target)
        {
            var sourcePath = source.TrimEnd('\\', ' ');
            var targetPath = target.TrimEnd('\\', ' ');

            var stack = new Stack<Folders>();
            stack.Push(new Folders(sourcePath, targetPath));

            while (stack.Count > 0)
            {
                var folders = stack.Pop();
                Directory.CreateDirectory(folders.Target);
                foreach (var file in Directory.GetFiles(folders.Source, "*.*"))
                {
                    File.Copy(file, Path.Combine(folders.Target, Path.GetFileName(file)));
                }

                foreach (var folder in Directory.GetDirectories(folders.Source))
                {
                    stack.Push(new Folders(folder, Path.Combine(folders.Target, Path.GetFileName(folder))));
                }
            }

            Directory.Delete(sourcePath, true);
        }

        public class Folders
        {
            public string Source { get; private set; }
            public string Target { get; private set; }

            public Folders(string source, string target)
            {
                Source = source;
                Target = target;
            }
        }

        

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

        private static string DefaultTokenStringValueIfNull(String tokenName, JToken parentToken, SimpleLogger logger,  string defaultValue = "")
        {
            JToken token = parentToken.SelectToken(tokenName);
            if (token != null)
            {
                string str = token.Value<string>();
                return str;
            }
            else
            {
                //logger.Warning(string.Format(Messages.STRING_TOKEN_NOT_FOUND, tokenName));
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
                    string strDouble = token.Value<string>().Replace(',','.');
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
            int defaultTimeDelay = 3;
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
                    logger.Warning(string.Format(Messages.INVALID_BOOLEAN_TOKEN, tokenTitle, defaultValue));
                    return defaultValue;
                }
            }
            else
            {
                logger.Warning(string.Format(Messages.INVALID_BOOLEAN_TOKEN, tokenTitle, defaultValue));
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
                                throw new Exception(errorMessage401.ToString());

                            case 500://CONTROLLER ERROR
                                StringBuilder errorMessage500 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                                errorMessage500.AppendLine(Messages.CONTROLLER_RESPONDED_WITH_ERRORS);
                                throw new Exception(errorMessage500.ToString());

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
                                throw new Exception(errorMessage400.ToString());

                            case 401:
                                StringBuilder errorMessage401 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                                errorMessage401.AppendLine(string.Format(Messages.INVALID_ACCESS_KEY));
                                throw new Exception(errorMessage401.ToString());

                            case 404:
                                StringBuilder errorMessage404 = new StringBuilder(String.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                                errorMessage404.AppendLine(string.Format(Messages.NO_SUCH_SCHEDULE_WAS_FOUND, scheduleTitle, scheduleId));
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
                catch (HttpRequestException e)
                {
                    String connectionErrorMessage = String.Format(Messages.COULD_NOT_CONNECT_TO_BUT_WAIT, e.Message);
                    logger.Error(connectionErrorMessage);
                    return Guid.Empty; //In case of problems with connection to controller the plugin will be waiting for connection reestablishment
                }
            }
            catch (Exception e)
            {
                String errorMessage = String.Format(Messages.SCHEDULE_RUN_FAILURE, scheduleTitle, scheduleId);
                logger.Error(errorMessage);
                logger.Error(e.Message);
                logger.Error(Messages.PLEASE_CONTACT_SUPPORT);
                leapworkRun.Error = string.Format("{0}\n{1}", errorMessage, e.StackTrace);
                leapworkRun.IncErrors();
                return Guid.Empty;
            }
        }

        private static async Task<List<Guid>> GetRunRunItems(HttpClient client, string controllerApiHttpAddress, Guid runId)
        {
            string uri = string.Format(Messages.GET_RUN_ITEMS_IDS_URI, controllerApiHttpAddress, runId.ToString());

            try
            {
                using (HttpResponseMessage response = await client.GetAsync(uri)) //Send PUT request and launch schedule
                {
                    int statusCode = (int) response.StatusCode;
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
            catch (Exception e)
            {
                throw e;
            }
        }

        private static async Task<RunItem> GetRunItem(HttpClient client, string controllerApiHttpAddress,  Guid runItemId, string scheduleName, SimpleLogger logger)
        {
            String uri = string.Format(Messages.GET_RUN_ITEM_URI, controllerApiHttpAddress, runItemId);

            try
            {
                using (HttpResponseMessage response = await client.GetAsync(uri))
                {
                    int statusCode = (int) response.StatusCode;
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

                                RunItem runItem;

                                if (!flowStatus.Equals("Initializing") &&
                                    !flowStatus.Equals("Connecting") &&
                                    !flowStatus.Equals("Connected") &&
                                    !flowStatus.Equals("Running") &&
                                    !flowStatus.Equals("NoStatus"))
                                {
                                    JArray jsonKeyframes = jsonRunItem.SelectToken("Keyframes") as JArray;
                                    StringBuilder fullKeyframes = new StringBuilder("");

                                    if (jsonKeyframes != null)
                                    {

                                        logger.Info(string.Format(Messages.CASE_INFORMATION, flowTitle, flowStatus, seconds));

                                        foreach (var jsonKeyFrame in jsonKeyframes)
                                        {

                                            string level =
                                                DefaultTokenStringValueIfNull("Level", jsonKeyFrame, logger, "Trace");
                                            if (!string.IsNullOrEmpty(level) && !level.Contains("Trace"))
                                            {
                                                JToken token = jsonKeyFrame.SelectToken("Timestamp").SelectToken("Value");
                                                var timeStampValue = (DateTime) token.ToObject(typeof(DateTime));
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

                                    }

                                    runItem = new RunItem(flowTitle, flowStatus, seconds, fullKeyframes.ToString(),
                                        scheduleName);

                                }
                                else
                                    runItem = new RunItem(flowTitle, flowStatus, seconds, scheduleName);

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
            catch (Exception e)
            {
                throw e;
            }
        }

        private static async Task<string> GetRunStatus(HttpClient client, string controllerApiHttpAddress, Guid runId, SimpleLogger logger)
        {
            String uri = string.Format(Messages.GET_RUN_STATUS_URI, controllerApiHttpAddress, runId);

            try
            {
                using (HttpResponseMessage response = await client.GetAsync(uri))
                {
                    int statusCode = (int) response.StatusCode;
                    string status = response.StatusCode.ToString();

                    switch (statusCode)
                    {
                        case 200:

                            using (HttpContent content = response.Content)
                            {
                                string responseContent = await content.ReadAsStringAsync();

                                JObject jsonRunStatus = JObject.Parse(responseContent);

                                string runStatus = DefaultTokenStringValueIfNull("Status", jsonRunStatus, logger,"Queued");

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
            catch (Exception e)
            {
                throw e;
            }

        }

        private static async Task<bool> StopRun(HttpClient client, string controllerApiHttpAddress, Guid runId, string scheduleName, SimpleLogger logger)
        {
            bool isSuccessfullyStopped = false;

            logger.Error(String.Format(Messages.STOPPING_RUN, scheduleName, runId));
            String uri = String.Format(Messages.STOP_RUN_URI, controllerApiHttpAddress, runId);
            try
            {
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

                                bool stopResult = DefaultTokenBooleanIfNull("OperationCompleted", jsonOperationResult, isSuccessfullyStopped, logger);

                                return stopResult;
                            }

                        case 401:
                            StringBuilder errorMessage401 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage401.AppendLine(Messages.INVALID_ACCESS_KEY);
                            throw new Exception(errorMessage401.ToString());

                        case 404:
                            StringBuilder errorMessage404 = new StringBuilder(string.Format(Messages.ERROR_CODE_MESSAGE, statusCode, status));
                            errorMessage404.AppendLine(string.Format(Messages.NO_SUCH_RUN_WAS_FOUND, runId, scheduleName));
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
            catch (Exception e)
            {
                logger.Error(String.Format(Messages.STOP_RUN_FAIL, scheduleName, runId));
                logger.Error(e.Message);
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

        public static string Call(string leapworkHostname,string leapworkPort, string leapworkAccessKey, string leapworkTime, string leapworkDoneStatus, string leapworkReport, string leapworkLog, string leapworkIds, string leapworkTitles)
        {

            SimpleLogger logger = new SimpleLogger(leapworkLog);

            

            logger.Info(String.Format("Leapwork hostname: {0}", leapworkHostname));
            logger.Info(String.Format("Leapwork port: {0}", leapworkPort));

            string controllerApiHttpAddress = GetControllerApiHttpAddress(leapworkHostname, leapworkPort, logger);

            logger.Info(String.Format("Leapwork Controller URL: {0}", controllerApiHttpAddress));
            logger.Info(String.Format("Leapwork Access Key: {0}", leapworkAccessKey));
            logger.Info(String.Format("Time Delay: {0}", leapworkTime));
            logger.Info(String.Format("Done Status As: {0}", leapworkDoneStatus));
            logger.Info(String.Format("Raw report file path: {0}", leapworkReport));
            logger.Info(String.Format("Log file path: {0}", leapworkLog));
            logger.Info(String.Format("Schedule ids: {0}", leapworkIds));
            logger.Info(String.Format("Schedule names: {0}", leapworkTitles));

            String junitReportPath = GetJunitReportFilePath(leapworkReport); //checks if .xml in the path exists
            logger.Info(String.Format("Full Report file path: {0}", junitReportPath));

            Dictionary<Guid, string> schedulesIdTitleDictionary; // Id-Title
            List<InvalidSchedule> invalidSchedules = new List<InvalidSchedule>();
            List<String> rawScheduleList = GetRawScheduleList(leapworkIds, leapworkTitles);

            int timeDelay = GetTimeDelay(leapworkTime,logger);

            Dictionary<Guid, LeapworkRun> resultsMap = new Dictionary<Guid, LeapworkRun>();

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("AccessKey", leapworkAccessKey);

                    schedulesIdTitleDictionary = GetSchedulesIdTitleDictionary(client,controllerApiHttpAddress, rawScheduleList, invalidSchedules, logger).Result;
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
                            CollectScheduleRunResults(client, controllerApiHttpAddress, runId, schTitle, timeDelay, leapworkDoneStatus, leapworkRun, logger);
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
                        logger.Info(string.Format("{0}: {1}", invalidSchedule.Name, invalidSchedule.StackTrace));
                        LeapworkRun notFoundLeapworkRun = new LeapworkRun(invalidSchedule.Name);
                        RunItem invalidRunItem = new RunItem("Error", "Error", 0, invalidSchedule.StackTrace, invalidSchedule.Name);
                        notFoundLeapworkRun.RunItems.Add(invalidRunItem);
                        notFoundLeapworkRun.Error = invalidSchedule.StackTrace;
                        buildResult.Runs.Add(notFoundLeapworkRun);
                    }
                }

                List<LeapworkRun> resultRuns = new List<LeapworkRun>(resultsMap.Values);

                foreach (LeapworkRun leapworkRun in resultRuns)
                {
                    buildResult.Runs.Add(leapworkRun);
                    buildResult.AddFailedTests(leapworkRun.Failed);
                    buildResult.AddPassedTests(leapworkRun.Passed);
                    buildResult.AddErrors(leapworkRun.Errors);
                    leapworkRun.Total = leapworkRun.Passed + leapworkRun.Failed;
                    buildResult.AddTotalTime(leapworkRun.Time);
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
                if (e is ThreadAbortException || e is ThreadInterruptedException)
                {
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

        private static void CollectScheduleRunResults(HttpClient client, string controllerApiHttpAddress,  Guid runId, string scheduleName, int timeDelay, string doneStatusAs, LeapworkRun resultRun, SimpleLogger logger)
        {
            List<Guid> runItemsId = new List<Guid>();

        try
        {
                bool isStillRunning = true;

                do
                {

                    Task.Delay(timeDelay * 1000).Wait();

                    List<Guid> executedRunItems = GetRunRunItems(client, controllerApiHttpAddress,  runId).Result;

                    foreach (Guid guid in runItemsId)
                    {
                        executedRunItems.Remove(guid); //left only new
                    }

                    executedRunItems.Reverse();

                    for (int i = executedRunItems.Count - 1; i >= 0; i--)
                    {
                        Guid runItemId = executedRunItems[i];
                        RunItem runItem = GetRunItem(client, controllerApiHttpAddress, runItemId, scheduleName, logger).Result;

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
                                if (doneStatusAs.Equals("Success"))
                                    resultRun.IncPassed();
                                else
                                    resultRun.IncFailed();
                                break;

                        }

                    }

                    runItemsId.AddRange(executedRunItems);

                    String runStatus = GetRunStatus(client, controllerApiHttpAddress,runId, logger).Result;
                    if (runStatus.Equals("Finished"))
                    {
                        List<Guid> allExecutedRunItems = GetRunRunItems(client, controllerApiHttpAddress, runId).Result;
                        if (allExecutedRunItems.Count > 0 && allExecutedRunItems.Count <= runItemsId.Count)//todo ==
                            isStillRunning = false;
                    }

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
                    throw e;
                }
                else
                {
                    logger.Error(e.StackTrace);
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

#Write-Output "##vso[task.addattachment type=Distributedtask.Core.Summary;name=Leaptest-TFS Integration Console Output;]$logFile"

Write-Output "##vso[build.uploadlog]$logFile"



