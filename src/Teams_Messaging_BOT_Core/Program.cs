using Teams_Messaging_BOT.Helpers;
using Teams_Messaging_BOT.Models;
using Teams_Messaging_BOT.MSGraphModels;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Teams_Messaging_BOT
{
    class Program
    {
        static string MessagesDirectoryPath;
        static string tenant_id;
        static string client_id;
        static string client_secret;
        static string username;
        static string password;
        static string GraphAPIUrl = "https://graph.microsoft.com";
        static string _accessToken;
        static string _defaultPermissionScope;
        static DateTime _tokenExpiryTime;
        //static bool IsFormHidden = true;
        static string ErrorNotifyChannel;
        static string ErrorNotifyTag;
        static string ErrorNotifySMSNo;
        static int RetryAttempts;
        static int RetryTimeInterval;
        static CommonFunc _helper = new CommonFunc();
        static Dictionary<string, DateTime> FileProcessed = new Dictionary<string, DateTime>();
        static void Main(string[] args)
        {
            bool instanceCountOne = false;

            using (Mutex mtex = new Mutex(true, "Teams Messaging BOT", out instanceCountOne))
            {
                if (instanceCountOne)
                {
                    _helper.WriteLog("Teams Messaging BOT Started.........", LogType.Info);

                    Console.WriteLine("Teams Messaging BOT Started.........");

                    if (!UpdateGlobalVars())
                    {
                        Console.ReadLine();
                    }
                    else
                    {

                        _helper.WriteLog($"Reading Files.........", LogType.Info);

                        Console.WriteLine("Reading Files.........");

                        ReadFilesOnStartUp();

                        cleanErrorList();

                        _helper.WriteLog($"Reading Retry Files.........", LogType.Info);

                        Console.WriteLine("Reading Retry Files.........");
                        RetryErrorFiles();

                        //Console.ReadLine();
                    }
                }
                else
                {
                    _helper.WriteLog("Teams Messaging BOT already running.........", LogType.Info);
                    Console.WriteLine("Already running.........");
                    //Console.ReadLine();
                }
            }
        }



        private static string GetAccessToken(string scope = "")
        {
            try
            {
                if (_tokenExpiryTime <= DateTime.Now || _accessToken == "")
                {
                    var client = new RestClient("https://login.microsoftonline.com");
                    var request = new RestRequest("/" + tenant_id + "/oauth2/v2.0/token", Method.Post);
                    request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
                    request.AddParameter("client_id", client_id);
                    request.AddParameter("scope", (scope != "") ? scope : _defaultPermissionScope);
                    request.AddParameter("client_secret", client_secret, true);
                    request.AddParameter("username", username);
                    request.AddParameter("password", password);
                    request.AddParameter("grant_type", "password");
                    var response = client.ExecuteAsync(request).Result.Content;

                    if (!response.StartsWith("{\"error"))
                    {
                        AccessTokenResponseModel model = JsonConvert.DeserializeObject<AccessTokenResponseModel>(response);
                        _tokenExpiryTime = DateTime.Now.AddSeconds(model.expires_in - 600);
                        _accessToken = model.access_token;
                        SaveTokenExpiryTime(_tokenExpiryTime);
                        SaveToken(_accessToken);
                        return model.access_token;
                    }
                    else
                    {
                        var errorResponse = JsonConvert.DeserializeObject<APIErrorResponse>(response);
                        throw new Exception(errorResponse.error.message);
                    }
                }
                else
                {
                    return _accessToken;
                }
            }
            catch (Exception ex)
            {
                throw new RetryMessageException(ex.Message);
            }
        }

        private static DateTime GetTokenExpiryTime()
        {
            try
            {
                if (File.Exists("TokenExpiry.txt"))
                {
                    string strDate = File.ReadAllText("TokenExpiry.txt");

                    try
                    {
                        return DateTime.Parse(strDate);
                    }
                    catch (Exception ex)
                    {
                        return DateTime.Now;
                    }
                }
                else
                {
                    return DateTime.Now;
                }
            }
            catch (Exception ex)
            {
                return DateTime.Now;
            }
        }

        private static void SaveTokenExpiryTime(DateTime TokenExpiryTime)
        {
            try
            {
                File.WriteAllText("TokenExpiry.txt", TokenExpiryTime.ToString());
            }
            catch { }
        }

        private static void SaveToken(string Token)
        {
            try
            {
                File.WriteAllText("Token.txt", _helper.Encrypt(Token));
            }
            catch { }
        }

        private static string GetToken()
        {
            try
            {
                if (File.Exists("Token.txt"))
                {
                    string strToken = File.ReadAllText("Token.txt");

                    try
                    {
                        return _helper.Decrypt(strToken);
                    }
                    catch (Exception ex)
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        private static bool UpdateGlobalVars()
        {
            var settings = ConfigurationManager.AppSettings;
            MessagesDirectoryPath = settings["MessagesDirectory"];
            tenant_id = _helper.Decrypt(settings["TenantId"]);
            client_id = _helper.Decrypt(settings["ClientId"]);
            client_secret = _helper.Decrypt(settings["ClientSecret"]);
            username = _helper.Decrypt(settings["UserName"]);
            password = _helper.Decrypt(settings["Password"]);
            ErrorNotifyChannel = settings["ErrorNotifyChannel"];
            ErrorNotifyTag = settings["ErrorNotifyTag"];
            ErrorNotifySMSNo = settings["ErrorNotifySMSNo"];
            RetryAttempts = Convert.ToInt32(settings["RetryAttempts"]);
            RetryTimeInterval = Convert.ToInt32(settings["RetryTimeInterval"]);
            _tokenExpiryTime = GetTokenExpiryTime();//DateTime.Now;
            _accessToken = GetToken();
            //_defaultPermissionScope = "Team.ReadBasic.All TeamSettings.Read.All TeamSettings.ReadWrite.All User.Read.All User.ReadWrite.All Directory.Read.All Directory.ReadWrite.All ChannelSettings.Read.Group Channel.ReadBasic.All ChannelSettings.Read.All ChannelSettings.ReadWrite.All Group.Read.All Group.ReadWrite.All Directory.Read.All ChatMessage.Send Chat.ReadWrite openid";
            _defaultPermissionScope = "TeamworkTag.Read TeamworkTag.ReadWrite Team.ReadBasic.All Channel.ReadBasic.All ChannelMember.Read.All ChatMessage.Send Chat.ReadWrite openid";

            if (string.IsNullOrEmpty(MessagesDirectoryPath))
            {
                Console.WriteLine("Please update [MessagesDirectory] in config file.");
                return false;
            }
            else if (string.IsNullOrEmpty(tenant_id))
            {
                Console.WriteLine("Please update [TenantId] in config file.");
                return false;
            }
            else if (string.IsNullOrEmpty(client_id))
            {
                Console.WriteLine("Please update [ClientId] in config file.");
                return false;
            }
            else if (string.IsNullOrEmpty(client_secret))
            {
                Console.WriteLine("Please update [ClientSecret] in config file.");
                return false;
            }
            else if (string.IsNullOrEmpty(username))
            {
                Console.WriteLine("Please update [UserName] in config file.");
                return false;
            }
            else if (string.IsNullOrEmpty(password))
            {
                Console.WriteLine("Please update [Password] in config file.");
                return false;
            }
            else if (string.IsNullOrEmpty(ErrorNotifyChannel))
            {
                Console.WriteLine("Please update [ErrorNotifyChannel] in config file.");
                return false;
            }
            else if (string.IsNullOrEmpty(ErrorNotifyTag))
            {
                Console.WriteLine("Please update [ErrorNotifyTag] in config file.");
                return false;
            }
            //else if (string.IsNullOrEmpty(ErrorNotifySMSNo))
            //{
            //    Console.WriteLine("Please update [ErrorNotifySMSNo] in config file.");
            //    return false;
            //}

            GetAccessToken();
            return true;

        }

        private static void ReadFilesOnStartUp()
        {
            DirectoryInfo dir = new DirectoryInfo(MessagesDirectoryPath);    //GetFiles(MessagesDirectoryPath, "*.*", SearchOption.TopDirectoryOnly);
            var files = dir.GetFiles("*.message", SearchOption.AllDirectories);

            int FilesFound = 0;
            foreach (var file in files)
            {
                string FolderName = Path.GetFileName(Path.GetDirectoryName(file.FullName)).ToLower();
                if (FolderName != "error" && FolderName != "send")
                {
                    FilesFound += 1;
                }
            }

            _helper.WriteLog(FilesFound.ToString() + " File(s) found.......", LogType.Info);
            Console.WriteLine(FilesFound.ToString() + " File(s) found.......");

            if (FilesFound == 0)
            {
                //Thread.Sleep(2000);
            }
            else
            {
                foreach (var file in files)
                {
                    string FolderName = Path.GetFileName(Path.GetDirectoryName(file.FullName)).ToLower();
                    if (FolderName != "error" && FolderName != "send")
                    {
                        _helper.WriteLog($"Processing message file, FilePath:{file.FullName}", LogType.Info);
                        if (CheckFileHasCopied(file.FullName))
                        {
                            _helper.WriteLog($"File accessible, reading file content, FilePath:{file.FullName}", LogType.Info);
                            SendMessage(file.FullName);
                        }
                    }
                }
            }
        }

        private static void cleanErrorList()
        {
            DirectoryInfo dir = new DirectoryInfo(MessagesDirectoryPath);
            var files = dir.GetFiles("*.retry", SearchOption.AllDirectories);

            var ErrorList = GetErrorList();

            ErrorList.RemoveAll(w =>
            !files.Select(s => Path.GetFileNameWithoutExtension(s.FullName))
            .Contains(w.FileName));

            UpdateErrorList(ErrorList);
        }

        private static void RetryErrorFiles()
        {
            DirectoryInfo dir = new DirectoryInfo(MessagesDirectoryPath);
            var files = dir.GetFiles("*.retry", SearchOption.AllDirectories);

            _helper.WriteLog($"{files.Count()} Retry File(s) Found.......", LogType.Info);

            foreach (var file in files)
            {
                string filenameoriginal = Path.GetFileNameWithoutExtension(file.FullName);

                var ErrorList = GetErrorList();

                var FileToProcess = ErrorList.Where(w => w.FileName == filenameoriginal).FirstOrDefault();

                if (FileToProcess != null)
                {
                    if (FileToProcess.RetryCount < RetryAttempts && FileToProcess.LastRetryDate.AddMinutes(RetryTimeInterval) <= DateTime.Now)
                    {
                        _helper.WriteLog($"Processing retry file, FilePath:{file.FullName}", LogType.Info);
                        SendMessage(file.FullName);
                    }
                    else if (FileToProcess.RetryCount >= RetryAttempts)
                    {
                        _helper.WriteLog($"Moving retry file to error after retry attempt completed, FilePath:{file.FullName}", LogType.Info);
                        MoveProcessedFile(file.FullName, false);
                        ErrorList.Remove(FileToProcess);
                        UpdateErrorList(ErrorList);
                    }
                }
            }
        }

        private static bool CheckFileHasCopied(string FilePath)
        {
            try
            {
                if (File.Exists(FilePath))
                    using (File.OpenRead(FilePath))
                    {
                        return true;
                    }
                else
                    return false;
            }
            catch (Exception)
            {
                Thread.Sleep(1000);
                return CheckFileHasCopied(FilePath);
            }
        }

        private static bool SendMessage(string FilePath, string EventType = "create", string OldFilePath = "")
        {
            string MessageContent = "";
            try
            {
                using (StreamReader sr = new StreamReader(FilePath, Encoding.Default))
                {
                    MessageContent = sr.ReadToEnd().Trim();
                }

                if (!MessageContent.EndsWith("File=\"No\""))
                {
                    _helper.WriteLog($"Incomplete or Corrupt file, FilePath:{FilePath}", LogType.Info);
                    throw new WithoutRetryException($"({Path.GetFileName(FilePath)})->Incomplete or Corrupt file.");
                }

                _helper.WriteLog($"Checking file already processed, FilePath:{FilePath}", LogType.Info);

                if (EventType == "create")
                {
                    if (FileProcessed.ContainsKey(FilePath))
                    {
                        var PreviousFile = FileProcessed[FilePath];

                        _helper.WriteLog($"Previous processed file found, Previous file time stamp: {PreviousFile.ToString("MM/dd/yyyy hh:mm:ss tt")}, FilePath:{FilePath}", LogType.Info);
                        if ((DateTime.Now - PreviousFile).TotalSeconds <= 1)
                        {
                            throw new WithoutRetryException("Duplicate file found, message sending stopped.");
                        }
                    }
                }
                else
                {
                    if (FileProcessed.ContainsKey(OldFilePath))
                    {
                        var PreviousFile = FileProcessed[OldFilePath];

                        _helper.WriteLog($"Previous file with Old fileName found, Previous file time stamp: {PreviousFile.ToString("MM/dd/yyyy hh:mm:ss tt")}, FilePath:{FilePath}", LogType.Info);
                        if ((DateTime.Now - PreviousFile).TotalSeconds <= 1)
                        {
                            throw new WithoutRetryException("Duplicate file found, message sending stopped.");
                        }
                    }
                }

                MessageInfoModel messageInfo = getMessageInfo(MessageContent);

                _helper.WriteLog($"Send message process starts, FilePath:{FilePath}", LogType.Info);
                var response = SendMessage(messageInfo);

                if (!response.StartsWith("{\"error"))
                {
                    if (FileProcessed.ContainsKey(FilePath))
                    {
                        FileProcessed[FilePath] = DateTime.Now;
                    }
                    else
                    {
                        FileProcessed.Add(FilePath, DateTime.Now);
                    }

                    _helper.WriteLog($"Message successfully send, moving file to send folder, FilePath:{FilePath}", LogType.Info);
                    MoveProcessedFile(FilePath, true);
                    _helper.WriteLog($"------------------------------------------------------------------------------------------", LogType.Info);
                }
                else
                {
                    var errorResponse = JsonConvert.DeserializeObject<APIErrorResponse>(response);
                    //throw new MessageNotSendExceptions(errorResponse.error.message);
                    throw new RetryMessageException("ResponseCode:Unauthorized, Errors:Unauthorized {\"error\":{\"code\":\"Unauthorized\",\"message\":\"Identity required\",\"innerError\":{\"message\":\"Identity required\",\"code\":\"NoAuthenticationToken\",\"innerError\":{},\"date\":\"2023-08-16T22:48:34\",\"request-id\":\"d4f6a732-21cc-4a13-9c08-608d1b32e93d\",\"client-request-id\":\"d4f6a732-21cc-4a13-9c08-608d1b32e93d\"}}}Request failed with status code Unauthorized");
                }

                return true;

            }
            catch (RetryMessageException exm)
            {
                _helper.WriteLog($"An error occured(check error log), FilePath:{FilePath}", LogType.Info);

                string ErrorMessage = $"FileName: {Path.GetFileName(FilePath)}{Environment.NewLine}Error Details:{Environment.NewLine}{exm.Message}";
                _helper.WriteLog($"An error occurred while sending message, {ErrorMessage}", LogType.Error);

                try
                {
                    string FileNameProcessing = Path.GetFileNameWithoutExtension(FilePath);
                    var ErrorList = GetErrorList();
                    if (ErrorList.Count == 0 || (ErrorList.Count > 0 && ErrorList.Where(w => w.FileName == FileNameProcessing).Count() == 0))
                    {
                        ErrorFilesModel errorFilesModel = new ErrorFilesModel();
                        errorFilesModel.ErroDetails = exm.Message;
                        errorFilesModel.FileName = FileNameProcessing;
                        errorFilesModel.LastRetryDate = DateTime.Now;
                        errorFilesModel.RecordDate = DateTime.Now;
                        errorFilesModel.RetryCount = 0;

                        SaveErrorList(errorFilesModel);

                        _helper.WriteLog($"Moving file for retry, FilePath:{FilePath}", LogType.Info);
                        MoveRetryFile(FilePath);
                    }
                    else
                    {
                        var RetryFile = ErrorList.Where(w => w.FileName == FileNameProcessing).FirstOrDefault();
                        if (RetryFile.RetryCount == RetryAttempts)
                        {
                            NotifyErrorToTeam(FilePath, exm.Message, MessageContent);
                            _helper.WriteLog($"Moving retry file to error after retry attempt completed, FilePath:{FilePath}", LogType.Info);
                            MoveProcessedFile(FilePath, false);
                            ErrorList.Remove(RetryFile);
                        }
                        else
                        {
                            ErrorList.ForEach(f =>
                            {
                                if (f.FileName == FileNameProcessing)
                                {
                                    f.LastRetryDate = DateTime.Now;
                                    f.RetryCount += 1;
                                    f.ErroDetails = exm.Message;
                                    _helper.WriteLog($"Error on retry, FilePath:{FilePath}", LogType.Info);
                                }
                            });
                        }

                        UpdateErrorList(ErrorList);
                    }

                }
                catch (Exception ex2)
                {
                    _helper.WriteLog($"An error occurred while moving the file to error folder(check error log), FilePath:{FilePath}", LogType.Info);
                    _helper.WriteLog($"An error occurred while moving the file to error folder, Error Details:{Environment.NewLine}{ex2.Message}", LogType.Error);
                }

                _helper.WriteLog($"------------------------------------------------------------------------------------------", LogType.Info);
                return false;
            }
            catch (WithoutRetryException exr)
            {
                _helper.WriteLog($"An error occured(check error log), FilePath:{FilePath}", LogType.Info);

                string ErrorMessage = $"FileName: {Path.GetFileName(FilePath)}{Environment.NewLine}Error Details:{Environment.NewLine}{exr.Message}";
                _helper.WriteLog($"An error occurred while sending message, {ErrorMessage}", LogType.Error);

                NotifyErrorToTeam(FilePath, exr.Message, MessageContent);
                try
                {
                    _helper.WriteLog($"Moving file to error folder, FilePath:{FilePath}", LogType.Info);
                    MoveProcessedFile(FilePath, false);
                }
                catch (Exception ex2)
                {
                    _helper.WriteLog($"An error occurred while moving the file to error folder(check error log), FilePath:{FilePath}", LogType.Info);
                    _helper.WriteLog($"An error occurred while moving the file to error folder, Error Details:{Environment.NewLine}{ex2.Message}", LogType.Error);
                }

                _helper.WriteLog($"------------------------------------------------------------------------------------------", LogType.Info);
                return false;
            }
            catch (Exception ex)
            {
                _helper.WriteLog($"An error occured(check error log), FilePath:{FilePath}", LogType.Info);

                string ErrorMessage = $"FileName: {Path.GetFileName(FilePath)}{Environment.NewLine}Error Details:{Environment.NewLine}{ex.Message}";
                _helper.WriteLog($"An error occurred while sending message, {ErrorMessage}", LogType.Error);

                NotifyErrorToTeam(FilePath, ex.Message, MessageContent);
                try
                {
                    _helper.WriteLog($"Moving file to error folder, FilePath:{FilePath}", LogType.Info);
                    MoveProcessedFile(FilePath, false);

                }
                catch (Exception ex2)
                {
                    _helper.WriteLog($"An error occurred while moving the file to error folder(check error log), FilePath:{FilePath}", LogType.Info);
                    _helper.WriteLog($"An error occurred while moving the file to error folder, Error Details:{Environment.NewLine}{ex2.Message}", LogType.Error);
                }

                _helper.WriteLog($"------------------------------------------------------------------------------------------", LogType.Info);
                return false;
            }

        }

        private static void NotifyErrorToTeam(string FilePath, string exMessage, string MessageContent)
        {
            string ErrorMessage = $"FileName: {Path.GetFileName(FilePath)}{Environment.NewLine}Error Details:{Environment.NewLine}{exMessage}";

            try
            {
                MessageInfoModel messageInfo = new MessageInfoModel();
                messageInfo.messageToType = MessageToType.Channel;
                messageInfo.To = ErrorNotifyChannel;
                messageInfo.Title = "An error occurred while sending message.";
                messageInfo.Message = $"{ErrorNotifyTag} {ErrorMessage}";

                _helper.WriteLog($"Sending error details to {ErrorNotifyChannel}, FilePath:{FilePath}", LogType.Info);

                SendMessage(messageInfo, true, ErrorNotifyTag, MessageContent);

                _helper.WriteLog($"Error details send successfully to {ErrorNotifyChannel}, FilePath:{FilePath}", LogType.Info);
            }
            catch (Exception iex)
            {
                _helper.WriteLog($"Unable to send error details to {ErrorNotifyChannel}(check error log), FilePath:{FilePath}", LogType.Info);

                _helper.WriteLog($"Unable to send error details to {ErrorNotifyChannel}, ErrorDetails:{Environment.NewLine}{iex.Message}", LogType.Error);

                _helper.WriteLog($"Sending error notification to SMS, FilePath:{FilePath}", LogType.Info);
                if (!string.IsNullOrEmpty(ErrorNotifySMSNo))
                    SendSMS(ErrorNotifySMSNo, $"Unable to send message to teams, FileName: {Path.GetFileName(FilePath)}", FilePath);
            }
        }

        public static MessageInfoModel getMessageInfo(string MessageContent)
        {

            MessageInfoModel model = new MessageInfoModel();

            string messagetype = getValueMessageData(MessageContent, "ChatOrChannel");

            if (messagetype.ToLower() == "chat")
            {
                model.messageToType = MessageToType.Chat;
            }
            else if (messagetype.ToLower() == "channel")
            {
                model.messageToType = MessageToType.Channel;
            }

            model.To = getValueMessageData(MessageContent, "TeamsNameOrUsername");
            model.Title = getValueMessageData(MessageContent, "Title");
            model.SubTitle = getValueMessageData(MessageContent, "SubTitle");
            model.Message = getValueMessageData(MessageContent, "Message").Replace(Environment.NewLine, $"<br>{Environment.NewLine}").Replace("	", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	");

            if (getValueMessageData(MessageContent, "Image").ToLower() != "no")
            {
                model.HasImage = true;
                model.ImagePath = getValueMessageData(MessageContent, "Image");
            }

            if (getValueMessageData(MessageContent, "File").ToLower() != "no")
            {
                model.HasFile = true;
                model.FilePath = getValueMessageData(MessageContent, "File");
            }

            return model;
        }

        private static string getValueMessageData(string MessageContent, string Key)
        {
            Regex CSVParser = new Regex($"{Environment.NewLine}(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");
            String[] Fields = CSVParser.Split(MessageContent);

            string Value = Fields.Where(w => w.ToLower().Split('=')[0] == Key.ToLower()).Select(s => s.Split('=')[1]).FirstOrDefault() ?? "";
            Value = Value.TrimStart('"').TrimEnd('"');
            Value = Value.Split(';')[0];

            return Value.Trim();

        }

        private static string SendMessage(MessageInfoModel messageInfo, bool IsErrorNotify = false, string MentionedTag = "", string OriginalMessage = "")
        {
            string AccessToken = GetAccessToken();

            //Chat Message
            if (messageInfo.messageToType == MessageToType.Chat)
            {
                //OneOnOne Chat
                if (_helper.IsEmailAddress(messageInfo.To))
                {
                    string ChatId = GetChatId(username, messageInfo.To);

                    if (string.IsNullOrWhiteSpace(ChatId))
                    {
                        throw new RetryMessageException("Chat username not found.");
                    }

                    ChatMessageBodyModel messagebody = new ChatMessageBodyModel();
                    messagebody.contentType = "html";
                    messagebody.content = "";

                    if (!string.IsNullOrEmpty(messageInfo.Title))
                        messagebody.content += $"<h1>{messageInfo.Title}</h1><br>" + Environment.NewLine;

                    if (!string.IsNullOrEmpty(messageInfo.SubTitle))
                        messagebody.content += $"<h3>{messageInfo.SubTitle}</h3><br>" + Environment.NewLine;

                    if (!string.IsNullOrEmpty(messageInfo.Message))
                        messagebody.content += messageInfo.Message;

                    List<ChatMessageHostedContentModel> imgs = new List<ChatMessageHostedContentModel>();

                    if (messageInfo.HasImage)
                    {
                        var img = new ChatMessageHostedContentModel();
                        img.contentBytes = Convert.ToBase64String(System.IO.File.ReadAllBytes(messageInfo.ImagePath));
                        img.contentType = "image/png";
                        img.MicrosoftGraphTemporaryId = "1";
                        imgs.Add(img);

                        messagebody.content += Environment.NewLine + "<br><img src=\"../hostedContents/1/$value\">";
                    }

                    object JsonToSend = null;

                    if (imgs.Count > 0)
                    {
                        JsonToSend = new { body = messagebody, hostedContents = imgs };
                    }
                    else
                    {
                        JsonToSend = new { body = messagebody };
                    }

                    AccessToken = GetAccessToken();

                    string response = _helper.getGraphResponse(AccessToken, GraphAPIUrl, $"/chats/{ChatId}/messages", Method.Post, JsonToSend);

                    return response;
                }
                //Group Chat
                else
                {
                    string ChatId = GetChatGroupId(username, messageInfo.To);

                    var messagebody = new ChatMessageBodyModel();
                    messagebody.contentType = "html";
                    messagebody.content = "";

                    if (!string.IsNullOrEmpty(messageInfo.Title))
                        messagebody.content += $"<h1>{messageInfo.Title}</h1>" + Environment.NewLine;

                    if (!string.IsNullOrEmpty(messageInfo.SubTitle))
                        messagebody.content += $"<h3>{messageInfo.SubTitle}</h3>" + Environment.NewLine;

                    if (!string.IsNullOrEmpty(messageInfo.Message))
                        messagebody.content += messageInfo.Message;

                    List<ChatMessageHostedContentModel> imgs = new List<ChatMessageHostedContentModel>();

                    if (messageInfo.HasImage)
                    {
                        var img = new ChatMessageHostedContentModel();
                        img.contentBytes = Convert.ToBase64String(System.IO.File.ReadAllBytes(messageInfo.ImagePath));
                        img.contentType = "image/png";
                        img.MicrosoftGraphTemporaryId = "1";
                        imgs.Add(img);

                        messagebody.content += Environment.NewLine + "<img src=\"../hostedContents/1/$value\">";
                    }

                    object JsonToSend = null;

                    if (imgs.Count > 0)
                    {
                        JsonToSend = new { body = messagebody, hostedContents = imgs };
                    }
                    else
                    {
                        JsonToSend = new { body = messagebody };
                    }

                    AccessToken = GetAccessToken();

                    string response = _helper.getGraphResponse(AccessToken, GraphAPIUrl, $"/chats/{ChatId}/messages", Method.Post, JsonToSend);

                    return response;
                }
            }
            //Channel Message
            else //if (messageInfo.messageToType == MessageToType.Channel)
            {
                string[] To = messageInfo.To.Split('/');
                string TeamName = To[0];
                string ChannelName = "General";
                if (To.Count() > 1)
                {
                    ChannelName = To[1];
                }

                string TeamId = GetTeamId(username, TeamName);
                string ChannelId = GetChannelId(username, TeamId, ChannelName);

                List<string> TagsList = ExtractTags(messageInfo.Message);
                List<string> mentionsList = ExtractEmails(messageInfo.Message);

                if (IsErrorNotify)
                {
                    TagsList = new List<string>();
                    mentionsList = new List<string>();

                    if (!string.IsNullOrEmpty(MentionedTag))
                        TagsList.Add(MentionedTag.Replace("@", ""));
                }



                var messagebody = new ChatMessageBodyModel();
                messagebody.contentType = "html";
                messagebody.content = "";

                if (!string.IsNullOrEmpty(messageInfo.Title))
                    messagebody.content += $"<h1>{messageInfo.Title}</h1>" + Environment.NewLine;

                if (!string.IsNullOrEmpty(messageInfo.SubTitle))
                    messagebody.content += $"<h3>{messageInfo.SubTitle}</h3>" + Environment.NewLine;

                var mentionsToAdd = new List<object>();

                if (!string.IsNullOrEmpty(messageInfo.Message))
                {
                    if (mentionsList.Count > 0 || TagsList.Count > 0)
                    {
                        string Message = messageInfo.Message;

                        int mentionIndex = 0;

                        foreach (var item in mentionsList)
                        {
                            var user = getUserDetails(item);
                            var mention1 = new ChatMessageMentionModel();
                            mention1.id = mentionIndex;
                            mention1.mentioned = new ChatMessageMentionedModel();
                            mention1.mentioned.user = new ChatMessageUserModel();
                            mention1.mentioned.user.id = user.id;
                            mention1.mentionText = user.displayName;

                            mentionsToAdd.Add(mention1);

                            Message = Message.Replace($"@{item}", $"<at id=\"{mentionIndex}\">{user.displayName}</at>");

                            mentionIndex++;
                        }

                        foreach (var item in TagsList)
                        {
                            var tag = getTagDetails(TeamId, item);
                            var mention1 = new ChatMessageMentionTagModel();
                            mention1.id = mentionIndex;
                            mention1.mentioned = new ChatMessageMentionedTagModel();
                            mention1.mentioned.tag = new ChatMessageTagModel();
                            mention1.mentioned.tag.id = tag.id;
                            mention1.mentioned.tag.displayName = tag.displayName;
                            mention1.mentionText = tag.displayName;

                            mentionsToAdd.Add(mention1);

                            Message = Message.Replace($"@{item}", $"<at id=\"{mentionIndex}\">{tag.displayName}</at>");

                            mentionIndex++;
                        }

                        messagebody.content += Message;
                    }
                    else
                    {
                        messagebody.content += messageInfo.Message;
                    }
                }

                var imgs = new List<ChatMessageHostedContentModel>();

                if (messageInfo.HasImage)
                {
                    var img = new ChatMessageHostedContentModel();
                    img.contentBytes = Convert.ToBase64String(System.IO.File.ReadAllBytes(messageInfo.ImagePath));
                    img.contentType = "image/png";
                    img.MicrosoftGraphTemporaryId = "1";
                    imgs.Add(img);

                    messagebody.content += Environment.NewLine + "<img src=\"../hostedContents/1/$value\">";
                }

                if (IsErrorNotify && !string.IsNullOrWhiteSpace(OriginalMessage))
                {
                    messagebody.content += $"{Environment.NewLine}{Environment.NewLine}Message file content:{Environment.NewLine}{OriginalMessage}";
                }


                object JsonToSend = null;
                if (imgs.Count > 0 && mentionsToAdd.Count > 0)
                {
                    JsonToSend = new { body = messagebody, hostedContents = imgs, mentions = mentionsToAdd };
                }
                else if (mentionsToAdd.Count > 0)
                {
                    JsonToSend = new { body = messagebody, mentions = mentionsToAdd };
                }
                else if (imgs.Count > 0)
                {
                    JsonToSend = new { body = messagebody, hostedContents = imgs };
                }
                else
                {
                    JsonToSend = new { body = messagebody };
                }

                string APIVersion = "v1.0";

                if (TagsList.Count > 0)
                {
                    APIVersion = "beta";
                }

                AccessToken = GetAccessToken();

                string response = _helper.getGraphResponse(AccessToken, GraphAPIUrl, $"/teams/{TeamId}/channels/{ChannelId}/messages", Method.Post, JsonToSend, null, APIVersion);

                return response;
            }
        }

        public static string GetChatId(string FromEmail, string ToEmail)
        {
            string AccessToken = GetAccessToken();
            string response = _helper.getGraphResponse(AccessToken, GraphAPIUrl, $"/users/{FromEmail}/chats?$expand=members", Method.Get);

            if (!response.StartsWith("{\"error"))
            {
                var model = JsonConvert.DeserializeObject<ChatListResponseModel>(response);

                string ChatId = "";

                if (model.OdataCount == 0)
                {
                    ChatId = CreateChat(FromEmail, ToEmail);
                }
                else
                {
                    var Chats = model.value.Where(w => w.chatType == "oneOnOne").ToList();
                    foreach (var chat in Chats)
                    {
                        if (chat.members.Where(w => w.email.ToLower() == ToEmail.ToLower()).Any())
                        {
                            ChatId = chat.id;
                            break;
                        }
                    }
                }

                return ChatId;
            }
            else
            {
                var errorResponse = JsonConvert.DeserializeObject<APIErrorResponse>(response);
                throw new RetryMessageException(errorResponse.error.message);
            }


        }

        public static string CreateChat(string FromEmail, string ToEmail)
        {
            string AccessToken = GetAccessToken();
            var chat = new CreateChatModel();
            chat.chatType = "oneOnOne";
            chat.members = new List<ChatMemberModel>();
            chat.members.Add(new ChatMemberModel($"{GraphAPIUrl}/v1.0/users/{FromEmail}"));
            chat.members.Add(new ChatMemberModel($"{GraphAPIUrl}/v1.0/users/{ToEmail}"));

            string response = _helper.getGraphResponse(AccessToken, GraphAPIUrl, "/chats", Method.Post, chat);

            if (!response.StartsWith("{\"error"))
            {
                var model = JsonConvert.DeserializeObject<ChatValueResponse>(response);
                return model.id;
            }
            else
            {
                var errorResponse = JsonConvert.DeserializeObject<APIErrorResponse>(response);
                throw new RetryMessageException(errorResponse.error.message);
            }
        }

        public static string GetChatGroupId(string FromEmail, string GroupName)
        {
            string AccessToken = GetAccessToken();
            string response = _helper.getGraphResponse(AccessToken, GraphAPIUrl, $"/users/{FromEmail}/chats", Method.Get);

            if (!response.StartsWith("{\"error"))
            {
                var model = JsonConvert.DeserializeObject<ChatListResponseModel>(response);

                if (model.OdataCount == 0)
                {
                    throw new RetryMessageException("No chat groups found.");
                }
                else
                {
                    var Chat = model.value.Where(w => w.chatType == "group" && w.topic.ToLower() == GroupName.ToLower()).FirstOrDefault();
                    if (Chat == null)
                    {
                        throw new RetryMessageException($"Chat group with the name '" + GroupName + "' not found.");
                    }
                    else
                    {
                        return Chat.id;
                    }
                }
            }
            else
            {
                var errorResponse = JsonConvert.DeserializeObject<APIErrorResponse>(response);
                throw new RetryMessageException(errorResponse.error.message);
            }


        }

        public static string GetTeamId(string FromEmail, string TeamName)
        {
            string AccessToken = GetAccessToken();
            string response = _helper.getGraphResponse(AccessToken, GraphAPIUrl, $"/users/{FromEmail}/joinedTeams", Method.Get);

            if (!response.StartsWith("{\"error"))
            {
                var model = JsonConvert.DeserializeObject<TeamListResponseModel>(response);

                var Team = model.value.Where(w => w.displayName.ToLower() == TeamName.ToLower()).FirstOrDefault();

                if (Team == null)
                {
                    throw new RetryMessageException($"Team with the name '" + TeamName + "' not found.");
                }
                else
                {
                    return Team.id;
                }
            }
            else
            {
                var errorResponse = JsonConvert.DeserializeObject<APIErrorResponse>(response);
                throw new RetryMessageException(errorResponse.error.message);
            }
        }

        public static string GetChannelId(string FromEmail, string TeamId, string ChannelName)
        {
            string AccessToken = GetAccessToken();
            string response = _helper.getGraphResponse(AccessToken, GraphAPIUrl, $"/teams/{TeamId}/channels", Method.Get);

            if (!response.StartsWith("{\"error"))
            {
                var model = JsonConvert.DeserializeObject<ChannelListResponseModel>(response);

                var Channel = model.value.Where(w => w.displayName.ToLower() == ChannelName.ToLower()).FirstOrDefault();

                if (Channel == null)
                {
                    throw new RetryMessageException($"Channel with the name '" + Channel + "' not found.");
                }
                else
                {
                    return Channel.id;
                }
            }
            else
            {
                var errorResponse = JsonConvert.DeserializeObject<APIErrorResponse>(response);
                throw new RetryMessageException(errorResponse.error.message);
            }
        }

        private static List<string> ExtractEmails(string textToScrape)
        {
            Regex reg = new Regex(@"[a-zA-Z0-9._%+-]+@[a-zA-Z]+(\.[a-zA-Z0-9]+)+", RegexOptions.IgnoreCase);
            //Regex reg = new Regex(@"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,6}", RegexOptions.IgnoreCase);
            Match match;

            List<string> results = new List<string>();
            for (match = reg.Match(textToScrape); match.Success; match = match.NextMatch())
            {
                if (!(results.Contains(match.Value)))
                    results.Add(match.Value);
            }

            return results;
        }
        private static List<string> ExtractTags(string textToScrape)
        {
            Regex reg = new Regex(@"@(.*?) ", RegexOptions.IgnoreCase);
            //Regex reg = new Regex(@"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,6}", RegexOptions.IgnoreCase);
            Match match;

            List<string> results = new List<string>();
            for (match = reg.Match(textToScrape); match.Success; match = match.NextMatch())
            {
                if (!(results.Contains(match.Value)))
                    results.Add(Regex.Replace(match.Value, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled).Trim());
            }

            return results;
        }

        public static UserResponseModel getUserDetails(string EmailAddress)
        {
            string AccessToken = GetAccessToken();
            string response = _helper.getGraphResponse(AccessToken, GraphAPIUrl, $"/users/{EmailAddress}", Method.Get);
            if (!response.StartsWith("{\"error"))
            {
                var model = JsonConvert.DeserializeObject<UserResponseModel>(response);
                return model;
            }
            else
            {
                var errorResponse = JsonConvert.DeserializeObject<APIErrorResponse>(response);
                throw new RetryMessageException(errorResponse.error.message);
            }
        }

        public static TagsResponseModel getTagDetails(string TeamId, string TagName)
        {
            string AccessToken = GetAccessToken();
            string response = _helper.getGraphResponse(AccessToken, GraphAPIUrl, $"/teams/{TeamId}/tags", Method.Get, null, null, "beta");

            if (!response.StartsWith("{\"error"))
            {
                var modelList = JsonConvert.DeserializeObject<TagsListResponseModel>(response);

                if (modelList.OdataCount == 0)
                {
                    throw new RetryMessageException($"No tags found in the for selected team.");
                }

                var model = modelList.value.Where(w => w.displayName.ToLower() == ("@" + TagName.ToLower()) || w.displayName.ToLower() == TagName.ToLower()).FirstOrDefault();

                if (model == null)
                {
                    throw new RetryMessageException($"Tag with the name '@{TagName}' not found in the selected team.");
                }

                return model;
            }
            else
            {
                var errorResponse = JsonConvert.DeserializeObject<APIErrorResponse>(response);
                throw new RetryMessageException(errorResponse.error.message);
            }
        }

        private static void SendSMS(string PhoneNo, string Message, string FilePath)
        {
            try
            {
                Message += " " + DateTime.Now.ToString("MM/dd/yyyy HH:mm");
                string Url = "";

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    try
                    {
                        using (Stream respstream = response.GetResponseStream())
                        {
                            StreamReader reader = new StreamReader(respstream, Encoding.UTF8);
                            String responseString = reader.ReadToEnd();
                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.LoadXml(responseString);
                            _helper.WriteLog($"SMS API Response: {xmlDoc.SelectSingleNode("//errortext").InnerText}", LogType.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        _helper.WriteLog($"SMS API Response Error, ErrorDetails: {ex.Message}", LogType.Error);
                        _helper.WriteLog($"Unable to send SMS(check error log), FilePath:{FilePath}", LogType.Info);
                    }
                }
            }
            catch (WebException we)
            {
                var resp = we.Response as HttpWebResponse;
                if (resp != null)
                {
                    try
                    {
                        using (Stream respstream = resp.GetResponseStream())
                        {
                            StreamReader reader = new StreamReader(respstream, Encoding.UTF8);
                            String responseString = reader.ReadToEnd();
                            _helper.WriteLog($"Unable to send SMS, ErrorDetails: {we.Message}{Environment.NewLine}{responseString}", LogType.Error);
                            _helper.WriteLog($"Unable to send SMS(check error log), FilePath:{FilePath}", LogType.Info);
                        }
                    }
                    catch (Exception ex)
                    {
                        _helper.WriteLog($"Unable to send SMS, ErrorDetails: {we.Message}", LogType.Error);
                        _helper.WriteLog($"Unable to send SMS(check error log), FilePath:{FilePath}", LogType.Info);
                    }
                }
                else
                {
                    _helper.WriteLog($"Unable to send SMS, ErrorDetails: {we.Message}", LogType.Error);
                    _helper.WriteLog($"Unable to send SMS(check error log), FilePath:{FilePath}", LogType.Info);
                }

            }
            catch (Exception ex)
            {
                _helper.WriteLog($"Unable to send SMS, ErrorDetails: {ex.Message}", LogType.Error);
                _helper.WriteLog($"Unable to send SMS(check error log), FilePath:{FilePath}", LogType.Info);
            }
        }

        public static void MoveProcessedFile(string FilePath, bool IsSuccess)
        {
            string path = Path.GetDirectoryName(FilePath);
            string filenameoriginal = Path.GetFileNameWithoutExtension(FilePath);
            string filenameoriginalExt = Path.GetExtension(FilePath);
            string newextention = ((IsSuccess) ? ".send" : ".error");

            string filename = filenameoriginal + newextention;
            string newpath = Path.Combine(path, ((IsSuccess) ? "Send" : "Error"));

            if (!Directory.Exists(newpath))
            {
                Directory.CreateDirectory(newpath);
            }

            int Index = 0;
            while (File.Exists(Path.Combine(newpath, filename)))
            {
                Index += 1;
                filename = filenameoriginal + $"({Index}){newextention}";
            }

            File.Move(Path.Combine(path, $"{filenameoriginal}{filenameoriginalExt}"), Path.Combine(newpath, filename));
        }

        public static void MoveRetryFile(string FilePath)
        {
            string path = Path.GetDirectoryName(FilePath);
            string filenameoriginal = Path.GetFileNameWithoutExtension(FilePath);
            string filenameoriginalExt = Path.GetExtension(FilePath);
            string newextention = ".retry";

            string filename = filenameoriginal + newextention;
            string newpath = path;//Path.Combine(path, "Error");

            //if (!Directory.Exists(newpath))
            //{
            //    Directory.CreateDirectory(newpath);
            //}

            if (!File.Exists(Path.Combine(newpath, filename)))
                File.Move(Path.Combine(path, $"{filenameoriginal}{filenameoriginalExt}"), Path.Combine(newpath, filename));
        }


        private static bool SaveErrorList(ErrorFilesModel ErrorData)
        {
            string FilePath = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ErrorFilesList.json");
            List<ErrorFilesModel> lstMasterData = null;
            if (File.Exists(FilePath))
            {
                lstMasterData = JsonConvert.DeserializeObject<List<ErrorFilesModel>>(File.ReadAllText(FilePath));
            }
            else
            {
                lstMasterData = new List<ErrorFilesModel>();
            }

            lstMasterData.Add(ErrorData);

            File.WriteAllText(FilePath, JsonConvert.SerializeObject(lstMasterData));

            return true;
        }

        private static List<ErrorFilesModel> GetErrorList()
        {
            string FilePath = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ErrorFilesList.json");
            List<ErrorFilesModel> lstMasterData = null;

            if (File.Exists(FilePath))
            {
                lstMasterData = JsonConvert.DeserializeObject<List<ErrorFilesModel>>(File.ReadAllText(FilePath));
            }
            else
            {
                lstMasterData = new List<ErrorFilesModel>();
            }

            return lstMasterData;
        }

        private static bool UpdateErrorList(List<ErrorFilesModel> ErrorData)
        {
            string FilePath = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ErrorFilesList.json");
            File.WriteAllText(FilePath, JsonConvert.SerializeObject(ErrorData));

            return true;
        }
    }
}
