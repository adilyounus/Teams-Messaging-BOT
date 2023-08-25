using Newtonsoft.Json;
using RestSharp;
using RestSharp.Authenticators.OAuth2;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Teams_Messaging_BOT.Helpers
{
    public enum LogType
    {
        Error = 1,
        Info = 2
    }
    public class CommonFunc
    {
        public string getGraphResponse(string accessToken, string APIUrl, string APIMethod, Method method, object JsonBody = null, Dictionary<string, string> reqParamters = null, string APIVersion = "v1.0")
        {
            try
            {
                var client = new RestClient(APIUrl);

                client.Authenticator = new OAuth2AuthorizationRequestHeaderAuthenticator(accessToken, "Bearer");

                var request = new RestRequest($"/{APIVersion}/{APIMethod}", method);
                request.AddHeader("Accept", "application/json");

                if (reqParamters != null)
                {
                    foreach (var param in reqParamters)
                    {
                        request.AddParameter(param.Key, param.Value, true);
                    }
                }

                if (JsonBody != null)
                    request.AddBody(JsonConvert.SerializeObject(JsonBody), "application/json");

                var response = client.ExecuteAsync(request).Result;
                if (response.IsSuccessful)
                {
                    return response.Content;
                }
                else
                {
                    throw new RetryMessageException("ResponseCode:" + response.StatusCode.ToString() + ", Errors:" + (response.StatusDescription ?? "") + Environment.NewLine + (response.Content ?? "") + Environment.NewLine + (response.ErrorMessage ?? "") + Environment.NewLine + (response.ErrorException?.Message ?? ""));
                }
            }
            catch(Exception ex)
            {
                throw new RetryMessageException(ex.Message);
            }
        }

        public void WriteLog(string LogData, LogType logType)
        {
            try
            {
                string LogPath = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),"Logs", logType.ToString());
                if (!Directory.Exists(LogPath))
                    Directory.CreateDirectory(LogPath);

                using (StreamWriter sw = new StreamWriter(LogPath + "\\" + DateTime.Now.ToString("dd-MMM-yyyy") + ".log", true))
                {
                    //sw.WriteLine("".PadRight(50, '-'));
                    sw.WriteLine(DateTime.Now.ToString() + " - " + LogData);
                    sw.Flush();
                }
            }
            catch { }
        }
        public bool IsEmailAddress(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
        public string Encrypt(string clearText)
        {
            string EncryptionKey = "1ed9ca38-67c8-4234-917c-3c4c22f6ac1e";
            byte[] clearBytes = Encoding.Unicode.GetBytes(clearText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    clearText = Convert.ToBase64String(ms.ToArray());
                }
            }
            return clearText;
        }
        public string Decrypt(string cipherText)
        {
            string EncryptionKey = "1ed9ca38-67c8-4234-917c-3c4c22f6ac1e";
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] { 0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76 });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }
    }
}
