using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Linq;

namespace EmailClassifierConsole
{
    /* Add class description */
    class Program
    {
        static string[] Scopes = { GmailService.Scope.MailGoogleCom};
        static string ApplicationName = "Auto Email Classification";

        static void Main(string[] args)
        {
            try
            { 
                UserCredential credential;

                using (var stream = new FileStream(@"C:\temp\AutoEmailClassification\Credentials\client_secret.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = @"c:\temp\AutoEmailClassification\Credentials";
                    //credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart.json");
                    credPath = Path.Combine(credPath, "gmail-api-credentials.json");

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;

                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Gmail API service.
                var gmailService = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                //get our emails
                var emailListRequest = gmailService.Users.Messages.List("me");
                emailListRequest.LabelIds = "INBOX";
                emailListRequest.IncludeSpamTrash = false;
                emailListRequest.Q = "is:unread";
                emailListRequest.MaxResults = 10;

                var emailListResponse = emailListRequest.Execute();

                if (emailListResponse != null && emailListResponse.Messages != null)
                {
                    //loop through each email and get what fields you want...
                    foreach (var email in emailListResponse.Messages)
                    {
                        ClasssifyMessage(gmailService, email, string.Join(" ", args));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception encountered! " + "Error details: " + ex.Message); 
            }
            Console.Read();
        }

        public static Message ModifyMessage(GmailService service, String userId, String messageId, List<String> labelsToAdd, List<String> labelsToRemove)
        {
            ModifyMessageRequest mods = new ModifyMessageRequest();
            mods.AddLabelIds = labelsToAdd;
            mods.RemoveLabelIds = labelsToRemove;

            try
            {
                return service.Users.Messages.Modify(mods, userId, messageId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }

            return null;
        }

        public static string StripHTML(string input)
        {
            return Regex.Replace(input, "<.*?>", String.Empty);
        }

        private static int GetTermFrequencyWeight(string topic, string doc)
        {
            int result = 0;

            string[] x = topic.Split(" ".ToArray()); // words in the query/topic
            string[] y = doc.Split(" ".ToArray());

            List<string> listX = new List<string>();
            listX.AddRange(x);

            List<string> listY = new List<string>();
            listY.AddRange(y);

            foreach (string w in listX.Distinct())
            {
                result += listX.Count(s => string.Compare(s, w, true) == 0) * listY.Count(t => string.Compare(t, w, true) == 0); 
            }

            return result;
        }

        private static double GetMatchScore(string topic, string doc)
        {
            // Set smoothing factor for topic words that don't appear in the doc
            double alpha = 0.0005;

            // word count
            double probablity = 1.00;

            string[] tmp = topic.Split(" ".ToCharArray());

            int docLength = doc.Length;
            List<string> docWordList = new List<string>();
            docWordList.AddRange(doc.Split(" ".ToCharArray()));

            foreach (string w in tmp)
            {
                int wordCount = docWordList.Where(s => string.Compare(s, w, true) == 0).Count();
                probablity *= wordCount > 0 ? ((double)wordCount / (double)docLength) : alpha; 
            }

            return probablity;
        }

        private static void ClasssifyMessage(GmailService gmailService, Message email, string topic)
        {
            var emailInfoRequest = gmailService.Users.Messages.Get("me", email.Id);
            //make another request for that email id...
            var emailInfoResponse = emailInfoRequest.Execute();

            if (emailInfoResponse != null)
            {
                String from = "";
                String date = "";
                String subject = "";
                String body = "";
                //loop through the headers and get the fields we need...
                foreach (var mParts in emailInfoResponse.Payload.Headers)
                {
                    if (mParts.Name == "Date")
                    {
                        date = mParts.Value;
                    }
                    else if (mParts.Name == "From")
                    {
                        from = mParts.Value;
                    }
                    else if (mParts.Name == "Subject")
                    {
                        subject = mParts.Value;
                    }

                    if (date != "" && from != "")
                    {
                        if (emailInfoResponse.Payload.Parts == null && emailInfoResponse.Payload.Body != null)
                            body = DecodeBase64String(emailInfoResponse.Payload.Body.Data);
                        else
                            body = GetNestedBodyParts(emailInfoResponse.Payload.Parts, "");
                    }
                }

                Console.WriteLine(string.Format("From: {0}, Subject: {1}", from, subject));

                int tf = GetTermFrequencyWeight(topic, subject + StripHTML(body));
                Console.WriteLine(string.Format("TF: {0}", tf));

                if (tf >= 1) // set threshold arbitrarily for testing purposes
                {
                    List<string> labelsToAdd = new List<string>();
                    labelsToAdd.Add(topic);
                    List<string> labelsToRemove = new List<string>();
                    
                    try
                    {
                        ModifyMessage(gmailService, "me", email.Id, labelsToAdd, labelsToRemove);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Unable to add label to message. Error details: " + ex.Message);
                    }
                }
            }
        }

        private static String DecodeBase64String(string s)
        {
            var ts = s.Replace("-", "+");
            ts = ts.Replace("_", "/");
            var bc = Convert.FromBase64String(ts);
            var tts = Encoding.UTF8.GetString(bc);

            return tts;
        }

        private static String GetNestedBodyParts(IList<MessagePart> part, string curr)
        {
            string str = curr;
            if (part == null)
            {
                return str;
            }
            else
            {
                foreach (var parts in part)
                {
                    if (parts.Parts == null)
                    {
                        if (parts.Body != null && parts.Body.Data != null)
                        {
                            var ts = DecodeBase64String(parts.Body.Data);
                            str += ts;
                        }
                    }
                    else
                    {
                        return GetNestedBodyParts(parts.Parts, str);
                    }
                }

                return str;
            }
        }
    }
}
