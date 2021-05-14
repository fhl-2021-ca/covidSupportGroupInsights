using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection; // to use Missing.Value

//TO DO: If you use the Microsoft Outlook 11.0 Object Library, uncomment the following line.
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Configuration;
using CsvHelper;
using Microsoft.CodeAnalysis;
using CsvHelper.Configuration;
using System.Globalization;

namespace Covidsupportgroup
{

    /**
     * 
     * 
     * Things to Note : 
     * 
     * 1. The Message.Body contains all the messages recieved earlier in the thread in text format.
     * 
     * 2. Finding cities in the email : If we find a city, we just return. If there is another city mentioned in the email, we won't get it.
     * 
     * 
     * TO DO :
     * 
     * 
     * How to put data in excel.
     * 
     **/
    class Program
    {

        private static int emailsToFilterLookbackDays;
        // Top cities by emails.
        private static int topCitiesCountNeeded;
        private static int topStateCountNeeded;
        private static int mailReceivedToday;
        private static Dictionary<string, int> citiesCountMap = new Dictionary<string, int>();
        private static Dictionary<string, int> stateCountMap = new Dictionary<string, int>();
        private static List<string> oxygenUniqueCities = new List<string>();
        private static List<string> remdesivirUniqueCities = new List<string>();
        private static List<string> cities = new List<string>();
        private static List<string> states = new List<string>();
        private static List<string> csvFileHeaders = new List<string> { "Date", "Active threads", "Closed threads", "Oxygen", "Remdesivir", "HospitalBeds", "ICUBeds", "Plasma", "RT-PCR", "Important Emails" };
        private static string applicationPath = "";
        private static string logFilePath = "";
        private static string logFileName = "";
        private static StringBuilder stringBuilder = new StringBuilder();

        public static int Main(string[] args)
        {
            applicationPath = Directory.GetCurrentDirectory().Replace("\\bin\\Debug", "");
            logFileName = $"log-{GetTodaysDay()}.{GetTodaysMonth()}.txt";
            logFilePath = Path.Combine(applicationPath, logFileName);
            cities = File.ReadAllLines(Path.Combine(applicationPath, "list_of_cities_and_towns.txt")).ToList();
            states = File.ReadAllLines(Path.Combine(applicationPath, "list_of_states.txt")).ToList();
            string folderName = ConfigurationManager.AppSettings.Get("folderName");
            try
            {
                emailsToFilterLookbackDays = Int32.Parse(ConfigurationManager.AppSettings.Get("emailsToFilterLookbackDays"));
                topCitiesCountNeeded = Int32.Parse(ConfigurationManager.AppSettings.Get("topCitiesCountNeeded"));
                topStateCountNeeded = Int32.Parse(ConfigurationManager.AppSettings.Get("topStateCountNeeded"));

                // Create the Outlook application.
                // in-line initialization
                Outlook.Application oApp = new Outlook.Application();

                // Get the MAPI namespace.
                Outlook.NameSpace oNS = oApp.GetNamespace("mapi");

                // Log on by using the default profile or existing session (no dialog box).
                oNS.Logon(Missing.Value, Missing.Value, false, false);

                // Alternate logon method that uses a specific profile name.
                // TODO: If you use this logon method, specify the correct profile name
                // and comment the previous Logon line.
                //oNS.Logon("profilename",Missing.Value,false,true);

                //The covid support folder has to be INSIDE Inbox folder.
                Outlook.MAPIFolder PRFolder = oApp.ActiveExplorer().Session.GetDefaultFolder
                (Outlook.OlDefaultFolders.olFolderInbox).Folders[folderName];

                // Not working correctly
                //DateTime dt = DateTime.Now.Subtract(new TimeSpan(0, 30, 0));
                DateTime end = DateTime.Now;
                DateTime start = end.Subtract(new TimeSpan(0, 50, 0));

                string filter = "[Start] >= '"
                    + start.ToString("g")
                    + "' AND [End] <= '"
                    + end.ToString("g") + "'";
                //stringBuilder.AppendLine(filter);

                //Get the Items collection in the Inbox folder.
                // DOES NOT WORK
                //Outlook.Items oItems = PRFolder.Items.Restrict("[ReceivedTime] > '" + dt.ToString("MM/dd/yyyy HH:mm") + "'"); ;
                Outlook.Items oItems = PRFolder.Items;
                stringBuilder.AppendLine($"Total emails retreived = {oItems.Count}");

                oItems.Sort("[ReceivedTime]", OlSortOrder.olAscending);
                List<Outlook.MailItem> filteredEmails = filterLastXDaysEmail(oItems, emailsToFilterLookbackDays);
                List<Outlook.MailItem> emailsReceivedToday = filterLastXDaysEmail(oItems, 1);
                mailReceivedToday = emailsReceivedToday.Count;
                int emailCount = 0;
                foreach (Outlook.MailItem oMsg in filteredEmails)
                {
                    emailCount++;
                    /*                    stringBuilder.AppendLine(oMsg.Subject);
                                        stringBuilder.AppendLine("SenderEmail");
                                        stringBuilder.AppendLine(getSenderEmailAddress(oMsg));
                                        stringBuilder.AppendLine("To email");
                                        GetSMTPAddressForRecipients(oMsg);
                                        stringBuilder.AppendLine("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!1");
                    */
                    /*  SenderEmail
                    /O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=MICROSOFT.ONMICROSOFT.COM-55760-ANURAG DWIVEDI (CSS)
                    To email
                    ROI COVID Support Group*/
                }

                stringBuilder.AppendLine($"Total emails to scan = {emailCount}");
                Dictionary<string, List<Outlook.MailItem>> conversations = Program.GroupByThread(filteredEmails);

                // Final insights
                stringBuilder.AppendLine($"Emails received today = {emailsReceivedToday.Count}");
                Program.findInsightsFromConversations(conversations);
            }

            //Error handler.
            catch (System.Exception e)
            {
                Console.WriteLine("{0} Exception caught: ", e);
            }

            WriteLog();
            Console.WriteLine("Program executed successfully. Press any key to exit.");
            Console.ReadLine();
            return 0;
        }

        private static List<Outlook.MailItem> filterLastXDaysEmail(Items oItems, int days)
        {
            List<Outlook.MailItem> keepEmails = new List<Outlook.MailItem>();
                        int i = 0;
            DateTime cutOffDate = DateTime.Now.Subtract(new TimeSpan(days, 0, 0, 0));
            foreach (Outlook.MailItem oMsg in oItems){
                if (oMsg.ReceivedTime.CompareTo(cutOffDate) > 0) {
                    keepEmails.Add(oMsg);
                }
                else {
                    stringBuilder.AppendLine($"Removing message with subject {oMsg.Subject}");
                }
                i++;
            }
            return keepEmails;
        }

        private static void GetSMTPAddressForRecipients(Outlook.MailItem mail)
        {
            const string PR_SMTP_ADDRESS =
                "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            Outlook.Recipients recips = mail.Recipients;
            foreach (Outlook.Recipient recip in recips)
            {
                Outlook.PropertyAccessor pa = recip.PropertyAccessor;
                string smtpAddress =
                    pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                stringBuilder.AppendLine(recip.Name + " SMTP=" + smtpAddress);
            }
        }
        private static string getSenderEmailAddress(Outlook.MailItem mail)
        {
            Outlook.AddressEntry sender = mail.Sender;
            string SenderEmailAddress = "";

            if (sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry || sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
            {
                Outlook.ExchangeUser exchUser = sender.GetExchangeUser();
                if (exchUser != null)
                {
                    SenderEmailAddress = exchUser.PrimarySmtpAddress;
                }
            }
            else
            {
                SenderEmailAddress = mail.SenderEmailAddress;
            }

            return SenderEmailAddress;
        }

        private static void findInsightsFromConversations(Dictionary<string, List<Outlook.MailItem>> conversations)
        {
            var uniqueThreads = conversations.Count;
            var oxygenCount = 0;
            var remdesivirCount = 0;
            var hospitalBedsCount = 0;
            var ICUBedsCount = 0;
            var plasmaCount = 0;
            var rtPCRCount = 0;
            var importantCount = 0;
            var closedThreads = 0;
            var activeThreads = 0;

            foreach (KeyValuePair<string, List<Outlook.MailItem>> entry in conversations)
            {
                List<Outlook.MailItem> messages = entry.Value;
                if (checkIfThreadClosed(messages)) {
                    closedThreads++;
                    continue;
                }
                else activeThreads++;

/*                // if the latest message in this conversation was sent before today, either it is closed or not replied to.
                if (messages[0].ReceivedTime.Day < GetTodaysDay())
                {
                    //lastActivityBeforeToday.Add(messages[0].WebLink);
                }
                else if (messages[0].ReceivedTime.Day == GetTodaysDay())
                {
                    //activeToday.Add(messages[0].Links);
                }
*/
                var isOxygenNeeded = findWordInThread(messages, "Oxygen");
                var isRemdesivirNeeded = findWordInThread(messages, "Remdesivir");
                var isHospitalBedNeeded = findWordInThread(messages, "hospital bed");
                var isPlasmaNeeded = findWordInThread(messages, "plasma");
                var isICUNeeded = findWordInThread(messages, "ICU bed");
                var isRtPCRNeeded = findWordInThread(messages, "rt pcr") || findWordInThread(messages, "rt-pcr");

                // Find the city and increase city count    
                Tuple<bool, string>  cityTuple = findKeywordInThread(messages, cities);
                Tuple<bool, string>  stateTuple = findKeywordInThread(messages, states);
                countHotCitiesOrState(cityTuple, citiesCountMap, messages);
                countHotCitiesOrState(stateTuple, stateCountMap, messages);

                if(isOxygenNeeded && cityTuple.Item1 && !oxygenUniqueCities.Contains(cityTuple.Item2)) {
                    oxygenUniqueCities.Add(cityTuple.Item2);
                }

                var isImportant = isImportantThread(messages);

                if (isOxygenNeeded)
                {
                    oxygenCount++;
                }

                if (isRemdesivirNeeded)
                {
                    remdesivirCount++;
                }

                if (isHospitalBedNeeded)
                {
                    hospitalBedsCount++;
                }

                if (isICUNeeded)
                {
                    ICUBedsCount++;
                }

                if (isPlasmaNeeded)
                {
                    plasmaCount++;
                }

                if (isRtPCRNeeded)
                {
                    rtPCRCount++;
                }

                if (isImportant)
                {
                    importantCount++;
                }
            }

            stringBuilder.AppendLine($" Today's date = {GetTodaysDay()}/{GetTodaysMonth()}");
            stringBuilder.AppendLine($" Active threads = {activeThreads}");
            stringBuilder.AppendLine($" Closed threads = {closedThreads}");
            stringBuilder.AppendLine($" oxygenCount = {oxygenCount}");
            stringBuilder.AppendLine($" remdesivirCount = {remdesivirCount}");
            stringBuilder.AppendLine($" hospitalBedsCount = {hospitalBedsCount}");
            stringBuilder.AppendLine($" ICUBedsCount = {ICUBedsCount}");
            stringBuilder.AppendLine($" plasmaCount = {plasmaCount}");
            stringBuilder.AppendLine($" RT-PCRCount = {rtPCRCount}");
            stringBuilder.AppendLine($" importantCount = {importantCount}");
            stringBuilder.AppendLine("\n\nHot cities:");
            printTopHitsDictionary(citiesCountMap, topCitiesCountNeeded);
            stringBuilder.AppendLine("\n\nHot states:");
            printTopHitsDictionary(stateCountMap, topStateCountNeeded);
            stringBuilder.AppendLine("\n\nOxygen is required in these cities:");
            printList(oxygenUniqueCities);

            //{ "Date", "Emails received today", "Active threads", "Closed threads", "Oxygen", "Remdesivir", "HospitalBeds", "ICUBeds", "Plasma", "RT-PCR", "Important Emails" };
            List<string> csvValues = new List<string>()
;           csvValues.Add($"{GetTodaysDay()}/{GetTodaysMonth()}");
            csvValues.Add(mailReceivedToday.ToString());
            csvValues.Add(activeThreads.ToString());
            csvValues.Add(closedThreads.ToString());
            csvValues.Add(oxygenCount.ToString());
            csvValues.Add(remdesivirCount.ToString());
            csvValues.Add(hospitalBedsCount.ToString());
            csvValues.Add(ICUBedsCount.ToString());
            csvValues.Add(plasmaCount.ToString());
            csvValues.Add(rtPCRCount.ToString());
            csvValues.Add(importantCount.ToString());
            WriteToCSV(csvValues);
        }

        private static bool checkIfThreadClosed(List<MailItem> messages)
        {
            if (isWordMatch(messages[0].Subject, "closed"))
                return true;
            return false;
        }

        private static void countHotCitiesOrState(Tuple<bool, string> tuple, Dictionary<string, int> countMap, List<MailItem> messages)
        {
            if (tuple.Item1 == true)
            {
                if (!countMap.ContainsKey(tuple.Item2))
                {
                    countMap.Add(tuple.Item2, 1);
                }
                else
                {
                    var value = 0;
                    countMap.TryGetValue(tuple.Item2, out value);
                    countMap[tuple.Item2] = value + 1;
                }
            }
            else
            {
                stringBuilder.AppendLine($"Unable to find the city/state used in {messages[0].Subject}");
            }
        }

        private static void printList(List<string> oxygenUniqueCities)
        {
            stringBuilder.AppendLine(String.Join("\n", oxygenUniqueCities));
        }

        private static void printTopHitsDictionary(Dictionary<string, int> dictionary, int topx)
        {
            var sortedDict = (from entry in dictionary orderby entry.Value descending select entry)
                           .ToDictionary(pair => pair.Key, pair => pair.Value).Take(topx);
            foreach (KeyValuePair<string, int> kvp in sortedDict)
            {
                stringBuilder.AppendLine($"{kvp.Key},{kvp.Value}");
            }
        }

        private static int GetTodaysDay()
        {
            return DateTime.UtcNow.Day;
        }

        private static int GetTodaysMonth()
        {
            return DateTime.UtcNow.Month;
        }

        private static Dictionary<string, List<Outlook.MailItem>> GroupByThread(List<Outlook.MailItem> toGroup)
        {
            Dictionary<string, List<Outlook.MailItem>> conversations = new Dictionary<string, List<Outlook.MailItem>>();
            foreach (Outlook.MailItem message in toGroup)
            {
                var key = message.ConversationID;
                if (conversations.ContainsKey(key))
                {
                    List<Outlook.MailItem> list = conversations[key];
                    if (list.Contains(message) == false)
                    {
                        list.Add(message);
                    }
                }
                else
                {
                    List<Outlook.MailItem> list = new List<Outlook.MailItem>();
                    list.Add(message);
                    conversations.Add(key, list);
                }
            }

            return conversations;
        }

        private static bool SentToEmail(List<Recipient> toRecipients, string email)
        {
            return null != toRecipients.Find(x => x.Address.Equals(email));
        }

        private static List<string> GetKeywords(Outlook.MailItem message)
        {
            List<string> list = new List<string>();
            if (isOxygen(message))
            {
                list.Add("oxygen");
            }

            if (message.Body.IndexOf("remdesivir", 0, StringComparison.CurrentCultureIgnoreCase) != -1)
            {
                list.Add("remdesivir");
            }

            return list;
        }

        private static bool isOxygenThread(List<Outlook.MailItem> messages)
        {
/*            foreach (Outlook.MailItem message in messages)
            {*/
                if (isOxygen(messages[0]))
                {
                    return true;
                }
//            }
            return false;
        }

        private static bool findWordInThread(List<Outlook.MailItem> messages, string word)
        {
/*            foreach (Outlook.MailItem message in messages)
            {*/
                if (isMessageBodyMatch(messages[0], word))
                {
                    return true;
//                }
            }
            return false;
        }
        private static Tuple<bool, string> findKeywordInThread(List<Outlook.MailItem> messages, List<string> keywordsList)
        {
            foreach (string keyWord in keywordsList) {
/*                foreach (Outlook.MailItem message in messages)
                {
*/                    if (isMessageBodyMatch(messages[0], keyWord))
                    {
                        return new Tuple<bool, string>(true, keyWord);
//                    }
                }
            }
            return new Tuple<bool, string>(false, "NA"); ;
        }

        private static bool isImportantThread(List<Outlook.MailItem> messages)
        {
            foreach (Outlook.MailItem message in messages)
            {
                if (isImportant(message))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool isOxygen(Outlook.MailItem message)
        {
            return message.Body.IndexOf("oxygen", 0, StringComparison.CurrentCultureIgnoreCase) != -1;
        }

        private static bool isDelhi(Outlook.MailItem message)
        {
            return message.Body.IndexOf("delhi", 0, StringComparison.CurrentCultureIgnoreCase) != -1;
        }

        private static bool isWordMatch(string text, string pattern)
        {
            return text.IndexOf(pattern, 0, StringComparison.CurrentCultureIgnoreCase) != -1;
        }

        private static bool isMessageBodyMatch(Outlook.MailItem message, string word)
        {
            return message.Body.IndexOf(word, 0, StringComparison.CurrentCultureIgnoreCase) != -1;
        }

        private static bool isImportant(Outlook.MailItem message)
        {
            return message.Importance.Equals(OlImportance.olImportanceHigh);
        }

        private static List<Outlook.MailItem> GetMessagesRecievedToday(List<Outlook.MailItem> allMessages)
        {
            var currentTime = DateTime.UtcNow;
            return allMessages.FindAll(x => x.ReceivedTime.Day == currentTime.Day);
        }
        private static void WriteLog()
        {
            File.WriteAllText(logFilePath, stringBuilder.ToString());            
        }

        private static void WriteToCSV(List<string> data)
        {
            string insightsFileName = Path.Combine(applicationPath, "insights.csv");
            if (!File.Exists(insightsFileName)) {
                using (StreamWriter sw = File.AppendText(insightsFileName)){
                    sw.WriteLine(String.Join(",", csvFileHeaders.ToArray()));
                }
            } else {
                using (StreamWriter sw = File.AppendText(insightsFileName))
                {
                    sw.WriteLine(String.Join(",", data.ToArray()));
                }
            }
        }
    }
}
