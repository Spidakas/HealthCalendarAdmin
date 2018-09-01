using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.Net.Mail;
using Microsoft.Exchange.WebServices.Data;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Exchange101;
using Microsoft.Exchange.WebServices.Autodiscover;

namespace HealthCalendarClasses
{
    class HealthCalendarClass
    {
        public string SubscriberID { get; set; }
        public string SubscriberOID { get; set; }
        public string Initials { get; set; }
        public string Title { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public DateTime ModifiedDTTM { get; set; }
        public DateTime EndDTTM { get; set; }
        public string SourceOID { get; set; }
        public string SourceType { get; set; }
        public string OwnerOrganisationOID { get; set; }
        public string RoneOID { get; set; }
        public string UITypCode{ get; set; }
        public string GoogleEmail { get; set; }
        public string GoogleCalendarID { get; set; }
        public string OutlookEmail { get; set; }
        public string OutlookCalendarID { get; set; }
        public string ExchangeEmail { get; set; }
        public string ExchangeCalendarID { get; set; }
        public string ExchangeCalendarName { get; set; }
        public string NHSNetEmail { get; set; }
        public string NHSNetCalendarID { get; set; }
        public string NHSNetCalendarName{ get; set; }
        public string AppleEmail { get; set; }
        public string AppleCalendarID { get; set; }

        public string HealthOrgCode { get; set; }
        public string HealthOrgName { get; set; }
        public string HealthOrgLocation { get; set; }
        public string GoogleOrgMasterAccount { get; set; }
        public string GoogleClientID { get; set; }
        public string GoogleClientSecret { get; set; }
        public string NHSNetExchangeServer { get; set; }
        public string NHSNetOrgMasterAccount { get; set; }
        public string NHSNetOrgMasterCredentials { get; set; }
        public string ExchangeExchangeServer { get; set; }
        public string ExchangeOrgMasterAccount { get; set; }
        public string ExchangeOrgMasterCredentials { get; set; }
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "HealthCalendar";
        public CalendarService GoogleCalenderService  { get; set; }
        public ExchangeService NHSNetCalendarService { get; set; }
        public ExchangeService ExchangeCalendarService { get; set; }

        static string MyConnString = ConfigurationManager.ConnectionStrings["connstrng"].ConnectionString;


        public void SendExchangeTestEmail(HealthCalendarClass c)
        {
            //ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            //service.Credentials = new WebCredentials("michael.georgiades@mbht.nhs.uk", "Lancaster345!!");
            //service.Credentials = new WebCredentials( c.NHSNetOrgMasterAccount, c.NHSNetOrgMasterCredentials);
            //service.TraceEnabled = true;
            //service.TraceFlags = TraceFlags.All;
            //service.AutodiscoverUrl("michael.georgiades@mbht.nhs.uk", RedirectionUrlValidationCallback);
            //service.Url = new Uri("https://canlfgh-mail01.xcanl.nhs.uk/EWS/Exchange.asmx");
            //service.Url = new Uri("https://mail.nhs.net/ews/exchange.asmx");


            //EmailMessage email = new EmailMessage(service);
            //email.ToRecipients.Add("michael.georgiades@mbht.nhs.uk");
            //email.ToRecipients.Add("michael.georgiades@nhs.net");
            //email.Subject = "HelloWorld";
            //email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
            //email.Send();
            CalendarFolder folder = new CalendarFolder(c.NHSNetCalendarService);
            folder.DisplayName = "New calendar folder";            
            folder.Save(WellKnownFolderName.Calendar);
        }

        public bool CreateShareGoogleDiary(HealthCalendarClass c)
        {
            bool isSuccess = false;

            if (string.IsNullOrEmpty (c.SubscriberID))
            {
                return isSuccess;
            }

            //Create new Calendar
            Calendar calendar = new Calendar();
            calendar.Summary = "Trust" + c.HealthOrgCode + c.Title + " " + c.FirstName + " " + c.LastName + " " + c.SubscriberID;
            calendar.TimeZone = "Europe/London";
            calendar.Location = c.HealthOrgLocation;
            calendar.Description = c.Title + " " + c.FirstName + " " + c.LastName + " " + c.HealthOrgName + " Activity Diary";
            Calendar createdCalendar = c.GoogleCalenderService.Calendars.Insert(calendar).Execute();
            c.GoogleCalendarID = createdCalendar.Id;

            AclRule rule = new AclRule()
            {
                Role = "reader",
                Scope = new AclRule.ScopeData()
                {
                    Type = "user",
                    Value = c.GoogleEmail
                }
            };
            AclRule createdrule = c.GoogleCalenderService.Acl.Insert(rule, c.GoogleCalendarID).Execute();

            //Send the user an email instructing on how to sync diary
            //MailMessage mail = new MailMessage();
            //SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            //mail.From = new MailAddress("trustrtx@gmail.com");
            //mail.To.Add(c.GoogleEmail);
            //mail.Subject = "Your new Google Activity diary has been created";
            //mail.Body = "This is the message body";

            //System.Net.Mail.Attachment attachment;
            //attachment = new System.Net.Mail.Attachment("c:/textfile.txt");
            //mail.Attachments.Add(attachment);

            //SmtpServer.Port = 587;
            //SmtpServer.Credentials = new System.Net.NetworkCredential("trustrtx@gmail.com", "Lancaster982!!");
            //SmtpServer.EnableSsl = true;

            //SmtpServer.Send(mail);


            // Update db Record
            SqlConnection conn = new SqlConnection(MyConnString);
            try { 
                string sql = "UPDATE Subscribers SET GoogleCalendarID=@GoogleCalendarID WHERE SubscriberID=@SubscriberID";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@GoogleCalendarID", c.GoogleCalendarID);
                cmd.Parameters.AddWithValue("@SubscriberID", c.SubscriberID);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }
            return isSuccess;

        }

        public bool CreateShareNHSNetDiary(HealthCalendarClass c)
        {
            bool isSuccess = false;

            if (string.IsNullOrEmpty(c.SubscriberID))
            {
                return isSuccess;
            }
                        
            CalendarFolder folder = new CalendarFolder(c.NHSNetCalendarService);            
            folder.DisplayName = "Trust" + c.HealthOrgCode + c.Title + " " + c.FirstName + " " + c.LastName + " " + c.SubscriberID;
            c.NHSNetCalendarName = folder.DisplayName;
            folder.Permissions.Add(new FolderPermission(c.NHSNetEmail,FolderPermissionLevel.Reviewer));
            folder.Save(WellKnownFolderName.Calendar);
            c.NHSNetCalendarID = folder.Id.ToString();

            // Bind to the folder
            /*Folder folderStoreInfo;
            folderStoreInfo = Folder.Bind(service, WellKnownFolderName.Calendar);
            string EwsID = folderStoreInfo.Id.UniqueId;

            // The value of folderidHex will be what we need to use for the FolderId in the xml file
            string folderidHex = GetConvertedEWSIDinHex(service, folderid, txtImpersonatedUser.Text);*/

            // Update db Record
            SqlConnection conn = new SqlConnection(MyConnString);
            try
            {
               string sql = "UPDATE Subscribers SET NHSNetCalendarID=@NHSNetCalendarID, NHSNetCalendarName=@NHSNetCalendarName WHERE SubscriberID=@SubscriberID";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@NHSNetCalendarID", c.NHSNetCalendarID);
                cmd.Parameters.AddWithValue("@NHSNetCalendarName", c.NHSNetCalendarName);
                cmd.Parameters.AddWithValue("@SubscriberID", c.SubscriberID);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {
            
            }
            finally
            {
                conn.Close();
            }
            return isSuccess;

        }
        public bool CreateShareExchangeDiary(HealthCalendarClass c)
        {
            bool isSuccess = false;

            if (string.IsNullOrEmpty(c.SubscriberID))
            {
                return isSuccess;
            }

            CalendarFolder folder = new CalendarFolder(c.ExchangeCalendarService);
            folder.DisplayName = "Trust" + c.HealthOrgCode + c.Title + " " + c.FirstName + " " + c.LastName + " " + c.SubscriberID;
            c.ExchangeCalendarName = folder.DisplayName;
            folder.Permissions.Add(new FolderPermission(c.ExchangeEmail, FolderPermissionLevel.Reviewer));
            folder.Save(WellKnownFolderName.Calendar);
            c.ExchangeCalendarID = folder.Id.ToString();

            // Bind to the folder
            /*Folder folderStoreInfo;
            folderStoreInfo = Folder.Bind(service, WellKnownFolderName.Calendar);
            string EwsID = folderStoreInfo.Id.UniqueId;

            // The value of folderidHex will be what we need to use for the FolderId in the xml file
            string folderidHex = GetConvertedEWSIDinHex(service, folderid, txtImpersonatedUser.Text);*/

            // Update db Record
            SqlConnection conn = new SqlConnection(MyConnString);
            try
            {
                string sql = "UPDATE Subscribers SET ExchangeCalendarID=@ExchangeCalendarID, ExchangeCalendarName=@ExchangeCalendarName WHERE SubscriberID=@SubscriberID";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@ExchangeCalendarID", c.ExchangeCalendarID);
                cmd.Parameters.AddWithValue("@ExchangeCalendarName", c.ExchangeCalendarName);
                cmd.Parameters.AddWithValue("@SubscriberID", c.SubscriberID);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }
            return isSuccess;

        }

        /*public Dictionary<string, string> GetUserSettings(AutodiscoverService autodiscoverService)
        {

            // Get the user settings.
            // Submit a request and get the settings. The response contains only the
            // settings that are requested, if they exist. 
            GetUserSettingsResponse userresponse = autodiscoverService.GetUserSettings(
                txtImpersonatedUser.Text,
                UserSettingName.UserDisplayName,
                UserSettingName.InternalMailboxServer,
                UserSettingName.UserDN
            );

            Dictionary<string, string> myUserSettings = new Dictionary<string, string>();
            // Obviously this should be cleaned up with a switch statement 
            // or something, but I was working through the problem hence the 
            // extra effort on the code
            foreach (KeyValuePair<UserSettingName, Object> usersetting in userresponse.Settings)
            {

                if (usersetting.Key.ToString() == "InternalMailboxServer")
                {
                    string[] arrResult = usersetting.Value.ToString().Split('.');

                    myUserSettings.Add("InternalMailboxServer", arrResult[0].ToString());

                }
                if (usersetting.Key.ToString() == "UserDisplayName")
                {
                    string[] arrResult = usersetting.Value.ToString().Split('.');

                    myUserSettings.Add("UserDisplayName", arrResult[0].ToString());

                }
                if (usersetting.Key.ToString() == "UserDN")
                {
                    string[] arrResult = usersetting.Value.ToString().Split('.');

                    myUserSettings.Add("UserDN", arrResult[0].ToString());

                }
            }

            return myUserSettings;
        }*/

        public String GetMailboxDN()
        {
            string result = "error";
            AutodiscoverService autodiscoverService = new AutodiscoverService();

            // The following RedirectionUrlValidationCallback required for httpsredirection.
            //autodiscoverService.RedirectionUrlValidationCallback = RedirectionUrlValidationCallback;
            //autodiscoverService.Credentials = new WebCredentials(webCredentialObject);

            // Get the user settings.
            // Submit a request and get the settings. The response contains only the
            // settings that are requested, if they exist.
            //GetUserSettingsResponse userresponse = autodiscoverService.GetUserSettings(
            //    txtImpersonatedUser.Text.ToString(),
            //    UserSettingName.UserDisplayName,
            //    UserSettingName.UserDN,
            //    );

            //foreach (KeyValuePair usersetting in userresponse.Settings)
            //{
            //
            //    if (usersetting.Key.ToString() == "UserDN")
            //    {
            //        result = usersetting.Value.ToString();
            //    }
            //}

            return result;

        }




        public void CreateSharingMessageAttachment(string folderid, string userSharing, string userSharingEntryID, string invitationMailboxID, string userSharedTo, string dataType)
        {

            XmlDocument sharedMetadataXML = new XmlDocument();

            try
            {
                // just logging stuff as well during my debugging
                using (StreamWriter w = new StreamWriter("SharingMessageMetaData.txt", false, Encoding.ASCII))
                {
                    // Create a String that contains our new sharing_metadata.xml file
                    StringBuilder metadataString = new StringBuilder("<?xml version=\"1.0\"?>");
                    metadataString.Append("<SharingMessage xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" ");
                    metadataString.Append("xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" ");
                    metadataString.Append("xmlns=\"http://schemas.microsoft.com/sharing/2008\">");
                    metadataString.Append("<DataType>" + dataType + "</DataType>");
                    metadataString.Append("<Initiator>");
                    //metadataString.Append("<Name>" + GetDisplayName(userSharing) + "</Name>");
                    metadataString.Append("<SmtpAddress>" + userSharing + "</SmtpAddress><EntryId>" + userSharingEntryID.Trim());
                    metadataString.Append("</EntryId>");
                    metadataString.Append("</Initiator>");
                    metadataString.Append("<Invitation>");
                    metadataString.Append("<Providers>");
                    metadataString.Append("<Provider Type=\"ms-exchange-internal\" TargetRecipients=\"" + userSharedTo + "\">");
                    metadataString.Append("<FolderId xmlns=\"http://schemas.microsoft.com/exchange/sharing/2008\">");
                    metadataString.Append(folderid);
                    metadataString.Append("</FolderId>");
                    metadataString.Append("<MailboxId xmlns=\"http://schemas.microsoft.com/exchange/sharing/2008\">");
                    metadataString.Append(invitationMailboxID);
                    metadataString.Append("</MailboxId>");
                    metadataString.Append("</Provider>");
                    metadataString.Append("</Providers>");
                    metadataString.Append("</Invitation>");
                    metadataString.Append("</SharingMessage>");
                    // MessageBox.Show(metadataString.ToString(), "metadataString before loading into soapEnvelope");
                    sharedMetadataXML.LoadXml(metadataString.ToString());
                    //ExchangeFolderPermissionsManager.form1.Log(metadataString.ToString(), w, "Generate XML");

                    // MessageBox.Show("returning SOAP envelope now");
                    w.Close();
                }

                string tmpPath = Application.StartupPath + "\\temp\\";
                sharedMetadataXML.Save(tmpPath + "sharing_metadata.xml");
            }
            catch (Exception eg)
            {

                MessageBox.Show("Exception:" + eg.Message.ToString(), "Error Try CreateSharedMessageInvitation()");

            }
        }
        public String GetConvertedEWSIDinHex(ExchangeService esb, String sID, String strSMTPAdd)
        {
            // Create a request to convert identifiers.
            AlternateId objAltID = new AlternateId();
            objAltID.Format = IdFormat.EwsId;
            objAltID.Mailbox = strSMTPAdd;
            objAltID.UniqueId = sID;

            //Convert  PR_ENTRYID identifier format to an EWS identifier.
            AlternateIdBase objAltIDBase = esb.ConvertId(objAltID, IdFormat.HexEntryId);
            AlternateId objAltIDResp = (AlternateId)objAltIDBase;
            return objAltIDResp.UniqueId.ToString();
        }

        /// <summary>
        /// The person sharing the Calandar
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public String GetIntiatorEntryID(HealthCalendarClass c)
        {
            String result = String.Empty;

            // Bind to EWS
            c.NHSNetCalendarService.ImpersonatedUserId =
            new ImpersonatedUserId(ConnectingIdType.SmtpAddress, c.NHSNetOrgMasterAccount);

            // Get LegacyDN Using the function above this one 
            //string sharedByLegacyDN = GetMailboxDN();
            // A conversion function from earlier
            //string legacyDNinHex = ConvertStringToHex(sharedByLegacyDN);

            

            // Note while I was debugging I logged this to a text file as well
            //using (StreamWriter w = new StreamWriter("BuildAddressBookEntryIdResult.txt"))
            //{

                //addBookEntryId.Append("00000000"); /* Flags */
                //addBookEntryId.Append("DCA740C8C042101AB4B908002B2FE182"); /* ProviderUID */
                //addBookEntryId.Append("01000000"); /* Version */
                //addBookEntryId.Append("00000000"); /* Type - 00 00 00 00  = Local Mail User */
                //addBookEntryId.Append(legacyDNinHex); /* Returns the userDN of the impersonated user */
                //addBookEntryId.Append("00"); /* terminator bit */

                //Log(addBookEntryId.ToString(), w, "GetAddBookEntryId");
                //w.Close();

            //}
            StringBuilder addBookEntryId = new StringBuilder();

            addBookEntryId.Append("00000000"); /* Flags */
            addBookEntryId.Append("DCA740C8C042101AB4B908002B2FE182"); /* ProviderUID */
            addBookEntryId.Append("01000000"); /* Version */
            addBookEntryId.Append("00000000"); /* Type - 00 00 00 00  = Local Mail User */
            //addBookEntryId.Append(legacyDNinHex); /* Returns the userDN of the impersonated user */
            addBookEntryId.Append("00"); /* terminator bit */
            var logger = NLog.LogManager.GetCurrentClassLogger();
            logger.Info(addBookEntryId.ToString());

            result = addBookEntryId.ToString();
            c.NHSNetCalendarService.ImpersonatedUserId = null;
            return result;
        }

        /// <summary>
        /// Get the Malbox of the user for whom the calandar will be shared
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public String GetInvitationMailboxId(HealthCalendarClass c)
        {
            c.NHSNetCalendarService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, c.NHSNetEmail);
            // Generate The Store Entry Id for the impersonated user
            StringBuilder MailboxIDPointer = new StringBuilder();

            // Notice again the logging for debugging the results
            using (StreamWriter w = new StreamWriter("BuildInvitationMailboxIdResult.txt"))
            {

                MailboxIDPointer.Append("00000000"); /* Flags */
                MailboxIDPointer.Append("38A1BB1005E5101AA1BB08002B2A56C2"); /* ProviderUID */
                MailboxIDPointer.Append("00"); /* Version */
                MailboxIDPointer.Append("00"); /* Flag */
                MailboxIDPointer.Append("454D534D44422E444C4C00000000"); /* DLLFileName */
                MailboxIDPointer.Append("00000000"); /* Wrapped Flags */
                MailboxIDPointer.Append("1B55FA20AA6611CD9BC800AA002FC45A"); /* WrappedProvider UID (Mailbox Store Object) */
                MailboxIDPointer.Append("0C000000"); /* Wrapped Type (Mailbox Store) */
                //MailboxIDPointer.Append(ConvertStringToHex(GetMailboxServer()).ToString()); /* ServerShortname (FQDN) */
                MailboxIDPointer.Append("00"); /* termination bit */
                //MailboxIDPointer.Append(ConvertStringToHex(GetMailboxDN()).ToString()); /* Returns the userDN of the impersonated user */
                MailboxIDPointer.Append("00"); /* terminator bit */

                //Log(MailboxIDPointer.ToString(), w, "GetInvitiationEntryID");
                w.Close();

            }
            c.NHSNetCalendarService.ImpersonatedUserId = null;
            return MailboxIDPointer.ToString();

        }

        public string ConvertStringToHex(string input)
        {
            // Take our input and break it into an array
            char[] arrInput = input.ToCharArray();
            String result = String.Empty;

            // For each set of characters
            foreach (char element in arrInput)
            {
                // My lazy way of dealing with whether or not anything had been added yet
                if (String.IsNullOrEmpty(result))
                {
                    result = String.Format("{0:X2}", Convert.ToUInt16(element)).ToString();
                }
                else
                {
                    result += String.Format("{0:X2}", Convert.ToUInt16(element)).ToString();
                }

            }
            return result.ToString();
        }

        private static byte[] HexStringToByteArray(string input)
        {

            byte[] Bytes;

            int ByteLength;

            string HexValue = "\x0\x1\x2\x3\x4\x5\x6\x7\x8\x9|||||||\xA\xB\xC\xD\xE\xF";

            ByteLength = input.Length / 2;

            Bytes = new byte[ByteLength];

            for (int x = 0, i = 0; i < input.Length; i += 2, x += 1)

            {
                Bytes[x] = (byte)(HexValue[Char.ToUpper(input[i + 0]) - '0'] << 4);
                Bytes[x] |= (byte)(HexValue[Char.ToUpper(input[i + 1]) - '0']);
            }

            return Bytes;
        }

        public bool DeleteSecondaryGoogleCalendar(HealthCalendarClass c)
        {
            bool isSuccess = false;
            if (string.IsNullOrEmpty(c.SubscriberID))
            {
                return isSuccess;
            }
            //Delete secondary google Calendar
            //If successful update selected db record with empty GoogleCalenderID
            c.GoogleCalenderService.Calendars.Delete(c.GoogleCalendarID).Execute();
            SqlConnection conn = new SqlConnection(MyConnString);
            try
            {
                string sql = "UPDATE Subscribers SET GoogleCalendarID=@GoogleCalendarID WHERE SubscriberID=@SubscriberID";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@GoogleCalendarID", "");
                cmd.Parameters.AddWithValue("@SubscriberID", c.SubscriberID);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }


            return isSuccess;
        }

        public bool DeleteNHSNetCalendar(HealthCalendarClass c)
        {
            bool isSuccess = false;
            if (string.IsNullOrEmpty(c.SubscriberID))
            {
                return isSuccess;
            }
            //Delete NHSNet
            //If successful update selected db record with empty GoogleCalenderID
            var view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            var filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.NHSNetCalendarName);
            var results = c.NHSNetCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);

            if (results.TotalCount == 1)
            {
                Folder folder = Folder.Bind(c.NHSNetCalendarService, results.Folders.Single().Id);
                folder.Delete(DeleteMode.HardDelete);
            }
            else
            {
                return isSuccess;
            }


            SqlConnection conn = new SqlConnection(MyConnString);
            try
            {
                string sql = "UPDATE Subscribers SET NHSNetCalendarID=@NHSNetCalendarID, NHSNetCalendarName=@NHSNetCalendarName WHERE SubscriberID=@SubscriberID";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@NHSNetCalendarID", "");
                cmd.Parameters.AddWithValue("@NHSNetCalendarName", "");
                cmd.Parameters.AddWithValue("@SubscriberID", c.SubscriberID);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }


            return isSuccess;
        }

        public bool DeleteExchangeCalendar(HealthCalendarClass c)
        {
            bool isSuccess = false;
            if (string.IsNullOrEmpty(c.SubscriberID))
            {
                return isSuccess;
            }
            //Delete NHSNet
            //If successful update selected db record with empty GoogleCalenderID
            var view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            var filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.ExchangeCalendarName);
            var results = c.ExchangeCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);

            if (results.TotalCount == 1)
            {
                Folder folder = Folder.Bind(c.ExchangeCalendarService, results.Folders.Single().Id);
                folder.Delete(DeleteMode.HardDelete);
            }
            else
            {
                return isSuccess;
            }


            SqlConnection conn = new SqlConnection(MyConnString);
            try
            {
                string sql = "UPDATE Subscribers SET ExchangeCalendarID=@ExchangeCalendarID, ExchangeCalendarName=@ExchangeCalendarName WHERE SubscriberID=@SubscriberID";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@ExchangeCalendarID", "");
                cmd.Parameters.AddWithValue("@ExchangeCalendarName", "");
                cmd.Parameters.AddWithValue("@SubscriberID", c.SubscriberID);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }


            return isSuccess;
        }

        public bool DeleteGoogleCalendarEvents(HealthCalendarClass c)
        {
            bool isSuccess = false;

            // Define parameters of request.
            EventsResource.ListRequest request = c.GoogleCalenderService.Events.List(c.GoogleCalendarID);
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 2500;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            Events events = request.Execute();

            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                   c.GoogleCalenderService.Events.Delete(c.GoogleCalendarID, eventItem.Id).Execute();

                }
            }

            isSuccess = true;
            return isSuccess;
        }

        public bool DeleteNHSNetCalendarEvents(HealthCalendarClass c)
        {
            bool isSuccess = false;

            var view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            var filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.NHSNetCalendarName);
            var results = c.NHSNetCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);
            if (results.TotalCount == 1)
            {            
                CalendarFolder calendar = results.Where(f => f.DisplayName == c.NHSNetCalendarName).Cast<CalendarFolder>().FirstOrDefault();
                CalendarView cView = new CalendarView(DateTime.Now.AddYears(-1), DateTime.Now.AddYears(1));
                cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Id);
                FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

                if (appointments.Items != null && appointments.Items.Count > 0)
                {
                    foreach (var eventItem in appointments.Items)
                    {
                        eventItem.Delete(DeleteMode.HardDelete);
                    }
                }
            }
            else
            {
                return isSuccess;
            }       
            isSuccess = true;
            return isSuccess;
        }

        public bool DeleteExchangeCalendarEvents(HealthCalendarClass c)
        {
            bool isSuccess = false;

            var view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            var filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.ExchangeCalendarName);
            var results = c.ExchangeCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);
            if (results.TotalCount == 1)
            {
                CalendarFolder calendar = results.Where(f => f.DisplayName == c.ExchangeCalendarName).Cast<CalendarFolder>().FirstOrDefault();
                CalendarView cView = new CalendarView(DateTime.Now.AddYears(-1), DateTime.Now.AddYears(1));
                cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Id);
                FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);

                if (appointments.Items != null && appointments.Items.Count > 0)
                {
                    foreach (var eventItem in appointments.Items)
                    {
                        eventItem.Delete(DeleteMode.HardDelete);
                    }
                }
            }
            else
            {
                return isSuccess;
            }
            isSuccess = true;
            return isSuccess;
        }

        public DataTable Select()
        {
            SqlConnection conn = new SqlConnection(MyConnString);
            DataTable dt = new DataTable();
            try
            {
                string sql = "SELECT * from Subscribers";
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                conn.Open();
                adapter.Fill(dt);

            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();

            }
            return dt;
        }

        public bool UpdateGoogle(HealthCalendarClass c)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(MyConnString);
            try
            {
                string sql = "UPDATE Subscribers SET GoogleEmail=@GoogleEmail WHERE SubscriberID=@SubscriberID";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@GoogleEmail", c.GoogleEmail);
                cmd.Parameters.AddWithValue("SubscriberID", c.SubscriberID);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }

        public bool UpdateNHSNet(HealthCalendarClass c)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(MyConnString);
            try
            {
                string sql = "UPDATE Subscribers SET NHSNetEmail=@NHSNetEmail WHERE SubscriberID=@SubscriberID";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@NHSNetEmail", c.NHSNetEmail);
                cmd.Parameters.AddWithValue("SubscriberID", c.SubscriberID);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }

        public bool UpdateExchange(HealthCalendarClass c)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(MyConnString);
            try
            {
                string sql = "UPDATE Subscribers SET ExchangeEmail=@ExchangeEmail WHERE SubscriberID=@SubscriberID";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@ExchangeEmail", c.ExchangeEmail);
                cmd.Parameters.AddWithValue("SubscriberID", c.SubscriberID);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }

        public bool IsValidEmail(string source)
        {
            return new EmailAddressAttribute().IsValid(source);
        }

        public bool CreateSampleGoogleCalendarData(HealthCalendarClass c)
        {
            bool isSuccess = false;
            DateTime today,nextMonday;
            int daysUntilMonday;
            DateTime dtEventStart, dtEventEnd;

            //Next Monday
            today = DateTime.Today;
            // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
            daysUntilMonday = ((int)DayOfWeek.Monday - (int)today.DayOfWeek + 7) % 7;
            nextMonday = today.AddDays(daysUntilMonday);

            //Add Events to Calendar from next Monday
            //Monday
            dtEventStart = nextMonday.AddMinutes(540);
            dtEventEnd = nextMonday.AddMinutes(750);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "Clinic", "Tue AM - Ortho Clinic - Slots 20, Booked 18, Avail 2", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Total Slots 20, Booked Slots 18, Available Slots 2", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(780);
            dtEventEnd = nextMonday.AddMinutes(960);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "Theatre", "Theatre	Theatre 04 List - 1 x Knee Replacement, 2 x Knee Arthroscopy", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "1 x Knee Replacement, 2 x Knee Arthroscopy", dtEventStart, dtEventEnd);
            //Tuesday
            dtEventStart = nextMonday.AddMinutes(1440 + 570);
            dtEventEnd = nextMonday.AddMinutes(1440 + 630);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "Contact", "Contact Community - M 19", "Galgate Health Centre, Highland Brow, Galgate, Lancaster, Lancashire, LA2 0NB", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 660);
            dtEventEnd = nextMonday.AddMinutes(1440 + 720);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "Contact", "Contact Community - M 68", "Garstang Medical Centre, Kepple Lane, Preston, Lancashire, PR3 1PB", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 780);
            dtEventEnd = nextMonday.AddMinutes(1440 + 840);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "Contact", "Contact Community - F 77", "Brookfield Surgery, Main Road, Bolton-Le-Sands, Carnforth, Lancashire, LA5 8DH", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 870);
            dtEventEnd = nextMonday.AddMinutes(1440 + 930);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "Contact", "Contact Community - F 88", "Owen Road Surgery, 67 Owen Road, Skerton, Lancaster, Lancashire, LA1 2LG", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 960);
            dtEventEnd = nextMonday.AddMinutes(1440 + 990);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "Review", "Review - F 88", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Meet at Trust", dtEventStart, dtEventEnd);

            //Wednesday
            dtEventStart = nextMonday.AddMinutes(2880 + 720);
            dtEventEnd = nextMonday.AddMinutes(2880 + 1230);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "Theatre", "Theatre 02 List - A&E Theatre", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Meet at Trust", dtEventStart, dtEventEnd);

            //Thursday
            dtEventStart = nextMonday.AddMinutes(4320 + 540);
            dtEventEnd = nextMonday.AddMinutes(4320 + 720);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "StudyLeave", "Diary - Study Leave", "", "Prep for Exam", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(4320 + 780);
            dtEventEnd = nextMonday.AddMinutes(4320 + 1050);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "Vacation", "Diary - Vacation", "", "", dtEventStart, dtEventEnd);

            //Friday
            dtEventStart = nextMonday.AddMinutes(5760 + 540);
            dtEventEnd = nextMonday.AddMinutes(5760 + 555);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "TCI", "TCI Ward 01 - F 55", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 555);
            dtEventEnd = nextMonday.AddMinutes(5760 + 570);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "TCI", "TCI Ward 01 - F 49", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 570);
            dtEventEnd = nextMonday.AddMinutes(5760 + 585);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "TCI", "TCI Ward 01 - F 56", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 585);
            dtEventEnd = nextMonday.AddMinutes(5760 + 600);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "TCI", "TCI Ward 01 - F 40", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 780);
            dtEventEnd = nextMonday.AddMinutes(5760 + 1020);
            AddGoogleCalenderEvent(c.GoogleCalenderService, c.GoogleCalendarID, "PreClinic", "Fri PM - PreOp Clinic Slots 10, Booked 10, Avail 0", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Slots 10, Booked 10, Avail 0", dtEventStart, dtEventEnd);

            isSuccess = true;
            return isSuccess;
        }

        public static void AddGoogleCalenderEvent(CalendarService service, String strGoogleCalendarID, String strActivityType, String strEventSummary, String strEventLocation, String strEventDescription, DateTime dtEventStart, DateTime dtEventEnd)
        {
            String strColour;

            switch (strActivityType)
            {
                case "Contact":
                    strColour = "2";
                    break;
                case "Review":
                    strColour = "10";
                    break;
                case "Vacation":
                    strColour = "7";
                    break;
                case "StudyLeave":
                    strColour = "9";
                    break;
                case "TCI":
                    strColour = "4";
                    break;
                case "Theatre":
                    strColour = "11";
                    break;
                case "Clinic":
                    strColour = "5";
                    break;
                case "PreClinic":
                    strColour = "6";
                    break;
                default:
                    strColour = "8";
                    break;
            }

            Event Event = new Event
            {
                Summary = strEventSummary,
                Location = strEventLocation,
                Description = strEventDescription,


                ColorId = strColour,
                Start = new EventDateTime()
                {
                    DateTime = new DateTime(dtEventStart.Year, dtEventStart.Month, dtEventStart.Day, dtEventStart.Hour, dtEventStart.Minute, dtEventStart.Second),
                    TimeZone = "Europe/London"
                },
                End = new EventDateTime()
                {
                    DateTime = new DateTime(dtEventEnd.Year, dtEventEnd.Month, dtEventEnd.Day, dtEventEnd.Hour, dtEventEnd.Minute, dtEventEnd.Second),
                    TimeZone = "Europe/London"
                },
            };
            Event ThisEvent = service.Events.Insert(Event, strGoogleCalendarID).Execute();
        }

        public bool CreateSampleNHSNetCalendarData(HealthCalendarClass c)
        {
            bool isSuccess = false;
            DateTime today, nextMonday;
            int daysUntilMonday;
            DateTime dtEventStart, dtEventEnd;

            //Next Monday
            today = DateTime.Today;
            // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
            daysUntilMonday = ((int)DayOfWeek.Monday - (int)today.DayOfWeek + 7) % 7;
            nextMonday = today.AddDays(daysUntilMonday);

            //Get the FolderID for the selected user
            var view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            var filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.NHSNetCalendarName);
            var results = c.NHSNetCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);
            //if (results.TotalCount < 1)
            //    throw new Exception("Cannot find Rejected folder");
            //if (results.TotalCount > 1)
            //    throw new Exception("Multiple Rejected folders");
            Folder folder = Folder.Bind(c.NHSNetCalendarService, results.Folders.Single().Id);

            //Add Events to Calendar from next Monday
            //Monday
            dtEventStart = nextMonday.AddMinutes(540);
            dtEventEnd = nextMonday.AddMinutes(750);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "Clinic", "Tue AM - Ortho Clinic - Slots 20, Booked 18, Avail 2", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Total Slots 20, Booked Slots 18, Available Slots 2", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(780);
            dtEventEnd = nextMonday.AddMinutes(960);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "Theatre", "Theatre	Theatre 04 List - 1 x Knee Replacement, 2 x Knee Arthroscopy", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "1 x Knee Replacement, 2 x Knee Arthroscopy", dtEventStart, dtEventEnd);
            //Tuesday
            dtEventStart = nextMonday.AddMinutes(1440 + 570);
            dtEventEnd = nextMonday.AddMinutes(1440 + 630);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "Contact", "Contact Community - M 19", "Galgate Health Centre, Highland Brow, Galgate, Lancaster, Lancashire, LA2 0NB", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 660);
            dtEventEnd = nextMonday.AddMinutes(1440 + 720);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "Contact", "Contact Community - M 68", "Garstang Medical Centre, Kepple Lane, Preston, Lancashire, PR3 1PB", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 780);
            dtEventEnd = nextMonday.AddMinutes(1440 + 840);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "Contact", "Contact Community - F 77", "Brookfield Surgery, Main Road, Bolton-Le-Sands, Carnforth, Lancashire, LA5 8DH", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 870);
            dtEventEnd = nextMonday.AddMinutes(1440 + 930);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "Contact", "Contact Community - F 88", "Owen Road Surgery, 67 Owen Road, Skerton, Lancaster, Lancashire, LA1 2LG", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 960);
            dtEventEnd = nextMonday.AddMinutes(1440 + 990);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "Review", "Review - F 88", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Meet at Trust", dtEventStart, dtEventEnd);

            //Wednesday
            dtEventStart = nextMonday.AddMinutes(2880 + 720);
            dtEventEnd = nextMonday.AddMinutes(2880 + 1230);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "Theatre", "Theatre 02 List - A&E Theatre", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Meet at Trust", dtEventStart, dtEventEnd);

            //Thursday
            dtEventStart = nextMonday.AddMinutes(4320 + 540);
            dtEventEnd = nextMonday.AddMinutes(4320 + 720);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "StudyLeave", "Diary - Study Leave", "", "Prep for Exam", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(4320 + 780);
            dtEventEnd = nextMonday.AddMinutes(4320 + 1050);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "Vacation", "Diary - Vacation", "", "", dtEventStart, dtEventEnd);

            //Friday
            dtEventStart = nextMonday.AddMinutes(5760 + 540);
            dtEventEnd = nextMonday.AddMinutes(5760 + 555);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "TCI", "TCI Ward 01 - F 55", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 555);
            dtEventEnd = nextMonday.AddMinutes(5760 + 570);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "TCI", "TCI Ward 01 - F 49", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 570);
            dtEventEnd = nextMonday.AddMinutes(5760 + 585);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "TCI", "TCI Ward 01 - F 56", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 585);
            dtEventEnd = nextMonday.AddMinutes(5760 + 600);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "TCI", "TCI Ward 01 - F 40", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 780);
            dtEventEnd = nextMonday.AddMinutes(5760 + 1020);
            AddNHSNetCalenderEvent(c.NHSNetCalendarService, folder, "PreClinic", "Fri PM - PreOp Clinic Slots 10, Booked 10, Avail 0", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Slots 10, Booked 10, Avail 0", dtEventStart, dtEventEnd);

            isSuccess = true;
            return isSuccess;
        }
        //private static ExtendedPropertyDefinition AppointmentColorProperty = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Appointment, 0x8214, MapiPropertyType.Integer);

        public static void AddNHSNetCalenderEvent(ExchangeService service, Folder folder, String strActivityType, String strEventSummary, String strEventLocation, String strEventDescription, DateTime dtEventStart, DateTime dtEventEnd)
        {
            String strColour;

            switch (strActivityType)
            {
                case "Contact":
                    strColour = "2";
                    break;
                case "Review":
                    strColour = "10";
                    break;
                case "Vacation":
                    strColour = "7";
                    break;
                case "StudyLeave":
                    strColour = "9";
                    break;
                case "TCI":
                    strColour = "4";
                    break;
                case "Theatre":
                    strColour = "11";
                    break;
                case "Clinic":
                    strColour = "5";
                    break;
                case "PreClinic":
                    strColour = "6";
                    break;
                default:
                    strColour = "8";
                    break;
            }


            Appointment appointment = new Appointment(service);
            // Set the properties on the appointment object to create the appointment.
            appointment.Subject = strEventSummary;
            appointment.Location = strEventLocation;
            appointment.Body = strEventDescription;

            appointment.Start = new DateTime(dtEventStart.Year, dtEventStart.Month, dtEventStart.Day, dtEventStart.Hour, dtEventStart.Minute, dtEventStart.Second);
            appointment.End = new DateTime(dtEventEnd.Year, dtEventEnd.Month, dtEventEnd.Day, dtEventEnd.Hour, dtEventEnd.Minute, dtEventEnd.Second);
            //appointment.Categories. = CategoryColor.DarkMaroon;
            TimeZoneInfo GMTTZ = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");
            appointment.StartTimeZone = GMTTZ;
            appointment.EndTimeZone = GMTTZ;
            appointment.IsReminderSet = false;
            //appointment.SetExtendedProperty(AppointmentColorProperty, strColour);

            appointment.Save(folder.Id, SendInvitationsMode.SendToNone);
            // Verify that the appointment was created by using the appointment's item ID.
            Item item = Item.Bind(service, appointment.Id, new PropertySet(ItemSchema.Subject));
        }
        public bool CreateSampleExchangeCalendarData(HealthCalendarClass c)
        {
            bool isSuccess = false;
            DateTime today, nextMonday;
            int daysUntilMonday;
            DateTime dtEventStart, dtEventEnd;

            //Next Monday
            today = DateTime.Today;
            // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
            daysUntilMonday = ((int)DayOfWeek.Monday - (int)today.DayOfWeek + 7) % 7;
            nextMonday = today.AddDays(daysUntilMonday);

            //Get the FolderID for the selected user
            var view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            var filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.ExchangeCalendarName);
            var results = c.ExchangeCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);
            //if (results.TotalCount < 1)
            //    throw new Exception("Cannot find Rejected folder");
            //if (results.TotalCount > 1)
            //    throw new Exception("Multiple Rejected folders");
            Folder folder = Folder.Bind(c.ExchangeCalendarService, results.Folders.Single().Id);

            //Add Events to Calendar from next Monday
            //Monday
            dtEventStart = nextMonday.AddMinutes(540);
            dtEventEnd = nextMonday.AddMinutes(750);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "Clinic", "Tue AM - Ortho Clinic - Slots 20, Booked 18, Avail 2", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Total Slots 20, Booked Slots 18, Available Slots 2", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(780);
            dtEventEnd = nextMonday.AddMinutes(960);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "Theatre", "Theatre	Theatre 04 List - 1 x Knee Replacement, 2 x Knee Arthroscopy", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "1 x Knee Replacement, 2 x Knee Arthroscopy", dtEventStart, dtEventEnd);
            //Tuesday
            dtEventStart = nextMonday.AddMinutes(1440 + 570);
            dtEventEnd = nextMonday.AddMinutes(1440 + 630);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "Contact", "Contact Community - M 19", "Galgate Health Centre, Highland Brow, Galgate, Lancaster, Lancashire, LA2 0NB", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 660);
            dtEventEnd = nextMonday.AddMinutes(1440 + 720);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "Contact", "Contact Community - M 68", "Garstang Medical Centre, Kepple Lane, Preston, Lancashire, PR3 1PB", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 780);
            dtEventEnd = nextMonday.AddMinutes(1440 + 840);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "Contact", "Contact Community - F 77", "Brookfield Surgery, Main Road, Bolton-Le-Sands, Carnforth, Lancashire, LA5 8DH", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 870);
            dtEventEnd = nextMonday.AddMinutes(1440 + 930);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "Contact", "Contact Community - F 88", "Owen Road Surgery, 67 Owen Road, Skerton, Lancaster, Lancashire, LA1 2LG", "Meet at GP Surgery", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(1440 + 960);
            dtEventEnd = nextMonday.AddMinutes(1440 + 990);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "Review", "Review - F 88", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Meet at Trust", dtEventStart, dtEventEnd);

            //Wednesday
            dtEventStart = nextMonday.AddMinutes(2880 + 720);
            dtEventEnd = nextMonday.AddMinutes(2880 + 1230);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "Theatre", "Theatre 02 List - A&E Theatre", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Meet at Trust", dtEventStart, dtEventEnd);

            //Thursday
            dtEventStart = nextMonday.AddMinutes(4320 + 540);
            dtEventEnd = nextMonday.AddMinutes(4320 + 720);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "StudyLeave", "Diary - Study Leave", "", "Prep for Exam", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(4320 + 780);
            dtEventEnd = nextMonday.AddMinutes(4320 + 1050);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "Vacation", "Diary - Vacation", "", "", dtEventStart, dtEventEnd);

            //Friday
            dtEventStart = nextMonday.AddMinutes(5760 + 540);
            dtEventEnd = nextMonday.AddMinutes(5760 + 555);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "TCI", "TCI Ward 01 - F 55", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 555);
            dtEventEnd = nextMonday.AddMinutes(5760 + 570);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "TCI", "TCI Ward 01 - F 49", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 570);
            dtEventEnd = nextMonday.AddMinutes(5760 + 585);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "TCI", "TCI Ward 01 - F 56", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 585);
            dtEventEnd = nextMonday.AddMinutes(5760 + 600);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "TCI", "TCI Ward 01 - F 40", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
            dtEventStart = nextMonday.AddMinutes(5760 + 780);
            dtEventEnd = nextMonday.AddMinutes(5760 + 1020);
            AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, "PreClinic", "Fri PM - PreOp Clinic Slots 10, Booked 10, Avail 0", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Slots 10, Booked 10, Avail 0", dtEventStart, dtEventEnd);

            isSuccess = true;
            return isSuccess;
        }
        //private static ExtendedPropertyDefinition AppointmentColorProperty = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Appointment, 0x8214, MapiPropertyType.Integer);

        public static void AddExchangeCalenderEvent(ExchangeService service, Folder folder, String strActivityType, String strEventSummary, String strEventLocation, String strEventDescription, DateTime dtEventStart, DateTime dtEventEnd)
        {
            String strColour;

            switch (strActivityType)
            {
                case "Contact":
                    strColour = "2";
                    break;
                case "Review":
                    strColour = "10";
                    break;
                case "Vacation":
                    strColour = "7";
                    break;
                case "StudyLeave":
                    strColour = "9";
                    break;
                case "TCI":
                    strColour = "4";
                    break;
                case "Theatre":
                    strColour = "11";
                    break;
                case "Clinic":
                    strColour = "5";
                    break;
                case "PreClinic":
                    strColour = "6";
                    break;
                default:
                    strColour = "8";
                    break;
            }


            Appointment appointment = new Appointment(service);
            // Set the properties on the appointment object to create the appointment.
            appointment.Subject = strEventSummary;
            appointment.Location = strEventLocation;
            appointment.Body = strEventDescription;

            appointment.Start = new DateTime(dtEventStart.Year, dtEventStart.Month, dtEventStart.Day, dtEventStart.Hour, dtEventStart.Minute, dtEventStart.Second);
            appointment.End = new DateTime(dtEventEnd.Year, dtEventEnd.Month, dtEventEnd.Day, dtEventEnd.Hour, dtEventEnd.Minute, dtEventEnd.Second);
            //appointment.Categories. = CategoryColor.DarkMaroon;
            TimeZoneInfo GMTTZ = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");
            appointment.StartTimeZone = GMTTZ;
            appointment.EndTimeZone = GMTTZ;
            appointment.IsReminderSet = false;
            //appointment.SetExtendedProperty(AppointmentColorProperty, strColour);

            appointment.Save(folder.Id, SendInvitationsMode.SendToNone);
            // Verify that the appointment was created by using the appointment's item ID.
            Item item = Item.Bind(service, appointment.Id, new PropertySet(ItemSchema.Subject));
        }


        public bool GetTrustSettings(HealthCalendarClass c)
        {

            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(MyConnString);

            string sql = "SELECT [OrganisationName],[OrganisationCode],[OrgLocation],[GoogleMasterAccount],[NHSNetExchangeServer],[NHSNetCredentials],[NHSNetMasterAccount],[ExchangeExchangeServer],[ExchangeCredentials],[ExchangeMasterAccount] " +
                "FROM[HealthCalendar].[dbo].[Settings] " +
                "WHERE[HealthCalendar].[dbo].[Settings].[OrganisationCode]='RTX'";
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader readerClientID = cmd.ExecuteReader();
            while (readerClientID.Read())
            {
                c.HealthOrgName = readerClientID.GetString(0);
                c.HealthOrgCode = readerClientID.GetString(1);
                c.HealthOrgLocation = readerClientID.GetString(2);
                c.GoogleOrgMasterAccount = readerClientID.GetString(3);
                c.NHSNetExchangeServer = readerClientID.GetString(4);
                c.NHSNetOrgMasterCredentials = readerClientID.GetString(5);
                c.NHSNetOrgMasterAccount = readerClientID.GetString(6);
                c.ExchangeExchangeServer = readerClientID.GetString(7);
                c.ExchangeOrgMasterCredentials = readerClientID.GetString(8);
                c.ExchangeOrgMasterAccount = readerClientID.GetString(9);

                isSuccess = true;
            }
            readerClientID.Close();

            return isSuccess;
        }

        public bool GetGoogleClientSecret(HealthCalendarClass c)
        {

            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(MyConnString);

            string sql = "SELECT [client_id],[client_secret] " +
                "FROM[HealthCalendar].[dbo].[ClientID] " +
                "WHERE[HealthCalendar].[dbo].[ClientID].[oAuthType]='GoogleCalendar'" +
                "AND [HealthCalendar].[dbo].[ClientID].[OrganisationCode]='RTX'";
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            SqlDataReader readerClientID = cmd.ExecuteReader();
            while (readerClientID.Read())
            {
                c.GoogleClientID = readerClientID.GetString(0);
                c.GoogleClientSecret = readerClientID.GetString(1);
                isSuccess = true;
            }
            readerClientID.Close();

            return isSuccess;
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        public bool GetGoogleAuthorization(HealthCalendarClass c)
        {
            bool isSuccess = false;
            UserCredential credential;

            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
            new ClientSecrets
            {
                ClientId = c.GoogleClientID,
                ClientSecret = c.GoogleClientSecret
            },
            Scopes,
            "user",
            CancellationToken.None,
            new FileDataStore("Calendar.ListMyLibrary")).Result;
            // Create Google Calendar API service.
            c.GoogleCalenderService = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            isSuccess = true;
            return isSuccess;
        }


        public bool GetNHSNetAuthorization(HealthCalendarClass c)
        {
            bool isSuccess = false;
            c.NHSNetCalendarService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            c.NHSNetCalendarService.Credentials = new WebCredentials(c.NHSNetOrgMasterAccount, c.NHSNetOrgMasterCredentials);
            c.NHSNetCalendarService.TraceEnabled = true;
            c.NHSNetCalendarService.TraceFlags = TraceFlags.All;
            c.NHSNetCalendarService.Url = new Uri(c.NHSNetExchangeServer);

            isSuccess = true;
            return isSuccess;
        }

        public bool GetExchangeAuthorization(HealthCalendarClass c)
        {
            bool isSuccess = false;
            CertificateCallback.Initialize();
            c.ExchangeCalendarService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            c.ExchangeCalendarService.Credentials = new WebCredentials(c.ExchangeOrgMasterAccount, c.ExchangeOrgMasterCredentials);
            c.ExchangeCalendarService.TraceEnabled = true;
            c.ExchangeCalendarService.TraceFlags = TraceFlags.All;
            c.ExchangeCalendarService.Url = new Uri(c.ExchangeExchangeServer);

            isSuccess = true;
            return isSuccess;
        }


        /*public bool Delete(HealthCalendarClass c)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(MyConnString);
            try
            {
                string sql = "DELETE FROM Subscribers WHERE SubscriberID=@SubscriberID";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@SubscriberID", c.SubscriberID);
                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }*/

        /*public bool Insert(HealthCalendarClass c)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(MyConnString);
            try
            {
                string sql = "INSERT INTO Subscribers (Title, FirstName, LastName,SubscriberID, GoogleEmail, GoogleCalendarID ) VALUES ('@Title, @Firstname, @Surname, @GoogleEmail, @GoogleCalenderID)";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@Title", c.Title);
                cmd.Parameters.AddWithValue("@Firstname", c.FirstName);
                cmd.Parameters.AddWithValue("@Surname", c.LastName);
                cmd.Parameters.AddWithValue("@GoogleEmail", c.GoogleEmail);
                cmd.Parameters.AddWithValue("@GoogleCalendarID", c.GoogleCalendarID);

                conn.Open();
                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }*/
    }


}
