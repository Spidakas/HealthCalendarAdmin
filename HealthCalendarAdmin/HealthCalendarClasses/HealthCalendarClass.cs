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
using Attachment = Microsoft.Exchange.WebServices.Data.Attachment;
using System.Collections.ObjectModel;
//using NLog;

//[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace HealthCalendarClasses
{
    
    class HealthCalendarClass
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public string SubscriberID { get; set; }
        public string SubscriberOID { get; set; }
        public string Initials { get; set; }
        public string Title { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public DateTime ModifiedDTTM { get; set; }
        public DateTime EndDTTM { get; set; }
        public string MainIdentifier { get; set; }
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

        public string HealthOrgName { get; set; }
        public string HealthOrgCode { get; set; }        
        public string HealthOrgLocation { get; set; }
        public bool bGoogleEnabled { get; set; }
        public string GoogleOrgMasterAccount { get; set; }
        public string GoogleClientID { get; set; }
        public string GoogleClientSecret { get; set; }
        public bool bOutlookEnabled { get; set; }
        public bool bExchangeEnabled { get; set; }
        public string ExchangeExchangeServer { get; set; }
        public string ExchangeOrgMasterAccount { get; set; }
        public string ExchangeOrgMasterCredentials { get; set; }
        public string ExchangeDisplayName { get; set; }
        public string ExchangeFQDN { get; set; }
        public string ExchangeUserDN { get; set; }
        public bool bNHSNetEnabled { get; set; }
        public string NHSNetExchangeServer { get; set; }
        public string NHSNetOrgMasterAccount { get; set; }
        public string NHSNetOrgMasterCredentials { get; set; }
        public string NHSNetDisplayName { get; set; }
        public string NHSNetFQDN { get; set; }
        public string NHSNetUserDN { get; set; }

        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "HealthCalendar";
        public CalendarService GoogleCalenderService { get; set; }
        public ExchangeService NHSNetCalendarService { get; set; }
        public ExchangeService ExchangeCalendarService { get; set; }
        public AutodiscoverService GoogleCalenderAutoDiscoverService { get; set; }
        public AutodiscoverService NHSNetCalendarAutoDiscoverService { get; set; }
        public AutodiscoverService ExchangeCalendarAutoDiscoverService { get; set; }
        public AutodiscoverService ExchangeAutoDiscoverService { get; set; }

        public string strPRBODYHTML { get; set; }
        public string strBODY { get; set; }
        public string hexPRBODYHTML { get; set; }
        public byte[] binPRBODYHTML { get; set; }

        static string MyConnString = ConfigurationManager.ConnectionStrings["connstrng"].ConnectionString;

        public string strSharingFolderIdHex { get; set; }
        public string strInitiatorEntryID { get; set; }
        public string strInvitationMailboxID { get; set; }
        public string strOwnerSMTPAddress { get; set; }
        public string strOwnerDisplayName { get; set; }

        public Guid PropertySetSharing { get; set; }
        public byte[] byteSharingProviderGuid { get; set; }
        public byte[] binInitiatorEntryId { get; set; }

        public ExtendedPropertyDefinition PidTagNormalizedSubject { get; set; }
        public ExtendedPropertyDefinition PidLidSharingCapabilities { get; set; }
        public ExtendedPropertyDefinition PidNameXSharingCapabilities { get; set; }
        //public ExtendedPropertyDefinition PidLidSharingConfigurationUrl { get; set; }
        //public ExtendedPropertyDefinition PidNameXSharingConfigUrl { get; set; }
        public ExtendedPropertyDefinition PidLidSharingFlavor { get; set; }
        public ExtendedPropertyDefinition PidNameXSharingFlavor { get; set; }
        public ExtendedPropertyDefinition PidLidSharingInitiatorEntryId { get; set; }
        public ExtendedPropertyDefinition PidLidSharingInitiatorName { get; set; }
        public ExtendedPropertyDefinition PidLidSharingInitiatorSMTP { get; set; }
        public ExtendedPropertyDefinition PidLidSharingLocalType { get; set; }
        public ExtendedPropertyDefinition PidNameXSharingLocalType { get; set; }
        public ExtendedPropertyDefinition PidLidSharingProviderGuid { get; set; }
        public ExtendedPropertyDefinition PidNameXSharingProviderGuid { get; set; }
        public ExtendedPropertyDefinition PidLidSharingProviderName { get; set; }
        public ExtendedPropertyDefinition PidNameXSharingProviderName { get; set; }
        public ExtendedPropertyDefinition PidLidSharingProviderUrl { get; set; }
        public ExtendedPropertyDefinition PidNameXSharingProviderUrl { get; set; }
        public ExtendedPropertyDefinition PidLidSharingRemoteName { get; set; }
        public ExtendedPropertyDefinition PidNameXSharingRemoteName { get; set; }
        public ExtendedPropertyDefinition PidLidSharingRemoteStoreUid { get; set; }
        public ExtendedPropertyDefinition PidNameXSharingRemoteStoreUid { get; set; }
        public ExtendedPropertyDefinition PidLidSharingRemoteType { get; set; }
        public ExtendedPropertyDefinition PidNameXSharingRemoteType { get; set; }
        public ExtendedPropertyDefinition PidLidSharingRemoteUid { get; set; }
        public ExtendedPropertyDefinition PidNameXSharingRemoteUid { get; set; }
        public ExtendedPropertyDefinition PidNameContentClass { get; set; }
        public ExtendedPropertyDefinition PidTagMessageClass { get; set; }
        public ExtendedPropertyDefinition PidTagPriority { get; set; }
        public ExtendedPropertyDefinition PidTagSensitivity { get; set; }
        public ExtendedPropertyDefinition PR_BODY_HTML { get; set; }
        public ExtendedPropertyDefinition PR_BODY { get; set; }
        public ExtendedPropertyDefinition RecipientReassignmentProhibited { get; set; }

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
            try
            {
                AclRule createdrule = c.GoogleCalenderService.Acl.Insert(rule, c.GoogleCalendarID).Execute();
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating a Google calendar " + c.GoogleCalendarID + ". Error Message: " + ex.ToString());
                return isSuccess;
            }

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
                LogHealthCalendarError("Error when updating db record when creating a Google calendar " + c.GoogleCalendarID + ". Error Message: " + ex.ToString());
                return isSuccess;
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

            if (!ExistingNHSNetUser(c))
            {
                return isSuccess;
            }
                        
            CalendarFolder folder = new CalendarFolder(c.NHSNetCalendarService);            
            folder.DisplayName = "Trust" + c.HealthOrgCode + c.Title + " " + c.FirstName + " " + c.LastName + " " + c.SubscriberID;
            c.NHSNetCalendarName = folder.DisplayName;
            folder.Permissions.Add(new FolderPermission(c.NHSNetEmail,FolderPermissionLevel.Reviewer));
            try
            {
                folder.Save(WellKnownFolderName.Calendar);
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating NHSNet Calendar " + c.NHSNetCalendarName + ". Error Message: " + ex.ToString());
                return isSuccess;
            }

            c.NHSNetCalendarID = folder.Id.ToString();

            // Bind to the folder ????Is this the root folder or the calendar folder??
            //Folder folderStoreInfo;
            //folderStoreInfo = Folder.Bind(c.NHSNetCalendarService, WellKnownFolderName.Calendar);

            //string EwsID2 = folder.Id.UniqueId;

            // The value of folderidHex will be what we need to use for the FolderId in the xml file
            //string folderidHex = GetConvertedEWSIDinHex(c.NHSNetCalendarService, EwsID2, c.NHSNetOrgMasterAccount);


            //string tmpPath = Application.StartupPath + "\\temp\\";
            //System.IO.Directory.CreateDirectory(tmpPath); // Will create folder if it does not exist otherwise this does nothing
            //c.strSharingFolderIdHex = GetConvertedEWSIDinHex(c.NHSNetCalendarService, EwsID2, c.NHSNetOrgMasterAccount);
            //c.strInitiatorEntryID = GetIntiatorEntryID(c,2);
            //c.strInvitationMailboxID = GetInvitationMailboxId(c,2);
            //c.strOwnerSMTPAddress = c.NHSNetOrgMasterAccount;
            //c.strOwnerDisplayName = c.NHSNetDisplayName;

            //c.CreateSharingMessageAttachment(c.NHSNetOrgMasterAccount, c.NHSNetEmail,"calendar",c,2);

            // This is where I need the binary value of that initiator ID we talked about
            //c.binInitiatorEntryId = HexStringToByteArray(c.strInitiatorEntryID);

            //SetupExtendedPropertyDefinition(c);
            //SetCalendarSharingMessageBody(c);

            // Create a new message
            //EmailMessage invitationRequest = new EmailMessage(c.NHSNetCalendarService);
            //invitationRequest.Subject = "This is your Lorenzo activity which is being shared with you";
            //invitationRequest.Body = "Health Calendar by Loch Roag Limited has automatically sent you this calendar sharing massage";
            //invitationRequest.From = c.NHSNetOrgMasterAccount;
            //invitationRequest.Culture = "en-GB";
            //invitationRequest.Sensitivity = Sensitivity.Normal;
            //invitationRequest.Sender = c.NHSNetOrgMasterAccount;

            // Set a sharing specific property on the message
            //invitationRequest.ItemClass = "IPM.Sharing"; /* Constant Required Value [MS-ProtocolSpec] */

            //c.SetExtendedProperties(c, invitationRequest, 2);


            // Add a file attachment by using a stream
            // We need to do the following in order to prevent 3 extra bytes from being prepended to the attachment
            //string sharMetadata = File.ReadAllText(tmpPath + "sharing_metadata.xml", Encoding.ASCII);
            //byte[] fileContents;
            //UTF8Encoding encoding = new System.Text.UTF8Encoding();
            //fileContents = encoding.GetBytes(sharMetadata);

            // fileContents is a Stream object that represents the content of the file to attach.
            //invitationRequest.Attachments.AddFileAttachment("sharing_metadata.xml", fileContents);

            // This is where we set those "special" headers and other pertinent
            // information I noted in Part 1 of this series...
            //Attachment thisAttachment = invitationRequest.Attachments[0];
            //thisAttachment.ContentType = "application/x-sharing-metadata-xml";
            //thisAttachment.Name = "sharing_metadata.xml";
            //thisAttachment.IsInline = false;

            // Add recipient info and send message
            //invitationRequest.ToRecipients.Add(c.NHSNetEmail);

            //try
            //{
            //    invitationRequest.SendAndSaveCopy();
            //}
            //catch (Exception ex)
            //{
            //    LogHealthCalendarError("Error when sending NHSNet Calendar sharing invite " + c.NHSNetCalendarName + ". Error Message: " + ex.ToString());
            //    return isSuccess;
            //}
            //invitationRequest.Send();


            // Create a new message
            EmailMessage invitationRequest = new EmailMessage(c.NHSNetCalendarService);
            invitationRequest.Subject = "Your activity calendar is now available for you.";
            invitationRequest.Body = "Health Calendar by Loch Roag Limited has automatically given you permission to your activity calendar";
            invitationRequest.From = c.NHSNetOrgMasterAccount;
            invitationRequest.Culture = "en-GB";
            invitationRequest.Sensitivity = Sensitivity.Normal;
            invitationRequest.Sender = c.NHSNetOrgMasterAccount;
            invitationRequest.ToRecipients.Add(c.NHSNetEmail);
            try
            {
                invitationRequest.Send();
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when sending NHSNet Calendar sharing invite " + c.NHSNetCalendarName + ". Error Message: " + ex.ToString());
                return isSuccess;
            }



            //invitationRequest.Send();


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
                LogHealthCalendarError("Error when updating db record when creating an NHSNet calendar " + c.NHSNetCalendarName + ". Error Message: " + ex.ToString());
                return isSuccess;
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

            // check to see if user exists on the exchange server
            if (!ExistingExchangeUser(c))
            {
                return isSuccess;
            }

            CalendarFolder folder = new CalendarFolder(c.ExchangeCalendarService);
            folder.DisplayName = "Trust" + c.HealthOrgCode + c.Title + " " + c.FirstName + " " + c.LastName + " " + c.SubscriberID;
            c.ExchangeCalendarName = folder.DisplayName;
            folder.Permissions.Add(new FolderPermission(c.ExchangeEmail, FolderPermissionLevel.Reviewer));
            try
            {
                folder.Save(WellKnownFolderName.Calendar);
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating Exchange Calendar " + c.ExchangeCalendarName + ". Error Message: " + ex.ToString());
                return isSuccess;
            }

            c.ExchangeCalendarID = folder.Id.ToString();

            string EwsID2 = folder.Id.UniqueId;
            //string tmpPath = Application.StartupPath + "\\temp\\";

            //System.IO.Directory.CreateDirectory(tmpPath); // Will create folder if it does not exist otherwise this does nothing

            string tmpPath = Application.StartupPath + "\\"; ;

            c.strSharingFolderIdHex = GetConvertedEWSIDinHex(c.ExchangeCalendarService, EwsID2, c.ExchangeOrgMasterAccount);
            c.strInitiatorEntryID = GetIntiatorEntryID(c, 1);
            c.strInvitationMailboxID = GetInvitationMailboxId(c, 1);
            c.strOwnerSMTPAddress = c.ExchangeOrgMasterAccount;
            c.strOwnerDisplayName = c.ExchangeDisplayName;

            c.CreateSharingMessageAttachment(c.ExchangeOrgMasterAccount, c.ExchangeEmail, "calendar", c, 1);

            // This is where I need the binary value of that initiator ID we talked about
            c.binInitiatorEntryId = HexStringToByteArray(c.strInitiatorEntryID);

            SetupExtendedPropertyDefinition(c);
            SetCalendarSharingMessageBody(c);

            // Create a new message
            EmailMessage invitationRequest = new EmailMessage(c.ExchangeCalendarService);
            invitationRequest.Subject = "This is your Lorenzo activity which is being shared with you";
            invitationRequest.Body = "Health Calendar by Loch Roag Limited has automatically sent you this calendar sharing massage";
            invitationRequest.From = c.ExchangeOrgMasterAccount;
            invitationRequest.Culture = "en-GB";
            invitationRequest.Sensitivity = Sensitivity.Normal;
            invitationRequest.Sender = c.ExchangeOrgMasterAccount;

            // Set a sharing specific property on the message
            invitationRequest.ItemClass = "IPM.Sharing"; /* Constant Required Value [MS-ProtocolSpec] */

            c.SetExtendedProperties(c, invitationRequest, 1);


            // Add a file attachment by using a stream
            // We need to do the following in order to prevent 3 extra bytes from being prepended to the attachment
            string sharMetadata = File.ReadAllText(tmpPath + "sharing_metadata.xml", Encoding.ASCII);
            byte[] fileContents;
            UTF8Encoding encoding = new System.Text.UTF8Encoding();
            fileContents = encoding.GetBytes(sharMetadata);

            // fileContents is a Stream object that represents the content of the file to attach.
            invitationRequest.Attachments.AddFileAttachment("sharing_metadata.xml", fileContents);

            // This is where we set those "special" headers and other pertinent
            // information I noted in Part 1 of this series...
            Attachment thisAttachment = invitationRequest.Attachments[0];
            thisAttachment.ContentType = "application/x-sharing-metadata-xml";
            thisAttachment.Name = "sharing_metadata.xml";
            thisAttachment.IsInline = false;

            // Add recipient info and send message
            invitationRequest.ToRecipients.Add(c.ExchangeEmail);
            try
            {
                invitationRequest.SendAndSaveCopy();
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when sending Exchange Calendar sharing invite " + c.ExchangeCalendarName + ". Error Message: " + ex.ToString());
                return isSuccess;
            }

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
                LogHealthCalendarError("Error when updating db record when creating an Exchange calendar " + c.ExchangeCalendarName + ". Error Message: " + ex.ToString());
                return isSuccess;
            }
            finally
            {
                conn.Close();
            }
            return isSuccess;

        }

        public void SetupExtendedPropertyDefinition(HealthCalendarClass c)
        {
            // This is the Guid of the Sharing Provider in Exchange, and it's value does not change
            //Guid binSharingProviderGuid = new Guid("{AEF00600-0000-0000-C000-000000000046}");
            c.PropertySetSharing = new Guid("{AEF00600-0000-0000-C000-000000000046}");

            // Even though I don't think setting this property is mandatory, it just seemed like the right thing to do                                   
            //byte[] byteSharingProviderGuid = binSharingProviderGuid.ToByteArray();
            c.byteSharingProviderGuid = PropertySetSharing.ToByteArray();

            // Sharing Properties (in order of reference according to protocol examples in: [MS-OXSHARE])
            // Common Message Object Properties            
            c.PidTagNormalizedSubject = new ExtendedPropertyDefinition(0x0E1D, MapiPropertyType.String); // [MS-OXSHARE] 2.2.1
            // The PidTagSubjectPrefix is a zero-length string, so I do not set it
            // ExtendedPropertyDefinition PidTagSubjectPrefix = new ExtendedPropertyDefinition(0x003D, MapiPropertyType.String);

            // Sharing Object Message Properties            
            c.PidLidSharingCapabilities = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A17, MapiPropertyType.Integer); // [MS-OXSHARE] 2.2.2.1
            c.PidNameXSharingCapabilities = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Capabilities", MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.2

            // Sections 2.3 and 2.4 are also zero-length strings, so I won't set those either
            //c.PidLidSharingConfigurationUrl = new   //ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A24, MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.3
            //c.PidNameXSharingConfigUrl = new //ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Config-Url", MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.4

            c.PidLidSharingFlavor = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A18, MapiPropertyType.Integer);  // [MS-OXSHARE] 2.2.2.5
            c.PidNameXSharingFlavor = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Flavor", MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.6
            c.PidLidSharingInitiatorEntryId = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A09, MapiPropertyType.Binary); // [MS-OXSHARE] 2.2.2.7
            c.PidLidSharingInitiatorName = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A07, MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.8
            c.PidLidSharingInitiatorSMTP = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A08, MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.9
            c.PidLidSharingLocalType = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A14, MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.10
            c.PidNameXSharingLocalType = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Local-Type", MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.11
            c.PidLidSharingProviderGuid = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A01, MapiPropertyType.Binary); // [MS-OXSHARE] 2.2.2.12
            c.PidNameXSharingProviderGuid = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Provider-Guid", MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.13
            c.PidLidSharingProviderName = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A02, MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.14
            c.PidNameXSharingProviderName = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Provider-Name", MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.15 
            c.PidLidSharingProviderUrl = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A03, MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.16
            c.PidNameXSharingProviderUrl = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Provider-URL", MapiPropertyType.String); // [MS-OXSHARE] 2.2.2.17
            c.PidLidSharingRemoteName = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A05, MapiPropertyType.String); // [MS-OXSHARE] 2.2.3.1
            c.PidNameXSharingRemoteName = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Remote-Name", MapiPropertyType.String); // [MS-OXSHARE] 2.2.3.2 
            c.PidLidSharingRemoteStoreUid = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A48, MapiPropertyType.String);// [MS-OXSHARE] 2.2.3.3
            c.PidNameXSharingRemoteStoreUid = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Remote-Store-Uid", MapiPropertyType.String); // [MS-OXSHARE] 2.2.3.4
            c.PidLidSharingRemoteType = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A1D, MapiPropertyType.String); // [MS-OXSHARE] 2.2.3.5 
            c.PidNameXSharingRemoteType = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Remote-Type", MapiPropertyType.String); // [MS-OXSHARE] 2.2.3.6 
            c.PidLidSharingRemoteUid = new ExtendedPropertyDefinition(c.PropertySetSharing, 0x8A06, MapiPropertyType.String); // [MS-OXSHARE] 2.2.3.7 
            c.PidNameXSharingRemoteUid = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "X-Sharing-Remote-Uid", MapiPropertyType.String); // [MS-OXSHARE] 2.2.3.8 

            // Additional Property Constraints            
            c.PidNameContentClass = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "Content-Class", MapiPropertyType.String); // [MS-OXSHARE] 2.2.5.1
            c.PidTagMessageClass = new ExtendedPropertyDefinition(0x001A, MapiPropertyType.String); // [MS-OXSHARE] 2.2.5.2

            // From troubleshooting I noticed I was missing
            c.PidTagPriority = new ExtendedPropertyDefinition(0x0026, MapiPropertyType.Integer);
            c.PidTagSensitivity = new ExtendedPropertyDefinition(0x0036, MapiPropertyType.Integer);
            c.PR_BODY_HTML = new ExtendedPropertyDefinition(0x1013, MapiPropertyType.Binary); //PR_BOD
            c.PR_BODY = new ExtendedPropertyDefinition(0x1000, MapiPropertyType.String);
            c.RecipientReassignmentProhibited = new ExtendedPropertyDefinition(0x002b, MapiPropertyType.Boolean);
        }

        public void SetExtendedProperties(HealthCalendarClass c, EmailMessage invitationRequest, int CalendarType)
        {
            // Section 2.2.1
            invitationRequest.SetExtendedProperty(c.PidTagNormalizedSubject, "This is your Lorenzo activity which is being shared with you"); /* Constant Required Value [MS-OXSHARE] 2.2.1 */
                                                                                                                        //invitationRequest.SetExtendedProperty(PidTagSubjectPrefix, String.Empty); /* Constant Required Value [MS-OXSHARE] 2.2.1 */
                                                                                                                        // Section 2.2.2
            invitationRequest.SetExtendedProperty(c.PidLidSharingCapabilities, 0x40290); /* value for Special Folders */
            invitationRequest.SetExtendedProperty(c.PidNameXSharingCapabilities, "40290"); /* Test representation of SharingCapabilities value */
                                                                                           //invitationRequest.SetExtendedProperty(PidLidSharingConfigurationUrl, String.Empty); /* Zero-Length String [MS-OXSHARE] 2.2.2.3 */
                                                                                           //invitationRequest.SetExtendedProperty(PidNameXSharingConfigUrl, String.Empty); /* Zero-Length String [MS-OXSHARE] 2.2.2.4 */
            invitationRequest.SetExtendedProperty(c.PidLidSharingFlavor, 20310); /* Indicates Invitation for a special folder [MS-OXSHARE] 2.2.2.5 */
            invitationRequest.SetExtendedProperty(c.PidNameXSharingFlavor, "20310"); /* Text representation of SharingFlavor value [MS-OXSHARE] 2.2.2.6 */
            invitationRequest.SetExtendedProperty(c.PidLidSharingInitiatorEntryId, c.binInitiatorEntryId); /* Value from the Initiator/EntryId value in the Sharing Message attachment .xml document */
            if (CalendarType == 1)
            {
                invitationRequest.SetExtendedProperty(c.PidLidSharingInitiatorSMTP, c.ExchangeOrgMasterAccount); /* Value from Initiator/Smtp Address in the Sharing message attachment .xml document */
            }
            if (CalendarType == 2)
            {
                invitationRequest.SetExtendedProperty(c.PidLidSharingInitiatorSMTP, c.NHSNetOrgMasterAccount); /* Value from Initiator/Smtp Address in the Sharing message attachment .xml document */
            }            
            invitationRequest.SetExtendedProperty(c.PidLidSharingInitiatorName, c.strOwnerDisplayName); /* Value from Initiator/Name Address in the Sharing message attachment .xml document */
            invitationRequest.SetExtendedProperty(c.PidLidSharingLocalType, "IPF.Appointment"); /* MUST be set to PidTagContainerClass of folder to be shared */
            invitationRequest.SetExtendedProperty(c.PidNameXSharingLocalType, "IPF.Appointment"); /* MUST be set to same value as PidLidSharingLocalType */
            invitationRequest.SetExtendedProperty(c.PidLidSharingProviderGuid, c.byteSharingProviderGuid); /* Constant Required Value [MS-OXSHARE] 2.2.2.12 */
            invitationRequest.SetExtendedProperty(c.PidNameXSharingProviderGuid, "AEF0060000000000C000000000000046"); /* Constant Required Value [MS-OXSHARE] 2.2.2.13 */
            invitationRequest.SetExtendedProperty(c.PidLidSharingProviderName, "Microsoft Exchange"); /* Constant Required Value [MS-OXSHARE] 2.2.2.14 */
            invitationRequest.SetExtendedProperty(c.PidNameXSharingProviderName, "Microsoft Exchange"); /* Constant Required Value [MS-OXSHARE] 2.2.2.15] */
            invitationRequest.SetExtendedProperty(c.PidLidSharingProviderUrl, "HTTP://www.microsoft.com/exchange"); /* Constant Required Value [MS-OXSHARE] 2.2.2.16 */
            invitationRequest.SetExtendedProperty(c.PidNameXSharingProviderUrl, "HTTP://www.microsoft.com/exchange"); /* Constant Required Value [MS-OXSHARE] 2.2.2.17 */
                                                                                                                      // Section 2.2.3
            invitationRequest.SetExtendedProperty(c.PidLidSharingRemoteName, "Calendar"); /* MUST be set to PidTagDisplayName of the folder being shared */
            invitationRequest.SetExtendedProperty(c.PidNameXSharingRemoteName, "Calendar"); /* MUST be set to same value as PidLidSharingRemoteName */
            invitationRequest.SetExtendedProperty(c.PidLidSharingRemoteStoreUid, c.strInvitationMailboxID); /* Must be set to PidTagStoreEntryId of the folder being shared */
            invitationRequest.SetExtendedProperty(c.PidNameXSharingRemoteStoreUid, c.strInvitationMailboxID); /* MUST be set to same value as PidLidSharingRemoteStoreUid */
            invitationRequest.SetExtendedProperty(c.PidLidSharingRemoteType, "IPF.Appointment"); /* Constant Required Value [MS-OXSHARE] 2.2.3.5 */
            invitationRequest.SetExtendedProperty(c.PidNameXSharingRemoteType, "IPF.Appointment"); /* Constant Required Value [MS-OXSHARE] 2.2.3.6 */
            invitationRequest.SetExtendedProperty(c.PidLidSharingRemoteUid, c.strSharingFolderIdHex); /* MUST be set to PidTagEntryId of folder being shared */
            invitationRequest.SetExtendedProperty(c.PidNameXSharingRemoteUid, c.strSharingFolderIdHex); /* Must be set to same value as PidLidSharingRemoteUid */
                                                                                                        // Section 2.2.5
            invitationRequest.SetExtendedProperty(c.PidNameContentClass, "Sharing"); /* Constant Required Value [MS-ProtocolSpec] */
            invitationRequest.SetExtendedProperty(c.PidTagMessageClass, "IPM.Sharing"); /* Constant Required Value [MS-ProtocolSpec] */


            // ********* ADDITIONAL MAPPED PROPERTIES IM FINDING AS I TROUBLESHOOT ********************** //
            //invitationRequest.SetExtendedProperty(c.PidTagPriority, 0); /* From troubleshooting I'm just trying to match up values that were missing */
            //invitationRequest.SetExtendedProperty(c.PidTagSensitivity, 0); /* From troubleshooting as well */
            //invitationRequest.SetExtendedProperty(c.PR_BODY_HTML, c.binPRBODYHTML); /* From troubleshooting OWA error pointing to serializing HTML failing */
            //invitationRequest.SetExtendedProperty(c.PR_BODY, c.strBODY);
            //invitationRequest.SetExtendedProperty(c.RecipientReassignmentProhibited, true); /* Because it seemed like a good idea */
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
        public String GetIntiatorEntryID(HealthCalendarClass c, int CalendarType)
        {
            String result = String.Empty;

            StringBuilder addBookEntryId = new StringBuilder();

            addBookEntryId.Append("00000000"); /* Flags */
            addBookEntryId.Append("DCA740C8C042101AB4B908002B2FE182"); /* ProviderUID */
            addBookEntryId.Append("01000000"); /* Version */
            addBookEntryId.Append("00000000"); /* Type - 00 00 00 00  = Local Mail User */
            if (CalendarType == 1)
            {
                addBookEntryId.Append(ConvertStringToHex(c.ExchangeUserDN)); /* Returns the userDN of the user */
            }
            if (CalendarType == 2)
            {
                addBookEntryId.Append(ConvertStringToHex(c.NHSNetUserDN)); /* Returns the userDN of the user */
            }
            addBookEntryId.Append("00"); /* terminator bit */
            //LogHealthCalendarError(addBookEntryId.ToString());
            result = addBookEntryId.ToString();
            return result;
        }

        /// <summary>
        /// Get the Malbox of the user for whom the calandar will be shared
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public String GetInvitationMailboxId(HealthCalendarClass c, int CalendarType)
        {
            // Generate The Store Entry Id for the impersonated user
            StringBuilder MailboxIDPointer = new StringBuilder();       

            MailboxIDPointer.Append("00000000"); /* Flags */
            MailboxIDPointer.Append("38A1BB1005E5101AA1BB08002B2A56C2"); /* ProviderUID */
            MailboxIDPointer.Append("00"); /* Version */
            MailboxIDPointer.Append("00"); /* Flag */
            MailboxIDPointer.Append("454D534D44422E444C4C00000000"); /* DLLFileName */
            MailboxIDPointer.Append("00000000"); /* Wrapped Flags */
            MailboxIDPointer.Append("1B55FA20AA6611CD9BC800AA002FC45A"); /* WrappedProvider UID (Mailbox Store Object) */
            MailboxIDPointer.Append("0C000000"); /* Wrapped Type (Mailbox Store) */
            if (CalendarType == 1)
            {
                MailboxIDPointer.Append(ConvertStringToHex(c.ExchangeFQDN)); /* ServerShortname (FQDN) */
            }
            if (CalendarType == 2)
            {
                MailboxIDPointer.Append(ConvertStringToHex(c.NHSNetFQDN)); /* ServerShortname (FQDN) */
            }
            MailboxIDPointer.Append("00"); /* termination bit */

            if (CalendarType == 1)
            {
                MailboxIDPointer.Append(ConvertStringToHex(c.ExchangeUserDN)); /* Returns the userDN of the user */
            }
            if (CalendarType == 2)
            {
                MailboxIDPointer.Append(ConvertStringToHex(c.NHSNetUserDN)); /* Returns the userDN of the user */
            }

            MailboxIDPointer.Append("00"); /* terminator bit */
            //LogHealthCalendarError(MailboxIDPointer.ToString());
            return MailboxIDPointer.ToString();
        }

        public void CreateSharingMessageAttachment(string userSharing, string userSharedTo, string dataType, HealthCalendarClass c, int CalendarType)
        {

            XmlDocument sharedMetadataXML = new XmlDocument();

            try
            {     
                // Create a String that contains our new sharing_metadata.xml file
                StringBuilder metadataString = new StringBuilder("<?xml version=\"1.0\"?>");
                //metadataString.Append("<SharingMessage xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" ");
                //metadataString.Append("xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" ");
                //metadataString.Append("xmlns=\"http://schemas.microsoft.com/sharing/2008\">");            
                metadataString.Append("<SharingMessage xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" ");
                metadataString.Append("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" ");
                metadataString.Append("xmlns=\"http://schemas.microsoft.com/sharing/2008\">");

                metadataString.Append("<DataType>" + dataType + "</DataType>");
                metadataString.Append("<Initiator>");
                if (CalendarType==1)
                {
                    metadataString.Append("<Name>" + c.ExchangeDisplayName + "</Name>");
                }
                if (CalendarType == 2)
                {
                    metadataString.Append("<Name>" + c.NHSNetDisplayName + "</Name>");
                }
                metadataString.Append("<SmtpAddress>" + c.strOwnerSMTPAddress + "</SmtpAddress><EntryId>" + c.strInitiatorEntryID.Trim());
                metadataString.Append("</EntryId>");
                metadataString.Append("</Initiator>");
                metadataString.Append("<Invitation>");
                if (CalendarType == 1)
                {
                    metadataString.Append("<Title>" + c.ExchangeCalendarName + "</Title>");
                }
                if (CalendarType == 2)
                {
                    metadataString.Append("<Title>" + c.NHSNetCalendarName + "</Title>");
                }
                metadataString.Append("<Providers>");

                metadataString.Append("<Provider Type=\"ms-exchange-internal\" TargetRecipients=\"" + userSharedTo + "\">");
                metadataString.Append("<FolderId xmlns=\"http://schemas.microsoft.com/exchange/sharing/2008\">");
                metadataString.Append(c.strSharingFolderIdHex);
                metadataString.Append("</FolderId>");
                metadataString.Append("<MailboxId xmlns=\"http://schemas.microsoft.com/exchange/sharing/2008\">");
                metadataString.Append(c.strInvitationMailboxID);
                metadataString.Append("</MailboxId>");
                metadataString.Append("</Provider>");
                metadataString.Append("</Providers>");
                metadataString.Append("</Invitation>");
                metadataString.Append("</SharingMessage>");
                sharedMetadataXML.LoadXml(metadataString.ToString());
                //LogHealthCalendarError(metadataString.ToString());                
                //string tmpPath = Application.StartupPath + "\\temp\\";
                string tmpPath = Application.StartupPath + "\\";

                sharedMetadataXML.Save(tmpPath + "sharing_metadata.xml");
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Exception:" + ex.Message.ToString() + "Error Try CreateSharedMessageInvitation()");    
            }
        }

        public void SetCalendarSharingMessageBody(HealthCalendarClass c)
        {
            c.strPRBODYHTML = "<html dir=\"ltr\">\r\n<head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=iso-8859-1\">\r\n<meta name=\"GENERATOR\" content=\"MSHTML 8.00.7601.17514\">\r\n<style id=\"owaParaStyle\">P {\r\n   MARGIN-TOP: 0px; MARGIN-BOTTOM: 0px \r\n}\r\n</style>\r\n</head>\r\n<body fPStyle=\"1\" ocsi=\"0\">\r\n<tt>\r\n<pre>SharedByUserDisplayName (SharedByUserSmtpAddress) has invited you to view his or her Microsoft Exchange Calendar.\r\n\r\nFor instructions on how to view shared folders on Exchange, see the following article:\r\n\r\nhttp://go.microsoft.com/fwlink/?LinkId=57561\r\n\r\n*~*~*~*~*~*~*~*~*~*\r\n\r\n</pre>\r\n</tt>\r\n<div>\r\n<div style=\"direction: ltr;font-family: Tahoma;color: #000000;font-size: 10pt;\">this is a test message</div>\r\n</div>\r\n</body>\r\n</html>\r\n";

            c.strBODY = @"
SharedByUserDisplayName (SharedByUserSmtpAddress) has invited you to view his or
her Microsoft Exchange Calendar.
 
For instructions on how to view shared folders on Exchange, see the
following article:
 
http://go.microsoft.com/fwlink/?LinkId=57561
 
*~*~*~*~*~*~*~*~*~*
 
 
test body
 
";

            // Convert these to hex and binary equivelants to assign to their relevant
            // extended properties
            c.hexPRBODYHTML = ConvertStringToHex(c.strPRBODYHTML);
            c.binPRBODYHTML = HexStringToByteArray(c.hexPRBODYHTML);

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
                LogHealthCalendarError("Error updating db record when deleting Google calendar  " + c.GoogleCalendarID + ". Error Message: " + ex.ToString());
                return isSuccess;
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
                LogHealthCalendarError("Error updating db record when deleting NHSNet calendar  " + c.NHSNetCalendarName + ". Error Message: " + ex.ToString());
                return isSuccess;
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
            //Delete Exchange
            //If successful update selected db record with empty ExchangeCalenderID
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
                LogHealthCalendarError("Error updating db record when deleting Exchange calendar  " + c.ExchangeCalendarName + ". Error Message: " + ex.ToString());
                return isSuccess;
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
                CalendarView cView = new CalendarView(DateTime.Now.AddMonths(-6), DateTime.Now.AddMonths(12));
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

        public bool BulkDeleteNHSNetCalendarEvents(HealthCalendarClass c)
        {
            bool isSuccess = false;
            FolderView view;
            SearchFilter filter;
            FindFoldersResults results;
            FindItemsResults<Appointment> DelAppointments;

            view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.NHSNetCalendarName);
            results = c.NHSNetCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);
            if (results.TotalCount >= 1)
            {
                CalendarFolder calendar = results.Where(f => f.DisplayName == c.NHSNetCalendarName).Cast<CalendarFolder>().FirstOrDefault();
                CalendarView cView = new CalendarView(DateTime.Now.AddMonths(-6), DateTime.Now.AddMonths(12));
                cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Id);

                DelAppointments = calendar.FindAppointments(cView);
                ItemView iv = new ItemView(DelAppointments.TotalCount);
                List<ItemId> idItemIds = new List<ItemId>();
                if (DelAppointments.Items != null && DelAppointments.Items.Count > 0)
                {
                    foreach (var eventItem in DelAppointments.Items)
                    {
                        idItemIds.Add(eventItem.Id);
                    }
                    c.NHSNetCalendarService.DeleteItems(idItemIds, DeleteMode.HardDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences);
                }
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
                CalendarView cView = new CalendarView(DateTime.Now.AddMonths(-6), DateTime.Now.AddMonths(12));
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


        public bool BulkDeleteExchangeCalendarEvents(HealthCalendarClass c)
        {
            bool isSuccess = false;
            FolderView view;
            SearchFilter filter;
            FindFoldersResults results;
            FindItemsResults<Appointment> DelAppointments;

            view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.ExchangeCalendarName);
            results = c.ExchangeCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);
            if (results.TotalCount == 1)
            {
                CalendarFolder calendar = results.Where(f => f.DisplayName == c.ExchangeCalendarName).Cast<CalendarFolder>().FirstOrDefault();
                CalendarView cView = new CalendarView(DateTime.Now.AddMonths(-6), DateTime.Now.AddMonths(12));
                cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Id);

                DelAppointments = calendar.FindAppointments(cView);
                if (DelAppointments.TotalCount > 0)
                {
                    ItemView iv = new ItemView(DelAppointments.TotalCount);
                    List<ItemId> idItemIds = new List<ItemId>();
                    if (DelAppointments.Items != null && DelAppointments.Items.Count > 0)
                    {
                        foreach (var eventItem in DelAppointments.Items)
                        {
                            idItemIds.Add(eventItem.Id);
                        }
                        c.ExchangeCalendarService.DeleteItems(idItemIds, DeleteMode.HardDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences);
                    }
                }
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
                LogHealthCalendarError("Error in Select Method. " + "Error Message: " + ex.ToString());
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
                LogHealthCalendarError("Error when updating Google subscriber details: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
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
                LogHealthCalendarError("Error when updating NHSNet subscriber details: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
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
                LogHealthCalendarError("Error when updating Exchange subscriber details: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
            }
            finally
            {
                conn.Close();
            }
            return isSuccess;
        }


        public bool SetNHSNetCalendarDataFromDataSource(HealthCalendarClass c)
        {
            bool isSuccess = false;
            long lCareProviderOID;
            string strEventType;
            string strTitle;
            string strLocation;
            string strDescription;
            DateTime dtEventStart;
            DateTime dtEventEnd;
            SqlDataReader readerSQLClientID;
            FolderView view;
            SearchFilter filter;
            FindFoldersResults results;
            FindItemsResults<Appointment> DelAppointments;
            Folder folder;
            Collection<Appointment> appointments;

            ExtendedPropertyDefinition AppointmentColorProperty = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Appointment, 0x8214, MapiPropertyType.Integer);

            //Read Care Provider events using uspRTXEvents stored procedure.            
            c.ExchangeEmail = "";
            c.GoogleEmail="";
            try
            {
                SqlConnection conn = new SqlConnection(MyConnString);
                string sql = "uspHealthCalendarEvents";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@CareProviderOID", c.SubscriberOID));
                cmd.Parameters.Add(new SqlParameter("@ExchangeEmail", c.ExchangeEmail));
                cmd.Parameters.Add(new SqlParameter("@NHSNetEmail", c.NHSNetEmail));
                cmd.Parameters.Add(new SqlParameter("@GoogleEmail", c.GoogleEmail));
                cmd.Parameters.Add(new SqlParameter("@MainIdentifier", c.MainIdentifier));
                cmd.Parameters.Add(new SqlParameter("@DaysHence", 100));
                cmd.CommandTimeout = 600;
                conn.Open();

                readerSQLClientID = cmd.ExecuteReader();
                if (!readerSQLClientID.HasRows)
                {
                    isSuccess = false;
                    LogHealthCalendarError("No activity found for NHSNet: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID);
                    return isSuccess;
                }
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when reading data for NHSNet: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
            }

            //Clear all Events for Current user
            view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.NHSNetCalendarName);
            results = c.NHSNetCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);
            if (results.TotalCount >= 1)
            {
                CalendarFolder calendar = results.Where(f => f.DisplayName == c.NHSNetCalendarName).Cast<CalendarFolder>().FirstOrDefault();
                CalendarView cView = new CalendarView(DateTime.Now.AddMonths(-6), DateTime.Now.AddMonths(12));
                cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Id);

                DelAppointments = calendar.FindAppointments(cView);                
                if (DelAppointments.TotalCount > 0)
                {
                    ItemView iv = new ItemView(DelAppointments.TotalCount);
                    List<ItemId> idItemIds = new List<ItemId>();
                    if (DelAppointments.Items != null && DelAppointments.Items.Count > 0)
                    {
                        foreach (var eventItem in DelAppointments.Items)
                        {
                            idItemIds.Add(eventItem.Id);
                        }
                        c.NHSNetCalendarService.DeleteItems(idItemIds, DeleteMode.HardDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences);
                    }
                }
            }

            // Build Appointment collection 
            try
            {
                folder = Folder.Bind(c.NHSNetCalendarService, results.Folders.Single().Id);
                appointments = new Collection<Appointment>();
                TimeZoneInfo GMTTZ = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");

                while (readerSQLClientID.Read())
                {
                    strEventType = "";
                    strTitle = "";
                    strLocation = "";
                    strDescription = "";
                    dtEventStart = DateTime.Today;
                    dtEventEnd = DateTime.Today;

                    lCareProviderOID = (long)readerSQLClientID.GetDecimal(0);
                    if (!readerSQLClientID.IsDBNull(1)) strEventType = readerSQLClientID.GetString(1);
                    if (!readerSQLClientID.IsDBNull(2)) strTitle = readerSQLClientID.GetString(2);

                    if (!readerSQLClientID.IsDBNull(3)) strLocation = readerSQLClientID.GetString(3);
                    if (!readerSQLClientID.IsDBNull(4)) strDescription = readerSQLClientID.GetString(4);
                    if (!readerSQLClientID.IsDBNull(5)) dtEventStart = readerSQLClientID.GetDateTime(5);
                    if (!readerSQLClientID.IsDBNull(6)) dtEventEnd = readerSQLClientID.GetDateTime(6);

                    Appointment appointment = new Appointment(c.NHSNetCalendarService);
                    // Set the properties on the appointment object to create the appointment.
                    appointment.Subject = strTitle;
                    appointment.Location = strLocation;
                    appointment.Body = strDescription;
                    appointment.Start = dtEventStart;
                    appointment.End = dtEventEnd;
                    appointment.StartTimeZone = GMTTZ;
                    appointment.EndTimeZone = GMTTZ;
                    appointment.IsReminderSet = false;
                    appointment.SetExtendedProperty(AppointmentColorProperty, MSCalendarColour(strEventType));
                    appointments.Add(appointment);
                }
                readerSQLClientID.Close();
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating list of NHSNet calendar items: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
            }


            //Bulk Write appointment collection to appropriate calendar
            try
            {
                var saveResult = c.NHSNetCalendarService.CreateItems(appointments, folder.Id, MessageDisposition.SaveOnly, SendInvitationsMode.SendToNone);
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating NHSNetcalendar items: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
            }

            isSuccess = true;
            return isSuccess;
        }



        public bool SetExchangeCalendarDataFromDataSource(HealthCalendarClass c)
        {
            bool isSuccess = false;
            long lCareProviderOID;
            string strEventType;
            string strTitle;
            string strLocation;
            string strDescription;
            DateTime dtEventStart;
            DateTime dtEventEnd;
            SqlDataReader readerSQLClientID;
            FolderView view;
            SearchFilter filter;
            FindFoldersResults results;
            FindItemsResults<Appointment> DelAppointments;
            Folder folder;
            Collection<Appointment> appointments;

            ExtendedPropertyDefinition AppointmentColorProperty = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Appointment,0x8214, MapiPropertyType.Integer);

            //Read Care Provider events using uspRTXEvents stored procedure.            
            try
            {
                c.NHSNetEmail = "";
                c.GoogleEmail = "";
                SqlConnection conn = new SqlConnection(MyConnString);
                string sql = "uspHealthCalendarEvents";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add(new SqlParameter("@CareProviderOID", c.SubscriberOID));
                cmd.Parameters.Add(new SqlParameter("@ExchangeEmail", c.ExchangeEmail));
                cmd.Parameters.Add(new SqlParameter("@NHSNetEmail", c.NHSNetEmail));
                cmd.Parameters.Add(new SqlParameter("@GoogleEmail", c.GoogleEmail));
                cmd.Parameters.Add(new SqlParameter("@MainIdentifier", c.MainIdentifier));
                cmd.Parameters.Add(new SqlParameter("@DaysHence", 100));
                cmd.CommandTimeout = 600;
                conn.Open();

                readerSQLClientID = cmd.ExecuteReader();
                if (!readerSQLClientID.HasRows)
                {
                    isSuccess = false;
                    LogHealthCalendarError("No activity found for Exchange: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID);
                    return isSuccess;
                }
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when reading data for Exchange: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
            }

            //Clear all Events for Current user
            view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.ExchangeCalendarName);
            results = c.ExchangeCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);
            if (results.TotalCount == 1)
            {
                CalendarFolder calendar = results.Where(f => f.DisplayName == c.ExchangeCalendarName).Cast<CalendarFolder>().FirstOrDefault();
                CalendarView cView = new CalendarView(DateTime.Now.AddMonths(-6), DateTime.Now.AddMonths(12));
                cView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Id);

                DelAppointments = calendar.FindAppointments(cView);
                if (DelAppointments.TotalCount > 0)
                {
                    ItemView iv = new ItemView(DelAppointments.TotalCount);
                    List<ItemId> idItemIds = new List<ItemId>();
                    if (DelAppointments.Items != null && DelAppointments.Items.Count > 0)
                    {
                        foreach (var eventItem in DelAppointments.Items)
                        {
                            idItemIds.Add(eventItem.Id);
                        }
                        c.ExchangeCalendarService.DeleteItems(idItemIds, DeleteMode.HardDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences);
                    }
                }
            }            

            // Build Appointment collection 
            try
            {
                folder = Folder.Bind(c.ExchangeCalendarService, results.Folders.Single().Id);
                appointments = new Collection<Appointment>();
                TimeZoneInfo GMTTZ = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");

                while (readerSQLClientID.Read())
                {
                    strEventType = "";               
                    strTitle = "";
                    strLocation = "";
                    strDescription = "";
                    dtEventStart = DateTime.Today;
                    dtEventEnd = DateTime.Today;

                    lCareProviderOID = (long)readerSQLClientID.GetDecimal(0);
                    if (!readerSQLClientID.IsDBNull(1)) strEventType = readerSQLClientID.GetString(1);
                    if (!readerSQLClientID.IsDBNull(2)) strTitle = readerSQLClientID.GetString(2);

                    if (!readerSQLClientID.IsDBNull(3)) strLocation = readerSQLClientID.GetString(3);
                    if (!readerSQLClientID.IsDBNull(4)) strDescription = readerSQLClientID.GetString(4);
                    if (!readerSQLClientID.IsDBNull(5)) dtEventStart = readerSQLClientID.GetDateTime(5);
                    if (!readerSQLClientID.IsDBNull(6)) dtEventEnd = readerSQLClientID.GetDateTime(6);

                    Appointment appointment = new Appointment(c.ExchangeCalendarService);
                    // Set the properties on the appointment object to create the appointment.
                    appointment.Subject = strTitle;
                    appointment.Location = strLocation;
                    appointment.Body = strDescription;
                    appointment.Start = dtEventStart;
                    appointment.End = dtEventEnd;
                    appointment.StartTimeZone = GMTTZ;
                    appointment.EndTimeZone = GMTTZ;
                    appointment.IsReminderSet = false;
                    appointment.SetExtendedProperty(AppointmentColorProperty, MSCalendarColour(strEventType));
                    appointments.Add(appointment);
                }
                readerSQLClientID.Close();
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating list of Exchange calendar items: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
            }

            //Bulk Write appointment collection to appropriate calendar
            try
            {
                var saveResult = c.ExchangeCalendarService.CreateItems(appointments, folder.Id, MessageDisposition.SaveOnly, SendInvitationsMode.SendToNone);
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating Exchange calendar items: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
            }

            isSuccess = true;
            return isSuccess;
        }

        public bool CreateSampleGoogleCalendarData(HealthCalendarClass c)
        {
            bool isSuccess = false;
            DateTime today,nextMonday;
            int daysUntilMonday;
            DateTime dtEventStart, dtEventEnd;

            try
            {
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
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating Google calendar item. " + "Error Message: " + ex.ToString());
            }
            isSuccess = true;
            return isSuccess;
        }

        public void AddGoogleCalenderEvent(CalendarService service, String strGoogleCalendarID, String strActivityType, String strEventSummary, String strEventLocation, String strEventDescription, DateTime dtEventStart, DateTime dtEventEnd)
        {
            try
            {
                Event Event = new Event
                {
                    Summary = strEventSummary,
                    Location = strEventLocation,
                    Description = strEventDescription,

                    ColorId = GoogleCalendarColour(strActivityType),
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
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating Google calendar item. " + "Error Message: " + ex.ToString());
            }

        }


        public bool CreateSampleNHSNetCalendarData(HealthCalendarClass c)
        {
            bool isSuccess = false;
            DateTime today, nextMonday;
            int daysUntilMonday;
            DateTime dtEventStart;
            DateTime dtEventEnd;
            FolderView view;
            SearchFilter filter;
            FindFoldersResults results;
            Folder folder;
            Collection<Appointment> appointments;

            appointments = new Collection<Appointment>();

            //Get the FolderID for the selected user
            view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.NHSNetCalendarName);
            results = c.NHSNetCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);
            folder = Folder.Bind(c.NHSNetCalendarService, results.Folders.Single().Id);

            try
            {
                //Next Monday
                today = DateTime.Today;
                // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
                daysUntilMonday = ((int)DayOfWeek.Monday - (int)today.DayOfWeek + 7) % 7;
                nextMonday = today.AddDays(daysUntilMonday);

                //Add Events to Calendar from next Monday

                //Monday
                Appointment appointment01 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(540);
                dtEventEnd = nextMonday.AddMinutes(750);
                SetNHSNetCalendarEvent(appointment01, "Clinic", "Tue AM - Ortho Clinic - Slots 20, Booked 18, Avail 2", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Total Slots 20, Booked Slots 18, Available Slots 2", dtEventStart, dtEventEnd);
                appointments.Add(appointment01);

                Appointment appointment02 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(780);
                dtEventEnd = nextMonday.AddMinutes(960);
                SetNHSNetCalendarEvent(appointment02, "Theatre", "Theatre	Theatre 04 List - 1 x Knee Replacement, 2 x Knee Arthroscopy", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "1 x Knee Replacement, 2 x Knee Arthroscopy", dtEventStart, dtEventEnd);
                appointments.Add(appointment02);

                //Tuesday
                Appointment appointment03 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(1440 + 570);
                dtEventEnd = nextMonday.AddMinutes(1440 + 630);
                SetNHSNetCalendarEvent(appointment03, "Contact", "Contact Community - M 19", "Galgate Health Centre, Highland Brow, Galgate, Lancaster, Lancashire, LA2 0NB", "Meet at GP Surgery", dtEventStart, dtEventEnd);
                appointments.Add(appointment03);
                Appointment appointment04 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(1440 + 660);
                dtEventEnd = nextMonday.AddMinutes(1440 + 720);
                SetNHSNetCalendarEvent(appointment04, "Contact", "Contact Community - M 76", "Galgate Health Centre, Highland Brow, Galgate, Lancaster, Lancashire, LA2 0NB", "Meet at GP Surgery", dtEventStart, dtEventEnd);
                appointments.Add(appointment04);
                Appointment appointment05 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(1440 + 780);
                dtEventEnd = nextMonday.AddMinutes(1440 + 840);
                SetNHSNetCalendarEvent(appointment05, "Contact", "Contact Community - F 77", "Brookfield Surgery, Main Road, Bolton-Le-Sands, Carnforth, Lancashire, LA5 8DH", "Meet at GP Surgery", dtEventStart, dtEventEnd);
                appointments.Add(appointment05);
                Appointment appointment06 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(1440 + 870);
                dtEventEnd = nextMonday.AddMinutes(1440 + 930);
                SetNHSNetCalendarEvent(appointment06, "Contact", "Contact Community - F 88", "Owen Road Surgery, 67 Owen Road, Skerton, Lancaster, Lancashire, LA1 2LG", "Meet at GP Surgery", dtEventStart, dtEventEnd);
                appointments.Add(appointment06);
                Appointment appointment07 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(1440 + 960);
                dtEventEnd = nextMonday.AddMinutes(1440 + 990);
                SetNHSNetCalendarEvent(appointment07, "Review", "Review - F 88", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Meet at Trust", dtEventStart, dtEventEnd);
                appointments.Add(appointment07);

                //Wednesday
                Appointment appointment08 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(2880 + 720);
                dtEventEnd = nextMonday.AddMinutes(2880 + 1230);
                SetNHSNetCalendarEvent(appointment08, "Theatre", "Theatre 02 List - A&E Theatre", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Meet at Trust", dtEventStart, dtEventEnd);
                appointments.Add(appointment08);

                //Thursday
                Appointment appointment09 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(4320 + 540);
                dtEventEnd = nextMonday.AddMinutes(4320 + 720);
                SetNHSNetCalendarEvent(appointment09, "StudyLeave", "Diary - Study Leave", "", "Prep for Exam", dtEventStart, dtEventEnd);
                appointments.Add(appointment09);
                Appointment appointment10 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(4320 + 780);
                dtEventEnd = nextMonday.AddMinutes(4320 + 1050);
                SetNHSNetCalendarEvent(appointment10, "Vacation", "Diary - Vacation", "", "", dtEventStart, dtEventEnd);
                appointments.Add(appointment10);

                //Friday
                Appointment appointment11 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(5760 + 540);
                dtEventEnd = nextMonday.AddMinutes(5760 + 555);
                SetNHSNetCalendarEvent(appointment11, "TCI", "TCI Ward 01 - F 55", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
                appointments.Add(appointment11);
                Appointment appointment12 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(5760 + 555);
                dtEventEnd = nextMonday.AddMinutes(5760 + 570);
                SetNHSNetCalendarEvent(appointment12, "TCI", "TCI Ward 01 - F 49", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
                appointments.Add(appointment12);
                Appointment appointment13 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(5760 + 570);
                dtEventEnd = nextMonday.AddMinutes(5760 + 585);
                SetNHSNetCalendarEvent(appointment13, "TCI", "TCI Ward 01 - F 56", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
                appointments.Add(appointment13);
                Appointment appointment14 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(5760 + 585);
                dtEventEnd = nextMonday.AddMinutes(5760 + 600);
                SetNHSNetCalendarEvent(appointment14, "TCI", "TCI Ward 01 - F 40", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
                appointments.Add(appointment14);
                Appointment appointment15 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(5760 + 780);
                dtEventEnd = nextMonday.AddMinutes(5760 + 1020);
                SetNHSNetCalendarEvent(appointment15, "PreClinic", "Fri PM - PreOp Clinic Slots 10, Booked 10, Avail 0", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Slots 10, Booked 10, Avail 0", dtEventStart, dtEventEnd);
                appointments.Add(appointment15);
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating NHSNet Sample calendar item. " + "Error Message: " + ex.ToString());
            }

            //Bulk Write appointment collection to appropriate calendar
            try
            {
                var saveResult = c.NHSNetCalendarService.CreateItems(appointments, folder.Id, MessageDisposition.SaveOnly, SendInvitationsMode.SendToNone);
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating NHSNet sample calendar items: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
            }

            isSuccess = true;
            return isSuccess;
        }

        public static void SetNHSNetCalendarEvent(Appointment app, String strActivityType, String strEventSummary, String strEventLocation, String strEventDescription, DateTime dtEventStart, DateTime dtEventEnd)
        {
            ExtendedPropertyDefinition AppointmentColorProperty = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Appointment, 0x8214, MapiPropertyType.Integer);

            // Set the properties on the appointment object to create the appointment.
            app.Subject = strEventSummary;
            app.Location = strEventLocation;
            app.Body = strEventDescription;

            app.Start = new DateTime(dtEventStart.Year, dtEventStart.Month, dtEventStart.Day, dtEventStart.Hour, dtEventStart.Minute, dtEventStart.Second);
            app.End = new DateTime(dtEventEnd.Year, dtEventEnd.Month, dtEventEnd.Day, dtEventEnd.Hour, dtEventEnd.Minute, dtEventEnd.Second);
            //appointment.Categories. = CategoryColor.DarkMaroon;
            TimeZoneInfo GMTTZ = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");
            app.StartTimeZone = GMTTZ;
            app.EndTimeZone = GMTTZ;
            app.IsReminderSet = false;
            app.SetExtendedProperty(AppointmentColorProperty, MSCalendarColour(strActivityType));
        }

        public bool CreateSampleExchangeCalendarData(HealthCalendarClass c)
        {
            bool isSuccess = false;
            DateTime today, nextMonday;
            int daysUntilMonday;
            DateTime dtEventStart;
            DateTime dtEventEnd;
            FolderView view;
            SearchFilter filter;
            FindFoldersResults results;
            Folder folder;
            Collection<Appointment> appointments;

            appointments = new Collection<Appointment>();

            //Get the FolderID for the selected user
            view = new FolderView(1);
            view.Traversal = FolderTraversal.Deep;
            filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, c.ExchangeCalendarName);
            results = c.ExchangeCalendarService.FindFolders(WellKnownFolderName.Root, filter, view);
            folder = Folder.Bind(c.ExchangeCalendarService, results.Folders.Single().Id);

            try
            {
                //Next Monday
                today = DateTime.Today;
                // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
                daysUntilMonday = ((int)DayOfWeek.Monday - (int)today.DayOfWeek + 7) % 7;
                nextMonday = today.AddDays(daysUntilMonday);

                //Add Events to Calendar from next Monday

                //Monday
                Appointment appointment01 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(540);
                dtEventEnd = nextMonday.AddMinutes(750);
                SetExchangeCalendarEvent (appointment01, "Clinic", "Tue AM - Ortho Clinic - Slots 20, Booked 18, Avail 2", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Total Slots 20, Booked Slots 18, Available Slots 2", dtEventStart, dtEventEnd);
                appointments.Add(appointment01);

                Appointment appointment02 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(780);
                dtEventEnd = nextMonday.AddMinutes(960);
                SetExchangeCalendarEvent(appointment02, "Theatre", "Theatre	Theatre 04 List - 1 x Knee Replacement, 2 x Knee Arthroscopy", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "1 x Knee Replacement, 2 x Knee Arthroscopy", dtEventStart, dtEventEnd);
                appointments.Add(appointment02);

                //Tuesday
                Appointment appointment03 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(1440 + 570);
                dtEventEnd = nextMonday.AddMinutes(1440 + 630);
                SetExchangeCalendarEvent(appointment03, "Contact", "Contact Community - M 19", "Galgate Health Centre, Highland Brow, Galgate, Lancaster, Lancashire, LA2 0NB", "Meet at GP Surgery", dtEventStart, dtEventEnd);
                appointments.Add(appointment03);
                Appointment appointment04 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(1440 + 660);
                dtEventEnd = nextMonday.AddMinutes(1440 + 720);
                SetExchangeCalendarEvent(appointment04, "Contact", "Contact Community - M 76", "Galgate Health Centre, Highland Brow, Galgate, Lancaster, Lancashire, LA2 0NB", "Meet at GP Surgery", dtEventStart, dtEventEnd);
                appointments.Add(appointment04);
                Appointment appointment05 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(1440 + 780);
                dtEventEnd = nextMonday.AddMinutes(1440 + 840);
                SetExchangeCalendarEvent(appointment05, "Contact", "Contact Community - F 77", "Brookfield Surgery, Main Road, Bolton-Le-Sands, Carnforth, Lancashire, LA5 8DH", "Meet at GP Surgery", dtEventStart, dtEventEnd);
                appointments.Add(appointment05);
                Appointment appointment06 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(1440 + 870);
                dtEventEnd = nextMonday.AddMinutes(1440 + 930);
                SetExchangeCalendarEvent(appointment06, "Contact", "Contact Community - F 88", "Owen Road Surgery, 67 Owen Road, Skerton, Lancaster, Lancashire, LA1 2LG", "Meet at GP Surgery", dtEventStart, dtEventEnd);
                appointments.Add(appointment06);
                Appointment appointment07 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(1440 + 960);
                dtEventEnd = nextMonday.AddMinutes(1440 + 990);
                SetExchangeCalendarEvent(appointment07, "Review", "Review - F 88", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Meet at Trust", dtEventStart, dtEventEnd);
                appointments.Add(appointment07);

                //Wednesday
                Appointment appointment08 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(2880 + 720);
                dtEventEnd = nextMonday.AddMinutes(2880 + 1230);
                SetExchangeCalendarEvent(appointment08, "Theatre", "Theatre 02 List - A&E Theatre", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Meet at Trust", dtEventStart, dtEventEnd);
                appointments.Add(appointment08);

                //Thursday
                Appointment appointment09 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(4320 + 540);
                dtEventEnd = nextMonday.AddMinutes(4320 + 720);
                SetExchangeCalendarEvent(appointment09, "StudyLeave", "Diary - Study Leave", "", "Prep for Exam", dtEventStart, dtEventEnd);
                appointments.Add(appointment09);
                Appointment appointment10 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(4320 + 780);
                dtEventEnd = nextMonday.AddMinutes(4320 + 1050);
                SetExchangeCalendarEvent(appointment10, "Vacation", "Diary - Vacation", "", "", dtEventStart, dtEventEnd);
                appointments.Add(appointment10);

                //Friday
                Appointment appointment11 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(5760 + 540);
                dtEventEnd = nextMonday.AddMinutes(5760 + 555);
                SetExchangeCalendarEvent(appointment11, "TCI", "TCI Ward 01 - F 55", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
                appointments.Add(appointment11);
                Appointment appointment12 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(5760 + 555);
                dtEventEnd = nextMonday.AddMinutes(5760 + 570);
                SetExchangeCalendarEvent(appointment12, "TCI", "TCI Ward 01 - F 49", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
                appointments.Add(appointment12);
                Appointment appointment13 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(5760 + 570);
                dtEventEnd = nextMonday.AddMinutes(5760 + 585);
                SetExchangeCalendarEvent(appointment13, "TCI", "TCI Ward 01 - F 56", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
                appointments.Add(appointment13);
                Appointment appointment14 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(5760 + 585);
                dtEventEnd = nextMonday.AddMinutes(5760 + 600);
                SetExchangeCalendarEvent(appointment14, "TCI", "TCI Ward 01 - F 40", "Furness General Hospital, Dalton Lane, Barrow-in-Furness, LA14 4LF", "", dtEventStart, dtEventEnd);
                appointments.Add(appointment14);
                Appointment appointment15 = new Appointment(c.ExchangeCalendarService);
                dtEventStart = nextMonday.AddMinutes(5760 + 780);
                dtEventEnd = nextMonday.AddMinutes(5760 + 1020);
                SetExchangeCalendarEvent(appointment15, "PreClinic", "Fri PM - PreOp Clinic Slots 10, Booked 10, Avail 0", "Lancaster Royal Infirmary, Ashton Rd, Lancaster LA1 4RP", "Slots 10, Booked 10, Avail 0", dtEventStart, dtEventEnd);
                appointments.Add(appointment15);
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating Exchange sample calendar item. " + "Error Message: " + ex.ToString());
            }


            //Bulk Write appointment collection to appropriate calendar
            try
            {
                var saveResult = c.ExchangeCalendarService.CreateItems(appointments, folder.Id, MessageDisposition.SaveOnly, SendInvitationsMode.SendToNone);
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Error when creating Exchange sample calendar items: " + c.FirstName + " " + c.LastName + " Ref: " + c.SubscriberOID + "Error Message: " + ex.ToString());
                return isSuccess;
            }


            isSuccess = true;
            return isSuccess;
        }

        public static void SetExchangeCalendarEvent(Appointment app, String strActivityType, String strEventSummary, String strEventLocation, String strEventDescription, DateTime dtEventStart, DateTime dtEventEnd)
        {
            ExtendedPropertyDefinition AppointmentColorProperty = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Appointment, 0x8214, MapiPropertyType.Integer);

                // Set the properties on the appointment object to create the appointment.
                app.Subject = strEventSummary;
                app.Location = strEventLocation;
                app.Body = strEventDescription;

                app.Start = new DateTime(dtEventStart.Year, dtEventStart.Month, dtEventStart.Day, dtEventStart.Hour, dtEventStart.Minute, dtEventStart.Second);
                app.End = new DateTime(dtEventEnd.Year, dtEventEnd.Month, dtEventEnd.Day, dtEventEnd.Hour, dtEventEnd.Minute, dtEventEnd.Second);
                //appointment.Categories. = CategoryColor.DarkMaroon;
                TimeZoneInfo GMTTZ = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");
                app.StartTimeZone = GMTTZ;
                app.EndTimeZone = GMTTZ;
                app.IsReminderSet = false;
                app.SetExtendedProperty(AppointmentColorProperty, MSCalendarColour(strActivityType));
        }


        public bool GetTrustSettings(HealthCalendarClass c)
        {
            bool isSuccess = false;
            try
            {
                using (SqlConnection conn = new SqlConnection(MyConnString))
                {
                    string sql = "SELECT [OrganisationName],[OrganisationCode],[OrgLocation],[GoogleEnabled],[GoogleMasterAccount],[OutlookEnabled],[ExchangeEnabled],[ExchangeExchangeServer],[ExchangeCredentials],[ExchangeMasterAccount],[ExchangeDisplayName],[ExchangeFQDN],[ExchangeUserDN],[NHSNetEnabled],[NHSNetExchangeServer],[NHSNetCredentials],[NHSNetMasterAccount],[NHSNetDisplayName],[NHSNetFQDN],[NHSNetUserDN] " +
                        "FROM[HealthCalendar].[dbo].[Settings] " +
                        "WHERE[HealthCalendar].[dbo].[Settings].[OrganisationCode]='RTX'";
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        SqlDataReader readerClientID = cmd.ExecuteReader();
                        while (readerClientID.Read())
                        {
                            c.HealthOrgName = readerClientID.GetString(0);
                            c.HealthOrgCode = readerClientID.GetString(1);
                            c.HealthOrgLocation = readerClientID.GetString(2);

                            c.bGoogleEnabled = readerClientID.GetBoolean(3);
                                
                            c.GoogleOrgMasterAccount = readerClientID.GetString(4);

                            c.bOutlookEnabled = readerClientID.GetBoolean(5);

                            c.bExchangeEnabled = readerClientID.GetBoolean(6);
                            c.ExchangeExchangeServer = readerClientID.GetString(7);
                            c.ExchangeOrgMasterCredentials = readerClientID.GetString(8);
                            c.ExchangeOrgMasterAccount = readerClientID.GetString(9);
                            c.ExchangeDisplayName = readerClientID.GetString(10);
                            c.ExchangeFQDN = readerClientID.GetString(11);
                            c.ExchangeUserDN = readerClientID.GetString(12);

                            c.bNHSNetEnabled = readerClientID.GetBoolean(13);
                            c.NHSNetExchangeServer = readerClientID.GetString(14);
                            c.NHSNetOrgMasterCredentials = readerClientID.GetString(15);
                            c.NHSNetOrgMasterAccount = readerClientID.GetString(16);
                            c.NHSNetDisplayName = readerClientID.GetString(17);
                            c.NHSNetFQDN = readerClientID.GetString(18);
                            c.NHSNetUserDN = readerClientID.GetString(19);


                            isSuccess = true;
                        }
                        readerClientID.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Failed to obtain trust settings. " + "Error Message: " + ex.ToString());
                return isSuccess;
            }
            return isSuccess;
        }

        public bool GetGoogleClientSecret(HealthCalendarClass c)
        {
            bool isSuccess = false;
            try
            {
                SqlConnection conn = new SqlConnection(MyConnString);

                string sql = "SELECT [client_id],[client_secret] " +
                    "FROM[HealthCalendar].[dbo].[ClientID] " +
                    "WHERE[HealthCalendar].[dbo].[ClientID].[oAuthType]='GoogleCalendar'" +
                    "AND [HealthCalendar].[dbo].[ClientID].[OrganisationCode]='RTX'";
                SqlCommand cmd = new SqlCommand(sql, conn);
                conn.Open();
                SqlDataReader readerSQLClientID = cmd.ExecuteReader();
                if (!readerSQLClientID.HasRows)
                {
                    isSuccess = false;
                    LogHealthCalendarError("Google secret not found.");
                    return isSuccess;
                }

                while (readerSQLClientID.Read())
                {
                    c.GoogleClientID = readerSQLClientID.GetString(0);
                    c.GoogleClientSecret = readerSQLClientID.GetString(1);
                    isSuccess = true;
                }
                readerSQLClientID.Close();
            }

            catch (Exception ex)
            {
                LogHealthCalendarError("Google secret not found. " + "Error Message: " + ex.ToString());
                return isSuccess;
            }

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

            try
            {
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
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Google authorisation has failed for: " + c.GoogleOrgMasterAccount + ". Error Message: " + ex.ToString());
                return isSuccess;
            }
            isSuccess = true;
            return isSuccess;
        }

        public bool GetNHSNetAuthorization(HealthCalendarClass c)
        {
            bool isSuccess = false;
            //bool isGetUserDetailsSuccess = false;

            try
            {
                c.NHSNetCalendarService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                c.NHSNetCalendarService.Credentials = new WebCredentials(c.NHSNetOrgMasterAccount, c.NHSNetOrgMasterCredentials);
                c.NHSNetCalendarService.TraceEnabled = false;
                //c.NHSNetCalendarService.TraceEnabled = true;
                //c.NHSNetCalendarService.TraceFlags = TraceFlags.All;
                c.NHSNetCalendarService.Timeout = 100000;
                c.NHSNetCalendarService.Url = new Uri(c.NHSNetExchangeServer);
                //isGetUserDetailsSuccess = GetNHSNetMasterUserDetails( c );
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("NHSNet authorisation has failed for: " + c.NHSNetOrgMasterAccount + ". Error Message: " + ex.ToString());
                return isSuccess;
            }

            isSuccess = true;
            return isSuccess;
        }

        public bool GetExchangeAuthorization(HealthCalendarClass c)
        {
            bool isSuccess = false;
            //bool isGetUserDetailsSuccess = false;
            
            CertificateCallback.Initialize();
            try
            {
                c.ExchangeCalendarService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                c.ExchangeCalendarService.UseDefaultCredentials = true;
                //c.ExchangeCalendarService.Credentials = new NetworkCredential(c.ExchangeOrgMasterAccount, c.ExchangeOrgMasterCredentials);
                //c.ExchangeCalendarService.Credentials = new WebCredentials(c.ExchangeOrgMasterAccount, c.ExchangeOrgMasterCredentials);

                c.ExchangeCalendarService.TraceEnabled = false;
                //c.ExchangeCalendarService.TraceEnabled = true;
                //c.ExchangeCalendarService.TraceFlags = TraceFlags.All;
                c.ExchangeCalendarService.Timeout=100000;
                c.ExchangeCalendarService.AutodiscoverUrl(c.ExchangeOrgMasterAccount, RedirectionUrlValidationCallback);

                //c.ExchangeCalendarService.AutodiscoverUrl(c.ExchangeOrgMasterAccount);
                //isGetUserDetailsSuccess = GetExchangeMasterUserDetails( c );
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Exchange authorisation has failed for: " + c.ExchangeOrgMasterAccount + ". Error Message: " + ex.ToString());
                return isSuccess;
            }
            
            isSuccess = true;
            return isSuccess;
        }

        /// <summary>
        /// When accessing this the GetUserSettings API this method takes around 2-3 minutes to exexute.
        /// The UserSettings are therefore taken from the HealthCalendar.Settings table
        /// At a later point the application could be extended to include functionality(e.g. dialogs) which is is used to configure these settings
        /// within the HealthCalendar.Settings table
        /// 
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public bool GetNHSNetMasterUserDetails(HealthCalendarClass c)
        {
            bool isSuccess = false;

            try
            {
                c.NHSNetCalendarAutoDiscoverService = new AutodiscoverService();
                c.NHSNetCalendarAutoDiscoverService.Credentials = new NetworkCredential(c.NHSNetOrgMasterAccount, c.NHSNetOrgMasterCredentials);
                UserSettingName[] allSettings = (UserSettingName[])Enum.GetValues(typeof(UserSettingName));
                GetUserSettingsResponse userresponse = c.NHSNetCalendarAutoDiscoverService.GetUserSettings(
                    c.NHSNetOrgMasterAccount,
                    UserSettingName.UserDisplayName,
                    UserSettingName.InternalMailboxServer,
                    UserSettingName.UserDN);
                Dictionary<string, string> myUserSettings = new Dictionary<string, string>();
                foreach (KeyValuePair<UserSettingName, Object> usersetting in userresponse.Settings)
                {
                    if (usersetting.Key.ToString() == "InternalMailboxServer")
                    {
                        string[] arrResult = usersetting.Value.ToString().Split('.');
                        c.NHSNetFQDN = arrResult[0].ToString();
                    }
                    if (usersetting.Key.ToString() == "UserDisplayName")
                    {
                        string[] arrResult = usersetting.Value.ToString().Split('.');
                        c.NHSNetDisplayName = arrResult[0].ToString();
                    }
                    if (usersetting.Key.ToString() == "UserDN")
                    {
                        string[] arrResult = usersetting.Value.ToString().Split('.');
                        c.NHSNetUserDN = arrResult[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("NHSNet master user details not found " + c.NHSNetOrgMasterAccount + ". Error Message: " + ex.ToString());
                return isSuccess;
            }
            isSuccess = true;
            return isSuccess;
        }

        public bool ExistingNHSNetUser(HealthCalendarClass c)
        {
            bool isSuccess = false;

            try
            {
                NameResolutionCollection ncCol = c.NHSNetCalendarService.ResolveName(c.NHSNetEmail, ResolveNameSearchLocation.DirectoryOnly, true);
                if (ncCol[0].Contact.DisplayName == null)
                {
                    return isSuccess;
                }
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Email: " + c.NHSNetEmail + " does not exist within NHS Mail" + " Error Message: " + ex.ToString());
                return isSuccess;
            }
            isSuccess = true;
            return isSuccess;
        }


        /// <summary>
        /// When accessing this the GetUserSettings API this method takes a long time to execute
        /// The UserSettings are therefore taken from the HealthCalendar.Settings table
        /// At a later point the application could be extended to include functionality(e.g. dialogs) which is is used to configure these settings
        /// within the HealthCalendar.Settings table
        /// 
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public bool GetExchangeMasterUserDetails(HealthCalendarClass c)
        {
            bool isSuccess = false;

            try
            {
                c.ExchangeCalendarAutoDiscoverService = new AutodiscoverService();
                c.ExchangeCalendarAutoDiscoverService.Credentials = new NetworkCredential(c.ExchangeOrgMasterAccount, c.ExchangeOrgMasterCredentials);
                UserSettingName[] allSettings = (UserSettingName[])Enum.GetValues(typeof(UserSettingName));
                GetUserSettingsResponse userresponse = c.ExchangeCalendarAutoDiscoverService.GetUserSettings(
                    c.ExchangeOrgMasterAccount,
                    UserSettingName.UserDisplayName,
                    UserSettingName.InternalMailboxServer,
                    UserSettingName.UserDN);
                Dictionary<string, string> myUserSettings = new Dictionary<string, string>();
                foreach (KeyValuePair<UserSettingName, Object> usersetting in userresponse.Settings)
                {
                    if (usersetting.Key.ToString() == "InternalMailboxServer")
                    {
                        string[] arrResult = usersetting.Value.ToString().Split('.');
                        c.ExchangeFQDN = arrResult[0].ToString();
                    }
                    if (usersetting.Key.ToString() == "UserDisplayName")
                    {
                        string[] arrResult = usersetting.Value.ToString().Split('.');
                        c.ExchangeDisplayName = arrResult[0].ToString();
                    }
                    if (usersetting.Key.ToString() == "UserDN")
                    {
                        string[] arrResult = usersetting.Value.ToString().Split('.');
                        c.ExchangeUserDN = arrResult[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Exchange master user details not found " + c.ExchangeOrgMasterAccount + ". Error Message: " + ex.ToString());
                return isSuccess;
            }
            isSuccess = true;
            return isSuccess;
        }

        public bool ExistingExchangeUser(HealthCalendarClass c)
        {
            bool isSuccess = false;

            try
            {
                NameResolutionCollection ncCol = c.ExchangeCalendarService.ResolveName(c.ExchangeEmail, ResolveNameSearchLocation.DirectoryOnly, true);
                if (ncCol[0].Contact.DisplayName == null)
                {
                    return isSuccess;
                }
            }
            catch (Exception ex)
            {
                LogHealthCalendarError("Email: " + c.ExchangeEmail + " does not exist in the local Email Exchange." + " Error Message: " + ex.ToString() );
                return isSuccess;
            }
            isSuccess = true;
            return isSuccess;
        }

        public static GetUserSettingsResponse GetUserSettings(
          AutodiscoverService service,
          string emailAddress,
          int maxHops,
          params UserSettingName[] settings)
        {
            Uri url = null;
            GetUserSettingsResponse response = null;

            for (int attempt = 0; attempt < maxHops; attempt++)
            {
                service.Url = url;
                service.EnableScpLookup = (attempt < 2);

                response = service.GetUserSettings(emailAddress, settings);

                if (response.ErrorCode == AutodiscoverErrorCode.RedirectAddress)
                {
                    url = new Uri(response.RedirectTarget);
                }
                else if (response.ErrorCode == AutodiscoverErrorCode.RedirectUrl)
                {
                    url = new Uri(response.RedirectTarget);
                }
                else
                {
                    return response;
                }
            }

            throw new Exception("No suitable Autodiscover endpoint was found.");
        }

        public bool IsValidEmail(string source)
        {
            return new EmailAddressAttribute().IsValid(source);
        }

        private static string MSCalendarColour(string input)
        {
            string strColour;

            //1-Red, 2-Dark Blue, 3-Green, 4-Grey, 5-Orange, 6-Blue, 7-Olive, 8-Purple, 9-Teal, 10-Yellow
            switch (input)
            {
                case "Contact":
                    strColour = "5";
                    break;
                case "TCI":
                    strColour = "3";
                    break;
                case "Vacation":
                    strColour = "7";
                    break;
                case "StudyLeave":
                    strColour = "9";
                    break;
                case "Day Case":
                    strColour = "2";
                    break;
                case "Theatre":
                    strColour = "1";
                    break;
                case "Clinic":
                    strColour = "10";
                    break;
                case "PreClinic":
                    strColour = "6";
                    break;
                default:
                    strColour = "8";
                    break;
            }

            return strColour;
        }

        private static string GoogleCalendarColour(string input)
        {
            string strColour;

            switch (input)
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

            return strColour;
        }




        public string ConvertStringToHex(string input)
        {
            // Take input and break it into an array
            char[] arrInput = input.ToCharArray();
            String result = String.Empty;

            // For each set of characters
            foreach (char element in arrInput)
            {
                // Dealing with whether or not anything had been added yet
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

        public void LogHealthCalendarError(string input)
        {
            log.Error(input);
            //var logger = NLog.LogManager.GetCurrentClassLogger();
            //logger.Info(input);
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
