﻿using Google.Apis.Auth.OAuth2;
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

            if (!ExistingNHSNetUser(c))
            {
                return isSuccess;
            }
                        
            CalendarFolder folder = new CalendarFolder(c.NHSNetCalendarService);            
            folder.DisplayName = "Trust" + c.HealthOrgCode + c.Title + " " + c.FirstName + " " + c.LastName + " " + c.SubscriberID;
            c.NHSNetCalendarName = folder.DisplayName;
            folder.Permissions.Add(new FolderPermission(c.NHSNetEmail,FolderPermissionLevel.Reviewer));
            folder.Save(WellKnownFolderName.Calendar);
            c.NHSNetCalendarID = folder.Id.ToString();
            // Bind to the folder ????Is this the root folder or the calendar folder??
            //Folder folderStoreInfo;
            //folderStoreInfo = Folder.Bind(c.NHSNetCalendarService, WellKnownFolderName.Calendar);

            string EwsID2 = folder.Id.UniqueId;

            // The value of folderidHex will be what we need to use for the FolderId in the xml file
            //string folderidHex = GetConvertedEWSIDinHex(c.NHSNetCalendarService, EwsID2, c.NHSNetOrgMasterAccount);


            string tmpPath = Application.StartupPath + "\\temp\\";
            //string folderid = GetSharedFolderId("Calendar");
            c.strSharingFolderIdHex = GetConvertedEWSIDinHex(c.NHSNetCalendarService, EwsID2, c.NHSNetOrgMasterAccount);
            c.strInitiatorEntryID = GetIntiatorEntryID(c,2);
            c.strInvitationMailboxID = GetInvitationMailboxId(c,2);
            c.strOwnerSMTPAddress = c.NHSNetOrgMasterAccount;
            c.strOwnerDisplayName = c.NHSNetDisplayName;

            c.CreateSharingMessageAttachment(c.NHSNetOrgMasterAccount, c.NHSNetEmail,"calendar",c,2);

            // This is where I need the binary value of that initiator ID we talked about
            c.binInitiatorEntryId = HexStringToByteArray(c.strInitiatorEntryID);

            SetupExtendedPropertyDefinition(c);
            SetCalendarSharingMessageBody(c);

            // Create a new message
            EmailMessage invitationRequest = new EmailMessage(c.NHSNetCalendarService);
            invitationRequest.Subject = "This is your Lorenzo activity which is being shared with you";
            invitationRequest.Body = "Health Calendar by Loch Roag Limited has automatically sent you this calendar sharing massage";
            invitationRequest.From = c.NHSNetOrgMasterAccount;
            invitationRequest.Culture = "en-GB";
            invitationRequest.Sensitivity = Sensitivity.Normal;
            invitationRequest.Sender = c.NHSNetOrgMasterAccount;

            // Set a sharing specific property on the message
            invitationRequest.ItemClass = "IPM.Sharing"; /* Constant Required Value [MS-ProtocolSpec] */

            c.SetExtendedProperties(c, invitationRequest, 2);

            
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
            invitationRequest.ToRecipients.Add(c.NHSNetEmail);
            invitationRequest.SendAndSaveCopy();
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
            folder.Save(WellKnownFolderName.Calendar);
            c.ExchangeCalendarID = folder.Id.ToString();

            string EwsID2 = folder.Id.UniqueId;
            string tmpPath = Application.StartupPath + "\\temp\\";
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
            invitationRequest.SendAndSaveCopy();

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

            // Bind to EWS
            //c.NHSNetCalendarService.ImpersonatedUserId =
            //new ImpersonatedUserId(ConnectingIdType.SmtpAddress, c.NHSNetOrgMasterAccount);

            // Get LegacyDN Using the function above this one 
            //string sharedByLegacyDN = GetMailboxDN();
            // A conversion function from earlier
            //string legacyDNinHex = ConvertStringToHex(sharedByLegacyDN);            

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
            //addBookEntryId.Append(legacyDNinHex); /* Returns the userDN of the impersonated user */
            addBookEntryId.Append("00"); /* terminator bit */

            //var logger = NLog.LogManager.GetCurrentClassLogger();
            //logger.Info(addBookEntryId.ToString());

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
            //if (CalendarType == 1)
            //{
            //    c.ExchangeCalendarService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, c.ExchangeEmail);
            //}
            //if (CalendarType == 2)
            //{
            //    c.NHSNetCalendarService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, c.NHSNetEmail);
            //}

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

            //var logger = NLog.LogManager.GetCurrentClassLogger();
            //logger.Info(MailboxIDPointer.ToString());

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

                //var logger = NLog.LogManager.GetCurrentClassLogger();
                //logger.Info(metadataString.ToString());                

                string tmpPath = Application.StartupPath + "\\temp\\";
                sharedMetadataXML.Save(tmpPath + "sharing_metadata.xml");
            }
            catch (Exception eg)
            {
                var logger = NLog.LogManager.GetCurrentClassLogger();
                logger.Info("Exception:" + eg.Message.ToString() + "Error Try CreateSharedMessageInvitation()");    
                MessageBox.Show("Exception:" + eg.Message.ToString(), "Error Try CreateSharedMessageInvitation()");
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


        public bool BulkDeleteExchangeCalendarEvents(HealthCalendarClass c)
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

                ItemView iv = new ItemView(10000);
                //FindItemsResults<Item> fiResults = null;

                List<ItemId> idItemIds = new List<ItemId>();

                if (appointments.Items != null && appointments.Items.Count > 0)
                {
                    foreach (var eventItem in appointments.Items)
                    {
                        idItemIds.Add(eventItem.Id);
                    }
                    c.ExchangeCalendarService.DeleteItems(idItemIds, DeleteMode.HardDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences);
                }




                //ItemView iv = new ItemView(1000);
                //FindItemsResults<Item> fiResults = null;

                //do
                //{
                //    fiResults = calendar.FindItems(WellKnownFolderName.Calendar, iv);
                //    List<ItemId> idItemIds = new List<ItemId>();
                //   foreach (Item itItem in fiResults.Items)
                //    {
                //        if (itItem is Appointment)
                //        {
                //            idItemIds.Add(itItem.Id);
                 //       }
                //    }
                //    service.DeleteItems(idItemIds, DeleteMode.SoftDelete, SendCancellationsMode.SendToNone, AffectedTaskOccurrence.AllOccurrences);
                //    iv.Offset += fiResults.Items.Count;
                //} while (fiResults.MoreAvailable == true);


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

        public bool SetCalendarDataFromDataSourceTest(HealthCalendarClass c)
        {
            bool isSuccess = false;
            bool isDeleteCalEvents = false;
            long lCareProviderOID;
            string strEventType;
            string strTitle;
            string strLocation;
            string strDescription;
            DateTime dtEventStart;
            DateTime dtEventEnd;



            //Contacts
            SqlConnection conn = new SqlConnection(MyConnString);

            string sql = "uspRTXEvents";
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@CareProviderOID", c.SubscriberOID));
            cmd.Parameters.Add(new SqlParameter("@DaysHence", 100));
            cmd.CommandTimeout = 600;
            SqlDataReader readerClientID = cmd.ExecuteReader();

            //Clear all Events for Current user
            isDeleteCalEvents = c.BulkDeleteExchangeCalendarEvents(c);

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



            while (readerClientID.Read())
            {
                strEventType = "";
                strTitle ="";
                strLocation="";
                strDescription="";
                dtEventStart = DateTime.Today;
                dtEventEnd = DateTime.Today;

                lCareProviderOID = (long)readerClientID.GetDecimal(0);
                if (!readerClientID.IsDBNull(1)) strEventType = readerClientID.GetString(1);
                if (!readerClientID.IsDBNull(2)) strTitle = readerClientID.GetString(2);

                if (!readerClientID.IsDBNull(3)) strLocation = readerClientID.GetString(3);
                if (!readerClientID.IsDBNull(4)) strDescription = readerClientID.GetString(4);
                if (!readerClientID.IsDBNull(5)) dtEventStart = readerClientID. GetDateTime(5);
                if (!readerClientID.IsDBNull(6)) dtEventEnd = readerClientID.GetDateTime(6);
                AddExchangeCalenderEvent(c.ExchangeCalendarService, folder, strEventType, strTitle, strLocation, strDescription, dtEventStart, dtEventEnd);

            }
            readerClientID.Close();

            isSuccess = true;
            return isSuccess;
        }

        public bool SetCalendarDataFromDataSource(HealthCalendarClass c)
        {
            bool isSuccess = false;
            bool isDeleteCalEvents = false;
            long lCareProviderOID;
            string strEventType;
            string strTitle;
            string strLocation;
            string strDescription;
            DateTime dtEventStart;
            DateTime dtEventEnd;



            //Contacts
            SqlConnection conn = new SqlConnection(MyConnString);

            string sql = "uspRTXEvents";
            SqlCommand cmd = new SqlCommand(sql, conn);
            conn.Open();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new SqlParameter("@CareProviderOID", c.SubscriberOID));
            cmd.Parameters.Add(new SqlParameter("@DaysHence", 100));
            cmd.CommandTimeout = 600;
            SqlDataReader readerClientID = cmd.ExecuteReader();

            //Clear all Events for Current user
            isDeleteCalEvents = c.BulkDeleteExchangeCalendarEvents(c);

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


            var appointments = new Collection<Appointment>();
            TimeZoneInfo GMTTZ = TimeZoneInfo.FindSystemTimeZoneById("GMT Standard Time");

            while (readerClientID.Read())
            {
                strEventType = "";
                strTitle = "";
                strLocation = "";
                strDescription = "";
                dtEventStart = DateTime.Today;
                dtEventEnd = DateTime.Today;

                lCareProviderOID = (long)readerClientID.GetDecimal(0);
                if (!readerClientID.IsDBNull(1)) strEventType = readerClientID.GetString(1);
                if (!readerClientID.IsDBNull(2)) strTitle = readerClientID.GetString(2);

                if (!readerClientID.IsDBNull(3)) strLocation = readerClientID.GetString(3);
                if (!readerClientID.IsDBNull(4)) strDescription = readerClientID.GetString(4);
                if (!readerClientID.IsDBNull(5)) dtEventStart = readerClientID.GetDateTime(5);
                if (!readerClientID.IsDBNull(6)) dtEventEnd = readerClientID.GetDateTime(6);

                Appointment appointment = new Appointment(c.ExchangeCalendarService);
                // Set the properties on the appointment object to create the appointment.
                appointment.Subject = strTitle;
                appointment.Location = strLocation;
                appointment.Body = strDescription;
                appointment.Start = new DateTime(dtEventStart.Year, dtEventStart.Month, dtEventStart.Day, dtEventStart.Hour, dtEventStart.Minute, dtEventStart.Second);
                appointment.End = new DateTime(dtEventEnd.Year, dtEventEnd.Month, dtEventEnd.Day, dtEventEnd.Hour, dtEventEnd.Minute, dtEventEnd.Second);                
                appointment.StartTimeZone = GMTTZ;
                appointment.EndTimeZone = GMTTZ;
                appointment.IsReminderSet = false;
                appointments.Add(appointment);
            }
            readerClientID.Close();

            var saveResult = c.ExchangeCalendarService.CreateItems(appointments, folder.Id, MessageDisposition.SaveOnly, SendInvitationsMode.SendToNone);

            isSuccess = true;
            return isSuccess;
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
            catch {
                return isSuccess;
            }
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
            bool isGetUserDetailsSuccess = false;
            try
            {
                c.NHSNetCalendarService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                c.NHSNetCalendarService.Credentials = new WebCredentials(c.NHSNetOrgMasterAccount, c.NHSNetOrgMasterCredentials);
                c.NHSNetCalendarService.TraceEnabled = true;
                c.NHSNetCalendarService.TraceFlags = TraceFlags.All;
                c.NHSNetCalendarService.Url = new Uri(c.NHSNetExchangeServer);
                //isGetUserDetailsSuccess = GetNHSNetMasterUserDetails( c );
            }
            catch
            {
                return isSuccess;
            }
            isSuccess = true;
            return isSuccess;
        }

        public bool GetExchangeAuthorization(HealthCalendarClass c)
        {
            bool isSuccess = false;
            bool isGetUserDetailsSuccess = false;
            
            CertificateCallback.Initialize();
            try
            {
                c.ExchangeCalendarService = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                c.ExchangeCalendarService.UseDefaultCredentials = true;
                c.ExchangeCalendarService.TraceEnabled = true;
                c.ExchangeCalendarService.TraceFlags = TraceFlags.All;
                c.ExchangeCalendarService.Timeout=100000;
                c.ExchangeCalendarService.AutodiscoverUrl(c.ExchangeOrgMasterAccount, RedirectionUrlValidationCallback);
                //isGetUserDetailsSuccess = GetExchangeMasterUserDetails( c );
            }
            catch
            {
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
            catch
            {
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
            catch
            {
                var logger = NLog.LogManager.GetCurrentClassLogger();
                logger.Info("Email: " + c.NHSNetEmail + " does not exist within NHS Mail");
                //MessageBox.Show("Email: " + c.NHSNetEmail + " does not exist within NHS Mail");
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
            catch
            {
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
            catch
            {
                var logger = NLog.LogManager.GetCurrentClassLogger();
                logger.Info("Email: " + c.ExchangeEmail + " does not exist in the local Email Exchange");
                //MessageBox.Show("Email: " + c.NHSNetEmail + " does not exist in the local Email Exchange");
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