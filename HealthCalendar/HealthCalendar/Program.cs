﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using HealthCalendarClasses;
//using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

[assembly: log4net.Config.XmlConfigurator(Watch=true)]

namespace HealthCalendar
{    
    class Program
    {
        static string MyConnString = ConfigurationManager.ConnectionStrings["connstrng"].ConnectionString;
        

        static void Main(string[] args)
        {


            
            //var configuration = LogManager.Configuration;
            //bool isSuccess = false;
            HealthCalendarClasses.HealthCalendarClass c = new HealthCalendarClass();
            c.LogHealthCalendarError("Starting HealthCalendar Execution.");


            //Startup

            // Trust Settings
            if (!c.GetTrustSettings(c))
            {
                c.LogHealthCalendarError("Cannot connect or read data from the Health Calendar database.");
                Environment.Exit(0);
            }

            // Google Calendar Settings
            if (c.bGoogleEnabled)
            {
                if (!c.GetGoogleClientSecret(c))
                {
                    c.LogHealthCalendarError("Unable to obtain Google Client Secret. Please contact your IT Department.");
                }

                if (!c.GetGoogleAuthorization(c))
                {
                    c.LogHealthCalendarError("Unable to obtain Google Authorization. Please contact your IT Department.");
                }
            }

            // NHSNet Calendar Settings
            if (c.bNHSNetEnabled)
            {
                if (!c.GetNHSNetAuthorization(c))
                {
                    c.LogHealthCalendarError("Unable to connect to NHS Net Server. Please contact your IT Department.");
                }
            }

            // Trust Exchange Calendar Settings
            if (c.bExchangeEnabled)
            {
                if (!c.GetExchangeAuthorization(c))
                {
                    c.LogHealthCalendarError("Unable to connect to Exchange Server. Please contact your IT Department.");
                }
            }

            //Exchange            
            if (c.bExchangeEnabled)
            {
                SetSubscribersExchangeCalendars(c);
                SetSubscribersExchangeCalendarData(c);
            }
            //NHSNet            
            if (c.bNHSNetEnabled)
            {
                SetSubscribersNHSNetCalendars(c);
                SetSubscribersNHSNetCalendarData(c);
            }

            c.LogHealthCalendarError("Ending HealthCalendar Execution.");
        }

        //// 
        //// For each subscriber with an Exchange Email and Exchange Calendar Name
        ////
        ////
        static private void SetSubscribersExchangeCalendars(HealthCalendarClass c)
        {
            SqlDataReader readerSQLClientID;
            long lSubscriberID;

            try
            {
                SqlConnection conn = new SqlConnection(MyConnString);
                string sql = "SELECT SubscriberID, ExchangeEmail, Title, Firstname, Surname FROM Subscribers WHERE (NOT (ExchangeEmail IS NULL)) AND (NOT (ExchangeEmail = '')) AND ((ExchangeCalendarName IS NULL) OR (ExchangeCalendarName = ''))";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandTimeout = 600;
                conn.Open();
                readerSQLClientID = cmd.ExecuteReader();
                if (readerSQLClientID.HasRows)
                {
                    while (readerSQLClientID.Read())
                    {
                        if (!readerSQLClientID.IsDBNull(0))
                        {
                            c.Title = "";
                            c.FirstName = "";
                            c.LastName = "";
                            lSubscriberID = (long)readerSQLClientID.GetDecimal(0);
                            c.SubscriberID = lSubscriberID.ToString();
                            c.ExchangeEmail = readerSQLClientID.GetString(1);
                            if (!readerSQLClientID.IsDBNull(2)) c.Title = readerSQLClientID.GetString(2);
                            if (!readerSQLClientID.IsDBNull(3)) c.FirstName = readerSQLClientID.GetString(3);
                            if (!readerSQLClientID.IsDBNull(4)) c.LastName = readerSQLClientID.GetString(4);
                            c.CreateShareExchangeDiary(c);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                c.LogHealthCalendarError("Error when reading Exchange records from Subscribers. " + "Error Message: " + ex.ToString());
            }
        }

        static private void SetSubscribersNHSNetCalendars(HealthCalendarClass c)
        {
            SqlDataReader readerSQLClientID;
            long lSubscriberID;

            try
            {
                SqlConnection conn = new SqlConnection(MyConnString);
                string sql = "SELECT SubscriberID, NHSNetEmail, Title, Firstname, Surname FROM Subscribers WHERE (NOT (NHSNetEmail IS NULL)) AND (NOT (NHSNetEmail = '')) AND ((NHSNetCalendarName IS NULL) OR (NHSNetCalendarName = ''))";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandTimeout = 600;
                conn.Open();
                readerSQLClientID = cmd.ExecuteReader();
                if (readerSQLClientID.HasRows)
                {
                    while (readerSQLClientID.Read())
                    {
                        if (!readerSQLClientID.IsDBNull(0))
                        {
                            c.Title = "";
                            c.FirstName = "";
                            c.LastName = "";
                            lSubscriberID = (long)readerSQLClientID.GetDecimal(0);
                            c.SubscriberID = lSubscriberID.ToString();
                            c.NHSNetEmail = readerSQLClientID.GetString(1);
                            if (!readerSQLClientID.IsDBNull(2)) c.Title = readerSQLClientID.GetString(2);
                            if (!readerSQLClientID.IsDBNull(3)) c.FirstName = readerSQLClientID.GetString(3);
                            if (!readerSQLClientID.IsDBNull(4)) c.LastName = readerSQLClientID.GetString(4);
                            c.CreateShareNHSNetDiary(c);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                c.LogHealthCalendarError("Error when reading NHSNet records from Subscribers. " + "Error Message: " + ex.ToString());
            }
        }

        //// 
        //// For each subscriber with an Exchange Email and Exchange Calendar Name
        ////
        ////
        static private void SetSubscribersExchangeCalendarData (HealthCalendarClass c)
        {
            SqlDataReader readerSQLClientID;
            long lCareProviderOID;


            try
            {
                SqlConnection conn = new SqlConnection(MyConnString);
                string sql = "SELECT SubscriberOID, ExchangeEmail, ExchangeCalendarID, ExchangeCalendarName,Title, Firstname, Surname, MainIdentifier FROM Subscribers WHERE (NOT (ExchangeEmail IS NULL)) AND (NOT (ExchangeEmail = '')) AND (NOT (ExchangeCalendarName IS NULL)) AND (NOT (ExchangeCalendarName = ''))";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandTimeout = 600;
                conn.Open();
                readerSQLClientID = cmd.ExecuteReader();
                if (readerSQLClientID.HasRows)
                {
                    while (readerSQLClientID.Read())
                    {
                        if (!readerSQLClientID.IsDBNull(0))
                        {
                            c.NHSNetEmail = "";
                            c.ExchangeEmail = "";
                            c.GoogleEmail = "";
                            c.Title = "";
                            c.FirstName = "";
                            c.LastName = "";
                            c.MainIdentifier = "";
                            lCareProviderOID = (long)readerSQLClientID.GetDecimal(0);
                            if (!readerSQLClientID.IsDBNull(1)) c.ExchangeEmail = readerSQLClientID.GetString(1);
                            c.SubscriberOID = lCareProviderOID.ToString();
                            c.ExchangeCalendarName = readerSQLClientID.GetString(3);
                            if (!readerSQLClientID.IsDBNull(4)) c.Title = readerSQLClientID.GetString(4);
                            if (!readerSQLClientID.IsDBNull(5)) c.FirstName = readerSQLClientID.GetString(5);
                            if (!readerSQLClientID.IsDBNull(6)) c.LastName = readerSQLClientID.GetString(6);
                            if (!readerSQLClientID.IsDBNull(7)) c.MainIdentifier = readerSQLClientID.GetString(7);

                            c.SetExchangeCalendarDataFromDataSource(c);
                        }
                    } 
                }
            }
            catch (Exception ex)
            {
                c.LogHealthCalendarError("Error when reading Exchange records from Subscribers. " + "Error Message: " + ex.ToString());
            }
        }

        static private void SetSubscribersNHSNetCalendarData(HealthCalendarClass c)
        {
            SqlDataReader readerSQLClientID;
            long lCareProviderOID;

            try
            {
                SqlConnection conn = new SqlConnection(MyConnString);
                string sql = "SELECT SubscriberOID, NHSNetEmail, NHSNetCalendarID, NHSNetCalendarName, Title, Firstname, Surname, MainIdentifier FROM Subscribers WHERE (NOT (NHSNetEmail IS NULL)) AND (NOT (NHSNetEmail = '')) AND (NOT (NHSNetCalendarName IS NULL)) AND (NOT (NHSNetCalendarName = ''))";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandTimeout = 600;
                conn.Open();
                readerSQLClientID = cmd.ExecuteReader();
                if (readerSQLClientID.HasRows)
                {
                    while (readerSQLClientID.Read())
                    {
                        if (!readerSQLClientID.IsDBNull(0))
                        {
                            c.NHSNetEmail = "";
                            c.ExchangeEmail = "";
                            c.GoogleEmail = "";
                            c.Title = "";
                            c.FirstName = "";
                            c.LastName = "";
                            c.MainIdentifier = "";
                            lCareProviderOID = (long)readerSQLClientID.GetDecimal(0);
                            if (!readerSQLClientID.IsDBNull(1)) c.NHSNetEmail = readerSQLClientID.GetString(1);
                            c.SubscriberOID = lCareProviderOID.ToString();
                            c.NHSNetCalendarName = readerSQLClientID.GetString(3);
                            if (!readerSQLClientID.IsDBNull(4)) c.Title = readerSQLClientID.GetString(4);
                            if (!readerSQLClientID.IsDBNull(5)) c.FirstName = readerSQLClientID.GetString(5);
                            if (!readerSQLClientID.IsDBNull(6)) c.LastName = readerSQLClientID.GetString(6);
                            if (!readerSQLClientID.IsDBNull(7)) c.MainIdentifier = readerSQLClientID.GetString(7);

                            c.SetNHSNetCalendarDataFromDataSource(c);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                c.LogHealthCalendarError("Error when reading Exchange records from Subscribers. " + "Error Message: " + ex.ToString());
            }
        }

    }
}
