using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using HealthCalendarClasses;
using NLog;
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

namespace HealthCalendar
{    
    class Program
    {
        static string MyConnString = ConfigurationManager.ConnectionStrings["connstrng"].ConnectionString;

        static void Main(string[] args)
        {
            //bool isSuccess = false;
            SqlDataReader readerSQLClientID;
            long lCareProviderOID;
            HealthCalendarClasses.HealthCalendarClass c = new HealthCalendarClass();
            var logstart = NLog.LogManager.GetCurrentClassLogger();
            logstart.Info("Starting HealthCalendar Execution.");


            //Startup

            // Trust Settings
            if (!c.GetTrustSettings(c))
            {
                var logger = NLog.LogManager.GetCurrentClassLogger();
                logger.Info("Cannot connect or read data from the Health Calendar database.");
                Environment.Exit(0);
            }

            // Google Calendar Settings
            if (c.bGoogleEnabled)
            {
                if (!c.GetGoogleClientSecret(c))
                {
                    var logger = NLog.LogManager.GetCurrentClassLogger();
                    logger.Info("Unable to obtain Google Client Secret. Please contact your IT Department.");
                }

                if (!c.GetGoogleAuthorization(c))
                {
                    var logger = NLog.LogManager.GetCurrentClassLogger();
                    logger.Info("Unable to obtain Google Authorization. Please contact your IT Department.");
                }
            }

            // NHSNet Calendar Settings
            if (c.bNHSNetEnabled)
            {
                if (!c.GetNHSNetAuthorization(c))
                {
                    var logger = NLog.LogManager.GetCurrentClassLogger();
                    logger.Info("Unable to connect to NHS Net Server. Please contact your IT Department.");
                }
            }

            // Trust Exchange Calendar Settings
            if (c.bExchangeEnabled)
            {
                if (!c.GetExchangeAuthorization(c))
                {
                    var logger = NLog.LogManager.GetCurrentClassLogger();
                    logger.Info("Unable to connect to Exchange Server. Please contact your IT Department.");
                }
            }


            //Exchange            
            if (c.bExchangeEnabled)
            {
                SetSubscribersExchangeCalendarData(c);
            }
            //NHSNet            
            if (c.bNHSNetEnabled)
            {
                SetSubscribersNHSNetCalendarData(c);
            }


            var logend = NLog.LogManager.GetCurrentClassLogger();
            logend.Info("Ending HealthCalendar Execution.");
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
                string sql = "SELECT SubscriberOID, ExchangeEmail, ExchangeCalendarID, ExchangeCalendarName FROM Subscribers WHERE (NOT (ExchangeEmail IS NULL)) AND (NOT (ExchangeEmail = '')) AND (NOT (ExchangeCalendarName IS NULL)) AND (NOT (ExchangeCalendarName = ''))";
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
                            lCareProviderOID = (long)readerSQLClientID.GetDecimal(0);
                            c.SubscriberOID = lCareProviderOID.ToString();
                            c.ExchangeCalendarName = readerSQLClientID.GetString(3);
                            c.SetExchangeCalendarDataFromDataSource(c);
                        }
                    } 

                }
            }
            catch (Exception ex)
            {
                var logger = NLog.LogManager.GetCurrentClassLogger();
                logger.Info("Error when reading Exchange records from Subscribers. " + "Error Message: " + ex.ToString());
            }
        }

        static private void SetSubscribersNHSNetCalendarData(HealthCalendarClass c)
        {
            SqlDataReader readerSQLClientID;
            long lCareProviderOID;

            try
            {
                SqlConnection conn = new SqlConnection(MyConnString);
                string sql = "SELECT SubscriberOID, NHSNetEmail, NHSNetCalendarID, NHSNetCalendarName FROM Subscribers WHERE (NOT (NHSNetEmail IS NULL)) AND (NOT (NHSNetEmail = '')) AND (NOT (NHSNetCalendarName IS NULL)) AND (NOT (NHSNetCalendarName = ''))";
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
                            lCareProviderOID = (long)readerSQLClientID.GetDecimal(0);
                            c.SubscriberOID = lCareProviderOID.ToString();
                            c.NHSNetCalendarName = readerSQLClientID.GetString(3);
                            c.SetNHSNetCalendarDataFromDataSource(c);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                var logger = NLog.LogManager.GetCurrentClassLogger();
                logger.Info("Error when reading Exchange records from Subscribers. " + "Error Message: " + ex.ToString());
            }
        }

    }
}
