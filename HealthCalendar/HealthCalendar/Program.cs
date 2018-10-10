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
        
        static string myconnstr = ConfigurationManager.ConnectionStrings["connstrng"].ConnectionString;

        static void Main(string[] args)
        {
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
            if (c.GetExchangeAuthorization(c))
            {
                if (!c.GetExchangeAuthorization(c))
                {
                    var logger = NLog.LogManager.GetCurrentClassLogger();
                    logger.Info("Unable to connect to Exchange Server. Please contact your IT Department.");
                }
            }






            //Exchange
            // For each user with a



            var logend = NLog.LogManager.GetCurrentClassLogger();
            logend.Info("Ending HealthCalendar Execution.");


        }
    }
}
