using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

namespace TicketFromEmail365
{
    class Config
    {
        string configFile;
        string error;

        string user;
        string password;
        int logLevel;

        public Config(string path)
        {
            configFile = path;

            loadConfig();
        }

        public void loadConfig()
        {
            try
            {
                if (File.Exists(configFile))
                {
                    TextReader tr = new StreamReader(configFile);

                    user = tr.ReadLine();
                    password = tr.ReadLine();

                    string logLevelString = tr.ReadLine().Split('=')[1];

                    switch (logLevelString)
                    {
                        case "low":
                            logLevel = 0;
                            break;
                        case "medium":
                            logLevel = 1;
                            break;
                        case "high":
                            logLevel = 2;
                            break;
                    }
                }
                else
                {
                    error = "File doese not exist: " + configFile;
                }
            }
            catch (Exception ex)
            {
                error = "Unable to process config file: " + configFile + System.Environment.NewLine + "error: " + ex.ToString();
            }

            error = "";
        }

        public string Error
        {
            get { return error; }
        }
    }
}
