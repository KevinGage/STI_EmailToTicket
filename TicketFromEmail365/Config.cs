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
        string _configFile;
        string _error;

        string _user;
        string _password;
        int _logLevel;

        public Config(string path)
        {
            _configFile = path;

            loadConfig();
        }

        public void loadConfig()
        {
            try
            {
                if (File.Exists(_configFile))
                {
                    TextReader tr = new StreamReader(_configFile);

                    _user = tr.ReadLine();
                    _password = tr.ReadLine();

                    string logLevelString = tr.ReadLine().Split('=')[1];

                    switch (logLevelString)
                    {
                        case "low":
                            _logLevel = 0;
                            break;
                        case "medium":
                            _logLevel = 1;
                            break;
                        case "high":
                            _logLevel = 2;
                            break;
                    }
                }
                else
                {
                    _error = "File doese not exist: " + _configFile;
                }
            }
            catch (Exception ex)
            {
                _error = "Unable to process config file: " + _configFile + System.Environment.NewLine + "error: " + ex.ToString();
            }

            _error = "";
        }

        public string Error
        {
            get { return _error; }
        }
        public string User
        {
            get { return _user; }
        }
        public string Password
        {
            get { return _password; }
        }
        public int LogLevel
        {
            get { return _logLevel; }
        }
    }
}
