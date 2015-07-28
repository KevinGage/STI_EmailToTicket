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

        string _user365;
        string _password365;
        string _emailForward;
        string _dbServer;
        string _userDb;
        string _passwordDb;
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

                    _user365 = tr.ReadLine();
                    _password365 = tr.ReadLine();

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
        public string User365
        {
            get { return _user365; }
        }
        public string Password365
        {
            get { return _user365; }
        }
        public string EmailForward
        {
            get { return _emailForward; }
        }
        public string DbServer
        {
            get { return _dbServer; }
        }
        public string UserDb
        {
            get { return _userDb; }
        }
        public string PasswordDb
        {
            get { return _passwordDb; }
        }
        public int LogLevel
        {
            get { return _logLevel; }
        }
    }
}
