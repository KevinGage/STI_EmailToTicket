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
        string _dbPort;
        string _dbName;
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
                    string[] readText = File.ReadAllLines(_configFile);
                    
                    foreach (string s in readText)
                    {
                        string[] split = s.Split('=');
                        if (split.Count() < 2)
                        {
                            _error = "Invalid value found in config file: " + s;
                        }
                        else
                        {
                            switch (split[0])
                            {
                                case "user365":
                                    _user365 = split[1];
                                    break;
                                case "pass365":
                                    _password365 = split[1];
                                    break;
                                case "forward":
                                    _emailForward = split[1];
                                    break;
                                case "dbserver":
                                    _dbServer = split[1];
                                    break;
                                case "dbport":
                                    _dbPort = split[1];
                                    break;
                                case "dbname":
                                    _dbName = split[1];
                                    break;
                                case "dbuser":
                                    _userDb = split[1];
                                    break;
                                case "dbpassword":
                                    _passwordDb = split[1];
                                    break;
                                case "log":
                                    switch (split[1])
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
                                    break;
                            }
                        }
                    }
                    if (_user365 == "" || _password365 == "" || _emailForward == "" || _dbServer == "" || _dbPort == "" || _dbName == "" || _userDb == "" || _passwordDb == "" || _logLevel < 0 || _logLevel > 2)
                    {
                        _error = "Missing required config file parameter.  Must have user365, pass365, forward, dbserver, dbport, dbname, dbuser, dbpassword, and log.  parameter name and vaue must be seperated with =";
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
        public string DbPort
        {
            get { return _dbPort; }
        }
        public string DbName
        {
            get { return _dbName; }
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
