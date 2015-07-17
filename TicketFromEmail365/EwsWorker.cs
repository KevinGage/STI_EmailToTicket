using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Exchange.WebServices.Data;

namespace TicketFromEmail365
{
    class EwsWorker
    {
        Config _currentConfig;

        public EwsWorker(Config config)
        {
            _currentConfig = config;

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            service.UseDefaultCredentials = false;
            service.Credentials = new WebCredentials(_currentConfig.User, _currentConfig.Password);

            if (_currentConfig.LogLevel > 0)
            {
                Logger.writeSingleLine("Attempting to login to ");
            }
            service.AutodiscoverUrl(_currentConfig.User, RedirectionUrlValidationCallback);
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl) // used as a callback to make sure autodiscover hits a https endpoint
        {
            Logger.writeSingleLine("Attempting to resolve Autodiscover URL");

            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            string logMessage = "Unable to connect to autodiscover address: " + redirectionUrl;

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
                logMessage = "Succesfully resolved autodiscover address: " + redirectionUrl;
            }

            Logger.writeSingleLine(logMessage);
            return result;
        }
    }
}
