using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Exchange.WebServices.Data;

namespace TicketFromEmail365
{
    class MyEwsWorker
    {
        MyConfig _currentConfig;
        string _error;

        public MyEwsWorker(MyConfig config)
        {
            _error = "";
            _currentConfig = config;

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            service.UseDefaultCredentials = false;
            service.Credentials = new WebCredentials(_currentConfig.User365, _currentConfig.Password365);

            if (_currentConfig.LogLevel > 0)
            {
                MyLogger.writeSingleLine("Attempting to login as: " + _currentConfig.User365);
            }
            try
            {
                service.AutodiscoverUrl(_currentConfig.User365, RedirectionUrlValidationCallback);
            }
            catch
            {
                MyLogger.writeSingleLine("Login attempt to email server failed");
                _error = ("Login attempt to email server failed");
            }

            if (_error == "")
            {
                ConnectToStream(service);
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl) // used as a callback to make sure autodiscover hits a https endpoint
        {
            MyLogger.writeSingleLine("Attempting to resolve Autodiscover URL");

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

            MyLogger.writeSingleLine(logMessage);
            return result;
        }

        private void ConnectToStream(ExchangeService authenticatedSession)
        {
            try
            {
                StreamingSubscription subscription = authenticatedSession.SubscribeToStreamingNotifications(
                    new FolderId[] { WellKnownFolderName.Inbox },
                    EventType.NewMail, // chose events that we want to listen for. could include deleted, modified, moved, etc
                    EventType.Created);

                // Create a streaming connection to the service object, over which events are returned to the client.
                // Keep the streaming connection open for 30 minutes.
                StreamingSubscriptionConnection connection = new StreamingSubscriptionConnection(authenticatedSession, 30);
                connection.AddSubscription(subscription);
                connection.OnNotificationEvent += new StreamingSubscriptionConnection.NotificationEventDelegate(OnEvent);
                connection.OnSubscriptionError += new StreamingSubscriptionConnection.SubscriptionErrorDelegate(OnError);
                connection.OnDisconnect += new StreamingSubscriptionConnection.SubscriptionErrorDelegate(OnDisconnect); 

                if (_currentConfig.LogLevel > 0)
                {
                    MyLogger.writeSingleLine("Attempting to connect to EWS using streaming method for 30 minutes");
                }
            
                connection.Open();
                if (_currentConfig.LogLevel > 0)
                {
                    MyLogger.writeSingleLine("Succesfully connected");
                }
            }
            catch (Exception ex)
            {
                _error = "Error connecting to EWS using streaming method: " + ex.ToString();
                MyLogger.writeSingleLine("Error connecting to EWS using streaming method");
                MyLogger.writeSingleLine("Error: " + ex.ToString());
            }
        }

        private void OnEvent(object sender, NotificationEventArgs args) 
        {
            StreamingSubscription subscription = args.Subscription;

            // Loop through all item-related events. 
            foreach (NotificationEvent notification in args.Events)
            {

                switch (notification.EventType)
                {
                    case EventType.NewMail:
                        if (_currentConfig.LogLevel > 0)
                        {
                            MyLogger.writeSingleLine("New Mail");
                        }

                        // Display the notification identifier. 
                        if (notification is ItemEvent)
                        {
                            // The NotificationEvent for an e-mail message is an ItemEvent. 
                            ItemEvent itemEvent = (ItemEvent)notification;
                            if (_currentConfig.LogLevel > 1)
                            {
                                MyLogger.writeSingleLine("ItemId: " + itemEvent.ItemId.UniqueId);
                            }

                            StreamingSubscriptionConnection senderConnection = (StreamingSubscriptionConnection)sender;

                            try
                            {
                                Item singleItem = Item.Bind(senderConnection.CurrentSubscriptions.First().Service, itemEvent.ItemId.UniqueId);

                                if (singleItem is EmailMessage)
                                {
                                    EmailMessage message = (EmailMessage)singleItem;
                                    PropertySet propertiesToLoad = new PropertySet(EmailMessageSchema.Sender, EmailMessageSchema.CcRecipients, EmailMessageSchema.BccRecipients, ItemSchema.Subject, ItemSchema.TextBody, ItemSchema.Body);

                                    message.Load(propertiesToLoad);
                                    /*
                                    if (_currentConfig.LogLevel > 0)
                                    {
                                        MyLogger.writeSingleLine("Sender: " + message.Sender.Address);
                                    }
                                    if (_currentConfig.LogLevel > 0)
                                    {
                                        MyLogger.writeSingleLine("Subject: " + message.Subject);
                                    }
                                    if (_currentConfig.LogLevel > 1)
                                    {
                                        MyLogger.writeSingleLine("Text Body: " + message.TextBody);
                                    }

                                    ForwardMessage(message, "I forwarded stuff", 123);
                                    */

                                    MyTicketWorker ticketWorker = new MyTicketWorker(message, _currentConfig);

                                    if (ticketWorker.TicketNumber != 0 && ticketWorker.Error == null)
                                    {
                                        //ticket exists in subject. no error
                                        //ticket should be updated
                                        //forward to sti
                                        //dont forward to client
                                        if (ForwardMessage(message, "ticket notes updated"))
                                        {
                                            if (_currentConfig.LogLevel > 1)
                                            {
                                                MyLogger.writeSingleLine("Ticket Succesfully updated and email forwarded: Ticket " + ticketWorker.TicketNumber.ToString());
                                            }
                                        }
                                        else
                                        {
                                            MyLogger.writeSingleLine("Ticket Updated but error forwarding email: Ticket " + ticketWorker.TicketNumber.ToString());
                                            MyLogger.writeSingleLine("message id: " + message.Id.ToString());
                                        }
                                    }
                                    else if (ticketWorker.TicketNumber == 0 && ticketWorker.Error == "")
                                    {
                                        //ticket doesn't exist. no error
                                        //check client domain
                                        if (ticketWorker.CheckClientDomain())
                                        {
                                            //is client
                                            //open ticket
                                            //response to client
                                            //forward to sti
                                        }
                                        else
                                        {
                                            //not a client
                                            //just forward email to sti
                                            
                                        }
                                    }
                                    else
                                    {
                                        //error
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MyLogger.writeSingleLine("Error retrieving message information.  Error: " + ex.ToString());
                            }

                        }
                        else
                        {
                            // The NotificationEvent for a folder is an FolderEvent. 
                            FolderEvent folderEvent = (FolderEvent)notification;
                            MyLogger.writeSingleLine("FolderId: " + folderEvent.FolderId.UniqueId);
                        }

                        break;
//                    case EventType.Created:
//                       Logger.writeSingleLine("New Item or Folder!");
//                        break;
                }         
            } 
        }

        private void OnDisconnect(object sender, SubscriptionErrorEventArgs args) 
        {

            if (_currentConfig.LogLevel > 1)
            {
                MyLogger.writeSingleLine("Disconnecting from EWS");
            }

            StreamingSubscriptionConnection connection = (StreamingSubscriptionConnection)sender;

            if (_currentConfig.LogLevel > 1)
            {
                MyLogger.writeSingleLine("Attempting reconnect to EWS");
            }
            connection.Open();

            if (_currentConfig.LogLevel > 1)
            {
                MyLogger.writeSingleLine("Connection to EWS re-opened");
            }
        }

        static private void OnError(object sender, SubscriptionErrorEventArgs args)
        {
            // Handle error conditions. 
            Exception e = args.Exception;
            MyLogger.writeSingleLine("Error Encountered in EWS Stream");
            MyLogger.writeSingleLine("Error: " + e.Message);
        }

        private bool ForwardMessage(EmailMessage message, string prefix)
        {
            ResponseMessage forwardMessage = message.CreateForward();

            forwardMessage.BodyPrefix = prefix;

            forwardMessage.ToRecipients.Add(_currentConfig.EmailForward);

            forwardMessage.SendAndSaveCopy();

            return false;
        }

        public string Error
        {
            get { return _error; }
        }
    }
}
