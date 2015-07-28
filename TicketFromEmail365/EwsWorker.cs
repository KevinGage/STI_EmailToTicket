﻿using System;
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
        string _error;

        public EwsWorker(Config config)
        {
            _error = "";
            _currentConfig = config;

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            service.UseDefaultCredentials = false;
            service.Credentials = new WebCredentials(_currentConfig.User365, _currentConfig.Password365);

            if (_currentConfig.LogLevel > 0)
            {
                Logger.writeSingleLine("Attempting to login as: " + _currentConfig.User365);
            }
            try
            {
                service.AutodiscoverUrl(_currentConfig.User365, RedirectionUrlValidationCallback);
            }
            catch
            {
                Logger.writeSingleLine("Login attempt to email server failed");
                _error = ("Login attempt to email server failed");
            }

            if (_error == "")
            {
                ConnectToStream(service);
            }
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
                    Logger.writeSingleLine("Attempting to connect to EWS using streaming method for 30 minutes");
                }
            
                connection.Open();
                if (_currentConfig.LogLevel > 0)
                {
                    Logger.writeSingleLine("Succesfully connected");
                }
            }
            catch (Exception ex)
            {
                _error = "Error connecting to EWS using streaming method: " + ex.ToString();
                Logger.writeSingleLine("Error connecting to EWS using streaming method");
                Logger.writeSingleLine("Error: " + ex.ToString());
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
                        Logger.writeSingleLine("New Mail!");

                        // Display the notification identifier. 
                        if (notification is ItemEvent)
                        {
                            // The NotificationEvent for an e-mail message is an ItemEvent. 
                            ItemEvent itemEvent = (ItemEvent)notification;
                            Logger.writeSingleLine("ItemId: " + itemEvent.ItemId.UniqueId);

                            StreamingSubscriptionConnection senderConnection = (StreamingSubscriptionConnection)sender;

                            if (_currentConfig.LogLevel > 0)
                            {
                                Logger.writeSingleLine(GetItemSubject(itemEvent.ItemId.UniqueId, senderConnection.CurrentSubscriptions.First().Service));
                            }
                            if (_currentConfig.LogLevel > 1)
                            {
                                Logger.writeSingleLine(GetItemBody(itemEvent.ItemId.UniqueId, senderConnection.CurrentSubscriptions.First().Service));
                            }
                        }
                        else
                        {
                            // The NotificationEvent for a folder is an FolderEvent. 
                            FolderEvent folderEvent = (FolderEvent)notification;
                            Logger.writeSingleLine("FolderId: " + folderEvent.FolderId.UniqueId);
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
                Logger.writeSingleLine("Disconnecting from EWS");
            }

            StreamingSubscriptionConnection connection = (StreamingSubscriptionConnection)sender;

            if (_currentConfig.LogLevel > 1)
            {
                Logger.writeSingleLine("Attempting reconnect to EWS");
            }
            connection.Open();

            if (_currentConfig.LogLevel > 1)
            {
                Logger.writeSingleLine("Connection to EWS re-opened");
            }
        }

        static private void OnError(object sender, SubscriptionErrorEventArgs args)
        {
            // Handle error conditions. 
            Exception e = args.Exception;
            Logger.writeSingleLine("Error Encountered in EWS Stream");
            Logger.writeSingleLine("Error: " + e.Message);
        }

        static private string GetItemSubject(ItemId itemId, ExchangeService service)
        {
            // Retrieve the subject for a given item
            string ItemInfo = "";
            Item singleItem;
            PropertySet singleItemPropertySet = new PropertySet(ItemSchema.Subject);

            try
            {
                singleItem = Item.Bind(service, itemId, singleItemPropertySet);
            }
            catch (Exception ex)
            {
                return "Error retrieving message subject Error message: " + System.Environment.NewLine + ex.Message + System.Environment.NewLine + " Message ID: " + itemId.ToString();
            }

            if (!(singleItem is Appointment))
            {
                ItemInfo += "Item subject=" + singleItem.Subject;
            }

            return ItemInfo;
        }

        static private string GetItemBody(ItemId itemId, ExchangeService service)
        {
            // Retrieve the body for a given item
            string ItemInfo = "";
            Item singleItem;
            PropertySet singleItemPropertySet = new PropertySet(ItemSchema.TextBody);

            try
            {
                singleItem = Item.Bind(service, itemId, singleItemPropertySet);
            }
            catch (Exception ex)
            {
                return "Error retrieving message body Error message: " + System.Environment.NewLine + ex.Message + System.Environment.NewLine + " Message ID: " + itemId.ToString();
            }

            if (!(singleItem is Appointment))
            {
                ItemInfo += "Item body=" + singleItem.TextBody;
            }

            return ItemInfo;
        }

        public string Error
        {
            get { return _error; }
        }
    }
}
