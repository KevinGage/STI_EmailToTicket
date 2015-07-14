# STI_EmailToTicket
A C# service created to check an office 365 inbox for messages and auto-open tickets in the STI ticket applicaiton

Initial Goals:
  Read mail server address, username, password, logging level from file. Exclude this file from git!  Have example file in git.
  
  Login using EWS managed API stream method
  
  On new email do the following
    1.  Is email from a clients domain?
      a.  No, forward to distribution group, delete email
      b.  Yes, Does the email have ticket info in the subject?
        i.  Yes, update relevant ticket notes, forward email to distribution group, delete email
        ii.  No, open ticket, forward email to distribution group with modified subject to include ticket, delete email
      c.  Log activity based on logging level
  
  EWS Managed API stream method has max connect time of 30 minutes.  On disconnect immediatly reconnect.
  
  Possibly eventually add option to reply to emails directly through ticket system?
