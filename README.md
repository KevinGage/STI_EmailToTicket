# STI_EmailToTicket
A C# service created to check an office 365 inbox for messages and auto-open tickets in the STI ticket applicaiton

<em>Notes:</em><br />

<b>To install the service</b>
  <OL>
    <LI>run developer command prompt for VS2012 as administrator</LI>
    <LI>change to bin folder</LI>
    <LI>run "installutil TicketFromEmail365.exe"</LI>
  </OL>
  
<b>To uninstall the service</b>
  <OL>
    <LI>run developer command prompt for VS2012 as administrator</LI>
    <LI>change to bin folder</LI>
    <LI>run "installutil /u TicketFromEmail365.exe"</LI>
  </OL>
</UL>

<em>Initial Goals:</em><br />
  Read mail server address, username, password, logging level from file. Exclude this file from git!  Have example file in git.
  
  Login using EWS managed API stream method
  
  On new email do the following 
  <OL>
    <LI>Does the email subject already have a ticket number in in?</LI>
    <OL>
      <LI>Yes, update relevant ticket notes, forward to distribution group, reply to client, delete email</LI>
      <LI>No, Is the message from a clients domain?</LI>
        <OL>
          <LI>No, Forward the email, delete the message</LI>
          <LI>Yes, open ticket, forward email to distribution group with modified subject to include ticket, reply to client,  delete email</LI>
        </OL>
      <LI>Log activity based on logging level</LI>
    </OL>
  </OL>
  EWS Managed API stream method has max connect time of 30 minutes.  On disconnect immediatly reconnect.
  
  Possibly eventually add option to reply to emails directly through ticket system?
