---
title: Attach an Outlook Contact Item to an Email Message
ms.prod: outlook
ms.assetid: ae5240ad-dc3e-4499-8fd0-d8c2d90aa9ba
ms.date: 06/08/2017
---


# Attach an Outlook Contact Item to an Email Message
This topic describes how you can programmatically attach a copy of a Microsoft Outlook item, such as a contact, calendar item, or another email message, to an email message before sending the message.

 **Provided by:** Ken Getz, [MCW Technologies, LLC](http://www.mcwtech.com/)

To attach one or more files or Outlook items to a mail message, you can use the  **Attachments** property of the **MailItem** object that represents your outgoing mail, and call the **Add(Object, Object, Object, Object)** method of the **Attachments** object for each of the attachments. The **Add** method allows you to specify the file name and how you want to associate the attachment. To attach Outlook items, such as the contact item shown in the example code for this topic, specify theType parameter of the **Add** method as the **Outlook.olAttachmentType.olEmbeddedItem** enumerated value.

The  `SendMailItem` sample procedure in the code example later in this topic accepts the following:


- A reference to the Outlook  **Application** object.
    
- Strings that contain the subject and body of the message. 
    
- A generic list of strings that contain the SMTP addresses for recipients of the message.
    
- A string that contains the SMTP address of the sender.
    
After creating a new mail item, the code loops through all the recipient addresses, adding each to the  **Recipients** collection of the message. Once the code calls the **ResolveAll()** method of the **Recipients** object, it sets the **Subject** and **Body** properties of the mail item. Next, the code creates a new Outlook **ContactItem** object and adds this new contact item as an attachment to the mail message, specifying the **Outlook.olAttachmentType.olEmbeddedItem** value as a parameter to the **Add** method call.
Before actually sending the email, you must specify the account from which to send the email message. One technique for finding this information is to use the SMTP address of the sender. The  `GetAccountForEmailAddress` function accepts a string that contains the sender's SMTP email address, and returns a reference for the corresponding **Account** object. This method compares the sender's SMTP address with the **SmtpAddress** for each configured email account defined for the session's profile. `application.Session.Accounts` returns an **Accounts** collection for the current profile, tracking information for all accounts including Exchange, IMAP, and POP3 accounts, each of which can be associated with a different delivery store. The **Account** object that has an associated **SmtpAddress** property value that matches the sender's SMTP address is the account to use to send the email message.
After identifying the appropriate account, the code completes by setting the  **SendUsingAccount** property of the mail item to that **Account** object, and then calling the **Send()** method.
The following managed code samples are written in C# and Visual Basic. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code samples in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.
The following code shows how to programmatically attach a copy of a contact item to a mail message. To demonstrate this functionality, in Visual Studio, create a new managed Outlook add-in named  `EmbedOutlookItemAddIn`, and replace the contents of the ThisAddIn.vb or ThisAddIn.cs file with the example code shown here. Modify the  `ThisAddIn_Startup` procedure and update the email addresses appropriately. The SMTP address included in the call to the `SendMailWithAttachments` procedure must correspond to the SMTP address of one of the outgoing email accounts you have previously configured in Outlook.



```C#
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
 
namespace EmbedOutlookItemAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            List<string> recipients = new List<string>();
            recipients.Add("john@contoso.com");
            recipients.Add("john@example.com");
 
            // Replace the SMTP address for sending.
            SendMailItem(Application, "Outlook started", "Outlook started at " + 
                DateTime.Now, recipients, "john@contoso.com");
        }
 
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
 
        public void SendMailItem(Outlook.Application application, 
            string subject, string body, 
            List<string> recipients, string smtpAddress)
        {
 
            Outlook.MailItem newMail = 
                application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
 
            // Set up all the recipients.
            foreach (var recipient in recipients)
            {
                newMail.Recipients.Add(recipient);
            }
            if (newMail.Recipients.ResolveAll())
            {
                // Set the details.
                newMail.Subject = subject;
                newMail.Body = body;
 
                Outlook.ContactItem contact = (Outlook.ContactItem)(application.CreateItem
                    Outlook.OlItemType.olContactItem));
 
                // Create a new contact. Use an existing contact instead, 
                // if you have one to work with.
                contact.FullName = "Kim Abercrombie";
                contact.LastName = "Kim";
                contact.FirstName = "Abercrombie";
                contact.HomeTelephoneNumber = "555-555-1212";
                contact.Save();
 
                newMail.Attachments.Add(contact, Outlook.OlAttachmentType.olEmbeddeditem);
                newMail.SendUsingAccount = GetAccountForEmailAddress(application, smtpAddress);
                newMail.Send();
             }
          }
 
        public Outlook.Account GetAccountForEmailAddress(Outlook.Application application, 
            string smtpAddress)
        {
            // Loop over the Accounts collection of the current Outlook session.
            Outlook.Accounts accounts = application.Session.Accounts;
            foreach (Outlook.Account account in accounts)
            {
                // When the email address matches, return the account.
                if (account.SmtpAddress == smtpAddress)
                {
                    return account;
                }
             }
             // If you get here, no matching account was found.
             throw new System.Exception(string.Format("No Account with SmtpAddress: {0} exists!",
                 smtpAddress));
        }
    }
 
    #region VSTO generated code
 
    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InternalStartup()
    {
        this.Startup += new System.EventHandler(ThisAddIn_Startup);
        this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
    }
        
    #endregion
 
}
```




```VB.net
Public Class ThisAddIn
 
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Dim recipients As New List(Of String)
        recipients.Add("john@contoso.com")
        recipients.Add("john@example.com")
     
        ' Replace the SMTP address for sending.
        SendMailItem(Application, "Outlook started",
            "Outlook started at " &; DateTime.Now, recipients,
            "john@contoso.com")
    End Sub
 
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
 
    End Sub
 
    Private Sub SendMailItem(ByVal application As Outlook.Application, _
        ByVal subject As String, ByVal body As String, ByVal recipients As List(Of String), _
        ByVal smtpAddress As String)
 
        Dim newMail As Outlook.MailItem = _
            DirectCast(application.CreateItem(Outlook.OlItemType.olMailItem), _
            Outlook.MailItem)
 
        ' Set up all the recipients.
        For Each recipient In recipients
            newMail.Recipients.Add(recipient)
        Next
        If newMail.Recipients.ResolveAll() Then
          ' Set the details.
          newMail.Subject = subject
          newMail.Body = body
 
          Dim contact As Outlook.ContactItem =_
             DirectCast(application.CreateItem(
             Outlook.OlItemType.olContactItem), Outlook.ContactItem)
 
          ' Create a new contact. Use an existing contact instead, 
          ' if you have one to work with.
          contact.FullName = "Kim Abercrombie"
          contact.LastName = "Kim"
          contact.FirstName = "Abercrombie"
          contact.HomeTelephoneNumber = "555-555-1212"
          contact.Save()
 
          newMail.Attachments.Add(contact, Outlook.OlAttachmentType.olEmbeddeditem)
          newMail.SendUsingAccount = GetAccountForEmailAddress(application, smtpAddress)
          newMail.Send()
        End If
    End Sub
 
    Private Function GetAccountForEmailAddress(ByVal application As Outlook.Application,_
        ByVal smtpAddress As String) As Outlook.Account
        ' Loop over the Accounts collection of the current Outlook session.
        Dim accounts As Outlook.Accounts = application.Session.Accounts
        For Each account In accounts
            ' When the email address matches, return the account.
            If account.SmtpAddress = smtpAddress Then
                Return account
            End If
        Next
        ' If you get here, no matching account was found.
        Throw New System.Exception(_
            String.Format("No Account with SmtpAddress: {0} exists!", smtpAddress))
    End Function
End Class
```


## See also


#### Concepts


 [Attach a File to a Mail Item](attach-a-file-to-a-mail-item.md)
 [Limit the Size of an Attachment to an Outlook Email Message](limit-the-size-of-an-attachment-to-an-outlook-email-message.md)
 [Modify an Attachment of an Outlook Email Message](modify-an-attachment-of-an-outlook-email-message.md)

