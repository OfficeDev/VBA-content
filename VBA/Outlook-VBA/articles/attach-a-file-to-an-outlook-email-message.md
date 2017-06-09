---
title: Attach a File to an Outlook Email Message
ms.prod: outlook
ms.assetid: 44721ad9-750c-4813-bcdb-585ffe8b32c5
ms.date: 06/08/2017
---


# Attach a File to an Outlook Email Message
This topic describes how to programmatically attach one or more files to an outgoing email message in Microsoft Outlook. 

 **Provided by:** Ken Getz, [MCW Technologies, LLC](http://www.mcwtech.com/)


## Object Model Support for Attachments

In Outlook, the  **Attachments** property of the **MailItem** object supports attaching one or more files to an email message. To attach one or more files to a mail item before sending the item, you call the **Add(Object, Object, Object, Object)** method of the **Attachments** object for each of the attachment files. The **Add** method allows you to specify the file name (theSource parameter) and the attachment type (theType parameter) by using the **OlAttachmentType** enumeration. For files in the file system, specify theType parameter as the **Outlook.olAttachmentType.olByValue** enumerated value.


 **Note**  Since Microsoft Office Outlook 2007, you would always use this value to attach a copy of a file in the file system;  **Outlook.olAttachmentType.olByReference** is no longer supported.

In addition, when you send an email in Rich Text Format (RTF), you can also specify two other optional parameters—Position andDisplayName—when you call the  **Add** method. ThePosition parameter allows you to specify the position within the email where the attachment should appear. Use one of the following values for thePosition parameter:


- The value 0 hides the attachment within the body of the email. 
    
- The value 1 places the attachment before the first character.
    
- A number that is larger than the number of characters in the body of the email item places the attachment at the end of the body text. 
    
For RTF email messages, you can also specify the DisplayName parameter, which provides the name that is displayed within the body of the message for the attachment. For plain text or HTML email messages, the attachment displays only the name of the file.


## Sending a Message with Files as Attachments

The  `SendEmailWithAttachments` sample procedure in the code example later in this topic accepts the following:


- A reference to the Outlook  **Application** object.
    
- Strings that contain the subject and body of the message. 
    
- A generic list of strings that contain a list of the SMTP addresses for recipients of the message.
    
- A string that contains the SMTP address of the sender.
    
- A generic list of strings that contain the paths for the files to be attached. 
    
After creating a new email item, the code adds each recipient to the  **Recipients** collection property of the mail item. Once the code calls the **ResolveAll()** method, it sets the **Subject** and **Body** properties of the mail item before looping through each item in the provided list of attachment paths, adding each to the **Attachments** property of the mail item.

Before actually sending the email, you must specify the account from which to send the email message. One technique for finding this information is to use the SMTP address of the sender. The  `GetAccountForEmailAddress` function accepts a string that contains the sender's SMTP email address, and returns a reference for the corresponding **Account** object. This method compares the sender's SMTP address with the **SmtpAddress** property for each configured email account defined for the session's profile. `application.Session.Accounts` returns an **Accounts** collection for the current profile, tracking information for all accounts including Exchange, IMAP, and POP3 accounts, each of which can be associated with a different delivery store. The **Account** object that has an associated **SmtpAddress** property value that matches the sender's SMTP address is the account to use to send the email message.

After identifying the appropriate account, the code completes by setting the  **SendUsingAccount** property of the mail item to that **Account** object, and then calling the **Send()** method.

The following managed code samples are written in C# and Visual Basic. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code samples in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code shows how to programmatically attach files to an outgoing email message in Outlook. To demonstrate this functionality, in Visual Studio, create a new managed Outlook add-in named  `AttachFileAddIn`, and replace the contents of the ThisAddIn.vb or ThisAddIn.cs file with the example code shown here. Modify the  `ThisAddIn_Startup` procedure to include a reference to a file in your file system, and update the email addresses appropriately. The SMTP address included in the call to the `SendMailWithAttachments` procedure must correspond to the SMTP address of one of the outgoing email accounts you have previously configured in Outlook.




```C#
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
 
namespace AttachFileAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            List<string> attachments = new List<string>();
            attachments.Add("c:\\somefile.txt");
 
            List<string> recipients = new List<string>();
            recipients.Add("john@contoso.com");
            recipients.Add("john@example.com");
            SendEmailWithAttachments(Application, "Test", "Body", recipients, "john@example.com", 
              attachments);
        }
 
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
 
        private void SendEmailWithAttachments(Outlook.Application application, 
            string subject, string body, List<string> recipients, 
            string smtpAddress, List<string> attachments)
        {
 
            // Create a new MailItem and set the To, Subject, and Body properties.
            var newMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
 
            // Set up all the recipients.
            foreach (var recipient in recipients)
            {
                newMail.Recipients.Add(recipient);
            }
            if (newMail.Recipients.ResolveAll())
            {
                newMail.Subject = subject;
                newMail.Body = body;
                foreach (string attachment in attachments)
                {
                    newMail.Attachments.Add(attachment, Outlook.OlAttachmentType.olByValue);
                }
            }
 
            // Retrieve the account that has the specific SMTP address.
            Outlook.Account account = GetAccountForEmailAddress(application, smtpAddress);
            // Use this account to send the e-mail.
            newMail.SendUsingAccount = account;
            newMail.Send();
        }
 
        private Outlook.Account GetAccountForEmailAddress(Outlook.Application application, 
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
}
```




```VB.net
Public Class ThisAddIn
 
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Dim attachments As New List(Of String)
        attachments.Add("c:\somefile.txt")
 
        Dim recipients As New List(Of String)
        recipients.Add("john@contoso.com")
        recipients.Add("john@example.com")
        SendEmailWithAttachments(Application, "Test", "Body", recipients, "john@contoso.com", attachments)
    End Sub
 
    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
 
    End Sub
 
    Private Sub SendEmailWithAttachments(ByVal application As Outlook.Application, _
        ByVal subject As String, ByVal body As String,
        ByVal recipients As List(Of String),
        ByVal smtpAddress As String,
        ByVal attachments As List(Of String))
 
        ' Create a new MailItem and set the To, Subject, and Body properties.
        Dim newMail As Outlook.MailItem =
            DirectCast(application.CreateItem(Outlook.OlItemType.olMailItem), 
            Outlook.MailItem)
 
        ' Set up all the recipients.
        For Each recipient In recipients
            newMail.Recipients.Add(recipient)
        Next
        If newMail.Recipients.ResolveAll() Then
            newMail.Subject = subject
            newMail.Body = body
            For Each attachment As String In attachments
                newMail.Attachments.Add(attachment, Outlook.OlAttachmentType.olByValue)
            Next
        End If
 
        ' Retrieve the account that has the specific SMTP address.
        Dim account As Outlook.Account = GetAccountForEmailAddress(application, smtpAddress)
        ' Use this account to send the e-mail.
        newMail.SendUsingAccount = account
        newMail.Send()
    End Sub
 
   
    Private Function GetAccountForEmailAddress(
        ByVal application As Outlook.Application,
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
        Throw New System.Exception(
            String.Format("No Account with SmtpAddress: {0} exists!", smtpAddress))
    End Function
End Class
```


## See also


#### Concepts


 [Attach an Outlook Contact Item to an Email Message](attach-an-outlook-contact-item-to-an-email-message.md)
 [Limit the Size of an Attachment to an Outlook Email Message](limit-the-size-of-an-attachment-to-an-outlook-email-message.md)
 [Modify an Attachment of an Outlook Email Message](modify-an-attachment-of-an-outlook-email-message.md)

