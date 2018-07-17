---
title: Send an E-mail Given the SMTP Address of an Account (Outlook)
ms.prod: outlook
ms.assetid: 5e5f707d-8771-bd5f-945b-58537732d99a
ms.date: 06/08/2017
---


# Send an E-mail Given the SMTP Address of an Account (Outlook)

This topic shows how to create an e-mail and send it from a Microsoft Outlook account, given the Simple Mail Transfer Protocol (SMTP) address of that account. 



|
![MVP Logo](./images/MVPLogo_Small_ZA10349011.jpg)

|Helmut Obertanner provided the following code samples. Helmut is a  [Microsoft Most Valuable Professional](https://mvp.microsoft.com/en-us/default.aspx
) with expertise in Microsoft Office development tools in Microsoft Visual Studio and Microsoft Office Outlook.|



The following managed code samples are written in C# and Visual Basic. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code samples in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.
The following code samples contain the  `SendEmailFromAccount` and `GetAccountForEmailAddress` methods of the `Sample` class, implemented as part of an Outlook add-in project. Each project adds a reference to the Outlook PIA, which is based on the **Microsoft.Office.Interop.Outlook** namespace. The `SendEmailFromAccount` method accepts as input arguments a trusted **[Application](application-object-outlook.md)** object, and strings that represent the subject, body, a semicolon-delimited list of recipients, and the SMTP address of an e-mail account. `SendEmailFromAccount` creates a **[MailItem](mailitem-object-outlook.md)** object and initializes the **[To](mailitem-to-property-outlook.md)**,  **[Subject](mailitem-subject-property-outlook.md)**, and  **[Body](mailitem-body-property-outlook.md)** properties with the given arguments. To find the **[Account](account-object-outlook.md)** object from which to send the e-mail, `SendEmailFromAccount` calls the `GetAccountForEmailAddress` method, which matches the given SMTP address with the **[SmtpAddress](account-smtpaddress-property-outlook.md)** property of each account for the current profile. The matching **Account** object is returned to `SendEmailFromAccount`, which then initializes the  **[SendUsingAccount](mailitem-sendusingaccount-property-outlook.md)** property of the **MailItem** with this **Account** object, and sends the **MailItem**.
The following is the C# code sample.



```C#
using System; 
using System.Text; 
using Outlook = Microsoft.Office.Interop.Outlook; 
 
namespace OutlookAddIn1 
{ 
    class Sample 
    { 
        public static void SendEmailFromAccount(Outlook.Application application, string subject, string body, string to, string smtpAddress) 
        { 
 
            // Create a new MailItem and set the To, Subject, and Body properties. 
            Outlook.MailItem newMail = (Outlook.MailItem)application.CreateItem(Outlook.OlItemType.olMailItem); 
            newMail.To = to; 
            newMail.Subject = subject; 
            newMail.Body = body; 
 
            // Retrieve the account that has the specific SMTP address. 
            Outlook.Account account = GetAccountForEmailAddress(application, smtpAddress); 
            // Use this account to send the e-mail. 
            newMail.SendUsingAccount = account; 
            newMail.Send(); 
        } 
 
 
        public static Outlook.Account GetAccountForEmailAddress(Outlook.Application application, string smtpAddress) 
        { 
 
            // Loop over the Accounts collection of the current Outlook session. 
            Outlook.Accounts accounts = application.Session.Accounts; 
            foreach (Outlook.Account account in accounts) 
            { 
                // When the e-mail address matches, return the account. 
                if (account.SmtpAddress == smtpAddress) 
                { 
                    return account; 
                } 
            } 
            throw new System.Exception(string.Format("No Account with SmtpAddress: {0} exists!", smtpAddress)); 
        } 
 
    } 
}
```

The following is the Visual Basic code sample.



```VB.net
Imports Outlook = Microsoft.Office.Interop.Outlook 
 
Namespace OutlookAddIn2 
    Class Sample 
         
        Shared Sub SendEmailFromAccount(ByVal application As Outlook.Application, _ 
            ByVal subject As String, ByVal body As String, ByVal recipients As String, ByVal smtpAddress As String) 
 
            ' Create a new MailItem and set the To, Subject and Body properties. 
            Dim newMail As Outlook.MailItem = DirectCast(application.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem) 
            newMail.To = recipients 
            newMail.Subject = subject 
            newMail.Body = body 
 
            ' Retrieve the account that has the specific SMTP address. 
            Dim account As Outlook.Account = GetAccountForEmailAddress(application, smtpAddress) 
            ' Use this account to send the e-mail. 
            newMail.SendUsingAccount = account 
            newMail.Send() 
        End Sub 
 
        Shared Function GetAccountForEmailAddress(ByVal application As Outlook.Application, ByVal smtpAddress As String) As Outlook.Account 
 
            ' Loop over the Accounts collection of the current Outlook session. 
            Dim accounts As Outlook.Accounts = application.Session.Accounts 
            Dim account As Outlook.Account 
            For Each account In accounts 
                ' When the e-mail address matches, return the account. 
                If account.SmtpAddress = smtpAddress Then 
                    Return account 
                End If 
            Next 
            Throw New System.Exception(String.Format("No Account with SmtpAddress: {0} exists!", smtpAddress)) 
        End Function 
 
    End Class 
End Namespace
```


