---
title: Account Object (Outlook)
keywords: vbaol11.chm3153
f1_keywords:
- vbaol11.chm3153
ms.prod: outlook
api_name:
- Outlook.Account
ms.assetid: f624438c-4e45-2822-18b6-bfe8074a33c0
ms.date: 06/08/2017
---


# Account Object (Outlook)

The  **Account** object represents an account that is defined for the current profile.


## Remarks

The purpose of the [Accounts](accounts-object-outlook.md) collection object and the **Account** object is to provide the capacity to enumerate **Account** objects in a given profile, to identify the type of **Account**, and to use a specific **Account** object to send mail.


 **Note**  Helmut Obertanner provided the following code samples. Helmut is a [Microsoft Most Valuable Professional](https://mvp.microsoft.com/en-us/default.aspx
) with expertise in Microsoft Office development tools in Microsoft Visual Studio and Microsoft Office Outlook.


## Example

The following managed code samples are written in C# and Visual Basic. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code samples in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code samples show the  `DisplayAccountInformation` method of the `Sample` class, implemented as part of an Outlook add-in project. Each project adds a reference to the Outlook PIA, which is based on the **Microsoft.Office.Interop.Outlook** namespace. The `DisplayAccountInformation` method takes as an input argument a trusted Outlook[Application](http://msdn.microsoft.com/library/797003e7-ecd1-eccb-eaaf-32d6ddde8348%28Office.15%29.aspx) object, and uses the **Account** object to display the details of each account that is available for the current Outlook profile.




```C#
using System; 
using System.Text; 
using Outlook = Microsoft.Office.Interop.Outlook; 
 
namespace OutlookAddIn1 
{ 
 class Sample 
 { 
 public static void DisplayAccountInformation(Outlook.Application application) 
 { 
 
 // The Namespace Object (Session) has a collection of accounts. 
 Outlook.Accounts accounts = application.Session.Accounts; 
 
 // Concatenate a message with information about all accounts. 
 StringBuilder builder = new StringBuilder(); 
 
 // Loop over all accounts and print detail account information. 
 // All properties of the Account object are read-only. 
 foreach (Outlook.Account account in accounts) 
 { 
 
 // The DisplayName property represents the friendly name of the account. 
 builder.AppendFormat("DisplayName: {0}\n", account.DisplayName); 
 
 // The UserName property provides an account-based context to determine identity. 
 builder.AppendFormat("UserName: {0}\n", account.UserName); 
 
 // The SmtpAddress property provides the SMTP address for the account. 
 builder.AppendFormat("SmtpAddress: {0}\n", account.SmtpAddress); 
 
 // The AccountType property indicates the type of the account. 
 builder.Append("AccountType: "); 
 switch (account.AccountType) 
 { 
 
 case Outlook.OlAccountType.olExchange: 
 builder.AppendLine("Exchange"); 
 break; 
 
 case Outlook.OlAccountType.olHttp: 
 builder.AppendLine("Http"); 
 break; 
 
 case Outlook.OlAccountType.olImap: 
 builder.AppendLine("Imap"); 
 break; 
 
 case Outlook.OlAccountType.olOtherAccount: 
 builder.AppendLine("Other"); 
 break; 
 
 case Outlook.OlAccountType.olPop3: 
 builder.AppendLine("Pop3"); 
 break; 
 } 
 
 builder.AppendLine(); 
 } 
 
 // Display the account information. 
 System.Windows.Forms.MessageBox.Show(builder.ToString()); 
 } 
 } 
}
```




```VB.net
Imports Outlook = Microsoft.Office.Interop.Outlook 
 
Namespace OutlookAddIn2 
 Class Sample 
 Shared Sub DisplayAccountInformation(ByVal application As Outlook.Application) 
 
 ' The Namespace Object (Session) has a collection of accounts. 
 Dim accounts As Outlook.Accounts = application.Session.Accounts 
 
 ' Concatenate a message with information about all accounts. 
 Dim builder As StringBuilder = New StringBuilder() 
 
 ' Loop over all accounts and print detail account information. 
 ' All properties of the Account object are read-only. 
 Dim account As Outlook.Account 
 For Each account In accounts 
 
 ' The DisplayName property represents the friendly name of the account. 
 builder.AppendFormat("DisplayName: {0}" &amp; vbNewLine, account.DisplayName) 
 
 ' The UserName property provides an account-based context to determine identity. 
 builder.AppendFormat("UserName: {0}" &amp; vbNewLine, account.UserName) 
 
 ' The SmtpAddress property provides the SMTP address for the account. 
 builder.AppendFormat("SmtpAddress: {0}" &amp; vbNewLine, account.SmtpAddress) 
 
 ' The AccountType property indicates the type of the account. 
 builder.Append("AccountType: ") 
 Select Case (account.AccountType) 
 
 Case Outlook.OlAccountType.olExchange 
 builder.AppendLine("Exchange") 
 
 
 Case Outlook.OlAccountType.olHttp 
 builder.AppendLine("Http") 
 
 
 Case Outlook.OlAccountType.olImap 
 builder.AppendLine("Imap") 
 
 
 Case Outlook.OlAccountType.olOtherAccount 
 builder.AppendLine("Other") 
 
 
 Case Outlook.OlAccountType.olPop3 
 builder.AppendLine("Pop3") 
 
 
 End Select 
 
 builder.AppendLine() 
 Next 
 
 
 ' Display the account information. 
 Windows.Forms.MessageBox.Show(builder.ToString()) 
 End Sub 
 
 
 End Class 
End Namespace
```


## Methods



|**Name**|
|:-----|
|[GetAddressEntryFromID](http://msdn.microsoft.com/library/5aa9c67e-579f-5519-ed38-c80009cf506b%28Office.15%29.aspx)|
|[GetRecipientFromID](http://msdn.microsoft.com/library/7b97ce67-6015-ece6-de1b-6d4226be83aa%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AccountType](http://msdn.microsoft.com/library/7e59f240-512d-eb20-69b2-b88ee52a9d27%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/47b2dd80-9b5f-6873-9d4a-c465641605da%28Office.15%29.aspx)|
|[AutoDiscoverConnectionMode](http://msdn.microsoft.com/library/d9089143-caff-6e08-cc7d-f8659384d36e%28Office.15%29.aspx)|
|[AutoDiscoverXml](http://msdn.microsoft.com/library/201c5aba-5cff-0934-a750-b4ac0cb30860%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/93add2b8-e71d-1d4f-f8e2-a5898d0242fc%28Office.15%29.aspx)|
|[CurrentUser](http://msdn.microsoft.com/library/e17ab6a9-344e-b3bf-543c-07590c406a2b%28Office.15%29.aspx)|
|[DeliveryStore](http://msdn.microsoft.com/library/181d52ff-7c48-af7b-dbec-3562f1c8801b%28Office.15%29.aspx)|
|[DisplayName](http://msdn.microsoft.com/library/20fd9286-c7d9-3bb3-5b33-137313f1c8d5%28Office.15%29.aspx)|
|[ExchangeConnectionMode](http://msdn.microsoft.com/library/40fee809-48ab-5788-819a-c61b6eb782a5%28Office.15%29.aspx)|
|[ExchangeMailboxServerName](http://msdn.microsoft.com/library/f75354c9-3374-140f-63a6-ca04ce6101cb%28Office.15%29.aspx)|
|[ExchangeMailboxServerVersion](http://msdn.microsoft.com/library/5bfd2c63-5a87-9225-a9a8-1771fc480f21%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/86d6bc88-6357-97b7-71e4-3c51eae01d74%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/92890235-402c-80c8-10b7-7339f153134e%28Office.15%29.aspx)|
|[SmtpAddress](http://msdn.microsoft.com/library/443beb7a-0ada-8e86-69d7-63880033abca%28Office.15%29.aspx)|
|[UserName](http://msdn.microsoft.com/library/3ab96240-b68c-e2f7-83b9-6d6663c4880d%28Office.15%29.aspx)|

## See also


#### Other resources


[Account Object Members](http://msdn.microsoft.com/library/37759c57-d1ec-775c-cbe6-75c8f314d196%28Office.15%29.aspx)
[How to: Send an E-mail Given the SMTP Address of an Account](http://msdn.microsoft.com/library/5e5f707d-8771-bd5f-945b-58537732d99a%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
