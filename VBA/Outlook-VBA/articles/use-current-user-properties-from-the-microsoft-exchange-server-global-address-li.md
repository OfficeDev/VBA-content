---
title: Use Current User Properties from the Microsoft Exchange Server Global Address List
keywords: olfm10.chm3077404
f1_keywords:
- olfm10.chm3077404
ms.prod: outlook
ms.assetid: fa3e6e11-a63e-fcf5-14f0-f16dc3b755dd
ms.date: 06/08/2017
---


# Use Current User Properties from the Microsoft Exchange Server Global Address List

In code, open an OLE messaging session and log on, and then use the following table to reference the property you want to use.


```vb
Set olemSession = Application.CreateObject("MAPI.Session") 
ReturnCode = olemSession.Logon( Application.GetNameSpace("MAPI").CurrentUser, "", False, False, 0 ) 
myPage = Item.GetInspector.ModifiedFormPages("Message") 
Set myUser = olemSession.CurrentUser 
Item.UserProperties.Find("Name") = myUser.Name 
Item.UserProperties.Find("Messaging Address") = myUser.Address 
Item.UserProperties.Find("MAPI First Name") = myUser.Fields.item(&;h3a06001e)
```



|**Address Book Property**|**Reference**|
|:-----|:-----|
|PidTagGivenName|&;h3a06001e|
|PidTagInitials|&;h3a0a001e|
|PidTagSurname|&;h3a11001e|
|PidTag7BitDisplayName|&;h39ff001e|
|PidTagStreetAddress|&;h3a29001e|
|PidTagLocality|&;h3a27001e|
|PidTagStateOrProvince|&;h3a28001e|
|PidTagPostalCode|&;h3a2a001e|
|PidTagCountry|&;h3a26001e|
|PidTagTitle|&;h3a17001e|
|PidTagCompanyName|&;h3a16001e|
|PidTagDepartmentName|&;h3a18001e|
|PidTagOfficeLocation|&;h3a19001e|
|PidTagAssistant|&;h3a30001e|
|PidTagBusinessTelephoneNumber|&;h3a08001e|

