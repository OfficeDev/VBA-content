---
title: Recipient Object (Outlook)
keywords: vbaol11.chm2339
f1_keywords:
- vbaol11.chm2339
ms.prod: outlook
api_name:
- Outlook.Recipient
ms.assetid: 8cee4d79-ec55-52a4-710b-6456944ca86d
ms.date: 06/08/2017
---


# Recipient Object (Outlook)

Represents a user or resource in Outlook, generally a mail or mobile message addressee.


## Remarks

Use the  **[Recipients](http://msdn.microsoft.com/library/7cfad374-519e-4312-9050-8a8b66b3911e%28Office.15%29.aspx)** ( _index_ ) method, where _index_ is the name or index number, to return a single **Recipient** object. The name can be a string that represents the display name, the alias, the full SMTP e-mail address, or the mobile phone number of the recipient. A good practice is to use the SMTP e-mail address for a mail message, and the mobile phone number for a mobile message.

Use the  **[Add](http://msdn.microsoft.com/library/7c285291-0f92-ca8d-1c7b-a71ace83ac84%28Office.15%29.aspx)** method to create a new **Recipient** object and add it to the **[Recipients](recipients-object-outlook.md)** object. The **[Type](http://msdn.microsoft.com/library/3bdc616c-f008-ec95-0a92-0f704eedee34%28Office.15%29.aspx)** property of a new **Recipient** object is set to the default value for the associated **[AppointmentItem](appointmentitem-object-outlook.md)**, **[JournalItem](http://msdn.microsoft.com/library/6e850295-39f9-47b8-e866-9622e9958c69%28Office.15%29.aspx)**, **[MailItem](http://msdn.microsoft.com/library/14197346-05d2-0250-fa4c-4a6b07daf25f%28Office.15%29.aspx)**, **[MeetingItem](meetingitem-object-outlook.md)**, or **[TaskItem](taskitem-object-outlook.md)** object and must be reset to indicate another recipient type.


## Example



The following Visual Basic for Applications (VBA) example creates a new  **MailItem** object and adds Jon Grande as the recipient by using the default type ("To").




```
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande")
```

The following VBA example creates the same  **MailItem** object as the preceding example, and then changes the type of the **Recipient** object from the default (To) to CC.




```
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande") 
 
myRecipient.Type = olCC
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/726577e1-b91d-0127-adb8-069a648ee220%28Office.15%29.aspx)|
|[FreeBusy](http://msdn.microsoft.com/library/eeb831bc-c369-10f1-fb0b-08a8105c48e6%28Office.15%29.aspx)|
|[Resolve](http://msdn.microsoft.com/library/2c4f9243-2e31-642e-78a7-fe74cd73b385%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Address](http://msdn.microsoft.com/library/8e14f39a-0000-1039-bb0b-7726d7828a68%28Office.15%29.aspx)|
|[AddressEntry](http://msdn.microsoft.com/library/3b2b524e-4dd5-9ff4-98cc-811746ea0453%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/6968733a-a307-49f5-ba78-c0a1ac573803%28Office.15%29.aspx)|
|[AutoResponse](http://msdn.microsoft.com/library/db6e0658-8e12-ac0b-4317-396cfe4620f6%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/1e6aa19a-16ee-7835-c2fb-f5523e8614c4%28Office.15%29.aspx)|
|[DisplayType](http://msdn.microsoft.com/library/1109138d-ef1b-deec-13cc-8443d03e825c%28Office.15%29.aspx)|
|[EntryID](http://msdn.microsoft.com/library/f71d384c-6e1c-f96c-1415-cf21a0c26712%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/fe2ef09a-0046-1f82-e2ad-2e4cbb5a403f%28Office.15%29.aspx)|
|[MeetingResponseStatus](http://msdn.microsoft.com/library/27f3e40a-b5e9-9f36-ae26-78cc85d160fa%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/c444a728-3c1d-efd5-036e-d14fb2e7164a%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/fa37d562-af43-26f7-b446-fccf510e925a%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/fe10f888-f17a-932e-988b-ed565d6a169f%28Office.15%29.aspx)|
|[Resolved](http://msdn.microsoft.com/library/09c7655b-5acd-b527-56f6-59bc994a5ca1%28Office.15%29.aspx)|
|[Sendable](http://msdn.microsoft.com/library/ba6c3f35-5e51-f502-fb74-5403de3411e9%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/0719e438-c9b0-ecca-1aa0-f25c9b21fe69%28Office.15%29.aspx)|
|[TrackingStatus](http://msdn.microsoft.com/library/15787403-de2c-ee9f-4f8b-587cf1ee6087%28Office.15%29.aspx)|
|[TrackingStatusTime](http://msdn.microsoft.com/library/906fec55-13da-5a83-c4c6-fa2cd07d6d7a%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/3bdc616c-f008-ec95-0a92-0f704eedee34%28Office.15%29.aspx)|

## See also


#### Other resources


[Recipient Object Members](http://msdn.microsoft.com/library/70e34018-95de-7fcf-1331-9be61a8675a2%28Office.15%29.aspx)
[How to: Obtain the E-mail Address of a Recipient](http://msdn.microsoft.com/library/b645c227-a7d2-2861-3bf7-4190a19abe81%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
