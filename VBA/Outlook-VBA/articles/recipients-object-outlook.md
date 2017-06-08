---
title: Recipients Object (Outlook)
keywords: vbaol11.chm225
f1_keywords:
- vbaol11.chm225
ms.prod: outlook
api_name:
- Outlook.Recipients
ms.assetid: 774f56b7-4de8-9584-60cd-4fbf361f4c85
ms.date: 06/08/2017
---


# Recipients Object (Outlook)

Contains a collection of  **[Recipient](recipient-object-outlook.md)** objects for an Outlook item.


## Remarks

Use the  **Recipients** property to return the **Recipients** object of an **[AppointmentItem](appointmentitem-object-outlook.md)**, **[JournalItem](http://msdn.microsoft.com/library/6e850295-39f9-47b8-e866-9622e9958c69%28Office.15%29.aspx)**, **[MailItem](http://msdn.microsoft.com/library/14197346-05d2-0250-fa4c-4a6b07daf25f%28Office.15%29.aspx)**, **[MeetingItem](meetingitem-object-outlook.md)**, or **[TaskItem](taskitem-object-outlook.md)** object.

Use the  **[Add](http://msdn.microsoft.com/library/7c285291-0f92-ca8d-1c7b-a71ace83ac84%28Office.15%29.aspx)** method to create a new **Recipient** object and add it to the **Recipients** object. The **[Type](http://msdn.microsoft.com/library/3bdc616c-f008-ec95-0a92-0f704eedee34%28Office.15%29.aspx)** property of a new **Recipient** object is set to the default for the associated **AppointmentItem**, **JournalItem**, **MailItem**, or **TaskItem** object and must be reset to indicate another recipient type.

Use  **Recipients** ( _index_ ), where _index_ is the name or index number, to return a single **Recipient** object. The name can be a string representing the display name, the alias, or the full SMTP e-mail address of the recipient.


## Example

The following example creates a new  **MailItem** object and adds Jon Grande as the recipient using the default type ("To").


```
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande")
```

The following example creates the same  **MailItem** object as the preceding example, and then changes the type of the **Recipient** object from the default ("To") to CC.




```
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande") 
 
myRecipient.Type = olCC
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/7c285291-0f92-ca8d-1c7b-a71ace83ac84%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/7cfad374-519e-4312-9050-8a8b66b3911e%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/f5357d32-4901-fb96-3555-f9ef4d5bf3b1%28Office.15%29.aspx)|
|[ResolveAll](http://msdn.microsoft.com/library/82404dc6-af4e-f32d-68b2-9451328b5ca6%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/e8f5d72b-d3f6-6f83-f3f9-496278377c84%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/d83f6ca2-e77f-bfa5-b32b-ed52f023c701%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/3e96321d-a329-7460-0d25-4dc928de0441%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/99dcaedf-f971-48f8-7d6b-75d1ab48d540%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/41ddda3c-ca79-9387-b416-8334aeecc102%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[Recipients Object Members](http://msdn.microsoft.com/library/958f9e6d-c499-4c19-0550-02506998b125%28Office.15%29.aspx)
