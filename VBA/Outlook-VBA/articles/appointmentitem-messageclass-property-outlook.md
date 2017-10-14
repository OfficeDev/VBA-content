---
title: AppointmentItem.MessageClass Property (Outlook)
keywords: vbaol11.chm857
f1_keywords:
- vbaol11.chm857
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.MessageClass
ms.assetid: e98318d9-72e9-0914-83c6-3a05f544874f
ms.date: 06/08/2017
---


# AppointmentItem.MessageClass Property (Outlook)

Returns or sets a  **String** representing the message class for the Outlook item. Read/write.


## Syntax

 _expression_ . **MessageClass**

 _expression_ A variable that represents an **AppointmentItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass** . The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)
#### Other resources


[Item Types and Message Classes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)


