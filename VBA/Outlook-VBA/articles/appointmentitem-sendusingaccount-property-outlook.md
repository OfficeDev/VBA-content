---
title: AppointmentItem.SendUsingAccount Property (Outlook)
keywords: vbaol11.chm923
f1_keywords:
- vbaol11.chm923
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.SendUsingAccount
ms.assetid: c3a73b32-c2e1-cb32-35e3-e278f78700ad
ms.date: 06/08/2017
---


# AppointmentItem.SendUsingAccount Property (Outlook)

Returns or sets an  **[Account](account-object-outlook.md)** object that represents the account under which the **[AppointmentItem](appointmentitem-object-outlook.md)** is to be sent. Read/write.


## Syntax

 _expression_ . **SendUsingAccount**

 _expression_ An expression that returns a **AppointmentItem** object.


## Remarks

The  **SendUsingAccount** property can be used to specify the account that should be used to send the **AppointmentItem** when the **[Send](taskitem-send-method-outlook.md)** method is called. This property returns **Null** ( **Nothing** in Visual Basic) if the account specified for the **AppointmentItem** no longer exists.


## See also


#### Concepts


[AppointmentItem Object](appointmentitem-object-outlook.md)

