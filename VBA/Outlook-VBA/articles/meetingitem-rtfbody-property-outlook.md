---
title: MeetingItem.RTFBody Property (Outlook)
keywords: vbaol11.chm3526
f1_keywords:
- vbaol11.chm3526
ms.prod: outlook
api_name:
- Outlook.MeetingItem.RTFBody
ms.assetid: 4bf67ee1-f0bc-92b8-948f-2de7807a1dd3
ms.date: 06/08/2017
---


# MeetingItem.RTFBody Property (Outlook)

Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.


## Syntax

 _expression_ . **RTFBody**

 _expression_ A variable that represents a **[MeetingItem](meetingitem-object-outlook.md)** object.


## Remarks

You can use the  **StrConv** function in Microsoft Visual Basic for Applications (VBA), or the **System.Text.Encoding.AsciiEncoding.GetString()** method in C# or Visual Basic to convert an array of bytes to a string.


## See also


#### Concepts


[MeetingItem Object](meetingitem-object-outlook.md)

