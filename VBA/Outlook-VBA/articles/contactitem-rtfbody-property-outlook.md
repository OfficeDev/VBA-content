---
title: ContactItem.RTFBody Property (Outlook)
keywords: vbaol11.chm3525
f1_keywords:
- vbaol11.chm3525
ms.prod: outlook
api_name:
- Outlook.ContactItem.RTFBody
ms.assetid: f8e7e632-113b-a50e-211b-dbd182221168
ms.date: 06/08/2017
---


# ContactItem.RTFBody Property (Outlook)

Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.


## Syntax

 _expression_ . **RTFBody**

 _expression_ A variable that represents a **[ContactItem](contactitem-object-outlook.md)** object.


## Remarks

You can use the  **StrConv** function in Microsoft Visual Basic for Applications (VBA), or the **System.Text.Encoding.AsciiEncoding.GetString()** method in C# or Visual Basic to convert an array of bytes to a string.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

