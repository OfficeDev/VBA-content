---
title: PostItem.RTFBody Property (Outlook)
keywords: vbaol11.chm3527
f1_keywords:
- vbaol11.chm3527
ms.prod: outlook
api_name:
- Outlook.PostItem.RTFBody
ms.assetid: 79d197b0-d994-374f-ff25-ed7146352ba9
ms.date: 06/08/2017
---


# PostItem.RTFBody Property (Outlook)

Returns or sets a  **Byte** array that represents the body of the Microsoft Outlook item in Rich Text Format. Read/write.


## Syntax

 _expression_ . **RTFBody**

 _expression_ A variable that represents a **[PostItem](postitem-object-outlook.md)** object.


## Remarks

You can use the  **StrConv** function in Microsoft Visual Basic for Applications (VBA), or the **System.Text.Encoding.AsciiEncoding.GetString()** method in C# or Visual Basic to convert an array of bytes to a string.


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

