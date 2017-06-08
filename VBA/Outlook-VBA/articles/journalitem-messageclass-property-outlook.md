---
title: JournalItem.MessageClass Property (Outlook)
keywords: vbaol11.chm1246
f1_keywords:
- vbaol11.chm1246
ms.prod: outlook
api_name:
- Outlook.JournalItem.MessageClass
ms.assetid: 1a47a08f-d7ba-5627-dfae-c918c74074c4
ms.date: 06/08/2017
---


# JournalItem.MessageClass Property (Outlook)

Returns or sets a  **String** representing the message class for the Outlook item. Read/write.


## Syntax

 _expression_ . **MessageClass**

 _expression_ A variable that represents a **JournalItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass** . The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


#### Concepts


[JournalItem Object](journalitem-object-outlook.md)
#### Other resources



[Item Types and Message Classes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)

