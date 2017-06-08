---
title: DocumentItem.MessageClass Property (Outlook)
keywords: vbaol11.chm1198
f1_keywords:
- vbaol11.chm1198
ms.prod: outlook
api_name:
- Outlook.DocumentItem.MessageClass
ms.assetid: 635ba15e-cacc-4e3e-0824-8ca4dfca2a82
ms.date: 06/08/2017
---


# DocumentItem.MessageClass Property (Outlook)

Returns or sets a  **String** representing the message class for the Outlook item. Read/write.


## Syntax

 _expression_ . **MessageClass**

 _expression_ A variable that represents a **DocumentItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass** . The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


#### Concepts


[DocumentItem Object](documentitem-object-outlook.md)
#### Other resources


[Item Types and Message Classes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)


