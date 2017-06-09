---
title: SharingItem.MessageClass Property (Outlook)
keywords: vbaol11.chm612
f1_keywords:
- vbaol11.chm612
ms.prod: outlook
api_name:
- Outlook.SharingItem.MessageClass
ms.assetid: d2991917-120f-9d69-156f-793e67f45ed9
ms.date: 06/08/2017
---


# SharingItem.MessageClass Property (Outlook)

Returns or sets a  **String** representing the message class for the **[SharingItem](sharingitem-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **MessageClass**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass** . The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.

The default value for this property is  `IPM.Sharing`.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)
#### Other resources



[Item Types and Message Classes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)

