---
title: RemoteItem.MessageClass Property (Outlook)
keywords: vbaol11.chm1601
f1_keywords:
- vbaol11.chm1601
ms.prod: outlook
api_name:
- Outlook.RemoteItem.MessageClass
ms.assetid: cdb17ebc-ea8a-31b1-ef32-e9e4dda872c7
ms.date: 06/08/2017
---


# RemoteItem.MessageClass Property (Outlook)

Returns or sets a  **String** representing the message class for the Outlook item. Read/write.


## Syntax

 _expression_ . **MessageClass**

 _expression_ A variable that represents a **RemoteItem** object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass** . The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


#### Concepts


[RemoteItem Object](remoteitem-object-outlook.md)
#### Other resources


[Item Types and Message Classes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)


