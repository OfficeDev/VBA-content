---
title: Action.MessageClass Property (Outlook)
keywords: vbaol11.chm16
f1_keywords:
- vbaol11.chm16
ms.prod: outlook
api_name:
- Outlook.Action.MessageClass
ms.assetid: a1a1eaeb-2772-babc-18ba-28ce9a66500b
ms.date: 06/08/2017
---


# Action.MessageClass Property (Outlook)

Returns or sets a  **String** representing the message class for the **[Action](action-object-outlook.md)** . Read/write.


## Syntax

 _expression_ . **MessageClass**

 _expression_ A variable that represents an **Action** object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass** . The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


#### Concepts


[Action Object](action-object-outlook.md)
#### Other resources


[Item Types and Message Classes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)


