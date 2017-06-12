---
title: FormDescription.MessageClass Property (Outlook)
keywords: vbaol11.chm191
f1_keywords:
- vbaol11.chm191
ms.prod: outlook
api_name:
- Outlook.FormDescription.MessageClass
ms.assetid: 51ab2c14-de92-b029-e5b8-2e158a626319
ms.date: 06/08/2017
---


# FormDescription.MessageClass Property (Outlook)

Returns a  **String** representing the message class for the **[FormDescription](formdescription-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **MessageClass**

 _expression_ A variable that represents a **FormDescription** object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass** . The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


#### Concepts


[FormDescription Object](formdescription-object-outlook.md)
#### Other resources


[Item Types and Message Classes](http://msdn.microsoft.com/library/15b709cc-7486-b6c7-88a3-4a4d8e0ab292%28Office.15%29.aspx)


