---
title: ContactItem.GetInspector Property (Outlook)
keywords: vbaol11.chm941
f1_keywords:
- vbaol11.chm941
ms.prod: outlook
api_name:
- Outlook.ContactItem.GetInspector
ms.assetid: d1f8530f-f797-413f-92cb-d0e8215de0e4
ms.date: 06/08/2017
---


# ContactItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **ContactItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

