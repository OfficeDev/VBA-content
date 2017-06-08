---
title: PostItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1524
f1_keywords:
- vbaol11.chm1524
ms.prod: outlook
api_name:
- Outlook.PostItem.GetInspector
ms.assetid: 705fe03b-2ff4-8ed8-e3c2-fb7d52444169
ms.date: 06/08/2017
---


# PostItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **PostItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

