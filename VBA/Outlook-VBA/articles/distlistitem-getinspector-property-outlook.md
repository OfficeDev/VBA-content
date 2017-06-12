---
title: DistListItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1125
f1_keywords:
- vbaol11.chm1125
ms.prod: outlook
api_name:
- Outlook.DistListItem.GetInspector
ms.assetid: 2ffab19b-17a3-0de0-f9dd-3a8fbfea8efd
ms.date: 06/08/2017
---


# DistListItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **DistListItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

