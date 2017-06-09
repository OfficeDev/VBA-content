---
title: RemoteItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1597
f1_keywords:
- vbaol11.chm1597
ms.prod: outlook
api_name:
- Outlook.RemoteItem.GetInspector
ms.assetid: 0f8e0621-7094-afd5-8913-9f42d55765e0
ms.date: 06/08/2017
---


# RemoteItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **RemoteItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[RemoteItem Object](remoteitem-object-outlook.md)

