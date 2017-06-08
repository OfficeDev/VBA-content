---
title: NoteItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1482
f1_keywords:
- vbaol11.chm1482
ms.prod: outlook
api_name:
- Outlook.NoteItem.GetInspector
ms.assetid: 80e5bdc5-8161-afa7-6aab-65356fc5d2ea
ms.date: 06/08/2017
---


# NoteItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **NoteItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[NoteItem Object](noteitem-object-outlook.md)

