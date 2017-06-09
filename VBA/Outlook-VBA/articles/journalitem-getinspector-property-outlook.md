---
title: JournalItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1242
f1_keywords:
- vbaol11.chm1242
ms.prod: outlook
api_name:
- Outlook.JournalItem.GetInspector
ms.assetid: 49d173ba-e4fd-e9c4-12b4-423a4c60ec46
ms.date: 06/08/2017
---


# JournalItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **JournalItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[JournalItem Object](journalitem-object-outlook.md)

