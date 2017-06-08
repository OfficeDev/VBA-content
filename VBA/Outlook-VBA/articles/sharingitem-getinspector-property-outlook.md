---
title: SharingItem.GetInspector Property (Outlook)
keywords: vbaol11.chm608
f1_keywords:
- vbaol11.chm608
ms.prod: outlook
api_name:
- Outlook.SharingItem.GetInspector
ms.assetid: 960f9b66-35dc-54ab-13c3-9ea54802bccf
ms.date: 06/08/2017
---


# SharingItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified **[SharingItem](sharingitem-object-outlook.md)** . Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **SharingItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

