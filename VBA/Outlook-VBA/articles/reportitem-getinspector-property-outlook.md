---
title: ReportItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1649
f1_keywords:
- vbaol11.chm1649
ms.prod: outlook
api_name:
- Outlook.ReportItem.GetInspector
ms.assetid: 2a9ec97b-56c5-f93c-eb42-7ddb93a4697e
ms.date: 06/08/2017
---


# ReportItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **ReportItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[ReportItem Object](reportitem-object-outlook.md)

