---
title: TaskRequestDeclineItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1834
f1_keywords:
- vbaol11.chm1834
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.GetInspector
ms.assetid: 8892e56a-275d-b9df-9d9d-bbfd39b98c33
ms.date: 06/08/2017
---


# TaskRequestDeclineItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **TaskRequestDeclineItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[TaskRequestDeclineItem Object](taskrequestdeclineitem-object-outlook.md)

