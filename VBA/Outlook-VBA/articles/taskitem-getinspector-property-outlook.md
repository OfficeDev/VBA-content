---
title: TaskItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1697
f1_keywords:
- vbaol11.chm1697
ms.prod: outlook
api_name:
- Outlook.TaskItem.GetInspector
ms.assetid: 2a2faad7-1030-cdd8-8a8d-8018aad3b667
ms.date: 06/08/2017
---


# TaskItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **TaskItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[TaskItem Object](taskitem-object-outlook.md)

