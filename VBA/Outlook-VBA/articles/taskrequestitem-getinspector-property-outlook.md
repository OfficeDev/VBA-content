---
title: TaskRequestItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1883
f1_keywords:
- vbaol11.chm1883
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.GetInspector
ms.assetid: 114a879a-9e5c-5f90-0621-082348dab1df
ms.date: 06/08/2017
---


# TaskRequestItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **TaskRequestItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[TaskRequestItem Object](taskrequestitem-object-outlook.md)

