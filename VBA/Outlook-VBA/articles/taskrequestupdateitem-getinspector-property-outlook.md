---
title: TaskRequestUpdateItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1932
f1_keywords:
- vbaol11.chm1932
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.GetInspector
ms.assetid: 9542e72b-9b9d-be7a-5c2f-1c4a653be4d7
ms.date: 06/08/2017
---


# TaskRequestUpdateItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **TaskRequestUpdateItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


#### Concepts


[TaskRequestUpdateItem Object](taskrequestupdateitem-object-outlook.md)

