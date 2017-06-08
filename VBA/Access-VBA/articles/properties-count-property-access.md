---
title: Properties.Count Property (Access)
keywords: vbaac10.chm10050
f1_keywords:
- vbaac10.chm10050
ms.prod: access
api_name:
- Access.Properties.Count
ms.assetid: 00a6039e-82bf-7cfe-d7b2-9e9bdb12aa44
ms.date: 06/08/2017
---


# Properties.Count Property (Access)

Returns the number of objects in the specified collection. Read-only.


## Syntax

 _expression_. **Count**

 _expression_ A variable that represents a **Properties** object.


## Remarks

Because members of a collection begin with 0, you should always code loops starting with the 0 member and ending with the value of the  **Count** property minus 1. If you want to loop through the members of a collection without checking the **Count** property, you can use a **For Each...Next** command.

The  **Count** property setting is never Null. If its value is 0, there are no objects in the collection.


