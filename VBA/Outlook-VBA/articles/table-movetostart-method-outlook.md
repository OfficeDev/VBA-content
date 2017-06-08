---
title: Table.MoveToStart Method (Outlook)
keywords: vbaol11.chm2233
f1_keywords:
- vbaol11.chm2233
ms.prod: outlook
api_name:
- Outlook.Table.MoveToStart
ms.assetid: af499471-dd21-9374-7399-3ce977368015
ms.date: 06/08/2017
---


# Table.MoveToStart Method (Outlook)

Moves the current row of the  **[Table](table-object-outlook.md)** to just before the first row of the **Table** .


## Syntax

 _expression_ . **MoveToStart**

 _expression_ A variable that represents a **Table** object.


## Remarks

 **MoveToStart** is equivalent to resetting the **Table** . If you call **[GetNextRow](table-getnextrow-method-outlook.md)** after you call **MoveToStart** , you will return a row representing the first row in the **Table** .


## See also


#### Concepts


[Table Object](table-object-outlook.md)

