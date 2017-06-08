---
title: Table.GetNextRow Method (Outlook)
keywords: vbaol11.chm2231
f1_keywords:
- vbaol11.chm2231
ms.prod: outlook
api_name:
- Outlook.Table.GetNextRow
ms.assetid: e01ddaa0-a869-2f52-5e46-84d4d4090e61
ms.date: 06/08/2017
---


# Table.GetNextRow Method (Outlook)

Moves the current row to the next row in the  **[Table](table-object-outlook.md)** and obtains that row in the **Table** .


## Syntax

 _expression_ . **GetNextRow**

 _expression_ A variable that represents a **Table** object.


### Return Value

A  **[Row](row-object-outlook.md)** object that represents the next valid row in the **Table** if there are additional rows available. If there are no additional rows available (where **[Table.EndOfTable](table-endoftable-property-outlook.md)** is **True** ), **GetNextRow** will return **Null** ( **Nothing** in Visual Basic).


## See also


#### Concepts


[Table Object](table-object-outlook.md)

