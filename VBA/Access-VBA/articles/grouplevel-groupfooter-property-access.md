---
title: GroupLevel.GroupFooter Property (Access)
keywords: vbaac10.chm12242
f1_keywords:
- vbaac10.chm12242
ms.prod: access
api_name:
- Access.GroupLevel.GroupFooter
ms.assetid: c10d30b2-da18-cd6f-8b00-e964cd4751d6
ms.date: 06/08/2017
---


# GroupLevel.GroupFooter Property (Access)

You can use the  **GroupFooter** property to create a group footer for a selected field or expression in a report. Read/write **Boolean**.


## Syntax

 _expression_. **GroupFooter**

 _expression_ A variable that represents a **GroupLevel** object.


## Remarks

You can use group headers and footers to label or summarize data in a group of records. For example, if you set the  **GroupHeader** property to Yes for the Categories field, each group of products will begin with its category name.


 **Note**  You can't set or refer to these properties directly in Visual Basic. To create a group header or footer for a field or expression in Visual Basic, use the  **[CreateGroupLevel](application-creategrouplevel-method-access.md)** method.

To set the grouping properties —  **[GroupOn](grouplevel-groupon-property-access.md)**, **[GroupInterval](grouplevel-groupinterval-property-access.md)**, and **[KeepTogether](grouplevel-keeptogether-property-access.md)** — to other than their default values, you must first set the **GroupHeader** or **GroupFooter** property or both to Yes for the selected field or expression.


## See also


#### Concepts


[GroupLevel Object](grouplevel-object-access.md)

