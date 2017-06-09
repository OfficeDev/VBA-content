---
title: Report.TextWidth Method (Access)
keywords: vbaac10.chm13786
f1_keywords:
- vbaac10.chm13786
ms.prod: access
api_name:
- Access.Report.TextWidth
ms.assetid: 98827373-8610-5e48-ab46-2c89f8e2d2a7
ms.date: 06/08/2017
---


# Report.TextWidth Method (Access)

The  **TextWidth** method returns the width of a text string as it would be printed in the current font of a **[Report](report-object-access.md)** object.


## Syntax

 _expression_. **TextWidth**( ** _Expr_** )

 _expression_ A variable that represents a **Report** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Expr_|Required|**String**|The text string for which the text width will be determined.|

### Return Value

Single


## Remarks

You can use the  **TextWidth** method to determine the amount of horizontal space a text string will require in the current font when the report is formatted and printed. For example, a text string formatted in 9-point Arial will require a different amount of space than one formatted in 12-point Courier. To determine the current font and font size for text in a report, check the settings for the report's **FontName** and **FontSize** properties.

The value returned by the  **TextWidth** method is expressed in terms of the coordinate system in effect for the report, as defined by the **Scale** method. You can use the **ScaleMode** property to determine the coordinate system currently in effect for the report.

If the  _strexpr_ argument contains embedded carriage returns, the **TextWidth** method returns the width of the longest line, from the beginning of the line to the carriage return. You can use the value returned by the **TextWidth** method to calculate the necessary space and positioning for multiple lines of text within a report.


## See also


#### Concepts


[Report Object](report-object-access.md)

