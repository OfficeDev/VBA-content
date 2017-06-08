---
title: Workbook.EndReview Method (Excel)
keywords: vbaxl10.chm199208
f1_keywords:
- vbaxl10.chm199208
ms.prod: excel
api_name:
- Excel.Workbook.EndReview
ms.assetid: cd4a445b-4731-43ba-e46a-f80f19ea5a17
ms.date: 06/08/2017
---


# Workbook.EndReview Method (Excel)

Terminates a review of a file that has been sent for review using the  **[SendForReview](workbook-sendforreview-method-excel.md)** method.


## Syntax

 _expression_ . **EndReview**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example terminates the review of the active workbook. When executed, this procedure displays a message asking if you want to end the review. This example assumes the active workbook has been sent for review.


```vb
Sub EndWorkbookRev() 
 
 ActiveWorkbook.EndReview 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

