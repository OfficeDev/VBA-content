---
title: Workbook.EnvelopeVisible Property (Excel)
keywords: vbaxl10.chm199191
f1_keywords:
- vbaxl10.chm199191
ms.prod: excel
api_name:
- Excel.Workbook.EnvelopeVisible
ms.assetid: d511a75a-ddd1-64f5-a09b-720657f64c09
ms.date: 06/08/2017
---


# Workbook.EnvelopeVisible Property (Excel)

 **True** if the e-mail composition header and the envelope toolbar are both visible. Read/write **Boolean** .


## Syntax

 _expression_ . **EnvelopeVisible**

 _expression_ A variable that represents a **Workbook** object.


## Example

This example checks to see whether the e-mail composition header and the envelope toolbar are visible in the first workbook. If they are visible, the example then sets the variable  `strSubject` to the text of the e-mail subject line.


```vb
If Workbooks(1).EnvelopeVisible = True Then 
 strSubject = "Please read: Review immediately" 
End If
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

