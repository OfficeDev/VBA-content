---
title: Worksheet.Hyperlinks Property (Excel)
keywords: vbaxl10.chm175140
f1_keywords:
- vbaxl10.chm175140
ms.prod: excel
api_name:
- Excel.Worksheet.Hyperlinks
ms.assetid: ac2fe50a-23a0-9982-d448-b18a91092624
ms.date: 06/08/2017
---


# Worksheet.Hyperlinks Property (Excel)

Returns a  **[Hyperlinks](hyperlinks-object-excel.md)** collection that represents the hyperlinks for the worksheet.


## Syntax

 _expression_ . **Hyperlinks**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example checks to see whether any of the hyperlinks on worksheet one contain the word "Microsoft."


```vb
For Each h in Worksheets(1).Hyperlinks 
 If Instr(h.Name, "Microsoft") <> 0 Then h.Follow 
Next
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

