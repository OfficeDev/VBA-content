---
title: Refer to More Than One Sheet
keywords: vbaxl10.chm5204432
f1_keywords:
- vbaxl10.chm5204432
ms.prod: excel
ms.assetid: 70641be2-04fc-d8d7-631b-c87e6c270957
ms.date: 06/08/2017
---


# Refer to More Than One Sheet

Use the  **Array** function to identify a group of sheets. The following example selects three sheets in the active workbook.


```vb
Sub Several() 
 Worksheets(Array("Sheet1", "Sheet2", "Sheet4")).Select 
End Sub
```


