---
title: Refer to All the Cells on the Worksheet
keywords: vbaxl10.chm5204419
f1_keywords:
- vbaxl10.chm5204419
ms.prod: excel
ms.assetid: fbed1840-e9eb-a7a0-f780-f98939e9bac6
ms.date: 06/08/2017
---


# Refer to All the Cells on the Worksheet

When you apply the  **Cells** property to a worksheet without specifying an index number, the method returns a **Range** object that represents all the cells on the worksheet. The following **Sub** procedure clears the contents from all the cells on Sheet1 in the active workbook.


```vb
Sub ClearSheet() 
 Worksheets("Sheet1").Cells.ClearContents 
End Sub
```


