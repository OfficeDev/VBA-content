---
title: Workbook.WriteReserved Property (Excel)
keywords: vbaxl10.chm199167
f1_keywords:
- vbaxl10.chm199167
ms.prod: excel
api_name:
- Excel.Workbook.WriteReserved
ms.assetid: 96cc86d1-0e77-b6f3-3045-f6346de0f969
ms.date: 06/08/2017
---


# Workbook.WriteReserved Property (Excel)

 **True** if the workbook is write-reserved. Read-only **Boolean** .


## Syntax

 _expression_ . **WriteReserved**

 _expression_ A variable that represents a **Workbook** object.


## Remarks

Use the  **[SaveAs](workbook-saveas-method-excel.md)** method to set this property.


## Example

If the active workbook is write-reserved, this example displays a message that contains the name of the user who saved the workbook as write-reserved.


```vb
With ActiveWorkbook 
 If .WriteReserved = True Then 
 MsgBox "Please contact " &; .WriteReservedBy &; Chr(13) &; _ 
 " if you need to insert data in this workbook." 
 End If 
End With
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

