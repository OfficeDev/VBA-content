---
title: Validation.InCellDropdown Property (Excel)
keywords: vbaxl10.chm532077
f1_keywords:
- vbaxl10.chm532077
ms.prod: excel
api_name:
- Excel.Validation.InCellDropdown
ms.assetid: 019cf85b-831f-38f0-ea69-a30066acf30e
ms.date: 06/08/2017
---


# Validation.InCellDropdown Property (Excel)

 **True** if data validation displays a drop-down list that contains acceptable values. Read/write **Boolean** .


## Syntax

 _expression_ . **InCellDropdown**

 _expression_ A variable that represents a **Validation** object.


## Remarks

This property is ignored if the validation type isn't  **xlValidateList** .

Use the  _Formula1_ argument of the **Add** or **Modify** method of the **Validation** object to specify the range that contains valid data.


## Example

This example adds data validation to cell E5. The range A1:A10 contains the acceptable values for the cell and the cell displays a drop-down list that contains those values.


```vb
With Range("e5").Validation 
 .Add xlValidateList, xlValidAlertStop, xlBetween, "=$A$1:$A$10" 
 .InCellDropdown = True 
End With
```


## See also


#### Concepts


[Validation Object](validation-object-excel.md)

