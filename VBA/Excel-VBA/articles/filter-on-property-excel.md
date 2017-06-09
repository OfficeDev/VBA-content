---
title: Filter.On Property (Excel)
keywords: vbaxl10.chm542073
f1_keywords:
- vbaxl10.chm542073
ms.prod: excel
api_name:
- Excel.Filter.On
ms.assetid: 3e325750-2fdc-631f-e116-90769958366c
ms.date: 06/08/2017
---


# Filter.On Property (Excel)

 **True** if the specified filter is on. Read-only **Boolean** .


## Syntax

 _expression_ . **On**

 _expression_ A variable that represents a **Filter** object.


## Example

The following example sets a variable to the value of the  **Criteria1** property of the filter for the first column in the filtered range on the Crew worksheet.


```vb
With Worksheets("Crew") 
 If .AutoFilterMode Then 
 With .AutoFilter.Filters(1) 
 If .On Then c1 = .Criteria1 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Filter Object](filter-object-excel.md)

