---
title: Slicer.Cut Method (Excel)
keywords: vbaxl10.chm905090
f1_keywords:
- vbaxl10.chm905090
ms.prod: excel
api_name:
- Excel.Slicer.Cut
ms.assetid: a8778661-612f-0031-78b0-d59bb87fdf62
ms.date: 06/08/2017
---


# Slicer.Cut Method (Excel)

Cuts the specified slicer and copies it to the clipboard.


## Syntax

 _expression_ . **Cut**

 _expression_ A variable that represents a **[Slicer](slicer-object-excel.md)** object.


## Example

The following code example accesses the Customer slicer by using the  **[Range](shapes-range-property-excel.md)** property of the **[Shapes](shapes-object-excel.md)** collection, and then cuts and pastes it into the active worksheet.


```vb
ActiveSheet.Shapes.Range(Array("Customer")).Select 
Selection.Cut 
ActiveSheet.Paste 

```

Alternatively, you can perform the same operation by using the  **[Slicers](slicercache-slicers-property-excel.md)** property of the **[SlicerCaches](slicercaches-object-excel.md)** collection to access the slicer, as shown in the following code example.




```vb
ActiveWorkbook.SlicerCaches("Slicer_Customer") _ 
 .Slicers("Customer").Cut 
ActiveSheet.Paste
```


## See also


#### Concepts


[Slicer Object](slicer-object-excel.md)

