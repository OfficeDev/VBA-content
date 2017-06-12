---
title: Slicer.Copy Method (Excel)
keywords: vbaxl10.chm905091
f1_keywords:
- vbaxl10.chm905091
ms.prod: excel
api_name:
- Excel.Slicer.Copy
ms.assetid: 265e7819-db8b-deab-5ab1-2cc9782cd800
ms.date: 06/08/2017
---


# Slicer.Copy Method (Excel)

Copies the specified slicer to the clipboard.


## Syntax

 _expression_ . **Copy**

 _expression_ A variable that represents a **[Slicer](slicer-object-excel.md)** object.


## Example

The following code example accesses the Customer slicer by using the  **[Range](shapes-range-property-excel.md)** property of the **[Shapes](shapes-object-excel.md)** collection, and then copies and pastes it into the active worksheet.


```vb
ActiveSheet.Shapes.Range(Array("Customer")).Select 
Selection.Copy 
ActiveSheet.Paste 

```

Alternatively, you can perform the same operation by using the  **[Slicers](slicercache-slicers-property-excel.md)** property of the **[SlicerCaches](slicercaches-object-excel.md)** collection to access the slicer, as shown in the following code example.




```vb
ActiveWorkbook.SlicerCaches("Slicer_Customer") _ 
 .Slicers("Customer").Copy 
ActiveSheet.Paste
```


## See also


#### Concepts


[Slicer Object](slicer-object-excel.md)

