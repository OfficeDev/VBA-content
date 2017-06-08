---
title: SlicerCaches Object (Excel)
keywords: vbaxl10.chm894072
f1_keywords:
- vbaxl10.chm894072
ms.prod: excel
api_name:
- Excel.SlicerCaches
ms.assetid: d6097f70-cdc7-3be7-575c-cf43a0765e10
ms.date: 06/08/2017
---


# SlicerCaches Object (Excel)

Represents the collection of slicer caches associated with the specified workbook.


## Remarks

Use the  **[Item](slicercaches-item-property-excel.md)** property of the **SlicerCaches** collection to return a **[SlicerCache](slicercache-object-excel.md)** object associated with the specified **[Workbook](workbook-object-excel.md)** object. A **SlicerCache** object can be retrieved by using either the value of the **[Index](slicercache-index-property-excel.md)** property or the **[Name](slicercache-name-property-excel.md)** property of the specified object.


## Example

The following code example retrieves the  **SlicerCache** object that represents the slicer cache associated with the Country slicer.


```
ActiveWorkbook.SlicerCaches("Slicer_Country")
```


## Methods



|**Name**|
|:-----|
|[Add2](slicercaches-add-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](slicercaches-application-property-excel.md)|
|[Count](slicercaches-count-property-excel.md)|
|[Creator](slicercaches-creator-property-excel.md)|
|[Item](slicercaches-item-property-excel.md)|
|[Parent](slicercaches-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
