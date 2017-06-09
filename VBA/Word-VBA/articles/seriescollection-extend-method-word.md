---
title: SeriesCollection.Extend Method (Word)
keywords: vbawd10.chm150405347
f1_keywords:
- vbawd10.chm150405347
ms.prod: word
api_name:
- Word.SeriesCollection.Extend
ms.assetid: 6358fc57-394c-4982-c9b4-8ed2b256f5ea
ms.date: 06/08/2017
---


# SeriesCollection.Extend Method (Word)

Adds new data points to an existing series collection.


## Syntax

 _expression_ . **Extend**( **_Source_** , **_Rowcol_** , **_CategoryLabels_** )

 _expression_ A variable that represents a **[SeriesCollection](seriescollection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **Variant**|The new data to be added to the  **SeriesCollection** object, represented as an A1-style range reference.|
| _Rowcol_|Optional| **Variant**|One of the  **[XlRowCol](xlrowcol-enumeration-word.md)** enumeration values that specifies whether the new values are in the rows or columns of the given range source. If this argument is omitted, Microsoft Word attempts to determine where the values are by the size and orientation of the selected range or by the dimensions of the array.|
| _CategoryLabels_|Optional| **Variant**| **True** to have the first row or column contain the name of the category labels. **False** to have the first row or column contain the first data point of the series. If this argument is omitted, Word attempts to determine the location of the category label from the contents of the first row or column.|

## Remarks

This method is not available for PivotChart reports.


## Example

The following example extends the series on the first chart in the active document by adding the data in cells B1:B6 from the linked workbook.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection.Extend _ 
 Source:="B1:B6" 
 End If 
End With
```


## See also


#### Concepts


[SeriesCollection Object](seriescollection-object-word.md)

