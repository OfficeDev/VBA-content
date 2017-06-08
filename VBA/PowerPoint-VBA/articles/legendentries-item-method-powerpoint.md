---
title: LegendEntries.Item Method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.LegendEntries.Item
ms.assetid: 67745179-84b3-a2b8-23d8-ceb393828af7
ms.date: 06/08/2017
---


# LegendEntries.Item Method (PowerPoint)

Returns a single object from a collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ A variable that represents a **[LegendEntries](legendentries-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The index number for the object.|

### Return Value

A  **[LegendEntry](legendentry-object-powerpoint.md)** object that the collection contains.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example changes the font for the text of the legend entry at the top of the legend (this is usually the legend for series one) for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Legend.LegendEntries.Item(1). _
            Font.Italic = True
    End If
End With
```


## See also


#### Concepts


[LegendEntries Object](legendentries-object-powerpoint.md)

