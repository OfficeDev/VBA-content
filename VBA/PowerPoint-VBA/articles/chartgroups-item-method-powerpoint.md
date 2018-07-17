---
title: ChartGroups.Item Method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroups.Item
ms.assetid: 0b04a471-d726-f400-062c-8d4a7dc9c752
ms.date: 06/08/2017
---


# ChartGroups.Item Method (PowerPoint)

Returns a single object from a collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ A variable that represents a **[ChartGroups](chartgroups-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The index number for the object.|

### Return Value

A  **[ChartGroup](chartgroup-object-powerpoint.md)** object contained by the collection.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds drop lines to chart group one for the first chart group of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartGroups.Item(1).HasDropLines = True

    End If

End With
```


## See also


#### Concepts


[ChartGroups Object](chartgroups-object-powerpoint.md)

