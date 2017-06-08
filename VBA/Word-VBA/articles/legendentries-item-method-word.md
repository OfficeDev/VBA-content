---
title: LegendEntries.Item Method (Word)
keywords: vbawd10.chm6815744
f1_keywords:
- vbawd10.chm6815744
ms.prod: word
api_name:
- Word.LegendEntries.Item
ms.assetid: 52c5b905-0f5b-38c9-edf3-46018e4f4ecb
ms.date: 06/08/2017
---


# LegendEntries.Item Method (Word)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **[LegendEntries](legendentries-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number for the object.|

### Return Value

A  **[LegendEntry](legendentry-object-word.md)** object that the collection contains.


## Example

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


[LegendEntries Object](legendentries-object-word.md)

