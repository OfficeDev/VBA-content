---
title: ListGallery.Reset Method (Word)
keywords: vbawd10.chm160694372
f1_keywords:
- vbawd10.chm160694372
ms.prod: word
api_name:
- Word.ListGallery.Reset
ms.assetid: 456ed895-6e6e-334d-7cab-9df4376d8025
ms.date: 06/08/2017
---


# ListGallery.Reset Method (Word)

Resets the list template specified by Index for the specified list gallery to the built-in list template format.


## Syntax

 _expression_ . **Reset**( **_Index_** )

 _expression_ Required. A variable that represents a **[ListGallery](listgallery-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The template to reset.|

## Example

This example sets the fourth format listed on the Numbered tab in the Bullets and Numbering dialog box back to the built-in numbering format, and then it applies the list template to the selection.


```
ListGalleries(wdNumberGallery).Reset(4) 
Selection.Range.ListFormat.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(2).ListTemplates(4)
```

This example resets all the list templates in the Bullets and Numbering dialog box back to the built-in formats.




```vb
For Each lg In ListGalleries 
 For i = 1 to 7 
 lg.Reset Index:=i 
 Next i 
Next lg 
 

```


## See also


#### Concepts


[ListGallery Object](listgallery-object-word.md)

