---
title: Selection.Expand Method (Word)
keywords: vbawd10.chm158662785
f1_keywords:
- vbawd10.chm158662785
ms.prod: word
api_name:
- Word.Selection.Expand
ms.assetid: 8b716453-7656-e8b8-f6b0-0dc97ef2714d
ms.date: 06/08/2017
---


# Selection.Expand Method (Word)

Expands the specified range or selection. Returns the number of characters added to the range or selection.  **Long** .


## Syntax

 _expression_ . **Expand**( **_Unit_** )

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Unit_|Optional| **Variant**|A  **[WdUnits](wdunits-enumeration-word.md)** constant that represents the unit by which to expand the range. The default value is **wdWord** .|

## Example

This example capitalizes the first character in the selection and then expands the selection to include the entire sentence.


```vb
With Selection 
 .Characters(1).Case = wdTitleSentence 
 .Expand Unit:=wdSentence 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

