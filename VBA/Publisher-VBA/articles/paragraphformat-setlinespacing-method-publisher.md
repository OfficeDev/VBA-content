---
title: ParagraphFormat.SetLineSpacing Method (Publisher)
keywords: vbapb10.chm5439511
f1_keywords:
- vbapb10.chm5439511
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.SetLineSpacing
ms.assetid: 32e5b233-8415-2373-7423-18b66df3a5ea
ms.date: 06/08/2017
---


# ParagraphFormat.SetLineSpacing Method (Publisher)

Formats the line spacing of specified paragraphs.


## Syntax

 _expression_. **SetLineSpacing**( **_Rule_**,  **_Spacing_**)

 _expression_A variable that represents a  **ParagraphFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Rule|Required| **PbLineSpacingRule**|The line spacing to use for the specified paragraphs.|
|Spacing|Optional| **Variant**|The spacing (in points) for the specified paragraphs.|

## Remarks

The Rule parameter can be one of the  **PbLineSpacingRule** constants declared in the Microsoft Publisher type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **pbLineSpacing1pt5**|Sets the spacing for specified paragraphs to one-and-a-half lines.|
| **pbLineSpacingDouble**| Double-spaces the specified paragraphs.|
| **pbLineSpacingExactly**| Sets the line spacing to exactly the value specified in the Spacing argument, even if a larger font is used within the paragraph.|
| **pbLineSpacingMixed**| A return value for the **[LineSpacing](paragraphformat-linespacing-property-publisher.md)** property that indicates that line spacing is a combination of values for the specified paragraphs.|
| **pbLineSpacingMultiple**|Sets the line spacing to the value specified in the Spacing argument.|
| **pbLineSpacingSingle**|Single spaces the specified paragraphs.|

## Example

This example sets the line spacing to double.


```vb
Sub SetLineSpacingForSelection() 
 Selection.TextRange.ParagraphFormat.SetLineSpacing _ 
 Rule:=pbLineSpacingDouble, Spacing:=12 
End Sub
```


