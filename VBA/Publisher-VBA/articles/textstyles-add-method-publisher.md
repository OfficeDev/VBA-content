---
title: TextStyles.Add Method (Publisher)
keywords: vbapb10.chm5898244
f1_keywords:
- vbapb10.chm5898244
ms.prod: publisher
api_name:
- Publisher.TextStyles.Add
ms.assetid: 56bb84a2-5632-1baa-4b97-3c48d43367bf
ms.date: 06/08/2017
---


# TextStyles.Add Method (Publisher)

Adds a new  **TextStyle** object to the specified **TextStyles** object and returns the new **TextStyle** object.


## Syntax

 _expression_. **Add**( **_Font_**,  **_ParagraphFormat_**,  **_StyleName_**,  **_BasedOn_**)

 _expression_A variable that represents a  **TextStyles** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|StyleName|Required| **String**|The name of the new text style. If the name matches an existing text style, the existing text style is overwritten.|
|BasedOn|Optional| **String**|The name of the text style on which the new text style is based. If the name does not match an existing text style, an error occurs.|
|Font|Optional| **Font**|The font settings to apply to the new text style.|
|ParagraphFormat|Optional| **ParagraphFormat**|The paragraph formatting to apply to the new text style.|

### Return Value

TextStyle


## Example

The following example adds a new text style to the active publication based on the Normal text style.


```vb
Dim tsNew As TextStyle 
 
Set tsNew = ActiveDocument.TextStyles _ 
 .Add(StyleName:="Title", BasedOn:="Normal")
```


