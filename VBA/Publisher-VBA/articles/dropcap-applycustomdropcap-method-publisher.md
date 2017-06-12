---
title: DropCap.ApplyCustomDropCap Method (Publisher)
keywords: vbapb10.chm5505041
f1_keywords:
- vbapb10.chm5505041
ms.prod: publisher
api_name:
- Publisher.DropCap.ApplyCustomDropCap
ms.assetid: 906cf476-3826-8510-315f-425f6f50a92a
ms.date: 06/08/2017
---


# DropCap.ApplyCustomDropCap Method (Publisher)

Applies custom formatting to the first letters of paragraphs in a text frame.


## Syntax

 _expression_. **ApplyCustomDropCap**( **_LinesUp_**,  **_Size_**,  **_Span_**,  **_FontName_**,  **_Bold_**,  **_Italic_**)

 _expression_A variable that represents a  **DropCap** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|LinesUp|Optional| **Long**|The number of lines to move up the drop cap. The default is 0. The maximum number cannot be more than the number entered for the Size argument less one.|
|Size|Optional| **Long**|The size of the drop cap letters in number of lines high. The default is 5.|
|Span|Optional| **Long**|The number of letters included in the drop cap. The default is 1.|
|FontName|Optional| **String**|The name of the font to format the drop cap. The default is the current font.|
|Bold|Optional| **Boolean**| **True** to bold the drop cap. The default is **False**.|
|Italic|Optional| **Boolean**| **True** to italicize the drop cap. The default is **False**.|

## Example

This example formats the first three letters of the paragraphs in the specified text box.


```vb
Sub CustDropCap() 
 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.DropCap _ 
 .ApplyCustomDropCap LinesUp:=1, Size:=6, Span:=3, _ 
 FontName:="Script MT Bold", Bold:=True, Italic:=True 
 
End Sub
```


