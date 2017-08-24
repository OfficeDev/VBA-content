---
title: TextRange.InsertSymbol Method (Publisher)
keywords: vbapb10.chm5308452
f1_keywords:
- vbapb10.chm5308452
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertSymbol
ms.assetid: 607d12da-5a2d-4e0e-b45e-92275ce97bab
ms.date: 06/08/2017
---


# TextRange.InsertSymbol Method (Publisher)

Returns a  **[TextRange](textrange-object-publisher.md)** object that represents a symbol inserted in place of the specified range or selection.


## Syntax

 _expression_. **InsertSymbol**( **_FontName_**,  **_CharIndex_**)

 _expression_A variable that represents a  **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|FontName|Required| **String**|The name of the font that contains the symbol.|
|CharIndex|Required| **Long**|The Unicode character for the specified symbol.|

### Return Value

TextRange


## Remarks

If you do not want to replace the range or selection, use the  [TextRange.Collapse Method (Publisher)](textrange-collapse-method-publisher.md) before you use this method.


## Example

This example inserts a double-headed arrow at the cursor.


```vb
Sub Insert Arrow() 
    ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
            .Paragraphs(Start:=1, Length:=1).Select
    With .TextFrame.TextRange 
            .InsertPageNumber 
            .Collapse Direction:= pbCollapseStart
            .InsertSymbol FontName:="Symbol", CharIndex:=171
        End With 
End Sub
```


