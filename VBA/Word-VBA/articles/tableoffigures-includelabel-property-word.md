---
title: TableOfFigures.IncludeLabel Property (Word)
keywords: vbawd10.chm153157634
f1_keywords:
- vbawd10.chm153157634
ms.prod: word
api_name:
- Word.TableOfFigures.IncludeLabel
ms.assetid: b31b8ecf-348d-33c4-ef20-92b2680f1a78
ms.date: 06/08/2017
---


# TableOfFigures.IncludeLabel Property (Word)

 **True** if the caption label and caption number are included in a table of figures. Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeLabel**

 _expression_ An expression that returns a **[TableOfFigures](tableoffigures-object-word.md)** object.


## Example

This example formats the first table of figures in the active document to exclude caption labels (Figure 1, for example).


```vb
If ActiveDocument.TablesOfFigures.Count >= 1 Then 
 ActiveDocument.TablesOfFigures(1).IncludeLabel = False 
End If
```

This example adds a table of figures in place of the selection and then formats the table to include caption labels.




```vb
Dim tofTemp As TableOfFigures 
 
Set tofTemp = ActiveDocument.TablesOfFigures _ 
 .Add(Range:=Selection.Range, _ 
 Caption:="Figure") 
 
tofTemp.IncludeLabel = True
```


## See also


#### Concepts


[TableOfFigures Object](tableoffigures-object-word.md)

