---
title: Rows.HeadingFormat Property (Word)
keywords: vbawd10.chm155975685
f1_keywords:
- vbawd10.chm155975685
ms.prod: word
api_name:
- Word.Rows.HeadingFormat
ms.assetid: 225464d2-cb93-f347-6f02-ace23c4177eb
ms.date: 06/08/2017
---


# Rows.HeadingFormat Property (Word)

 **True** if the specified row or rows are formatted as a table heading. Read/write **Long** .


## Syntax

 _expression_ . **HeadingFormat**

 _expression_ A variable that represents a **[Rows](rows-object-word.md)** collection.


## Remarks

Rows formatted as table headings are repeated when a table spans more than one page. Can be  **True** , **False** or **wdUndefined** .


## Example

This example creates a 5x5 table at the beginning of the active document and then adds the table heading format to the first table row.


```vb
Dim rngTemp As Range 
Dim tableNew As Table 
 
Set rngTemp = ActiveDocument.Range(0, 0) 
Set tableNewe = ActiveDocument.Tables.Add(rngTemp, 5, 5) 
 
tableNew.Rows.HeadingFormat = True
```

This example determines whether the row that contains the insertion point is formatted as a table heading.




```vb
If Selection.Information(wdWithInTable) = True Then 
 If Selection.Rows.HeadingFormat = True Then _ 
 MsgBox "The current row is a table heading" 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

