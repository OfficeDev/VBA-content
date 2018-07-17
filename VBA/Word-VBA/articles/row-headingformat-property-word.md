---
title: Row.HeadingFormat Property (Word)
keywords: vbawd10.chm156237829
f1_keywords:
- vbawd10.chm156237829
ms.prod: word
api_name:
- Word.Row.HeadingFormat
ms.assetid: 18b0161c-ad04-57a2-02fb-870fabed158b
ms.date: 06/08/2017
---


# Row.HeadingFormat Property (Word)

 **True** if the specified row or rows are formatted as a table heading. Rows formatted as table headings are repeated when a table spans more than one page. Can be **True** , **False** or **wdUndefined** . Read/write **Long** .


## Syntax

 _expression_ . **HeadingFormat**

 _expression_ A variable that represents a **[Row](row-object-word.md)** object.


## Example

This example creates a 5x5 table at the beginning of the active document and then adds the table heading format to the first table row.


```vb
Dim rngTemp As Range 
Dim tableNew As Table 
 
Set rngTemp = ActiveDocument.Range(0, 0) 
Set tableNewe = ActiveDocument.Tables.Add(rngTemp, 5, 5) 
 
tableNew.Rows(1).HeadingFormat = True
```

This example determines whether the row that contains the insertion point is formatted as a table heading.




```vb
If Selection.Information(wdWithInTable) = True Then 
 If Selection.Rows(1).HeadingFormat = True Then _ 
 MsgBox "The current row is a table heading" 
Else 
 MsgBox "The insertion point is not in a table." 
End If
```


## See also


#### Concepts


[Row Object](row-object-word.md)

