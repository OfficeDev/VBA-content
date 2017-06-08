---
title: TableStyle.AllowBreakAcrossPage Property (Word)
keywords: vbawd10.chm244776973
f1_keywords:
- vbawd10.chm244776973
ms.prod: word
api_name:
- Word.TableStyle.AllowBreakAcrossPage
ms.assetid: 22ca3964-79ba-dd92-1898-0746f73f4d8b
ms.date: 06/08/2017
---


# TableStyle.AllowBreakAcrossPage Property (Word)

Sets or returns a  **Long** indicating whether lines in the rows of tables formatted with a specified style break across pages. Read/write.


## Syntax

 _expression_ . **AllowBreakAcrossPage**

 _expression_ A variable that represents a **[TableStyle](tablestyle-object-word.md)** object.


## Remarks

 **True** to break the lines in table rows across page breaks. **False** to keep the lines in a row of a table all on the same page. The default setting is **True** .


## Example

This example formats rows in tables formatted with the "Table Grid" style to not break at page breaks.


```vb
Sub DontSplitRows() 
 ActiveDocument.Styles("Table Grid") _ 
 .Table.AllowBreakAcrossPage = False 
End Sub
```


## See also


#### Concepts


[TableStyle Object](tablestyle-object-word.md)

