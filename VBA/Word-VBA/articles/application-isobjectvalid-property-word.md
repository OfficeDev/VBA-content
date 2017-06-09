---
title: Application.IsObjectValid Property (Word)
keywords: vbawd10.chm158335085
f1_keywords:
- vbawd10.chm158335085
ms.prod: word
api_name:
- Word.Application.IsObjectValid
ms.assetid: 94cb08e4-2a4f-5ebf-25b8-6492e35f5695
ms.date: 06/08/2017
---


# Application.IsObjectValid Property (Word)

 **True** if the specified variable that references an object is valid. Read-only **Boolean** .


## Syntax

 _expression_ . **IsObjectValid**( **_Object_** )

 _expression_ Optional. A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

The  **IsObjectValid** property returns **False** if the object referenced by the variable has been deleted.


## Example

This example adds a table to the active document and assigns it to the variable  `aTable`. The example then deletes the first table from the document. If the table that aTable refers to was not the first table in the document (that is, if  `aTable` is still a valid object), the example also removes any borders from that table.


```vb
Dim aTable As Table 
 
Set aTable = ActiveDocument.Tables.Add(Range:=Selection.Range, _ 
 NumRows:=2, NumColumns:=3) 
 
ActiveDocument.Tables(1).Delete 
If IsObjectValid(aTable) = True Then _ 
 aTable.Borders.Enable = False
```


## See also


#### Concepts


[Application Object](application-object-word.md)

