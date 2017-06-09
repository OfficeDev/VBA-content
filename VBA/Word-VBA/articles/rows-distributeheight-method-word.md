---
title: Rows.DistributeHeight Method (Word)
keywords: vbawd10.chm155975886
f1_keywords:
- vbawd10.chm155975886
ms.prod: word
api_name:
- Word.Rows.DistributeHeight
ms.assetid: f5fe9eea-debc-c1e4-b9a0-81c5f9a0c04a
ms.date: 06/08/2017
---


# Rows.DistributeHeight Method (Word)

Adjusts the height of the specified rows or cells so that they're equal.


## Syntax

 _expression_ . **DistributeHeight**

 _expression_ Required. A variable that represents a **[Rows](rows-object-word.md)** collection.


## Example

This example adjusts the height of the rows in the first table in the active document so that they're equal.


```vb
ActiveDocument.Tables(1).Rows.DistributeHeight
```

This example adjusts the height of the first three rows in the first table so that they're equal.




```vb
Dim rngTemp As Range 
 
If ActiveDocument.Tables.Count >= 1 Then 
 Set rngTemp = ActiveDocument.Range(Start:=ActiveDocument _ 
 .Tables(1).Rows(1).Range.Start, _ 
 End:=ActiveDocument.Tables(1).Rows(3).Range.End) 
 rngTemp.Rows.DistributeHeight 
End If
```


## See also


#### Concepts


[Rows Collection Object](rows-object-word.md)

