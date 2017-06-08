---
title: Global.IsObjectValid Property (Word)
keywords: vbawd10.chm163119213
f1_keywords:
- vbawd10.chm163119213
ms.prod: word
api_name:
- Word.Global.IsObjectValid
ms.assetid: 73115443-ad95-8e58-cd35-b9a34c6e641d
ms.date: 06/08/2017
---


# Global.IsObjectValid Property (Word)

 **True** if the specified variable that references an object is valid. Read-only **Boolean** .


## Syntax

 _expression_ . **IsObjectValid**( **_Object_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Object_|Required| **Object**|A variable that references an object.|

## Remarks

This property returns  **False** if the object referenced by the variable has been deleted.


## Example

This example adds a table to the active document and assigns it to the variable  _aTable_ . The example then deletes the first table from the document. If the table that _aTable_ refers to was not the first table in the document (that is, if _aTable_ is still a valid object), the example also removes any borders from that table.


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


[Global Object](global-object-word.md)

