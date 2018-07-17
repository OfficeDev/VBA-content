---
title: Variable.Index Property (Word)
keywords: vbawd10.chm157679618
f1_keywords:
- vbawd10.chm157679618
ms.prod: word
api_name:
- Word.Variable.Index
ms.assetid: 203623e2-61ba-a6d4-e1a2-cdb7a05d2857
ms.date: 06/08/2017
---


# Variable.Index Property (Word)

Returns a  **Long** that represents the ordinal position of a variable with in the collection of variables. Read-only.


## Syntax

 _expression_ . **Index**

 _expression_ Required. A variable that represents a **[Variable](variable-object-word.md)** object.


## Example

This example adds a document variable to the active document and then returns the position of the specified variable in the  **Variables** collection.


```vb
Set myVar = ActiveDocument.Variables.Add(Name:="Name", _ 
 Value:="Joe") 
num = myVar.Index
```


## See also


#### Concepts


[Variable Object](variable-object-word.md)

