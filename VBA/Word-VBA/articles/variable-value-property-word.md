---
title: Variable.Value Property (Word)
keywords: vbawd10.chm157679616
f1_keywords:
- vbawd10.chm157679616
ms.prod: word
api_name:
- Word.Variable.Value
ms.assetid: 6a687fff-062a-4e27-abc7-2f49d6f9c76b
ms.date: 06/08/2017
---


# Variable.Value Property (Word)

Returns or sets the value of a document variable. Read/write  **String** .


## Syntax

 _expression_ . **Value**

 _expression_ Required. A variable that represents a **[Variable](variable-object-word.md)** object.


## Example

This example adds a document variable to the active document and then displays the value of the new variable.


```vb
ActiveDocument.Variables.Add Name:="Temp2", Value:="10" 
MsgBox ActiveDocument.Variables("Temp2").Value
```


 **Caution**  The value of a  **[Variable](variable-object-word.md)** object cannot be set to a zero-length string. Setting a **[Variable](variable-object-word.md)** object to a zero-length string results in a run time error.


## See also


#### Concepts


[Variable Object](variable-object-word.md)

