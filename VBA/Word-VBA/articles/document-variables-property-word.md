---
title: Document.Variables Property (Word)
keywords: vbawd10.chm158007322
f1_keywords:
- vbawd10.chm158007322
ms.prod: word
api_name:
- Word.Document.Variables
ms.assetid: 93af7b84-f172-6ebd-2147-e7ebc92449c5
ms.date: 06/08/2017
---


# Document.Variables Property (Word)

Returns a  **[Variables](variables-object-word.md)** collection that represents the variables stored in the specified document. Read-only.


## Syntax

 _expression_ . **Variables**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example adds a document variable named "Value1" to the active document. The example then retrieves the value from the Value1 variable, adds 3 to the value, and displays the results.


```vb
ActiveDocument.Variables.Add Name:="Value1", Value:="1" 
MsgBox ActiveDocument.Variables("Value1") + 3
```

This example displays the name and value of each document variable in the active document.




```vb
For Each myVar In ActiveDocument.Variables 
 MsgBox "Name =" &; myVar.Name &; vbCr &; "Value = " &; myVar.Value 
Next myVar
```


## See also


#### Concepts


[Document Object](document-object-word.md)

