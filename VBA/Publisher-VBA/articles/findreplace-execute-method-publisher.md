---
title: FindReplace.Execute Method (Publisher)
keywords: vbapb10.chm8323086
f1_keywords:
- vbapb10.chm8323086
ms.prod: publisher
api_name:
- Publisher.FindReplace.Execute
ms.assetid: 351a64ab-3c6c-c9c9-7ffe-b60b73d390ae
ms.date: 06/08/2017
---


# FindReplace.Execute Method (Publisher)

Performs the specified Find or Replace operation.


## Syntax

 _expression_. **Execute**

 _expression_A variable that represents a  **FindReplace** object.


### Return Value

Boolean


## Example

This example executes a Find and Replace operation on the active document.


```vb
Sub ExecuteFindReplace() 
 Dim objFindReplace As FindReplace 
 Set objFindReplace = ActiveDocument.Find 
 With objFindReplace 
 .Clear 
 .FindText = "library" 
 .Execute 
 End With 
End Sub
```


