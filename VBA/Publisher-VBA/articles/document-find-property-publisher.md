---
title: Document.Find Property (Publisher)
keywords: vbapb10.chm196725
f1_keywords:
- vbapb10.chm196725
ms.prod: publisher
api_name:
- Publisher.Document.Find
ms.assetid: e9b31937-4504-79b5-5913-b2ef0a23f2a7
ms.date: 06/08/2017
---


# Document.Find Property (Publisher)

## Syntax

 _expression_. **Find**

 _expression_A variable that represents a  **Document** object.


## Example

As it applies to the  **Document** object.

The following example sets an object variable to the  **FindReplace** object of the active document. A search operation is executed that applies bold formatting to every occurrence of the word "important".




```vb
Dim objFind as FindReplace 
Dim fFound as Boolean 
 
Set objFind = ActiveDocument.Find 
fFound = True 
 
With objFind 
 .Clear 
 .FindText = "important" 
 Do While fFound = True 
 fFound = .Execute 
 If Not .FoundTextRange Is Nothing Then 
 .FoundTextRange.Font.Bold = True 
 End If 
 Loop 
End With 
```

As it applies to the  **TextRange** object.

The following example sets an object variable to the  **FindReplace** object of the text range of the first shape in the active document. A search operation is executed that applies bold formatting to every occurrence of the word "urgent" in the text range.




```vb
Dim objFind as FindReplace 
Dim fFound as Boolean 
 
Set objFind = ActiveDocument.Pages(1) _ 
 .Shapes(1).TextFrame.TextRange.Find 
fFound = True 
 
With objFind 
 .Clear 
 .FindText = "urgent" 
 Do While fFound = True 
 fFound = .Execute 
 If Not .FoundTextRange Is Nothing Then 
 .FoundTextRange.Font.Bold = True 
 End If 
 Loop 
End With
```


