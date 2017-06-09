---
title: Application.XMLValidationError Event (Word)
keywords: vbawd10.chm4000026
f1_keywords:
- vbawd10.chm4000026
ms.prod: word
api_name:
- Word.Application.XMLValidationError
ms.assetid: bb75a555-fb5e-fb7b-f152-4c6436ecb1c7
ms.date: 06/08/2017
---


# Application.XMLValidationError Event (Word)

Occurs when there is a validation error in the document.


## Syntax

Private Sub  _expression_ _**XMLValidationError**( **_XMLNode As XMLNode_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object. An object of type **Application** that has been declared in a class module by using the **WithEvents** keyword. For more information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XMLNode_|Required| **XMLNode**|The XML element that is invalid.|

## Example

The following example displays an error message to the user when a node is invalid.


```vb
Private Sub Wrd_XMLValidationError(ByVal XMLNode As XMLNode) 
 MsgBox "The " &; UCase(XMLNode.BaseName) &; " element is invalid." &; _ 
 vbCrLf &; XMLNode.ValidationErrorText 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

