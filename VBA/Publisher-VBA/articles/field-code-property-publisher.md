---
title: Field.Code Property (Publisher)
keywords: vbapb10.chm6094851
f1_keywords:
- vbapb10.chm6094851
ms.prod: publisher
api_name:
- Publisher.Field.Code
ms.assetid: bb2f3b23-dea1-bdfb-90bf-4b4ea09570f6
ms.date: 06/08/2017
---


# Field.Code Property (Publisher)

Returns a  **String** that represents the text displayed when the page view is set to show field codes. Read-only.


## Syntax

 _expression_. **Code**

 _expression_A variable that represents a  **Field** object.


### Return Value

String


## Example

This example loops through all the fields in the active publication, and then displays a message as to whether the string "www" was found in the code of any of the fields.


```vb
Sub FindWWWHyperlinks() 
 Dim intItem As Integer 
 Dim intField As Integer 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.Fields 
 Do 
 intItem = intItem + 1 
 If InStr(1, .Item(intItem).Code, "www") > 0 Then 
 intField = intField + 1 
 End If 
 Loop Until intItem = .Count 
 End With 
 
 If intField > 0 Then 
 MsgBox "You have " &; intField &; " World Wide Web " &; _ 
 "hyperlinks in your publication." 
 Else 
 MsgBox "You have no hyperlink fields in your publication." 
 End If 
End Sub
```


