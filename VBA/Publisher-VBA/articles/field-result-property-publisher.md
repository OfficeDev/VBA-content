---
title: Field.Result Property (Publisher)
keywords: vbapb10.chm6094855
f1_keywords:
- vbapb10.chm6094855
ms.prod: publisher
api_name:
- Publisher.Field.Result
ms.assetid: 213e123e-90a7-32b8-1dcf-37da61a8a7e7
ms.date: 06/08/2017
---


# Field.Result Property (Publisher)

Returns a  **String** that represents the result of the specified field. Read-only.


## Syntax

 _expression_. **Result**

 _expression_A variable that represents a  **Field** object.


### Return Value

String


## Example

This example applies bold formatting to the first field in the selection. This example assumes that either text or a shape with text is selected in the active publication.


```vb
Sub GetFieldResults() 
 If Selection.TextRange.Fields.Count > 0 Then 
 MsgBox "The result of the first field is " &; _ 
 Selection.TextRange.Fields(1).Result &; "." 
 End If 
End Sub
```


