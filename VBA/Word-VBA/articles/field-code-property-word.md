---
title: Field.Code Property (Word)
keywords: vbawd10.chm154075136
f1_keywords:
- vbawd10.chm154075136
ms.prod: word
api_name:
- Word.Field.Code
ms.assetid: 4273619f-184c-a964-6c0d-14fec927ec01
ms.date: 06/08/2017
---


# Field.Code Property (Word)

Returns a  **[Range](range-object-word.md)** object that represents a field's code. Read/write.


## Syntax

 _expression_ . **Code**

 _expression_ A variable that represents a **[Field](field-object-word.md)** object.


## Remarks

A field's code is everything that's enclosed by the field characters ( **{ }** ) including the leading space and trailing space characters. You can access a field's code without changing the view from field results.


## Example

This example displays the field code for each field in the active document.


```vb
Dim fieldLoop As Field 
 
For Each fieldLoop In ActiveDocument.Fields 
 MsgBox Chr(34) &; fieldLoop.Code.Text &; Chr(34) 
Next fieldLoop
```

This example changes the field code for the first field in the active document to CREATEDATE.




```vb
Dim rngTemp As Range 
 
Set rngTemp = ActiveDocument.Fields(1).Code 
rngTemp.Text = " CREATEDATE " 
ActiveDocument.Fields(1).Update
```

This example determines whether the active document includes a mail merge field named "Title."




```vb
Dim fieldLoop As Field 
 
For Each fieldLoop In ActiveDocument.MailMerge.Fields 
 If InStr(1, fieldLoop.Code.Text, "Title", 1) Then 
 MsgBox "A Title merge field is in this document" 
 End If 
Next fieldLoop
```


## See also


#### Concepts


[Field Object](field-object-word.md)

