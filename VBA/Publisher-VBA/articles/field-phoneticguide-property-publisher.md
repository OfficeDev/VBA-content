---
title: Field.PhoneticGuide Property (Publisher)
keywords: vbapb10.chm6094856
f1_keywords:
- vbapb10.chm6094856
ms.prod: publisher
api_name:
- Publisher.Field.PhoneticGuide
ms.assetid: c68dba00-56c0-abba-0be8-5bd1a725f678
ms.date: 06/08/2017
---


# Field.PhoneticGuide Property (Publisher)

Returns a  **PhoneticGuide** object that represents the properties of phonetic text applied to a text range.


## Syntax

 _expression_. **PhoneticGuide**

 _expression_A variable that represents a  **Field** object.


### Return Value

PhoneticGuide


## Example

This example adds phonetic text to the selection and displays the text to which the phonetic text applies, which is the originally selected text. This example assumes text is selected. If no text is selected, the message box will be blank.


```vb
Sub AddPhoneticText() 
 With Selection.TextRange.Fields.AddPhoneticGuide _ 
 (Range:=Selection.TextRange, Text:="ver-E nIs") 
 MsgBox "The base text is " &; .PhoneticGuide.BaseText 
 End With 
End Sub
```


