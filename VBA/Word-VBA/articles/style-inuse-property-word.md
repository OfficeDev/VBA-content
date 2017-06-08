---
title: Style.InUse Property (Word)
keywords: vbawd10.chm153878534
f1_keywords:
- vbawd10.chm153878534
ms.prod: word
api_name:
- Word.Style.InUse
ms.assetid: 6fbba751-f549-4175-6c1a-ec1f9abb478a
ms.date: 06/08/2017
---


# Style.InUse Property (Word)

 **True** if the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document. Read-only **Boolean** .


## Syntax

 _expression_ . **InUse**

 _expression_ An expression that returns a **[Style](style-object-word.md)** object.


## Remarks

The  **InUse** property doesn't necessarily indicate whether the style is currently applied to any text in the document. For instance, if text that's been formatted with a style is deleted, the **InUse** property of the style remains **True** . For built-in styles that have never been used in the document, this property returns **False** .


## Example

This example displays a message box that lists the names of all the styles that are currently being used in the active document.


```vb
Dim docActive As Document 
Dim strMessage As String 
Dim styleLoop As Style 
 
Set docActive = ActiveDocument 
 
strMessage = "Styles in use:" &; vbCr 
 
For Each styleLoop In docActive.Styles 
 If styleLoop.InUse = True Then 
 With docActive 
 .Content.Find 
 .ClearFormatting 
 .Text = "" 
 .Style = styleLoop 
 .Execute Format:=True 
 If .Found = True Then 
 strMessage = strMessage &; styleLoop.Name &; vbCr 
 End If 
 End With 
 End If 
Next styleLoop 
 
MsgBox strMessage
```


## See also


#### Concepts


[Style Object](style-object-word.md)

