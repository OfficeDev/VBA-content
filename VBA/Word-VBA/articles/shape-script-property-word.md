---
title: Shape.Script Property (Word)
keywords: vbawd10.chm161481207
f1_keywords:
- vbawd10.chm161481207
ms.prod: word
api_name:
- Word.Shape.Script
ms.assetid: d98f64f8-e097-fb56-736f-1247dcbdd3af
ms.date: 06/08/2017
---


# Shape.Script Property (Word)

Returns a  **Script** object, which represents a block of script or code for an image on a Web page.


## Syntax

 _expression_ . **Script**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

If the Web page contains no script, nothing is returned.


## Example

This example displays the type of scripting language used in the first shape in the active document.


```vb
Set objScr = ActiveDocument.Shapes(1).Script 
If Not (objScr Is Nothing) Then 
 Select Case objScr.Language 
 Case msoScriptLanguageVisualBasic 
 MsgBox "VBScript" 
 Case msoScriptLanguageJava 
 MsgBox "JavaScript" 
 Case msoScriptLanguageASP 
 MsgBox "Active Server Pages" 
 Case Else 
 Msgbox "Other scripting language" 
 End Select 
End If
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

