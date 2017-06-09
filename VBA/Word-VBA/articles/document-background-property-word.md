---
title: Document.Background Property (Word)
keywords: vbawd10.chm158007365
f1_keywords:
- vbawd10.chm158007365
ms.prod: word
api_name:
- Word.Document.Background
ms.assetid: 0425d9e6-1c26-3df7-bac6-6bc314a3ca47
ms.date: 06/08/2017
---


# Document.Background Property (Word)

Returns a  **Shape** object that represents the background image for the specified document. Read-only.


## Syntax

 _expression_ . **Background**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Backgrounds are visible only in Web layout view.


## Example

This example sets the background color for Web layout view to light gray for the active window.


```vb
ActiveDocument.ActiveWindow.View.Type = wdWebView 
With ActiveDocument.Background.Fill 
 .Visible = True 
 .ForeColor.RGB = RGB(192, 192, 192) 
End With
```

This example sets the background bitmap image of Web layout view to Bubbles.bmp.




```vb
ActiveDocument.ActiveWindow.View.Type = wdWebView 
ActiveDocument.Background.Fill.UserPicture _ 
 PictureFile:="C:\Windows\Bubbles.bmp"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

