---
title: ContentControl.Range Property (Word)
keywords: vbawd10.chm266534913
f1_keywords:
- vbawd10.chm266534913
ms.prod: word
api_name:
- Word.ContentControl.Range
ms.assetid: e83efa5d-edd7-2cdc-ee6f-925db82e3d40
ms.date: 06/08/2017
---


# ContentControl.Range Property (Word)

Returns a  **[Range](range-object-word.md)** that represents the contents of the content control in the active document. Read-only.


## Syntax

 _expression_ . **Range**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

Use the  **Range** property to access the text, the formatting for the text, and other text properties. For more information, see[Working with Range Objects](http://msdn.microsoft.com/library/9e240aa7-8608-9d70-aee3-2e202687459e%28Office.15%29.aspx).


## Example

The following example inserts a date content control into the active document, and then sets the contents of the content control and specifies that the user cannot edit the contents or delete the control from the document.


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls _ 
 .Add(wdContentControlDate) 
 
objCC.Range.Text = "January 1, 2007" 
objCC.LockContents = True 
objCC.LockContentControl = True
```


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

