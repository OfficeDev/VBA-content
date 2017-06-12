---
title: ContentControl.LockContents Property (Word)
keywords: vbawd10.chm266534915
f1_keywords:
- vbawd10.chm266534915
ms.prod: word
api_name:
- Word.ContentControl.LockContents
ms.assetid: 8d4a68dc-01c8-0f0f-5adf-7b53b4fe3ffc
ms.date: 06/08/2017
---


# ContentControl.LockContents Property (Word)

Returns or sets a  **Boolean** that represents whether the user can edit the contents of a content control. Read/write.


## Syntax

 _expression_ . **LockContents**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

The default value of this property is  **False** . This property corresponds to the **Contents cannot be edited** check box in the **Content Control Properties** dialog box.


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

