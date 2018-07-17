---
title: ContentControl.LockContentControl Property (Word)
keywords: vbawd10.chm266534914
f1_keywords:
- vbawd10.chm266534914
ms.prod: word
api_name:
- Word.ContentControl.LockContentControl
ms.assetid: a567f2a5-a3db-446c-e694-50dbfbb3e9db
ms.date: 06/08/2017
---


# ContentControl.LockContentControl Property (Word)

Returns or sets a  **Boolean** that represents whether the user can delete a content control from the active document. Read/write.


## Syntax

 _expression_ . **LockContentControl**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

The default value of this property is  **False** . This property corresponds to the **Content control cannot be deleted** check box in the **Content Control Properties** dialog box.


 **Note**  You cannot set this property if the  **[Temporary](contentcontrol-temporary-property-word.md)** property is set to **True** .


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

