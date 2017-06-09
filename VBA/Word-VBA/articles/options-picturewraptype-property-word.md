---
title: Options.PictureWrapType Property (Word)
keywords: vbawd10.chm162988468
f1_keywords:
- vbawd10.chm162988468
ms.prod: word
api_name:
- Word.Options.PictureWrapType
ms.assetid: bb0cc23d-d58c-c506-c6f9-e4ccf5f2a8ac
ms.date: 06/08/2017
---


# Options.PictureWrapType Property (Word)

Sets or returns a  **WdWrapTypeMerged** that indicates how Microsoft Word wraps text around pictures. Read/write.


## Syntax

 _expression_ . **PictureWrapType**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Remarks

This is a default option setting and affects all pictures inserted unless picture wrapping is individually defined for a picture.


## Example

This example sets Word to insert and paste all pictures inline with the text if inline is not already specified.


```vb
Sub PicWrap() 
 With Application.Options 
 If .PictureWrapType <> wdWrapMergeInline Then 
 .PictureWrapType = wdWrapMergeInline 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

