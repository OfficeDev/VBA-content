---
title: LinkFormat.SavePictureWithDocument Property (Word)
keywords: vbawd10.chm154206230
f1_keywords:
- vbawd10.chm154206230
ms.prod: word
api_name:
- Word.LinkFormat.SavePictureWithDocument
ms.assetid: 5aacc0de-7a95-1f95-2797-d84a722526a6
ms.date: 06/08/2017
---


# LinkFormat.SavePictureWithDocument Property (Word)

 **True** if the specified picture is saved with the document. Read/write **Boolean** .


## Syntax

 _expression_ . **SavePictureWithDocument**

 _expression_ An expression that returns a **[LinkFormat](linkformat-object-word.md)** object.


## Remarks

This property works only with shapes and inline shapes that are linked pictures.


## Example

This example saves the linked picture that's defined as the first inline shape in the active document when the document is saved.


```vb
Set myPic = ActiveDocument.InlineShapes(1) 
If myPic.Type = wdInlineShapeLinkedPicture Then 
 myPic.LinkFormat.SavePictureWithDocument = True 
End If
```


## See also


#### Concepts


[LinkFormat Object](linkformat-object-word.md)

