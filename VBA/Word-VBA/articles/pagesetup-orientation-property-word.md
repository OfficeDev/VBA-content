---
title: PageSetup.Orientation Property (Word)
keywords: vbawd10.chm158400619
f1_keywords:
- vbawd10.chm158400619
ms.prod: word
api_name:
- Word.PageSetup.Orientation
ms.assetid: 7761b95d-b6dc-7f2f-94b9-7e1d45a85498
ms.date: 06/08/2017
---


# PageSetup.Orientation Property (Word)

Returns or sets the orientation of the page. Read/write  **[WdOrientation](wdorientation-enumeration-word.md)** .


## Syntax

 _expression_ . **Orientation**

 _expression_ Required. A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Remarks

Some of the  **WdOrientation** constants may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed.


## Example

This example changes the orientation of the document named "MyDocument.doc" and then prints the document. The example then changes the orientation of the document back to portrait.


```vb
Set myDoc = Documents("MyDocument.doc") 
With myDoc 
 .PageSetup.Orientation = wdOrientLandscape 
 .PrintOut 
 .PageSetup.Orientation = wdOrientPortrait 
End With
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

