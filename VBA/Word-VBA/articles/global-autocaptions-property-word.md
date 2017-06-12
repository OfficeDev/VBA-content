---
title: Global.AutoCaptions Property (Word)
keywords: vbawd10.chm163119125
f1_keywords:
- vbawd10.chm163119125
ms.prod: word
api_name:
- Word.Global.AutoCaptions
ms.assetid: 88fac2d9-ac54-6f8a-aefd-100438a0ae1e
ms.date: 06/08/2017
---


# Global.AutoCaptions Property (Word)

Returns an  **[AutoCaptions](autocaptions-object-word.md)** collection that represents the captions that are automatically added when items such as tables and pictures are inserted into a document. Read-only.


## Syntax

 _expression_ . **AutoCaptions**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the name of each item that automatically gets a caption when inserted into the document.


```vb
Dim captionLoop as AutoCaption 
 
For Each captionLoop In AutoCaptions 
 If captionLoop.AutoInsert Then MsgBox captionLoop.Name 
Next captionLoop
```


## See also


#### Concepts


[Global Object](global-object-word.md)

