---
title: ShapeRange.Copy Method (Publisher)
keywords: vbapb10.chm2293778
f1_keywords:
- vbapb10.chm2293778
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Copy
ms.assetid: 11b9da00-85e4-fc7a-fa93-4a451b7bd15a
ms.date: 06/08/2017
---


# ShapeRange.Copy Method (Publisher)

Copies the specified object to the Clipboard.


## Syntax

 _expression_. **Copy**

 _expression_A variable that represents a  **ShapeRange** object.


### Return Value

Nothing


## Remarks

Use the  **Paste**method to paste the contents of the Clipboard.

The  **Copy** method can be used on **Shape** objects, but the **Paste** method cannot.


## Example

This example copies shapes one and two on page one of the active publication to the Clipboard and then pastes the copies onto page two.


```vb
With ActiveDocument 
 .Pages(1).Shapes.Range(Array(1, 2)).Copy 
 .Pages(2).Shapes.Paste 
End With
```

This example copies shape one on page one of the active publication to the Clipboard.




```vb
ActiveDocument.Pages(1).Shapes(1).Copy
```

This example copies the text in shape one on page one of the active publication to the Clipboard.




```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.Copy
```


