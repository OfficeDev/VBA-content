---
title: TextRange.Copy Method (Publisher)
keywords: vbapb10.chm5308480
f1_keywords:
- vbapb10.chm5308480
ms.prod: publisher
api_name:
- Publisher.TextRange.Copy
ms.assetid: e0d92492-fa0e-9424-471d-09866402702c
ms.date: 06/08/2017
---


# TextRange.Copy Method (Publisher)

Copies the specified object to the Clipboard.


## Syntax

 _expression_. **Copy**

 _expression_A variable that represents a  **TextRange** object.


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


