---
title: TabStops2 Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.TabStops2
ms.assetid: 1d1d8054-19eb-cd65-f37d-36e93e7fc347
---


# TabStops2 Object (Office)

The collection of  **TabStop2** objects.


## Remarks

Tab stops are indexed numerically from left to right along the ruler.


## Example

 The following example removes the first custom tab stop from the first paragraph in the active Microsoft Publisher publication.


```vb
Sub ClearTabStop() 
    ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
        .ParagraphFormat.Tabs(1).Clear 
End Sub
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

