---
title: TabStop2 Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.TabStop2
ms.assetid: fee461a9-684b-e6c2-a74a-d0aa161d0d9c
---


# TabStop2 Object (Office)

Represents a single tab stop. The  **TabStop2** object is a member of the **TabStops2** collection.


## Remarks

Tab stops are indexed numerically from left to right along the ruler.


## Example

The following example removes the first custom tab stop from the selected paragraphs.


```vb
Sub ClearTabStop() 
 Selection.TextRange.ParagraphFormat.Tabs(1).Clear 
End Sub
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

