---
title: Font2 Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.Font2
ms.assetid: 8e892c52-56d9-72bd-2893-b15a17cd59ae
---


# Font2 Object (Office)

Contains font attributes (for example, font name, font size, and color) for an object.


## Example

The following example changes the formatting of the Heading 2 style in the active document to Arial and bold.


```vb
With ActiveDocument.Styles(wdStyleHeading2).Font2 
 .Name = "Arial" 
 .Italic = True 
End With 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

