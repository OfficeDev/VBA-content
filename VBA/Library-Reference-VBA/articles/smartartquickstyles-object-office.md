---
title: SmartArtQuickStyles Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArtQuickStyles
ms.assetid: d488ac12-160b-c518-2b56-cc0a3a45c6b7
---


# SmartArtQuickStyles Object (Office)

Represents a collection of Smart Art quick styles.


## Example

The following code changes the quick style of a Smart Art diagram in Microsoft PowerPoint.


```vb
ActivePresentation.Slides(1).Shapes(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles(i)
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

