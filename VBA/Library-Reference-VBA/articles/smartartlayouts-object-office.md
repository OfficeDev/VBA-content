---
title: SmartArtLayouts Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArtLayouts
ms.assetid: 25e33439-fb5e-01d7-1b85-01884a42ba68
---


# SmartArtLayouts Object (Office)

Represents a collection of Smart Art layout diagrams.


## Remarks

Choices include Basic Block List, Picture Caption List, Vertical Bulleted List, etc.


## Example

The following code changes the diagram style of a Smart Art diagram in Microsoft PowerPoint.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## Methods



|**Name**|
|:-----|
|[Item](http://msdn.microsoft.com/library/8741eb7f-21d4-dfff-ef02-a87959d8a841%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/884b8508-1860-f21f-a3f7-b236909b9efa%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/caf73afe-63e5-0832-deb9-c608b7b1b41a%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/d68e64ff-541e-7276-b04e-a33a002e73bc%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/cb32827a-8109-ea95-6f49-abd34a391770%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[SmartArtLayouts Object Members](http://msdn.microsoft.com/library/29154639-17b7-7999-a9e1-b16cf9b2ada6%28Office.15%29.aspx)
