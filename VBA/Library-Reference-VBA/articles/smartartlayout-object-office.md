---
title: SmartArtLayout Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.SmartArtLayout
ms.assetid: f8d9db83-86f7-4830-096d-5d15368ab6b1
---


# SmartArtLayout Object (Office)

Represents a Smart Art diagram.


## Remarks

Choices include Basic Block List, Picture Caption List, Vertical Bulleted List, etc.


## Example

The following code changes the diagram style of a Smart Art diagram in Microsoft PowerPoint.


```
ActivePresentation.Slides(1).Shapes(1).SmartArt.Layout = Application.SmartArtLayouts(1)
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/4834cf2d-413e-bfcc-e824-d95b4a33c6c1%28Office.15%29.aspx)|
|[Category](http://msdn.microsoft.com/library/1981f073-1407-b27c-d388-55d9cb51c7f1%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/6951dc2d-92d2-5359-5f32-b22d24385d94%28Office.15%29.aspx)|
|[Description](http://msdn.microsoft.com/library/233e9a68-a546-b97f-5e88-8f338bb351e7%28Office.15%29.aspx)|
|[Id](http://msdn.microsoft.com/library/c8cd4332-6011-3ab7-a65c-f4f60240b2fd%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/66246cd6-7c1d-8777-7505-bec29d2678b7%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/d7d8c5b0-63de-bda1-8376-5587abbf971f%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[SmartArtLayout Object Members](http://msdn.microsoft.com/library/addb351f-b586-c4a1-e3d2-ad170e0ed750%28Office.15%29.aspx)
