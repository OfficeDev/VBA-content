---
title: CustomLayout Object (PowerPoint)
keywords: vbapp10.chm672000
f1_keywords:
- vbapp10.chm672000
ms.prod: powerpoint
api_name:
- PowerPoint.CustomLayout
ms.assetid: 67829704-0314-aed2-5415-6736cefc197e
ms.date: 06/08/2017
---


# CustomLayout Object (PowerPoint)

Represents a custom layout associated with a presentation design. The  **CustomLayout** object is a member of the **[CustomLayouts](customlayouts-object-powerpoint.md)** collection.


## Remarks

Use the  **CustomLayout** property of the **[Slide](slide-object-powerpoint.md)** or **[SlideRange](http://msdn.microsoft.com/library/440ab59d-744a-209f-bf28-d0acd3a21e1a%28Office.15%29.aspx)** objects to access a **CustomLayout** object, for example:


```
ActiveWindow.Selection.SlideRange(1).CustomLayout
```


```
ActivePresentation.Slides(1).CustomLayout
```

Use the  **[Add](http://msdn.microsoft.com/library/d22dc23a-cb03-ab32-fd27-e360377369a9%28Office.15%29.aspx)** method of the **CustomLayouts** collection to add a new custom layout to the presentation design's custom layouts. Use the **[Item](http://msdn.microsoft.com/library/d22dc23a-cb03-ab32-fd27-e360377369a9%28Office.15%29.aspx)** method to refer to a custom layout. Use the **[Paste](http://msdn.microsoft.com/library/d4fcd2db-3d6b-0c59-6ea3-f9aadf90ed04%28Office.15%29.aspx)** method to paste the slides on the Clipboard into a custom layout and add the custom layout to the **CustomLayouts** collection.


## Methods



|**Name**|
|:-----|
|[Copy](http://msdn.microsoft.com/library/6ad8ab68-0e94-761e-d352-96eb2f8f795c%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/e27f9ba5-d933-5e2d-e71c-e1757941bde1%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/31f678ea-768c-d7c7-7ea9-7007f6e12ad4%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/c4e0703e-5cd8-c305-bbc9-71b845ff4aba%28Office.15%29.aspx)|
|[MoveTo](http://msdn.microsoft.com/library/0efa5d50-0dd8-bcaa-5c05-1493c40c5b45%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/066d394e-2e5d-0d34-7bf5-438e3b72d908%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/91126d20-bfbb-fcad-72e0-fb10d78a17ab%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/61141722-d851-b3ff-f426-0865a6e31850%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/a589362b-d987-f2ed-79f2-0e0afd9ae051%28Office.15%29.aspx)|
|[Design](http://msdn.microsoft.com/library/9630b24c-57fb-29a6-0126-cebf384015bd%28Office.15%29.aspx)|
|[DisplayMasterShapes](http://msdn.microsoft.com/library/07790f9c-fad7-7086-5d18-80fd6bf0658b%28Office.15%29.aspx)|
|[FollowMasterBackground](http://msdn.microsoft.com/library/9554e610-8d9a-ab32-411e-0f4aa40a7f19%28Office.15%29.aspx)|
|[Guides](http://msdn.microsoft.com/library/30230637-f357-506b-2cb3-621fb08bd36c%28Office.15%29.aspx)|
|[HeadersFooters](http://msdn.microsoft.com/library/e8a53212-99cb-26df-12dd-ec6a6c7b7116%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/7ba167ab-72dc-f482-aa7d-f0804cac895d%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/834448c1-2acf-33b4-15c9-eb485d9c176c%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/bdbb922f-db6d-034e-b08b-08c9dd500a3b%28Office.15%29.aspx)|
|[MatchingName](http://msdn.microsoft.com/library/ff661ecd-37c7-5ea1-3bba-93e0d56aa66e%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/3ad36d6e-1b85-8ff2-9b76-f50a372e0f07%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/373ab10a-71c8-fefb-1d5f-67c19abbc679%28Office.15%29.aspx)|
|[Preserved](http://msdn.microsoft.com/library/8a686dd4-2a03-6e56-650c-fc9b52f14b24%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/ed8c332c-c69e-93e4-2611-96b015a0114d%28Office.15%29.aspx)|
|[SlideShowTransition](http://msdn.microsoft.com/library/f165346b-4ad3-035b-a9be-141dc7666958%28Office.15%29.aspx)|
|[ThemeColorScheme](http://msdn.microsoft.com/library/c60258b6-5119-ee70-0d81-60c7a7869c34%28Office.15%29.aspx)|
|[TimeLine](http://msdn.microsoft.com/library/641ccad6-2a91-64d7-2884-1ab436c58b9e%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/cddb5c12-7ee9-9ad3-6534-45f0388f2d08%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
