---
title: Slide.PublishSlides Method (PowerPoint)
keywords: vbapp10.chm531040
f1_keywords:
- vbapp10.chm531040
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.PublishSlides
ms.assetid: 76f7bd2a-f48c-33e5-52dc-ae9757a880db
ms.date: 06/08/2017
---


# Slide.PublishSlides Method (PowerPoint)

Publishes the specified slide to the specified location.


## Syntax

 _expression_. **PublishSlides**( **_SlideLibraryUrl_**, **_Overwrite_**, **_UseSlideOrder_** )

 _expression_ An expression that returns a **Slide** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SlideLibraryUrl_|Required|**String**|The URL to which to publish the slide library.|
| _Overwrite_|Optional|**Boolean**|Whether to overwrite existing content at SlideLibraryURL. The default is  **False**.|
| _UseSlideOrder_|Optional|**Boolean**|Whether to use the existing slide order. The default is  **False**.|

### Return Value

Nothing


## Example

The following example shows how to publish slide one in the active presentation to a specific URL. Before running this code, substitute a valid URL for  _myURL_.


```vb
Public Sub PublishSlides_Example()



    Dim pptSlide As Slide

        

    Set pptSlide = ActivePresentation.Slides(1)

    pptSlide.PublishSlides ("http://myURL ")



End Sub
```


## See also


#### Concepts


[Slide Object](slide-object-powerpoint.md)

