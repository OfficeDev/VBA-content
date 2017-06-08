---
title: Slides.AddSlide Method (PowerPoint)
keywords: vbapp10.chm530009
f1_keywords:
- vbapp10.chm530009
ms.prod: powerpoint
api_name:
- PowerPoint.Slides.AddSlide
ms.assetid: e8981122-325b-f1c3-c8d5-8e44427961ce
ms.date: 06/08/2017
---


# Slides.AddSlide Method (PowerPoint)

Creates a new slide, adds it to the  **[Slides](slides-object-powerpoint.md)** collection, and returns the slide.


## Syntax

 _expression_. **AddSlide**( **_Index_**, **_pCustomLayout_** )

 _expression_ An expression that returns a **Slides** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Int**|The index of the slide to be added.|
| _pCustomLayout_|Required|**CustomLayout**|The layout of the slide.|

### Return Value

Slide


## Example

The following example shows how to use the  **Add** method to add a new slide to the **Slides** collection. It adds a new slide in index position 2 that has the same layout as the first slide in the active presentation.


```vb
Public Sub Add_Example() 
 
    Dim pptSlide As Slide 
    Dim pptLayout As CustomLayout 
 
    Set pptLayout = ActivePresentation.Slides(1).CustomLayout 
    Set pptSlide = ActivePresentation.Slides.AddSlide(2, pptLayout) 
 
End Sub
```


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.Slides.Add(int, Microsoft.Office.Interop.PowerPoint.PpSlideLayout)**
    
-  **Microsoft.Office.Interop.PowerPoint.Slides.AddSlide(int, Microsoft.Office.Interop.PowerPoint.CustomLayout)**
    

## See also


#### Concepts


[Slides Object](slides-object-powerpoint.md)

