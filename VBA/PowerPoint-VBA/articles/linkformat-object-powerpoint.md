---
title: LinkFormat Object (PowerPoint)
keywords: vbapp10.chm563000
f1_keywords:
- vbapp10.chm563000
ms.prod: powerpoint
api_name:
- PowerPoint.LinkFormat
ms.assetid: e89ee344-4197-ac0d-dd53-966e4672a3ce
ms.date: 06/08/2017
---


# LinkFormat Object (PowerPoint)

Contains properties and methods that apply to linked OLE objects, linked pictures, and IIRC media objects. 


## Example

Use the  **LinkFormat** property to return a **LinkFormat** object. The following example loops through all the shapes on all the slides in the active presentation and sets all linked Microsoft Excel worksheets to be updated manually.


```
For Each sld In ActivePresentation.Slides

    For Each sh In sld.Shapes

        If sh.Type = msoLinkedOLEObject Then

            If sh.OLEFormat.ProgID = "Excel.Sheet" Then

                sh.LinkFormat.AutoUpdate = ppUpdateOptionManual

            End If

        End If

    Next

Next
```


## Methods



|**Name**|
|:-----|
|[BreakLink](http://msdn.microsoft.com/library/cc177e67-8664-7273-2339-7d9c01f65ba6%28Office.15%29.aspx)|
|[Update](http://msdn.microsoft.com/library/c1ce2e2f-53ca-9c64-4ce5-1e0d0bed6c54%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/a0854949-7bbf-5af7-7c32-a2d67be468ec%28Office.15%29.aspx)|
|[AutoUpdate](http://msdn.microsoft.com/library/de142aa6-2414-61c3-62d1-1226a0f9209f%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/49bc1179-6fc4-c11f-c0a2-e35d95704622%28Office.15%29.aspx)|
|[SourceFullName](http://msdn.microsoft.com/library/6a7fb694-609a-77c5-eabc-d95693a87299%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
