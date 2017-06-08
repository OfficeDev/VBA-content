---
title: OLEFormat Object (PowerPoint)
keywords: vbapp10.chm562000
f1_keywords:
- vbapp10.chm562000
ms.prod: powerpoint
api_name:
- PowerPoint.OLEFormat
ms.assetid: fbb6d6dd-4dbb-461b-986e-5095c6dc1486
ms.date: 06/08/2017
---


# OLEFormat Object (PowerPoint)

Contains properties and methods that apply to OLE objects. 


## Remarks

The  **[LinkFormat](linkformat-object-powerpoint.md)** object contains properties and methods that apply to linked OLE objects only. The **[PictureFormat](pictureformat-object-powerpoint.md)** object contains properties and methods that apply to pictures and OLE objects.


## Example

Use the  **OLEFormat** property to return an **OLEFormat** object. The following example loops through all the shapes on all the slides in the active presentation and sets all linked Microsoft Excel worksheets to be updated manually.


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
|[Activate](http://msdn.microsoft.com/library/cc4691a3-726f-5093-6345-f688b68ac15a%28Office.15%29.aspx)|
|[DoVerb](http://msdn.microsoft.com/library/1ee39c5d-3646-81de-79e9-f8cff869308d%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/095419ed-7d4b-16d0-a306-dc0da5c21d9c%28Office.15%29.aspx)|
|[FollowColors](http://msdn.microsoft.com/library/5f4c3f3d-0332-646f-de45-6854497f5782%28Office.15%29.aspx)|
|[Object](http://msdn.microsoft.com/library/fcaef43d-590e-179f-6698-4a8c191b92f9%28Office.15%29.aspx)|
|[ObjectVerbs](http://msdn.microsoft.com/library/895becb3-de86-638c-88e9-b9e72b6c713e%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/2eb7c4bf-5d11-d0e6-74b3-bde215ca3701%28Office.15%29.aspx)|
|[ProgID](http://msdn.microsoft.com/library/7564f3e1-4e14-9038-a836-5665518b0d09%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
