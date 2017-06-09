---
title: CustomLayouts Object (PowerPoint)
keywords: vbapp10.chm671000
f1_keywords:
- vbapp10.chm671000
ms.prod: powerpoint
api_name:
- PowerPoint.CustomLayouts
ms.assetid: 9ce682fb-545c-55cb-e9ac-3475f7556af1
ms.date: 06/08/2017
---


# CustomLayouts Object (PowerPoint)

Represents a set of custom layouts associated with a presentation design.


## Remarks

Use the  **[CustomLayouts](http://msdn.microsoft.com/library/8364388f-71be-c6b7-5ab0-4150e6f62feb%28Office.15%29.aspx)** property of the slide **[Master](master-object-powerpoint.md)** object to return a **CustomLayouts** collection. Use **CustomLayouts** ( _index_ ), where index is the color scheme index number, to return a single **[CustomLayout](customlayout-object-powerpoint.md)** object.

Use the  **[Add](http://msdn.microsoft.com/library/d22dc23a-cb03-ab32-fd27-e360377369a9%28Office.15%29.aspx)** method to create a new custom layout and add it to the **CustomLayouts** collection. Use the **[Paste](http://msdn.microsoft.com/library/d4fcd2db-3d6b-0c59-6ea3-f9aadf90ed04%28Office.15%29.aspx)** method to past slides from the Clipboard as a **CustomLayout** object into the **CustomLayouts** collection.

Use the  **CustomLayout** property of a **[Slide](slide-object-powerpoint.md)** or **[SlideRange](http://msdn.microsoft.com/library/440ab59d-744a-209f-bf28-d0acd3a21e1a%28Office.15%29.aspx)** object to return a custom layout for a slide or set of slides.


## Example

The following example adds a custom layout to the slide master of the active presentation.


```
Sub AddCustomLayout()

    With ActivePresentation.SlideMaster

        .CustomLayouts.Add (1)

        .CustomLayouts(1).Name = "MyLayout"

    End With

End Sub
```

The following example displays the name of the custom layout for the first slide of the active presentation.




```
MsgBox ActivePresentation.Slides(1).CustomLayout.Name
```


## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/d22dc23a-cb03-ab32-fd27-e360377369a9%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/1b88423a-0dc4-d45e-fe54-ee6ab6acfc62%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/d4fcd2db-3d6b-0c59-6ea3-f9aadf90ed04%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/56cea099-6d63-c0f7-6af2-c74a649ecb83%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/9267940e-244b-6f22-a517-2ec5728f40fa%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/90d228bc-edc3-2911-3629-892843970746%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
