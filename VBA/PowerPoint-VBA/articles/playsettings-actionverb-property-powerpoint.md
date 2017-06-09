---
title: PlaySettings.ActionVerb Property (PowerPoint)
keywords: vbapp10.chm568003
f1_keywords:
- vbapp10.chm568003
ms.prod: powerpoint
api_name:
- PowerPoint.PlaySettings.ActionVerb
ms.assetid: 769383e0-94b0-3baf-0211-e92987b139ce
ms.date: 06/08/2017
---


# PlaySettings.ActionVerb Property (PowerPoint)

Returns or sets a string that contains the OLE verb that will be run when the specified OLE object is animated during a slide show. Read/write.


## Syntax

 _expression_. **ActionVerb**

 _expression_ A variable that represents a **PlaySettings** object.


## Remarks

The default verb specifies the action that the OLE object runs — such as playing a wave file or displaying data so that the user can modify it — after the previous animation or slide transition. 


## Example

This example specifies that shape three on slide one in the active presentation will automatically open for editing when it is animated. Shape three must be an OLE object that contains a sound or movie object and that supports the "Edit" verb.


```vb
Set OLEobj = ActivePresentation.Slides(1).Shapes(3)

With OLEobj.AnimationSettings.PlaySettings

    .PlayOnEntry = True

    .ActionVerb = "Edit"

End With
```


## See also


#### Concepts


[PlaySettings Object](playsettings-object-powerpoint.md)

