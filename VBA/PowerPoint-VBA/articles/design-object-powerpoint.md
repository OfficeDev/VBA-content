---
title: Design Object (PowerPoint)
keywords: vbapp10.chm644000
f1_keywords:
- vbapp10.chm644000
ms.prod: powerpoint
api_name:
- PowerPoint.Design
ms.assetid: 3b02c779-8313-9512-c8d9-cf8a3883229f
ms.date: 06/08/2017
---


# Design Object (PowerPoint)

Represents an individual slide design template. The  **Design** object is a member of the **[Designs](designs-object-powerpoint.md)** and **[SlideRange](sliderange-object-powerpoint.md)** collections and the **[Master](master-object-powerpoint.md)** and **[Slide](slide-object-powerpoint.md)** objects.


## Remarks

Use the  **Design** property of the **Master**, **Slide**, or **SlideRange** objects to access a **Design** object, for example:


-  `ActivePresentation.SlideMaster.Design`
    
-  `ActivePresentation.Slides(1).Design`
    
-  `ActivePresentation.Slides.Range.Design`
    
Use the [Add](designs-add-method-powerpoint.md), [Item](designs-item-method-powerpoint.md), [Clone](designs-clone-method-powerpoint.md), or [Load](designs-load-method-powerpoint.md)methods of the  **Designs** collection to add, refer to, clone, or load a **Design** object, respectively. For example, to add a design template, use `ActivePresentation.Designs.Add designName:="MyDesign"`


## Example

The  **Design** object's[AddTitleMaster](presentation-addtitlemaster-method-powerpoint.md)method and [HasTitleMaster](presentation-hastitlemaster-property-powerpoint.md)property can be used to add and / or query the status of a title slide master. For example:


```vb
Sub AddQueryTitleMaster(dsn As Design)

    dsn.AddTitleMaster

    MsgBox dsn.HasTitleMaster

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

