---
title: Designs Object (PowerPoint)
keywords: vbapp10.chm643000
f1_keywords:
- vbapp10.chm643000
ms.prod: powerpoint
api_name:
- PowerPoint.Designs
ms.assetid: 9b02ed6d-9a84-3464-5669-f614e0f33b10
ms.date: 06/08/2017
---


# Designs Object (PowerPoint)

Represents a collection of slide design templates.


## Remarks

Use the [Designs](slide-design-property-powerpoint.md)property of the  **[Presentation](presentation-object-powerpoint.md)** object to reference a design template.

To add or clone an individual design template, use the  **Designs** collection's[Add](designs-add-method-powerpoint.md)or [Clone](designs-clone-method-powerpoint.md)methods, respectively. To refer to an individual design template, use the [Item](designs-item-method-powerpoint.md)method.

To load a design template, use the [Load](designs-load-method-powerpoint.md)method.


## Example

The following example adds a new design template to the  **Designs** collection and confirms it was added correctly.


```vb
Sub AddDesignMaster()

    With ActivePresentation.Designs

        .Add designName:="MyDesignName"

        MsgBox .Item("MyDesignName").Name

    End With

End Sub
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

