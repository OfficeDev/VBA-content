---
title: BulletFormat.Number Property (PowerPoint)
keywords: vbapp10.chm577013
f1_keywords:
- vbapp10.chm577013
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat.Number
ms.assetid: 90f92c4e-4a15-7efe-1251-5394a148db72
ms.date: 06/08/2017
---


# BulletFormat.Number Property (PowerPoint)

Returns the bullet number of a paragraph when the  **[Type](bulletformat-type-property-powerpoint.md)** property of the **BulletFormat** object is set to **ppBulletNumbered**. Read-only.


## Syntax

 _expression_. **Number**

 _expression_ A variable that represents a **BulletFormat** object.


### Return Value

Long


## Remarks

If this property is queried for multiple paragraphs with different numbers, then the value  **ppBulletMixed** is returned. If this property is queried for a paragraph with a type other than **ppBulletNumbered**, then a run-time error occurs.


## Example

This example returns the bullet number of paragraph one in the selected text range to a variable named  `myParnum`.


```vb
With ActiveWindow.Selection

    If .Type = ppSelectionTextRange Then

        With .TextRange.Paragraphs(1).ParagraphFormat.Bullet

            If .Type = ppBulletNumbered Then

                myParnum = .Number

            End If

        End With

    End If

End With
```


## See also


#### Concepts


[BulletFormat Object](bulletformat-object-powerpoint.md)

