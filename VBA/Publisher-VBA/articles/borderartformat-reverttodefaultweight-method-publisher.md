---
title: BorderArtFormat.RevertToDefaultWeight Method (Publisher)
keywords: vbapb10.chm7602180
f1_keywords:
- vbapb10.chm7602180
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat.RevertToDefaultWeight
ms.assetid: 3e46637f-3fce-3346-9193-063be40844bd
ms.date: 06/08/2017
---


# BorderArtFormat.RevertToDefaultWeight Method (Publisher)

Sets the BorderArt on the specified shape back to its default thickness.


## Syntax

 _expression_. **RevertToDefaultWeight**

 _expression_A variable that represents a  **BorderArtFormat** object.


## Remarks

The  **RevertToDefaultWeight** method has the same effect as the **Always apply at default size** control on the **BorderArt** dialog box.

Use the  **[Weight](borderartformat-weight-property-publisher.md)** property of the **[BorderArtFormat](borderartformat-object-publisher.md)** object to set the specified BorderArt to a thickness other than the default.


## Example

The following example tests for the existence of BorderArt on each shape for each page of the active document. If BorderArt exists, its weight is set to the default thickness and original color.


```vb
Sub RestoreBorderArtDefaults() 
 
Dim anyPage As Page 
Dim anyShape As Shape 
 
For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .RevertToDefaultWeight 
 .RevertToOriginalColor 
 End If 
 End With 
 Next anyShape 
Next anyPage 
End Sub
```


## See also


#### Concepts


 [BorderArtFormat Object](borderartformat-object-publisher.md)

