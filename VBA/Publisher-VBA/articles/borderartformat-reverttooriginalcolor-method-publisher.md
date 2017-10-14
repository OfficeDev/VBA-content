---
title: BorderArtFormat.RevertToOriginalColor Method (Publisher)
keywords: vbapb10.chm7602192
f1_keywords:
- vbapb10.chm7602192
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat.RevertToOriginalColor
ms.assetid: 6b966576-eac4-3e55-ffdc-c064341474c0
ms.date: 06/08/2017
---


# BorderArtFormat.RevertToOriginalColor Method (Publisher)

Sets the BorderArt on the specified shape back to its default color.


## Syntax

 _expression_. **RevertToOriginalColor**

 _expression_A variable that represents a  **BorderArtFormat** object.


## Remarks

The  **RevertToOriginalColor** method has the same effect as the **Default** selection on the **Color** control on the **Format <Shape&gt;** dialog box.

Use the  **[Color](borderartformat-color-property-publisher.md)** property of the **[BorderArtFormat](borderartformat-object-publisher.md)** object to set the BorderArt to a color other than the original color.


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

