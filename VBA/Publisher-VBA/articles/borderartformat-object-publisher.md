---
title: BorderArtFormat Object (Publisher)
keywords: vbapb10.chm7667711
f1_keywords:
- vbapb10.chm7667711
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat
ms.assetid: ba066b2e-fe40-aeef-9275-2cc2810f63ca
ms.date: 06/08/2017
---


# BorderArtFormat Object (Publisher)

Represents the formatting of the BorderArt applied to the specified shape.
 


## Remarks

BorderArt are picture borders that can be applied to text boxes, picture frames, or rectangles.
 

 

## Example

Use the  **[BorderArt](shape-borderart-property-publisher.md)** property of a shape to return a **BorderArtFormat** object.
 

 
The following example returns the BorderArt of the first shape on the first page of the active publication, and displays the name of the BorderArt in a message box.
 

 



```
Dim bdaTemp As BorderArtFormat 
 
Set bdaTemp = ActiveDocument.Pages(1).Shapes(1).BorderArt 
MsgBox "BorderArt name is: " &amp;bdaTemp.Name
```

Use the  **[Set](borderartformat-set-method-publisher.md)** method to specify which type of BorderArt you want applied to a picture. The following example tests for the existence of BorderArt on each shape for each page of the active document. Any BorderArt found is set to the same type.
 

 



```
Sub SetBorderArt() 
Dim anyPage As Page 
Dim anyShape As Shape 
Dim strBorderArtName As String 
 
strBorderArtName = Document.BorderArts(1).Name 
 
For Each anyPage in ActiveDocument.Pages 
For Each anyShape in anyPage.Shapes 
With anyShape.BorderArt 
If .Exists = True Then 
.Set(strBorderArtName) 
End If 
End With 
Next anyShape 
Next anyPage 
End Sub
```

You can also use the  **[Name](borderartformat-name-property-publisher.md)** property to specify which type of BorderArt you want applied to a picture. The following example sets all the BorderArt in a document to the same type using the **Name** property.
 

 



```
Sub SetBorderArtByName() 
Dim anyPage As Page 
Dim anyShape As Shape 
Dim strBorderArtName As String 
 
strBorderArtName = Document.BorderArts(1).Name 
 
For Each anyPage in ActiveDocument.Pages 
For Each anyShape in anyPage.Shapes 
With anyShape.BorderArt 
If .Exists = True Then 
.Name = strBorderArtName 
End If 
End With 
Next anyShape 
Next anyPage 
End Sub
```


 **Note**  Because  **Name** is the default property of both the **[BorderArt](borderart-object-publisher.md)** and **BorderArtFormat** objects, you do not need to state it explicitly when setting the BorderArt type. The statement `Shape.BorderArtFormat = Document.BorderArts(1)`is equivalent to  `Shape.BorderArtFormat.Name = Document.BorderArts(1).Name`
 

Use the  **[Delete](borderartformat-delete-method-publisher.md)** method to remove BorderArt from a picture. The following example tests for the existence of border art on each shape for each page of the active document. If border art exists, it is deleted.
 

 



```
Sub DeleteBorderArt() 
Dim anyPage As Page 
Dim anyShape As Shape 
 
For Each anyPage in ActiveDocument.Pages 
For Each anyShape in anyPage.Shapes 
With anyShape.BorderArt 
If .Exists = True Then 
.Delete 
End If 
End With 
Next anyShape 
Next anyPage 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](borderartformat-delete-method-publisher.md)|
|[RevertToDefaultWeight](borderartformat-reverttodefaultweight-method-publisher.md)|
|[RevertToOriginalColor](borderartformat-reverttooriginalcolor-method-publisher.md)|
|[Set](borderartformat-set-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](borderartformat-application-property-publisher.md)|
|[Color](borderartformat-color-property-publisher.md)|
|[Exists](borderartformat-exists-property-publisher.md)|
|[Name](borderartformat-name-property-publisher.md)|
|[Parent](borderartformat-parent-property-publisher.md)|
|[StretchPictures](borderartformat-stretchpictures-property-publisher.md)|
|[Weight](borderartformat-weight-property-publisher.md)|

