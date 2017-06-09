---
title: BorderArt Object (Publisher)
keywords: vbapb10.chm7733247
f1_keywords:
- vbapb10.chm7733247
ms.prod: publisher
api_name:
- Publisher.BorderArt
ms.assetid: 464bec0f-7912-ab27-9593-7f1cb53da342
ms.date: 06/08/2017
---


# BorderArt Object (Publisher)

Represents an available type of BorderArt. BorderArt is picture borders that can be applied to text boxes, picture frames, or rectangles. The  **BorderArt** object is a member of the **[BorderArts](borderarts-object-publisher.md)** collection. The **BorderArts** collection contains all BorderArt available for use in the specified publication.
 


## Remarks

The  **BorderArts** collection includes any custom BorderArt types created by the user for the specified publication.
 

 

## Example

Use the  **[Item](borderarts-item-method-publisher.md)** property of a **BorderArts** collection to return a specific BorderArt object. The Index argument of the **Item** property can be the number or name of the BorderArt object.
 

 
This example returns the BorderArt "Apples" from the active publication. 
 

 



```
Dim bdaTemp As BorderArt 
 
Set bdaTemp = ActiveDocument.BorderArts.Item (Index:="Apples") 
```

Use the  **[Name](borderart-name-property-publisher.md)** property to specify which type of BorderArt you want applied to a picture. The following example sets all the BorderArt in a document to the same type using the **Name** property.
 

 



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


 

 

 **Note**  Because  **Name** is the default property of both the **BorderArt** object and the **BorderArtFormat** object, you do not need to state it explicitly when setting the BorderArt type. The statement `Shape.BorderArtFormat = Document.BorderArts(1)`is equivalent to `Shape.BorderArtFormat.Name = Document.BorderArts(1).Name`
 


## Properties



|**Name**|
|:-----|
|[Application](borderart-application-property-publisher.md)|
|[Name](borderart-name-property-publisher.md)|
|[Parent](borderart-parent-property-publisher.md)|

