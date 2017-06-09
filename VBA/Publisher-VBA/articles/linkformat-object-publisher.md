---
title: LinkFormat Object (Publisher)
keywords: vbapb10.chm4456447
f1_keywords:
- vbapb10.chm4456447
ms.prod: publisher
api_name:
- Publisher.LinkFormat
ms.assetid: 5b588edd-b026-cfc7-4acb-77290ae4d297
ms.date: 06/08/2017
---


# LinkFormat Object (Publisher)

Represents the linking characteristics for an OLE object or picture.
 


## Remarks

Not all types of shapes and fields can be linked to a source. Use the  **[Type](shape-type-property-publisher.md)** property for the **[Shape](shape-object-publisher.md)** object to determine whether a particular shape can be linked.
 

 
Use the  **[Update](linkformat-update-method-publisher.md)** method to update links. To return or set the full path for a particular link's source file, use the **[SourceFullName](linkformat-sourcefullname-property-publisher.md)** property.
 

 

## Example

Use the  **[LinkFormat](shape-linkformat-property-publisher.md)** property for a shape or field to return a **LinkFormat** object. The following example updates the links to all linked OLE objects on the first page of the active publication.
 

 

```
Sub FindOLEObjects() 
 Dim shpShape As Shape 
 
 For Each shpShape In ActiveDocument.Pages(1).Shapes 
 If shpShape.Type = pbLinkedOLEObject Then 
 shpShape.LinkFormat.Update 
 End If 
 Next shpShape 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Update](linkformat-update-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](linkformat-application-property-publisher.md)|
|[Parent](linkformat-parent-property-publisher.md)|
|[SourceFullName](linkformat-sourcefullname-property-publisher.md)|

