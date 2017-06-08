---
title: OLEFormat Object (Publisher)
keywords: vbapb10.chm4521983
f1_keywords:
- vbapb10.chm4521983
ms.prod: publisher
api_name:
- Publisher.OLEFormat
ms.assetid: e5b72d6b-dff8-3882-549f-e376c1e4d372
ms.date: 06/08/2017
---


# OLEFormat Object (Publisher)

Represents the OLE characteristics, other than linking (see the  **[LinkFormat](linkformat-object-publisher.md)** object), for an OLE object, ActiveX control, or field.
 


## Remarks

Not all types of shapes and fields have OLE capabilities. Use the  **[Type](shape-type-property-publisher.md)** property for the **[Shape](shape-object-publisher.md)** object to determine into which category the specified shape falls.
 

 
Use the  **[Activate](oleformat-activate-method-publisher.md)** and **[DoVerb](oleformat-doverb-method-publisher.md)** methods to automate an OLE object.
 

 

## Example

Use the  **[OLEFormat](shape-oleformat-property-publisher.md)** property for a shape or field to return an **OLEFormat** object. The following example activates all OLE objects in the active publication.
 

 

```
Sub ActivateOLEObjects() 
 Dim shpShape As Shape 
 
 For Each shpShape In ActiveDocument.Pages(1).Shapes 
 If shpShape.Type = pbLinkedOLEObject Then 
 shpShape.OLEFormat.Activate 
 End If 
 Next 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Activate](oleformat-activate-method-publisher.md)|
|[DoVerb](oleformat-doverb-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](oleformat-application-property-publisher.md)|
|[Object](oleformat-object-property-publisher.md)|
|[ObjectVerbs](oleformat-objectverbs-property-publisher.md)|
|[Parent](oleformat-parent-property-publisher.md)|
|[ProgId](oleformat-progid-property-publisher.md)|

