---
title: Plate Object (Publisher)
keywords: vbapb10.chm2949119
f1_keywords:
- vbapb10.chm2949119
ms.prod: publisher
api_name:
- Publisher.Plate
ms.assetid: f7d7dbb1-a6a4-780f-814e-8e95aaaeeeea
ms.date: 06/08/2017
---


# Plate Object (Publisher)

Represents a single printer's plate. The  **Plate** object is a member of the **[Plates](plates-object-publisher.md)** collection.
 


## Example

Use the  **[Add](plates-add-method-publisher.md)** method of the **[Plates](plates-object-publisher.md)** collection to create a new plate. This example creates a new spot-color plate collection and adds a plate to it.
 

 

```
Sub AddNewPlates() 
 Dim plts As Plates 
 Set plts = ActiveDocument.CreatePlateCollection(Mode:=pbColorModeSpot) 
 plts.Add 
 With plts(1) 
 .Color.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 .Luminance = 4 
 End With 
End Sub
```

Use the  **[FindPlateByInkName](plates-findplatebyinkname-method-publisher.md)** method to return a specific plate by referencing its ink name. Process colors are assigned different index numbers in the **Plates** collection than in the **[PrintablePlates](printableplates-object-publisher.md)** collection. Use the **FindPlateByInkName** method to insure the desired **Plate** or **[PrintablePlate](printableplate-object-publisher.md)** object is accessed.
 

 

## Methods



|**Name**|
|:-----|
|[ConvertToProcess](plate-converttoprocess-method-publisher.md)|
|[Delete](plate-delete-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](plate-application-property-publisher.md)|
|[Color](plate-color-property-publisher.md)|
|[Index](plate-index-property-publisher.md)|
|[InkName](plate-inkname-property-publisher.md)|
|[InUse](plate-inuse-property-publisher.md)|
|[Luminance](plate-luminance-property-publisher.md)|
|[Name](plate-name-property-publisher.md)|
|[Parent](plate-parent-property-publisher.md)|

