---
title: Plates Object (Publisher)
keywords: vbapb10.chm2883583
f1_keywords:
- vbapb10.chm2883583
ms.prod: publisher
api_name:
- Publisher.Plates
ms.assetid: 7da44b06-c94f-dadc-da91-09b757d5a076
ms.date: 06/08/2017
---


# Plates Object (Publisher)

A collection of  **Plate** objects in a publication.
 


## Example

The  **Plates** collection is made up of **Plate** objects for the various publication color modes. Each publication can only use one color mode. For example, you can't specify the spot-color mode in a procedure and then later specify the process-color mode. Use the **[CreatePlateCollection](http://msdn.microsoft.com/library/339c2c90-d1b7-808e-2b3c-c52c000e4908%28Office.15%29.aspx)** method of the **[Document](document-object-publisher.md)** object to specify which color mode to use in a publication's plate collection. Use the **[Add](plates-add-method-publisher.md)** method of the **Plates** collection to add a new plate to the **Plates** collection. This example creates a new spot-color plate collection and adds a plate to it.
 

 

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

Use the  **[EnterColorMode](http://msdn.microsoft.com/library/3c04275d-d274-f681-7391-139a54232a3b%28Office.15%29.aspx)** method of the **[Document](document-object-publisher.md)** object to the specify the color mode and the **Plates** collection to use with the color mode. Use the **[ColorMode](http://msdn.microsoft.com/library/58befa97-9d9b-9294-18b2-ae10dc87f51c%28Office.15%29.aspx)** property to determine which color mode is in use in a publication. This example creates a spot-color plate collection, adds two plates to it, and then enters those plates into the spot-color mode.
 

 



```
Sub CreateSpotColorMode() 
 Dim plArray As Plates 
 
 With ThisDocument 
 'Creates a color plate collection, 
 'which contains one black plate by default 
 Set plArray = .CreatePlateCollection(Mode:=pbColorModeSpot) 
 
 'Sets the plate color to red 
 plArray(1).Color.RGB = RGB(255, 0, 0) 
 
 'Adds another plate, black by default and 
 'sets the plate color to green 
 plArray.Add 
 plArray(2).Color.RGB = RGB(0, 255, 0) 
 
 'Enters spot-color mode with above 
 'two plates in the plates array 
 .EnterColorMode Mode:=pbColorModeSpot, Plates:=plArray 
 End With 
End Sub
```

Use the  **[FindPlateByInkName](plates-findplatebyinkname-method-publisher.md)** method to return a specific plate by referencing its ink name. Process colors are assigned different index numbers in the **Plates** collection than in the **[PrintablePlates](printableplates-object-publisher.md)** collection. Use the **FindPlateByInkName** method to insure the desired **[Plate](plate-object-publisher.md)** or **[PrintablePlate](printableplate-object-publisher.md)** object is accessed.
 

 

## Methods



|**Name**|
|:-----|
|[Add](plates-add-method-publisher.md)|
|[FindPlateByInkName](plates-findplatebyinkname-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](plates-application-property-publisher.md)|
|[Count](plates-count-property-publisher.md)|
|[Item](plates-item-property-publisher.md)|
|[Parent](plates-parent-property-publisher.md)|

