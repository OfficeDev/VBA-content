
# Plate Object (Publisher)

 **Last modified:** July 28, 2015

Represents a single printer's plate. The  **Plate** object is a member of the ** [Plates](7da44b06-c94f-dadc-da91-09b757d5a076.md)** collection.

## Example

Use the  ** [Add](7fb7b602-8797-e275-4ff7-2e87cf1db11f.md)**method of the  ** [Plates](7da44b06-c94f-dadc-da91-09b757d5a076.md)**collection to create a new plate. This example creates a new spot-color plate collection and adds a plate to it.


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

Use the  ** [FindPlateByInkName](4ebbc826-468b-7cd7-806e-056e4cbb488c.md)** method to return a specific plate by referencing its ink name. Process colors are assigned different index numbers in the **Plates** collection than in the ** [PrintablePlates](40766b1a-64b3-e18a-2c67-c3db4c4ceb26.md)** collection. Use the **FindPlateByInkName** method to insure the desired **Plate** or ** [PrintablePlate](cea95f22-9c02-b66b-05b7-e11f1145a505.md)** object is accessed.

