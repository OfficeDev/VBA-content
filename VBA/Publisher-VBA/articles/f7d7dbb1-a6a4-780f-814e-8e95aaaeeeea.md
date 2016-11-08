
# Plate Object (Publisher)

Represents a single printer's plate. The  **Plate** object is a member of the **[Plates](7da44b06-c94f-dadc-da91-09b757d5a076.md)** collection.
 


## Example

Use the  **[Add](7fb7b602-8797-e275-4ff7-2e87cf1db11f.md)** method of the **[Plates](7da44b06-c94f-dadc-da91-09b757d5a076.md)** collection to create a new plate. This example creates a new spot-color plate collection and adds a plate to it.
 

 

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

Use the  **[FindPlateByInkName](4ebbc826-468b-7cd7-806e-056e4cbb488c.md)** method to return a specific plate by referencing its ink name. Process colors are assigned different index numbers in the **Plates** collection than in the **[PrintablePlates](40766b1a-64b3-e18a-2c67-c3db4c4ceb26.md)** collection. Use the **FindPlateByInkName** method to insure the desired **Plate** or **[PrintablePlate](cea95f22-9c02-b66b-05b7-e11f1145a505.md)** object is accessed.
 

 

## Methods



|**Name**|
|:-----|
|[ConvertToProcess](26476701-aa82-ca44-20c8-55a332a6539a.md)|
|[Delete](fadaba7c-6636-f1e2-e360-3fcf8700ab36.md)|

## Properties



|**Name**|
|:-----|
|[Application](12817b6a-18f4-66b3-a6a5-6fbea8dc9987.md)|
|[Color](4c4897f5-90bb-cb84-e9b8-47df1a912916.md)|
|[Index](7a16bd86-f0c4-d2df-832e-e9a55fed9068.md)|
|[InkName](248c1529-2706-5458-a13f-def479d16132.md)|
|[InUse](6c98ada2-ff05-30c9-0043-afbe892dab3d.md)|
|[Luminance](8d84fe74-8421-4ec2-bf6e-a156a0c0018b.md)|
|[Name](47453f6b-2f5b-17ba-5787-4701636ccd72.md)|
|[Parent](d5f31725-826e-f636-93b7-253884a90927.md)|
