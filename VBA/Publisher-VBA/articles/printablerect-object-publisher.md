---
title: PrintableRect Object (Publisher)
keywords: vbapb10.chm7602175
f1_keywords:
- vbapb10.chm7602175
ms.prod: publisher
api_name:
- Publisher.PrintableRect
ms.assetid: fd99e9d4-81d9-63ae-78ca-f7a16b031239
ms.date: 06/08/2017
---


# PrintableRect Object (Publisher)

Represents the sheet area within which the specified printer will print. The printable rectangle is determined by the printer based on the sheet size specified. The printable rectangle of the printer sheet should not be confused with the area within the margins of the publication page; it may be larger or smaller than the publication page.
 


## Remarks

In cases in which the printer sheet and the publication page size are identical, the publication page is centered on the printer sheet and none of the printer's marks print, even if they are selected.
 

 

## Example

Use the  **[PrintableRect](printer-printablerect-property-publisher.md)** property of the **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** object to return a **PrintableRect** object. The following example returns printable rectangle boundaries for the printer sheet of the active publication.
 

 

```
Sub ListPrintableRectBoundaries() 
 
With ActiveDocument.AdvancedPrintOptions.PrintableRect 
 
 Debug.Print "Printable area is " &amp; _ 
 PointsToInches(.Width) &amp; _ 
 " by " &amp; PointsToInches(.Height) &amp; " inches." 
 Debug.Print "Left Boundary: " &amp; PointsToInches(.Left) &amp; _ 
 " inches (from left)." 
 Debug.Print "Right Boundary: " &amp; PointsToInches(.Left + .Width) &amp; _ 
 " inches (from left)." 
 Debug.Print "Top Boundary: " &amp; PointsToInches(.Top) &amp; _ 
 " inches(from top)." 
 Debug.Print "Bottom Boundary: " &amp; PointsToInches(.Top + .Height) &amp; _ 
 " inches(from top)." 
 
End With 
 
End Sub 

```


## Properties



|**Name**|
|:-----|
|[Application](printablerect-application-property-publisher.md)|
|[Height](printablerect-height-property-publisher.md)|
|[Left](printablerect-left-property-publisher.md)|
|[Parent](printablerect-parent-property-publisher.md)|
|[Top](printablerect-top-property-publisher.md)|
|[Width](printablerect-width-property-publisher.md)|

