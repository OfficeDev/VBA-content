---
title: Printer.PrintableRect Property (Publisher)
keywords: vbapb10.chm8978450
f1_keywords:
- vbapb10.chm8978450
ms.prod: publisher
api_name:
- Publisher.Printer.PrintableRect
ms.assetid: 9d5b8264-9213-3d89-0613-421a4872c158
ms.date: 06/08/2017
---


# Printer.PrintableRect Property (Publisher)

Returns a  **[PrintableRect](printablerect-object-publisher.md)** object that represents the printer sheet area within which the specified printer will print. Read-only.


## Syntax

 _expression_. **PrintableRect**

 _expression_A variable that represents a  **Printer** object.


### Return Value

PrintableRect


## Remarks

The printable rectangle is determined by the printer based on the sheet size specified. The printable rectangle of the printer sheet should not be confused with the area within the margins of the publication page. The printable rectangle might be larger or smaller than the publication page.


 **Note**  When the printer sheet and the publication page size are identical, the publication page is centered on the printer sheet and none of the printer's marks print, even if they are selected.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **PrintableRect** property to get the boundaries of the printable rectangle for the printer sheet of the active printer.


```vb
Public Sub PrintableRect_Example() 
 
 Dim pubInstalledPrinters As Publisher.InstalledPrinters 
 Dim pubApplication As Publisher.Application 
 Dim pubPrinter As Publisher.Printer 
 
 Set pubApplication = ThisDocument.Application 
 Set pubInstalledPrinters = pubApplication.InstalledPrinters 
 
 For Each pubPrinter In pubInstalledPrinters 
 If pubPrinter.IsActivePrinter Then 
 With pubPrinter.PrintableRect 
 Debug.Print "Printable area is " &; PointsToInches(.Width) &; " by " &; PointsToInches(.Height) &; " inches." 
 Debug.Print "Left Boundary: " &; PointsToInches(.Left) &; " inches (from left)." 
 Debug.Print "Right Boundary: " &; PointsToInches(.Left + .Width) &; " inches (from left)." 
 Debug.Print "Top Boundary: " &; PointsToInches(.Top) &; " inches(from top)." 
 Debug.Print "Bottom Boundary: " &; PointsToInches(.Top + .Height) &; " inches (from top)." 
 End With 
 End If 
 Next 
 
End Sub
```


