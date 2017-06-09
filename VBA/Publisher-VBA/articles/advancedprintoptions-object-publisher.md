---
title: AdvancedPrintOptions Object (Publisher)
keywords: vbapb10.chm7143423
f1_keywords:
- vbapb10.chm7143423
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions
ms.assetid: 61f776cc-dc3e-61b6-057a-125ad15146c8
ms.date: 06/08/2017
---


# AdvancedPrintOptions Object (Publisher)

Represents the advanced print settings for a publication.
 


## Remarks

The properties of the  **AdvancedPrintOptions** object correspond to the options available on the tabs of the **Advanced Print Settings** dialog box.
 

 

## Example

Use the  **AdvancedPrintOptions** property of the **Document** object to return an **AdvancedPrintOptions** object. The following example tests to determine if the active publication has been set to print as separations. If it has, it is set to print only plates for the inks actually used in the publication, and to not print plates for any pages where a color is not used.
 

 

```
Sub PrintOnlyInksUsed 
 With ActiveDocument.AdvancedPrintOptions 
 If .PrintMode = pbPrintModeSeparations Then 
 .InksToPrint = pbInksToPrintUsed 
 .PrintBlankPlates = False 
 End If 
 End With 
End Sub
```


## Properties



|**Name**|
|:-----|
|[AllowBleeds](advancedprintoptions-allowbleeds-property-publisher.md)|
|[Application](advancedprintoptions-application-property-publisher.md)|
|[BackSideInsertFaceUp](advancedprintoptions-backsideinsertfaceup-property-publisher.md)|
|[GraphicsResolution](advancedprintoptions-graphicsresolution-property-publisher.md)|
|[HorizontalFlip](advancedprintoptions-horizontalflip-property-publisher.md)|
|[IsPostscriptPrinter](advancedprintoptions-ispostscriptprinter-property-publisher.md)|
|[ManualFeedAlign](advancedprintoptions-manualfeedalign-property-publisher.md)|
|[ManualFeedDirection](advancedprintoptions-manualfeeddirection-property-publisher.md)|
|[NegativeImage](advancedprintoptions-negativeimage-property-publisher.md)|
|[PageRotated](advancedprintoptions-pagerotated-property-publisher.md)|
|[Parent](advancedprintoptions-parent-property-publisher.md)|
|[PrintBleedMarks](advancedprintoptions-printbleedmarks-property-publisher.md)|
|[PrintCropMarks](advancedprintoptions-printcropmarks-property-publisher.md)|
|[PrintDensityBars](advancedprintoptions-printdensitybars-property-publisher.md)|
|[PrintJobInformation](advancedprintoptions-printjobinformation-property-publisher.md)|
|[PrintRegistrationMarks](advancedprintoptions-printregistrationmarks-property-publisher.md)|
|[Resolution](advancedprintoptions-resolution-property-publisher.md)|
|[UseOnlyPublicationFonts](advancedprintoptions-useonlypublicationfonts-property-publisher.md)|
|[VerticalFlip](advancedprintoptions-verticalflip-property-publisher.md)|

