---
title: Document.AdvancedPrintOptions Property (Publisher)
keywords: vbapb10.chm196713
f1_keywords:
- vbapb10.chm196713
ms.prod: publisher
api_name:
- Publisher.Document.AdvancedPrintOptions
ms.assetid: 33c075e0-f813-9bb4-e199-96e5e9ed4ba8
ms.date: 06/08/2017
---


# Document.AdvancedPrintOptions Property (Publisher)

Returns an  **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** object that represents the advanced print settings for a publication. Read-only.


## Syntax

 _expression_. **AdvancedPrintOptions**

 _expression_A variable that represents a  **Document** object.


### Return Value

AdvancedPrintOptions


## Remarks

The properties of the  **AdvancedPrintOptions** object correspond to the options in the **Advanced Print Settings** dialog box.


## Example

The following example tests to determine if the active publication has been set to print as separations. If it has, it is set to print only plates for the inks actually used in the publication, and to not print plates for any pages where a color is not used.


```vb
Sub PrintOnlyInksUsed 
 With ActiveDocument.AdvancedPrintOptions 
 If .PrintMode = pbPrintModeSeparations Then 
 .InksToPrint = pbInksToPrintUsed 
 .PrintBlankPlates = False 
 End If 
 End With 
End Sub
```


