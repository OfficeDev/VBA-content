---
title: AdvancedPrintOptions.PrintJobInformation Property (Publisher)
keywords: vbapb10.chm7077897
f1_keywords:
- vbapb10.chm7077897
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.PrintJobInformation
ms.assetid: c4494804-6dfa-8647-a72d-591f90624c1c
ms.date: 06/08/2017
---


# AdvancedPrintOptions.PrintJobInformation Property (Publisher)

 **True** to print information about the print job on each plate. The default is **True**. Read/write  **Boolean**.


## Syntax

 _expression_. **PrintJobInformation**

 _expression_A variable that represents a  **AdvancedPrintOptions** object.


### Return Value

Boolean


## Remarks

The  **PrintJobInformation** property can be set regardless of the print mode selected for the publication. However, it is ignored (and no job information is printed) when the print mode is set as composite RGB.

Job information includes the file name of the printed publication, the date it was printed, the page number, and which color ink the plate is for (cyan, magenta, yellow, black, or a spot color).

This property corresponds to the  **Job information** control on the **Page Settings** tab of the **Advanced Print Settings** dialog box.

These printer's marks print outside the publication and can be printed only if the size of the paper being printed to is larger than the publication page size.


## Example

The following example sets crop marks and job information to print with the publication. If the publication is printed as separations, the additional types of printer's marks are also set to print. This example assumes that the size of the paper being printed to is larger than the publication page size.


```vb
Sub SetPrintersMarksToPrint() 
 With ActiveDocument.AdvancedPrintOptions 
 .PrintCropMarks = True 
 .PrintJobInformation = True 
 If PrintMode = pbPrintModeSeparations Then 
 .PrintRegistrationMarks = True 
 .PrintDensityBars = True 
 .PrintColorBars = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


 [AdvancedPrintOptions Object](advancedprintoptions-object-publisher.md)

