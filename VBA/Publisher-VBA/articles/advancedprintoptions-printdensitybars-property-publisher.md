---
title: AdvancedPrintOptions.PrintDensityBars Property (Publisher)
keywords: vbapb10.chm7077904
f1_keywords:
- vbapb10.chm7077904
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.PrintDensityBars
ms.assetid: b98baed0-e2ba-bf69-78e2-d60125d7f57a
ms.date: 06/08/2017
---


# AdvancedPrintOptions.PrintDensityBars Property (Publisher)

 **True** to print a density bar for the specified publication. The default is **True**. Read/write  **Boolean**.


## Syntax

 _expression_. **PrintDensityBars**

 _expression_A variable that represents a  **AdvancedPrintOptions** object.


### Return Value

Boolean


## Remarks

Returns "Permission Denied" if any print mode other than separations is selected for the specified publication.

The density bar printed is graded from a 10 percent screen to a 100 percent fill. A commercial printer can use this bar to determine proper exposure time for plate burning, and to test dot gain in the printed pages.

This property corresponds to the  **Density bars** control on the **Page Settings** tab of the **Advanced Print Settings** dialog box.

These printer's marks print outside the publication and can be printed only if the size of the paper being printed on is larger than the publication page size.


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

