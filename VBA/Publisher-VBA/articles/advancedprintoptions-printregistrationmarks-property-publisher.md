---
title: AdvancedPrintOptions.PrintRegistrationMarks Property (Publisher)
keywords: vbapb10.chm7077896
f1_keywords:
- vbapb10.chm7077896
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.PrintRegistrationMarks
ms.assetid: 24928459-0158-b7a9-46c0-c1a6116518d5
ms.date: 06/08/2017
---


# AdvancedPrintOptions.PrintRegistrationMarks Property (Publisher)

 **True** to print registration marks for the specified publication. The default is **True**. Read/write  **Boolean**.


## Syntax

 _expression_. **PrintRegistrationMarks**

 _expression_A variable that represents a  **AdvancedPrintOptions** object.


### Return Value

Boolean


## Remarks

Returns "Permission Denied" if any print mode other than separations is selected for the specified publication.

This property corresponds to the  **Registration marks** control on the **Page Settings** tab of the **Advanced Print Settings** dialog box.

Registration marks are used to align (register) the printing of two or more press plates on a single page.

These printer's marks print outside the publication and can only be printed if the size of the paper being printed to is larger than the publication page size.


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

