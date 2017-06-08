---
title: AdvancedPrintOptions.AllowBleeds Property (Publisher)
keywords: vbapb10.chm7077906
f1_keywords:
- vbapb10.chm7077906
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.AllowBleeds
ms.assetid: 0c12a611-4e1e-468b-ada2-f07d01fd4445
ms.date: 06/08/2017
---


# AdvancedPrintOptions.AllowBleeds Property (Publisher)

 **True** to allow bleeds to print for the specified publication. The default is **True**. Read/write  **Boolean**.


## Syntax

 _expression_. **AllowBleeds**

 _expression_A variable that represents an  **AdvancedPrintOptions** object.


### Return Value

Boolean


## Remarks

When bleeds are allowed, objects that are partially off the page print to one eighth inch outside the defined page size.

If you allow bleeds in a document, you can specify whether bleed marks are printed by using the  **[PrintBleedMarks](advancedprintoptions-printbleedmarks-property-publisher.md)** property of the **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** object.

This property corresponds to the  **Allow bleeds** control on the **Page Settings** tab of the **Advanced Print Settings** dialog box.


## Example

The following example sets the publication to allow bleeds, and to print bleed marks.


```vb
Sub AllowBleedsAndPrintMarks() 
 With ActiveDocument.AdvancedPrintOptions 
 .AllowBleeds = True 
 .PrintBleedMarks = True 
 End With 
End Sub
```


## See also


#### Concepts


 [AdvancedPrintOptions Object](advancedprintoptions-object-publisher.md)

