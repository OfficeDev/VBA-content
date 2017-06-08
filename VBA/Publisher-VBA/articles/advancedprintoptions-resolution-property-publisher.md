---
title: AdvancedPrintOptions.Resolution Property (Publisher)
keywords: vbapb10.chm7077910
f1_keywords:
- vbapb10.chm7077910
ms.prod: publisher
api_name:
- Publisher.AdvancedPrintOptions.Resolution
ms.assetid: 6105287e-a0af-2fd6-e0de-5bedb2458010
ms.date: 06/08/2017
---


# AdvancedPrintOptions.Resolution Property (Publisher)

Returns or sets a  **String** that represents the resolution, in dots per inch (dpi), at which to print the specified publication. Default is dependent on the printer driver, but is usually "(default)". Read/write.


## Syntax

 _expression_. **Resolution()**

 _expression_A variable that represents a  **AdvancedPrintOptions** object.


### Return Value

String


## Remarks

Valid values for the  **Resolution** property depend on the printer driver being used. Printers have preset resolutions that cannot be customized. Values must be formatted in the following manner, including spacing:


- " _HorizontalDotsPerInch_ x _VerticalDotsPerInch_" 
    
HorizontalDotsPerInch and VerticalDotsPerInch are numeric values, separated by one space, a lowercase x, and another space.

For example, to set the resolution of a printer to 600 horizontal dpi by 600 vertical dpi, a valid string would read "600 x 600".

The  **Resolution** property also accepts the string "(default)" to specify the printer's default resolution setting. If the printer driver presents a language other than English, the **Resolution** property also accepts the string that denotes the default setting in that language.

If the  **Resolution** property is set to the default printer driver setting, using a **Get** statement returns the English string "(default)", regardless of whether the resolution was set to default using a non-English string.

This property corresponds to the  **Resolution** control on the **Separations** tab of the **Advanced Print Settings** dialog box.


## Example

The following example sets the resolution of the active publication at 300 dpi by 300 dpi. The example assumes that "300 x 300" is a valid string for the printer driver used.


```vb
ActiveDocument.AdvancedPrintOptions.Resolution = "300 x 300"
```


## See also


#### Concepts


 [AdvancedPrintOptions Object](advancedprintoptions-object-publisher.md)

