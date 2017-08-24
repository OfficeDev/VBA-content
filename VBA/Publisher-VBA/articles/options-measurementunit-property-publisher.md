---
title: Options.MeasurementUnit Property (Publisher)
keywords: vbapb10.chm1048594
f1_keywords:
- vbapb10.chm1048594
ms.prod: publisher
api_name:
- Publisher.Options.MeasurementUnit
ms.assetid: 49221e4e-c84a-6706-8f9a-3853283ebb18
ms.date: 06/08/2017
---


# Options.MeasurementUnit Property (Publisher)

Returns or sets a  **PbUnitType** constant representing the standard measurement unit for Microsoft Publisher. Read/write.


## Syntax

 _expression_. **MeasurementUnit**

 _expression_A variable that represents a  **Options** object.


### Return Value

PbUnitType


## Remarks

The  **MeasurementUnit** property value can be one of the **PbUnitType** constants declared in the Publisher type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **pbUnitCM**| Sets the unit of measurement to centimeters.|
| **pbUnitEmu**| Doesn't apply to this property; returns an error if used..|
| **pbUnitFeet**|Doesn't apply to this property; returns an error if used.|
| **pbUnitHa**|Doesn't apply to this property; returns an error if used.|
| **pbUnitInch**|Sets the unit of measurement to inches.|
| **pbUnitKyu**| Doesn't apply to this property; returns an error if used.|
| **pbUnitMeter** .|Doesn't apply to this property; returns an error if used.|
| **pbUnitPica**|Sets the unit of measurement to picas.|
| **pbUnitPoint**|Sets the unit of measurement to points.|
| **pbUnitTwip**|Doesn't apply to this property; returns an error if used.|

## Example

This example sets the standard measurement unit for Publisher to points.


```vb
Sub SetUnitOfMeasurement() 
 Options.MeasurementUnit = pbUnitPoint 
End Sub
```

This example displays the current unit of measurement.




```vb
Sub GetUnitOfMeasurement() 
 Dim measUnit As PbUnitType 
 Dim strUnit As String 
 
 measUnit = Options.MeasurementUnit 
 
 Select Case measUnit 
 Case 0 
 strUnit = "inches" 
 Case 1 
 strUnit = "centimeters" 
 Case 2 
 strUnit = "picas" 
 Case 3 
 strUnit = "points" 
 End Select 
 
 MsgBox "The current unit of measurement is " &; strUnit 
 
End Sub
```


