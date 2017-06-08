---
title: WizardProperty.CurrentValueId Property (Publisher)
keywords: vbapb10.chm1572869
f1_keywords:
- vbapb10.chm1572869
ms.prod: publisher
api_name:
- Publisher.WizardProperty.CurrentValueId
ms.assetid: d8a2eeb0-f6e7-2687-5952-cddd2cc3914b
ms.date: 06/08/2017
---


# WizardProperty.CurrentValueId Property (Publisher)

Returns or sets a  **Long** indicating the value of a setting in the specified publication design or Design Gallery object's wizard. Read/write.


## Syntax

 _expression_. **CurrentValueId**

 _expression_A variable that represents a  **WizardProperty** object.


### Return Value

Long


## Remarks

Accessing this property for a publication design setting whose  **[Enabled](wizardproperty-enabled-property-publisher.md)** property is **False** causes an error.


## Example

The following example changes the settings of the current publication design (Newsletter Wizard) so that the publication has a region dedicated to the customer's address.


```vb
Dim wizTemp As Wizard 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp.Properties 
 .FindPropertyById(ID:=901).CurrentValueId = 1 
End With
```


