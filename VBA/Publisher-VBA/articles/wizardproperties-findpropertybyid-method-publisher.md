---
title: WizardProperties.FindPropertyById Method (Publisher)
keywords: vbapb10.chm1507332
f1_keywords:
- vbapb10.chm1507332
ms.prod: publisher
api_name:
- Publisher.WizardProperties.FindPropertyById
ms.assetid: 9d13ffa2-f251-0e7d-2f36-c747413143d0
ms.date: 06/08/2017
---


# WizardProperties.FindPropertyById Method (Publisher)

Returns a  **[WizardProperty](wizardproperty-object-publisher.md)** object, based on the specified ID, from the collection of wizard properties associated with a publication design or a Design Gallery object's wizard.


## Syntax

 _expression_. **FindPropertyById**( **_ID_**)

 _expression_A variable that represents a  **WizardProperties** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ID|Required| **Long**|The ID of the the wizard property to return; corresponds to the  **[ID](wizardproperty-id-property-publisher.md)** property of the **WizardProperty** object.|

### Return Value

WizardProperty


## Example

The following example changes the settings of the current publication design (Newsletter Wizard) so that the publication has a region dedicated to the customer's address (Customer Address).


```vb
Sub SetWizardProperties 
 Dim wizTemp As Wizard 
 Dim wizproTemp As WizardProperty 
 
 Set wizTemp = ActiveDocument.Wizard 
 
 With wizTemp.Properties 
 Set wizproTemp = .FindPropertyById(ID:=901) 
 wizproTemp.CurrentValueId = 1 
 End With 
 
End Sub
```


