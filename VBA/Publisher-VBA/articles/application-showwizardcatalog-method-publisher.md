---
title: Application.ShowWizardCatalog Method (Publisher)
keywords: vbapb10.chm131189
f1_keywords:
- vbapb10.chm131189
ms.prod: publisher
api_name:
- Publisher.Application.ShowWizardCatalog
ms.assetid: a8307ff9-a6c1-7655-8127-284f3781dae9
ms.date: 06/08/2017
---


# Application.ShowWizardCatalog Method (Publisher)

Displays the  **Publication Types** catalog for the wizard of the specified type.


## Syntax

 _expression_. **ShowWizardCatalog**( **_Wizard_**)

 _expression_A variable that represents an  **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Wizard|Optional| **PbWizard**|The type of wizard catalog to be displayed.|

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ShowWizardCatalog** method to show the wizard catalog for brochures.


```vb
Public Sub ShowWizardCatalog_Example() 
 Application.ShowWizardCatalog (pbWizardBrochures) 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

