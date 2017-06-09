---
title: Documents.Add Method (Publisher)
keywords: vbapb10.chm8650756
f1_keywords:
- vbapb10.chm8650756
ms.prod: publisher
api_name:
- Publisher.Documents.Add
ms.assetid: 1e3536c8-8fc0-8c95-3a4c-b16fe8a99098
ms.date: 06/08/2017
---


# Documents.Add Method (Publisher)

Adds a new  **Document** object that represents a new publication to the **Documents** collection.


## Syntax

 _expression_. **Add**( **_PbWizard_**,  **_desid_**)

 _expression_An expression that returns a  **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PbWizard|Optional| **PbWizard**|The wizard to use to create the new publication.|
|desid|Optional| **Long**|The ID of the design to apply to the new publication.|

### Return Value

Document


## Remarks

The PbWizard parameter value should be a constant from the  **[PbWizard](pbwizard-enumeration-publisher.md)** enumeration, declared in the Microsoft Publisher 2007 type library.

The desid parameter value should be the ID of the design to apply. You can determine the design ID by creating a new publication that uses the wizard and design you want in the Publisher user interface and then running the following Visual Basic for Applications (VBA) macro.




```vb
Public Sub FindDesignID() 
 
 Dim pbWizard As Wizard 
 Dim pbWizardProperty As WizardProperty 
 
 Set pbWizard = ThisDocument.Wizard 
 Set pbWizardProperty = pbWizard.Properties(1) 
 
 Debug.Print pbWizardProperty.CurrentValueId 
 
End Sub
```


