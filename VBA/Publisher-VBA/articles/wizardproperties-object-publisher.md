---
title: WizardProperties Object (Publisher)
keywords: vbapb10.chm1572863
f1_keywords:
- vbapb10.chm1572863
ms.prod: publisher
api_name:
- Publisher.WizardProperties
ms.assetid: b3feecf2-ffbb-79de-8586-6a64df1b816a
ms.date: 06/08/2017
---


# WizardProperties Object (Publisher)

Represents the settings available in a publication design or in a Design Gallery object's wizard.
 


## Example

Use the  **[Properties](wizard-properties-property-publisher.md)** property with a **Wizard** object to return a **WizardProperties** collection. The following example reports on the publication design associated with the active publication, displaying its name and current settings.
 

 

```
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 MsgBox "Publication Design associated with " _ 
 &amp; "current publication: " .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Wizard property: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Note**  Depending on the language version of Microsoft Publisher that you are using, you may receive an error when using the above code. If this occurs, you will need to build in error handlers to circumvent the errors. For more information, see  **[Wizard Object](wizard-object-publisher.md)**.
 


## Methods



|**Name**|
|:-----|
|[FindPropertyById](wizardproperties-findpropertybyid-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](wizardproperties-application-property-publisher.md)|
|[Count](wizardproperties-count-property-publisher.md)|
|[Item](wizardproperties-item-property-publisher.md)|
|[Parent](wizardproperties-parent-property-publisher.md)|

