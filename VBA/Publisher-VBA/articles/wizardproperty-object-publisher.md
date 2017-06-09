---
title: WizardProperty Object (Publisher)
keywords: vbapb10.chm1638399
f1_keywords:
- vbapb10.chm1638399
ms.prod: publisher
api_name:
- Publisher.WizardProperty
ms.assetid: 9f059422-5454-1902-a092-76e21e36a3f7
ms.date: 06/08/2017
---


# WizardProperty Object (Publisher)

Represents a setting that is part of a specific publication design or a Design Gallery object's wizard.
 


## Example

Use the  **[Item](wizardproperties-item-property-publisher.md)** property or the **[FindByPropertyID](wizardproperties-findpropertybyid-method-publisher.md)** method with the **WizardProperties** collection to return a single **WizardProperty** object. The following example reports on the publication design associated with the active publication, displaying its name and current settings.
 

 

```
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 Debug.Print "Publication Design associated with " _ 
 &amp; "current publication: " _ 
 &amp; .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Wizard property: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Note**  Depending on the language version of Microsoft Publisher that you are using, you may receive an error when using the above code. If this occurs, you will need to build in error handlers to circumvent the errors. For more information, see  **[Wizard Object](wizard-object-publisher.md)**.
 


## Properties



|**Name**|
|:-----|
|[Application](wizardproperty-application-property-publisher.md)|
|[CurrentValueId](wizardproperty-currentvalueid-property-publisher.md)|
|[Enabled](wizardproperty-enabled-property-publisher.md)|
|[ID](wizardproperty-id-property-publisher.md)|
|[Name](wizardproperty-name-property-publisher.md)|
|[Parent](wizardproperty-parent-property-publisher.md)|
|[Values](wizardproperty-values-property-publisher.md)|

