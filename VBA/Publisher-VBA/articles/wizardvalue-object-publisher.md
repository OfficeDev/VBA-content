---
title: WizardValue Object (Publisher)
keywords: vbapb10.chm2162687
f1_keywords:
- vbapb10.chm2162687
ms.prod: publisher
api_name:
- Publisher.WizardValue
ms.assetid: 15b60632-d1b1-c62b-0264-72d65bd1fe82
ms.date: 06/08/2017
---


# WizardValue Object (Publisher)

Represents a possible value for the specified wizard property.
 


## Example

Use the  **[Item](wizardvalues-item-property-publisher.md)** property of the **WizardValues** collection to return a **WizardValue** object. The following example displays the current value for the first wizard property in the active publication and then lists all the other possible values.
 

 

```
Dim valAll As WizardValues 
Dim valLoop As WizardValue 
 
With ActiveDocument.Wizard 
 Set valAll = .Properties(1).Values 
 
 MsgBox "Wizard: " &amp; .Name &amp; vbLf &amp; _ 
 "Property: " &amp; .Properties(1).Name &amp; vbLf &amp; _ 
 "Current value: " &amp; .Properties(1).CurrentValueId 
 
 For Each valLoop In valAll 
 MsgBox "Possible value: " &amp; valLoop.ID &amp; " (" &amp; valLoop.Name &amp; ")" 
 Next valLoop 
End With
```


## Properties



|**Name**|
|:-----|
|[Application](wizardvalue-application-property-publisher.md)|
|[ID](wizardvalue-id-property-publisher.md)|
|[Name](wizardvalue-name-property-publisher.md)|
|[Parent](wizardvalue-parent-property-publisher.md)|

