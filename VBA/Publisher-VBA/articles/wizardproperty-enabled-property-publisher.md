---
title: WizardProperty.Enabled Property (Publisher)
keywords: vbapb10.chm1572871
f1_keywords:
- vbapb10.chm1572871
ms.prod: publisher
api_name:
- Publisher.WizardProperty.Enabled
ms.assetid: c66741c8-1493-ac90-4ecb-ed8d58743c69
ms.date: 06/08/2017
---


# WizardProperty.Enabled Property (Publisher)

 **True** if a wizard property is enabled. Read-only **Boolean**.


## Syntax

 _expression_. **Enabled**

 _expression_A variable that represents an  **WizardProperty** object.


### Return Value

Boolean


## Example

This example displays the name of each enabled wizard property in the active publication.


```vb
Sub SetEnabledProperty() 
 Dim wizProperty As WizardProperty 
 For Each wizProperty In ActiveDocument.Wizard.Properties 
 If wizProperty.Enabled = True Then 
 MsgBox "The name of the wizard property is " &; wizProperty.Name 
 End If 
 Next 
End Sub
```


