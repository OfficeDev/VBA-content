---
title: Application.WizardCatalogVisible Property (Publisher)
keywords: vbapb10.chm131173
f1_keywords:
- vbapb10.chm131173
ms.prod: publisher
api_name:
- Publisher.Application.WizardCatalogVisible
ms.assetid: 99323335-aabd-6799-b6aa-c5d95b88064f
ms.date: 06/08/2017
---


# Application.WizardCatalogVisible Property (Publisher)

Returns or sets a  **Boolean** indicating whether the Wizard Catalog is visible. Read/write.


## Syntax

 _expression_. **WizardCatalogVisible**

 _expression_A variable that represents a  **Application** object.


### Return Value

Boolean


## Example

The following example stores the current state of the Wizard Catalog so that it can restore it later.


```vb
Sub WizardCatalogExample() 
 Dim blnWizardCatalog As Boolean 
 
 ' Store current state of Wizard Catalog. 
 blnWizardCatalog = Application.WizardCatalogVisible 
 
 ' Code can run here that shows or hides the Wizard 
 ' Catalog as necessary; the original setting 
 ' will be restored at the end of the procedure. 
 
 ' Restore original state of Wizard Catalog. 
 Application.WizardCatalogVisible = blnWizardCatalog 
End Sub
```


## See also


#### Concepts


 [Application Object](application-object-publisher.md)

