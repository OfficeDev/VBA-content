---
title: Pages.AddWizardPage Method (Publisher)
keywords: vbapb10.chm458758
f1_keywords:
- vbapb10.chm458758
ms.prod: publisher
api_name:
- Publisher.Pages.AddWizardPage
ms.assetid: c56db218-d0f4-4f13-dfde-6198dc63cc81
ms.date: 06/08/2017
---


# Pages.AddWizardPage Method (Publisher)

Adds the specified new wizard page to a specified location in a publication.


## Syntax

 _expression_. **AddWizardPage**( **_After_**,  **_PageType_**,  **_AddHyperlinkToWebNavBar_**)

 _expression_A variable that represents a  **Pages** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|After|Required| **Long**|The page after which to place the new wizard page.|
|PageType|Optional| **PbWizardPageType**|The type of wizard page to add.|
|AddHyperlinkToWebNavBar|Optional| **Boolean**|Specifies whether a link to the new page will be added to the automatic navigation bars of existing pages. Default is  **False**, which means that if this argument is omitted, links to this page will not be added to the automatic navigation bars of existing pages.|

## Remarks

You can add wizard pages only to similar wizard publications. For example, you can add a Catalog Calendar Wizard page to a catalog but not to a newsletter. An error occurs if you try to add a wizard page to a different type of publication.

The PageType parameter can be one of the  **[PbWizardPageType](pbwizardpagetype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example creates a new catalog publication, adds the wizard calendar page after the first page of the catalog, and adds the page as a link to each Web navigation bar set of the publication.


```vb
Sub AddNewWizardPage() 
 Dim PubApp As Publisher.Application 
 Dim PubDoc As Publisher.Document 
 Set PubApp = New Publisher.Application 
 Set PubDoc = PubApp.NewDocument(Wizard:=pbWizardCatalogs, _ 
 Design:=7) 
 PubDoc.Pages.AddWizardPage After:=1, _ 
 PageType:=pbWizardPageTypeCatalogCalendar, _ 
 AddHyperLinkToWebNavBar:=True 
 PubApp.ActiveWindow.Visible = True 
End Sub
```

This example verifies that the active document is a catalog and, if it is, adds a catalog form after the first page but does not add the page as a link in any Web navigation bar sets.




```vb
Sub InsertCatalogWizardPage() 
 With ActiveDocument 
 If .Wizard.ID = 161 Then 
 .Pages.AddWizardPage After:=1, _ 
 PageType:=pbWizardPageTypeCatalogForm, _ 
 AddHyperLinkToWebNavBar:=False 
 End If 
 End With 
End Sub
```


