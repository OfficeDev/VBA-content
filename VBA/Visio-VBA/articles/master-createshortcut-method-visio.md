---
title: Master.CreateShortcut Method (Visio)
keywords: vis_sdr.chm10716150
f1_keywords:
- vis_sdr.chm10716150
ms.prod: visio
api_name:
- Visio.Master.CreateShortcut
ms.assetid: e808ba09-b85a-52bb-55e2-ced37f426a3b
ms.date: 06/08/2017
---


# Master.CreateShortcut Method (Visio)

Creates a shortcut for a master.


## Syntax

 _expression_ . **CreateShortcut**

 _expression_ A variable that represents a **Master** object.


### Return Value

MasterShortcut


## Remarks

The new master shortcut is created in the same document as the target master and is added to the document's  **MasterShortcuts** collection. If you are trying to create a shortcut in a stencil, the stencil must therefore be editable for this method to succeed.




 **Note**  Starting with Microsoft Office Visio 2003, only user-created stencils are editable. By default, Visio stencils are not editable. 

The new shortcut's name is "Shortcut to X", where "X" is the name of the target master. The shortcut's  **TargetDocumentName** and **TargetMasterName** properties identify the target master. So once a shortcut has been created, it can be moved or copied into other documents.

You cannot create a shortcut to a master in an unsaved stencil. If you try to do so, the  **CreateShortcut** method returns an error.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **CreateShortcut** method to create a shortcut to a master. This example assumes that you have created a stencil named "SampleStencil.vss" containing a master named "SampleMaster" and saved it to the C drive on your computer.


```vb
 
Public Sub CreateShortcut_Example() 
 
 Dim vsoApplication As Visio.Application 
 Dim vsoMasterShortcut As MasterShortcut 
 
 Set vsoApplication = ActiveDocument.Application 
 Set vsoMasterShortcut = vsoApplication.Documents("C:\SampleStencil.vss").Masters("SampleMaster").CreateShortcut 
 
End Sub
```


