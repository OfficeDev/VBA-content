---
title: MasterShortcut.TargetMasterName Property (Visio)
keywords: vis_sdr.chm16014500
f1_keywords:
- vis_sdr.chm16014500
ms.prod: visio
api_name:
- Visio.MasterShortcut.TargetMasterName
ms.assetid: 6c59e85e-9ee8-afb5-c631-c0d790dd666e
ms.date: 06/08/2017
---


# MasterShortcut.TargetMasterName Property (Visio)

Gets or sets the name of the master to which the master shortcut refers. Read/write.


## Syntax

 _expression_ . **TargetMasterName**

 _expression_ A variable that represents a **MasterShortcut** object.


### Return Value

String


## Remarks

The name specified by this property must be the target master's universal name, not its localized name. For a given master, the universal name is specified by the master's  **NameU** property, and the local name by its **Name** property.

When the user drops a master shortcut onto a drawing page, the application first locates the document identified by the shortcut's  **TargetDocumentName** property, and then it searches that document for a master whose universal name matches the shortcut's **TargetMasterName** property. Once located, the target master (not the shortcut) is used to create the new shape instance on the drawing page.


