---
title: MasterShortcut.TargetDocumentName Property (Visio)
keywords: vis_sdr.chm16014495
f1_keywords:
- vis_sdr.chm16014495
ms.prod: visio
api_name:
- Visio.MasterShortcut.TargetDocumentName
ms.assetid: 0f4ede53-403c-f88e-b76f-fbf620ff54d1
ms.date: 06/08/2017
---


# MasterShortcut.TargetDocumentName Property (Visio)

Gets and sets the path and file name of a Microsoft Visio document (usually a stencil) that contains the master to which a master shortcut refers. Read/write.


## Syntax

 _expression_ . **TargetDocumentName**

 _expression_ A variable that represents a **MasterShortcut** object.


### Return Value

String


## Remarks

If the target document is moved, deleted, or renamed, or if the property is set to the path of a nonexistent file, the application will not be able to access the shortcut's target master. As a result, the end user will not be able to use the shortcut to drop shapes onto their drawing.

If the  **TargetDocumentName** property contains a file name but no path, the application looks for the target document in the file path set in the **File Locations** dialog box (click the **File** tab, click **Options**, click  **Advanced**, and then, under  **General**, click  **File Locations**). The name may refer either to the document's file name or to one of its alternate file names. To set an alternate name for a document, use the  **AlternateNames** property.

The  **TargetDocumentName** property does not support file names that have relative paths.


