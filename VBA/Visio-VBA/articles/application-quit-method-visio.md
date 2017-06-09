---
title: Application.Quit Method (Visio)
keywords: vis_sdr.chm10016460
f1_keywords:
- vis_sdr.chm10016460
ms.prod: visio
api_name:
- Visio.Application.Quit
ms.assetid: 1f8b73cd-10bd-e571-eee4-db05d9aa12cc
ms.date: 06/08/2017
---


# Application.Quit Method (Visio)

Closes the indicated instance of Microsoft Visio.


## Syntax

 _expression_ . **Quit**

 _expression_ A variable that represents an **Application** object.


### Return Value

Nothing


## Remarks

If the  **Quit** method is invoked when any open document has unsaved changes, a dialog box appears asking if you want to save the document. To quit the application without saving and seeing the dialog box, set the **Saved** property of the **Document** object representing the document to **True** immediately before quitting. Set the **Saved** property to **True** only if you are sure you want to close the document without saving changes, because you will lose any unsaved changes.


