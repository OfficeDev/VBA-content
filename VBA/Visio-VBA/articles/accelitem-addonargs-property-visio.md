---
title: AccelItem.AddOnArgs Property (Visio)
keywords: vis_sdr.chm14513045
f1_keywords:
- vis_sdr.chm14513045
ms.prod: visio
api_name:
- Visio.AccelItem.AddOnArgs
ms.assetid: ebc91b1e-7780-1cdd-04dc-4a859c8929ff
ms.date: 06/08/2017
---


# AccelItem.AddOnArgs Property (Visio)

Gets or sets the argument string that you send to the add-on associated with a particular accelerator key. Read/write.


## Syntax

 _expression_ . **AddOnArgs**

 _expression_ An expression that returns a **AccelItem** object.


### Return Value

String


## Remarks

An argument's string can be anything appropriate for the add-on. However, the arguments are packaged together with other information into a command string, which cannot exceed 127 characters. For best results, limit arguments to 50 characters.

An object's  **AddOnName** property indicates the name of the add-on to which the arguments are sent.

 Beginning with Visio 2002, the **AddOnName** property used in the following example cannot execute a string that contains arbitrary Microsoft Visual Basic code. To call code that in previous versions of Visio you would have passed to the **AddOnName** property, move it to a procedure in a document's Visual Basic project that is called from the **AddOnName** property, as shown in the following example.


