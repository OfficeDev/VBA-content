---
title: Window.Paste Method (Visio)
ms.prod: visio
api_name:
- Visio.Window.Paste
ms.assetid: e5535c75-5a43-48dc-bd77-50db003809ba
ms.date: 06/08/2017
---


# Window.Paste Method (Visio)

This object, member, or enumeration is deprecated and is not intended to be used in your code. Pastes the contents of the Clipboard into an object.


## Version Information

Version Added: Visio 2.0 


### Syntax

 _expression_ . **Paste**

 _expression_ A variable that represents a **Window** object.


## Remarks

The  **Window** object's **Paste** method is now obsolete. Use the **Paste** or **PasteSpecial** method of the[Page](page-object-visio.md), [Master](master-object-visio.md), or [Shape](shape-object-visio.md) object. (Use the **Shape** object in the case of group shapes.)

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVWindow.Paste()**
    

