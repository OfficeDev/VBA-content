---
title: Application.Window Property (Visio)
keywords: vis_sdr.chm10051480
f1_keywords:
- vis_sdr.chm10051480
ms.prod: visio
api_name:
- Visio.Application.Window
ms.assetid: fd996e7d-a204-ab0d-538a-439c61532bb9
ms.date: 06/08/2017
---


# Application.Window Property (Visio)

Returns the window associated with the current instance of Microsoft Visio. Read-only.


## Syntax

 _expression_ . **Window**

 _expression_ A variable that represents an **Application** object.


### Return Value

Window


## Remarks

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVApplication.Window**
    

## Example

The following macro shows how to use the  **Window** property to print the caption of the window associated with the current instance of Visio in the Immediate window.


```vb
 
Public Sub Window_Example()  
  
    Debug.Print  Application.Window.Caption 
 
End Sub
```


