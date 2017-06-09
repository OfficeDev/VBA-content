---
title: Window.Activate Method (Visio)
keywords: vis_sdr.chm11616000
f1_keywords:
- vis_sdr.chm11616000
ms.prod: visio
api_name:
- Visio.Window.Activate
ms.assetid: e34a74e0-8a47-a0bb-4ac5-6fdc8c9e5e08
ms.date: 06/08/2017
---


# Window.Activate Method (Visio)

Activates a window.


## Syntax

 _expression_ . **Activate**

 _expression_ An expression that returns a **Window** object.


### Return Value

Nothing


## Remarks

Microsoft Visio can have more than one window open at a time; however, only one window is active. Activating a window can change the objects returned by the  **ActiveWindow** , **ActivePage** , and **ActiveDocument** properties.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVWindow.Activate()**
    

## Example

The following macro creates two windows and then shows how to activate one of the windows.


```vb
Public Sub Activate_Example() 
 
    Dim vsoDocument As Visio.Document  
    Dim vsoWindow As Visio.Window  
      
    'Create two new windows by adding documents. 
    Set vsoDocument = Documents.Add("")  
    Set vsoWindow = ActiveWindow  
    Set vsoDocument = Documents.Add("")  
     
    'Use the Activate method to make the first 
    'window created the active window. 
    vsoWindow.Activate 
 
End Sub  
```


