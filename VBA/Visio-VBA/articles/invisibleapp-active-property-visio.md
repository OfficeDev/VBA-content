---
title: InvisibleApp.Active Property (Visio)
keywords: vis_sdr.chm17513020
f1_keywords:
- vis_sdr.chm17513020
ms.prod: visio
api_name:
- Visio.InvisibleApp.Active
ms.assetid: 69277382-436e-3241-ccbf-1da229e04c3d
ms.date: 06/08/2017
---


# InvisibleApp.Active Property (Visio)

Indicates whether the instance of Microsoft Visio represented by the  **Application** object is the active application on the Microsoft Windows desktop—the application that has the highlighted title bar. Read-only.


## Syntax

 _expression_ . **Active**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Integer


## Remarks

The active application on the Windows desktop is distinct from the active Visio instance, which is returned by a call to the OLE  **GetActiveObject** method ( **GetObject** method in Microsoft Visual Basic). The **GetObject** method retrieves the instance of Visio that was most recently activated, which may or may not be the active application on the desktop at that moment. Of all instances of Visio that are currently running, only one is the active Visio instance.

For example, suppose you start one instance of Visio and one of another application, such as Microsoft Excel.




- If the instance of Visio is the active application on your desktop,  **GetObject** (, "visio.application") retrieves that instance, and its **Active** property is **True** .
    
- If you activate the instance of Microsoft Excel,  **GetObject** (, "visio.application") retrieves the same instance of Visio, but its **Active** property is **False** .
    


If an  **Application** object's **Active** property is **True** , you can assume that the corresponding instance of Visio is the active instance of Visio unless the **InPlace** property is also **True** . If an instance of Visio is activated for in-place editing in a container application, that instance may not necessarily report itself as the active instance of Visio.


## Example

The following Visual Basic program shows how to get the active instance of Visio.


```vb
 
Public Sub Active_Example() 
 
 Dim vsoApplication1 As Visio.Application 
 Dim vsoApplication2 As Visio.Application 
 
 'Create two new instances of Visio. 
 Set vsoApplication1 = CreateObject("visio.application") 
 Set vsoApplication2 = CreateObject("visio.application") 
 
 'Use the Active property to determine whether 
 'the instance of Visio is active. 'Result = False. Prints "0" in the Immediate window 
 Debug.Print vsoApplication1.Active 
 
 'Result = True. Prints "-1" in the Immediate window. 
 Debug.Print vsoApplication2.Active 
 
End Sub
```


