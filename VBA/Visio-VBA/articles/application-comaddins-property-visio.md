---
title: Application.COMAddIns Property (Visio)
keywords: vis_sdr.chm10050535
f1_keywords:
- vis_sdr.chm10050535
ms.prod: visio
api_name:
- Visio.Application.COMAddIns
ms.assetid: 182ea1e1-f896-f619-1bf0-df4a57b39abf
ms.date: 06/08/2017
---


# Application.COMAddIns Property (Visio)

Returns a reference to the  **COMAddIns** collection that represents all the Component Object Model (COM) add-ins currently registered in Microsoft Visio. Read-only.


## Syntax

 _expression_ . **COMAddIns**

 _expression_ A variable that represents an **Application** object.


### Return Value

Object


## Remarks

The COM add-ins that are currently registered are listed in the  **COM Add-Ins** dialog box (click the **File** tab, click **Options**, click  **Add-Ins**, and then click  **Go**).

To get information about the object returned by the  **COMAddIns** property:


1. In the  **Code** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab, click **Visual Basic**.
    
2. On the  **View** menu, click **Object Browser**.
    
3. In the  **Project/Library** list, click **Office**.
    
4. Under  **Classes**, examine the class named  **COMAddIns** .
    

## Example

This macro shows how to use the  **COMAddIns** property to list the COM add-ins registered with Visio.


```vb
 
Public Sub COMAddIns_Example()  
 
    Dim vsoCOMAddIns As COMAddIns  
    Dim vsoCOMAddIn As COMAddIn  
 
    'Get the set of COM add-ins.  
    Set vsoCOMAddIns = Application.COMAddIns  
 
    'List each COM add-in in the Immediate window. 
    For Each vsoCOMAddIn In vsoCOMAddIns  
        Debug.Print vsoCOMAddIn.Description  
    Next 
 
End Sub
```


