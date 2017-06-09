---
title: Application.AutoRecoverInterval Property (Visio)
keywords: vis_sdr.chm10014705
f1_keywords:
- vis_sdr.chm10014705
ms.prod: visio
api_name:
- Visio.Application.AutoRecoverInterval
ms.assetid: 06aa731b-b426-a1a2-a25b-8ac32133eb1a
ms.date: 06/08/2017
---


# Application.AutoRecoverInterval Property (Visio)

Represents the time interval (in minutes) for how often you want to save copies of open documents that have unsaved changes in case of a power failure or an application error. Read/write.


## Syntax

 _expression_ . **AutoRecoverInterval**

 _expression_ A variable that represents an **Application** object.


### Return Value

Integer


## Remarks

Must be an integer value from zero (0) to 120, representing the interval in minutes. The default is 0. If the value of the  **AutoRecoverInterval** property is less than or equal to 0, no automatic recovery copies are created.

If the value of the  **AutoRecoverInterval** property is greater than 0, automatic recovery is enabled for all documents in the Microsoft Visio instance. To disable automatic recovery for a particular document, set its **AutoRecover** property to **False** .


## Example

The following Microsoft Visual Basic for Applications (VBA) macros show how to set the  **AutoRecoverInterval** property and how to use it to disable automatic recovery.


```vb
 
Public Sub AutoRecoverInterval_Example() 
  
    'Save automatic recovery copies of unsaved files 
    'every 10 minutes.  
    Application.AutoRecoverInterval = 10  
 
End Sub   
 
Public Sub DisableAutoRecover_Example() 
  
    'Tell Visio not to save automatic recovery copies of unsaved files.  
    Application.AutoRecoverInterval = 0  
 
End Sub
```


