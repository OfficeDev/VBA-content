---
title: Application.WindowActivate Event (PowerPoint)
keywords: vbapp10.chm621009
f1_keywords:
- vbapp10.chm621009
ms.prod: powerpoint
api_name:
- PowerPoint.Application.WindowActivate
ms.assetid: 0d83fda3-b0ad-18df-57bf-c34dafcf782f
ms.date: 06/08/2017
---


# Application.WindowActivate Event (PowerPoint)

Occurs when the application window or any document window is activated.


## Syntax

 _expression_. **WindowActivate**( **_Pres_**, **_Wn_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The presentation displayed in the activated window.|
| _Wn_|Required|**DocumentWindow**|The activated document window.|

## Remarks

For information about using events with the  **Application** object, see[How to: Use Events with the Application Object](use-events-with-the-application-object.md).

If your Visual Studio solution includes the  **Microsoft.Office.Interop.PowerPoint** reference, this event maps to the following types:


-  **Microsoft.Office.Interop.PowerPoint.EApplication_WindowActivateEventHandler** (the **WindowActivate** delegate.)
    
-  **Microsoft.Office.Interop.PowerPoint.EApplication_Event.WindowActivate** (the **WindowActivate** event.)
    

## Example

This example opens every activated presentation in slide sorter view.


```vb
Private Sub App_WindowActivate (ByVal Pres As Presentation, ByVal Wn As DocumentWindow) 
    Wn.ViewType = ppViewSlideSorter 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

