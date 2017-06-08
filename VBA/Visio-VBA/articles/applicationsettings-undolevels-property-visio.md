---
title: ApplicationSettings.UndoLevels Property (Visio)
keywords: vis_sdr.chm16251490
f1_keywords:
- vis_sdr.chm16251490
ms.prod: visio
api_name:
- Visio.ApplicationSettings.UndoLevels
ms.assetid: 5d4ad370-254d-3b99-21d9-2cbdf60842a6
ms.date: 06/08/2017
---


# ApplicationSettings.UndoLevels Property (Visio)

Determines the number of consecutive actions the user can undo in Microsoft Visio. Read/write.


## Syntax

 _expression_ . **UndoLevels**

 _expression_ A variable that represents an **ApplicationSettings** object.


### Return Value

 **Long**


## Remarks

Setting the  **UndoLevels** property is equivalent to setting the **Undo levels** option on the **General** tab in the **Visio Options** dialog box (click the **File** tab, and then click **Options**).

The minimum possible value for  **UndoLevels** is 0 (zero); the maximum is 99. The default value is 20.

The higher the value of  **UndoLevels** , the more memory is required to store the actions.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **UndoLevels** property to print the current number of undo levels in the Immediate window. It also shows how to get an **ApplicationSettings** object from the Visio **Application** object.


```vb
Public Sub UndoLevels_Example() 
 
    Dim vsoApplicationSettings As Visio.ApplicationSettings 
    Dim lngUndoLevels As Long 
 
    Set vsoApplicationSettings = Visio.Application.Settings 
    lngUndoLevels = vsoApplicationSettings.UndoLevels 
 
    Debug.Print lngUndoLevels 
 
End Sub
```


