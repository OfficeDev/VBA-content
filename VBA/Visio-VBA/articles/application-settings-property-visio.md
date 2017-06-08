---
title: Application.Settings Property (Visio)
keywords: vis_sdr.chm10051635
f1_keywords:
- vis_sdr.chm10051635
ms.prod: visio
api_name:
- Visio.Application.Settings
ms.assetid: b62413cb-a038-2679-8701-47ba700a93c4
ms.date: 06/08/2017
---


# Application.Settings Property (Visio)

Returns an  **ApplicationSettings** object, which you can use to set Microsoft Visio application properties. Read-only.


## Syntax

 _expression_ . **Settings**

 _expression_ A variable that represents an **Application** object.


### Return Value

 **ApplicationSettings**


## Remarks

Use the  **Settings** property of the **Application** object to get an **ApplicationSettings** object that you can then use to set various application properties corresponding to those in the **Options** dialog box (click the **File** tab, and then click **Options**) and the  **Snap &; Glue** dialog box (on the **View** tab, click the **Visual Aids** arrow).


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Settings** property to get an **ApplicationSettings** object. It also shows how to use the **RecentFilesListSize** property to get the number of entries in the list of recently used files in Visio.


```vb
Public Sub Settings_Example() 
 
    Dim vsoApplicationSettings As Visio.ApplicationSettings 
    Dim lngListSize As Long 
 
    Set vsoApplicationSettings = Visio.Application.Settings 
    lngListSize = vsoApplicationSettings.RecentFilesListSize 
 
    Debug.Print lngListSize 
 
End Sub
```


