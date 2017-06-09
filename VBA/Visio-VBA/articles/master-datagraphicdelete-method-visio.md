---
title: Master.DataGraphicDelete Method (Visio)
keywords: vis_sdr.chm10760140
f1_keywords:
- vis_sdr.chm10760140
ms.prod: visio
api_name:
- Visio.Master.DataGraphicDelete
ms.assetid: aa84af70-975c-3747-1976-b872a6c2fa36
ms.date: 06/08/2017
---


# Master.DataGraphicDelete Method (Visio)

Deletes the  **Master** of type **visTypeDataGraphic** from the **Masters** collection of the document.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **DataGraphicDelete**

 _expression_ An expression that returns a **Master** object.


### Return Value

Nothing


## Remarks

The  **DataGraphicDelete** method deletes the data graphic master and removes the data graphic from all shapes in the drawing to which it has been applied. The **Master.Delete** method deletes only the data graphic master, leaving data graphics based on the master and already applied to shapes in the drawing intact.

Calling the  **DataGraphicDelete** method is the equivalent of right-clicking a data graphic in the **Data Graphics** task pane in the Microsoft Visio user interface and then clicking **Delete** on the context menu.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **DataGraphicDelete** to delete a **Master** of type **visTypeDataGraphic** from the **Masters** collection of the active document. The data graphic master deleted in this example is named "Data Graphic." You can determine the name of a data graphic master by pausing your mouse over a data graphic thumbnail in the **Data Graphics** task pane.


```vb
Public Sub DataGraphicDelete_example() 
 
    Dim vsoMaster As Visio.Master 
   
    Set vsoMaster = ActiveDocument.Masters("Data Graphic")    
    vsoMaster.DataGraphicDelete 
 
End Sub
```


