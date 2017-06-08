---
title: Application.WindowBeforeDoubleClick Event (PowerPoint)
keywords: vbapp10.chm621003
f1_keywords:
- vbapp10.chm621003
ms.prod: powerpoint
api_name:
- PowerPoint.Application.WindowBeforeDoubleClick
ms.assetid: 9b270238-1658-df56-4208-9cb98666519c
ms.date: 06/08/2017
---


# Application.WindowBeforeDoubleClick Event (PowerPoint)

Occurs when you double-click the items in the views listed in the following table.


## Syntax

 _expression_. **WindowBeforeDoubleClick**( **_Sel_**, **_Cancel_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sel_|Required|**Selection**|The selection below the mouse pointer when the double-click occurs.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the default double-click action isn't performed when the procedure is finished.|

## Remarks

For information about using events with the  **Application** object, see[How to: Use Events with the Application Object](use-events-with-the-application-object.md).



|**View**|**Item**|
|:-----|:-----|
|Normal or slide view|Shape|
|Slide sorter view|Slide|
|Notes page view|Slide image|
The default double-click action occurs after this event unless the Cancel argument is set to  **True**.


## Example

In slide sorter view, the default double-click event for any slide is to change to slide view. In this example, if the active presentation is displayed in slide sorter view, the default action is preempted by the  **WindowBeforeDoubleClick** event. The event procedure changes the view to normal view and then cancels the change to slide view by setting the Cancel argument to **True**.


```vb
Private Sub App_WindowBeforeDoubleClick (ByVal Sel As Selection, ByVal Cancel As Boolean)

    With Application.ActiveWindow

        If .ViewType = ppViewSlideSorter Then

           .ViewType = ppViewNormal

            Cancel = True

        End If

    End With

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

