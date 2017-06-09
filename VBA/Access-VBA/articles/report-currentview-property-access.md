---
title: Report.CurrentView Property (Access)
keywords: vbaac10.chm13826
f1_keywords:
- vbaac10.chm13826
ms.prod: access
api_name:
- Access.Report.CurrentView
ms.assetid: d1c33390-75f1-4e11-0735-a8860211b4ce
ms.date: 06/08/2017
---


# Report.CurrentView Property (Access)

You can use the  **CurrentView** property to determine how a report is currently displayed. Read/write **Integer**.


## Syntax

 _expression_. **CurrentView**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **CurrentView** property uses the following settings.



|**Setting**|**Report Displayed In:**|
|:-----|:-----|
|0|Design view|
|5|Print Preview|
|6|Report view|
|7|Layout view|
Use this property to perform different tasks depending on the current view. For example, an event procedure could determine which view the form is displayed in and perform one task if the form is displayed in Form view or another task if it's displayed in Datasheet view.


## See also


#### Concepts


[Report Object](report-object-access.md)

