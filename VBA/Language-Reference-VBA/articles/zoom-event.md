---
title: Zoom Event
keywords: fm20.chm5224952
f1_keywords:
- fm20.chm5224952
ms.prod: office
api_name:
- Office.Zoom
ms.assetid: 8716a59d-2d1c-88e6-bf0c-f062dc11b1b5
ms.date: 06/08/2017
---


# Zoom Event



Occurs when the value of the  **Zoom** property changes.
 **Syntax**
For Frame **Private Sub**_object_ _**Zoom(**_Percent_**As Integer)**
For MultiPage **Private Sub**_object_ _**Zoom(**_index_**As Long**, _Percent_**As Integer)**
The  **Zoom** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _index_|Required. The index of the page in a  **MultiPage** associated with this event.|
| _Percent_|Required. The percentage the form is to be zoomed. Valid values range from 10 percent to 400 percent.|
 **Remarks**
The value of the  **Zoom** property identifies how the size of the form or **Page** changes. The value of the property indicates how the size of the control should change relative to its current size. Values less than 100 reduce the displayed size of the form; values greater than 100 increase the displayed size of the form.
You can set this property to any integer from 10 to 400.

