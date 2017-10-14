---
title: Window.ReviewerMarkupVisible Property (Visio)
keywords: vis_sdr.chm11660011
f1_keywords:
- vis_sdr.chm11660011
ms.prod: visio
api_name:
- Visio.Window.ReviewerMarkupVisible
ms.assetid: 7b13a89c-4835-93cc-aece-fcbad1a7ed22
ms.date: 06/08/2017
---


# Window.ReviewerMarkupVisible Property (Visio)

Determines whether reviewer markup, for a particular reviewer or all reviewers, is visible in a Microsoft Visio window that displays a drawing page. Read/write.


## Syntax

 _expression_ . **ReviewerMarkupVisible**( **_ReviewerID_** )

 _expression_ A variable that represents a **Window** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ReviewerID_|Optional| **Long**|The ID of a particular reviewer. If you do not pass a reviewer ID, the  **ReviewerMarkupVisible** property value specifies visibility for all reviewers.|

### Return Value

Boolean


## Remarks

Use the  **ReviewerMarkupVisible** property to get and set the current status (visible or not) of reviewer markup, either for one or all reviewers, in a window that can display markup overlays. Setting the **ReviewerMarkupVisible** property corresponds to setting reviewer visibility status in the **Show Markup Overlays** section of the **Reviewing** task pane in the user interface. For example, setting **ReviewerMarkupVisible** to **True** without passing a value for _ReviewerID_ is equivalent to clicking **Show All** in the **Reviewing** task pane. And setting **ReviewerMarkupVisible** to **False** while passing the ID of a particular reviewer is equivalent to clearing the box for that reviewer in the taskpane.

The  **ReviewerMarkupVisible** property is enabled only when the parent window displays a Visio drawing page, and not another type of Visio window, such as a stencil or ShapeSheet window, for example.

When viewing markup is disabled in the user interface, setting the  **ReviewerMarkupVisible** property is also disabled. If you attempt to set **ReviewerMarkupVisible** when viewing markup is disabled, Visio will display an error message. To be able to set **ReviewerMarkupVisible** , you must enable viewing markup by clicking **Show Markup** on the **Review** tab. Alternatively, you can enable viewing markup on existing markup overlays by changing the value of the ViewMarkup cell in the Document Properties section of the document's ShapeSheet. Use the following code:




```vb
ActiveDocument.DocumentSheet.CellsSRC(visSectionObject, visRowDoc, visDocViewMarkup).FormulaU = True
```


## Example

This Microsoft Visual Basic for Applications (VBA) macro uses the  **ReviewerMarkupVisible** property to get the current visibility status of reviewer markup for all reviewers in the active Visio drawing window. Then it switches the status to the opposite value. This example assumes that the active window contains markup overlays.


```vb
Public Sub ReviewerMarkupVisible_Example() 
 
 ActiveWindow.Document.DocumentSheet.CellsSRC(visSectionObject, visRowDoc, visDocViewMarkup).FormulaU = True 
 
 Debug.Print ActiveWindow.ReviewerMarkupVisible 
 ActiveWindow.ReviewerMarkupVisible = Not ActiveWindow.ReviewerMarkupVisible 
 Debug.Print ActiveWindow.ReviewerMarkupVisible 
 
End Sub
```


