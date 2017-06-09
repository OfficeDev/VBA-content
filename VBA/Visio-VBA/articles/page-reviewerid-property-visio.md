---
title: Page.ReviewerID Property (Visio)
keywords: vis_sdr.chm10951670
f1_keywords:
- vis_sdr.chm10951670
ms.prod: visio
api_name:
- Visio.Page.ReviewerID
ms.assetid: f3de7746-f1f7-4a94-6fcb-e3c2775ed748
ms.date: 06/08/2017
---


# Page.ReviewerID Property (Visio)

Gets the reviewer ID associated with the markup overlay. Read-only.


## Syntax

 _expression_ . **ReviewerID**

 _expression_ A variable that represents a **Page** object.


### Return Value

Long


## Remarks

The  **ReviewerID** property is valid only for markup overlays. To determine if a page is a markup overlay, check to see whether **Page.Type** = **visTypeMarkup** (3). If you attempt to get the **ReviewerID** value for foreground pages and for background pages that are not markup overlays, Microsoft Visio returns an error.

The reviewer ID returned by the  **ReviewerID** property corresponds to one of the rows in the Reviewer section of the document's ShapeSheet. The ShapeSheet cell that contains the reviewer ID is hidden in the document ShapeSheet user interface, but you can determine the user name and initials associated with each reviewer ID by using the **Document.DocumentSheet.CellsSRC** property of the page. See the example that follows.




 **Note**  To view a document's ShapeSheet, on the  **Developer** tab, select **Drawing Explorer**, right-click the document's name, and then click  **Show ShapeSheet**.


## Example

This Microsoft Visual Basic for Applications (VBA) macro uses the  **ReviewerID** property to get the ID of the reviewer associated with a markup overlay and then prints the reviewer's name in the Immediate window. It first determines if the active page is a markup overlay, and if so, gets the reviewer ID. Before running this macro, make sure there is an active drawing page in the Visio drawing window.


```vb
Public Sub ReviewerID_Example() 
 Dim vsoPage As Visio.Page 
 Dim intCounter As Integer 
 
 Set vsoPage = ActivePage 
 
 If vsoPage.Type = visTypeMarkup Then 
 
 For intCounter = 0 To vsoPage.Document.DocumentSheet.RowCount(visSectionReviewer) - 1 
 
 If vsoPage.ReviewerID = vsoPage.Document.DocumentSheet.CellsSRC(visSectionReviewer, visRowReviewer + intCounter, visReviewerReviewerID).ResultStr(0) Then 
 
 Debug.Print "Reviewer name is: "; vsoPage.Document.DocumentSheet.CellsSRC(visSectionReviewer, visRowReviewer + intCounter, visReviewerName).ResultStr(0) 
 
 End If 
 
 Next intCounter 
 
 Else 
 
 Debug.Print "Active page is not a markup overlay." 
 
 End If 
 
End Sub
```


