---
title: Search.Results Property (Outlook)
keywords: vbaol11.chm2255
f1_keywords:
- vbaol11.chm2255
ms.prod: outlook
api_name:
- Outlook.Search.Results
ms.assetid: 405166fa-d0bc-33d2-f4aa-908fb821edd6
ms.date: 06/08/2017
---


# Search.Results Property (Outlook)

Returns a  **[Results](results-object-outlook.md)** collection that specifies the results of the search. Read-only.


## Syntax

 _expression_ . **Results**

 _expression_ A variable that represents a **Search** object.


## Example

The following Visual Basic for Applications (VBA) example searches the  **Inbox** for items with a subject that equals "Test" and displays the names of the senders of the e-mail items returned by the search. The `AdvanceSearchComplete` event procedure sets the boolean `blnSearchComp` to **True** when the search is complete. This boolean variable is used by the `TestAdvancedSearchComplete()` procedure to determine when the search is complete. The sample code must be placed in a class module, such as **ThisOutlookSession**, and the  `TestAdvancedSearchComplete()` procedure must be called before the event procedure can be called by Outlook.


```vb
Public blnSearchComp As Boolean 
 
 
 
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 MsgBox "The AdvancedSearchComplete Event fired" 
 
 blnSearchComp = True 
 
End Sub 
 
 
 
Sub TestAdvancedSearchComplete() 
 
 Dim sch As Outlook.Search 
 
 Dim rsts As Outlook.Results 
 
 Dim i As Integer 
 
 blnSearchComp = False 
 
 Const strF As String = "urn:schemas:mailheader:subject = 'Test'" 
 
 Const strS As String = "Inbox" 
 
 Set sch = Application.AdvancedSearch(strS, strF) 
 
 While blnSearchComp = False 
 
 DoEvents 
 
 Wend 
 
 Set rsts = sch.Results 
 
 For i = 1 To rsts.Count 
 
 MsgBox rsts.Item(i).SenderName 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[Search Object](search-object-outlook.md)

