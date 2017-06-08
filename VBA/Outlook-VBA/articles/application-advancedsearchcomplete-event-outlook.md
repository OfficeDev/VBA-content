---
title: Application.AdvancedSearchComplete Event (Outlook)
keywords: vbaol11.chm435
f1_keywords:
- vbaol11.chm435
ms.prod: outlook
api_name:
- Outlook.Application.AdvancedSearchComplete
ms.assetid: 4f33ad44-20a3-62cd-aa1b-db74581ebb3c
ms.date: 06/08/2017
---


# Application.AdvancedSearchComplete Event (Outlook)

Occurs when the  **[AdvancedSearch](application-advancedsearch-method-outlook.md)** method has completed.


## Syntax

 _expression_ . **AdvancedSearchComplete**( **_SearchObject_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SearchObject_|Required| **[Search](search-object-outlook.md)**|The  **Search** object returned by the **[AdvancedSearch](application-advancedsearch-method-outlook.md)** method.|

## Remarks

The  **AdvancedSearchComplete** event is used to return the object that was created by the **AdvancedSearch** method. This event only fires when the **AdvancedSearch** method is executed programmatically.


## Example

The following Visual Basic for Applications (VBA) example searches the  **Inbox** for items where the subject is equal to "Test" and displays the names of the senders of the e-mail items returned by the search. The `AdvanceSearchComplete` event procedure sets the boolean `blnSearchComp` to **True** when the search is complete. This boolean variable is used by the `TestAdvancedSearchComplete()` procedure to determine when the search is complete. The sample code must be placed in a class module such as `ThisOutlookSession`. The  `TestAdvancedSearchComplete()` procedure must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public blnSearchComp As Boolean 
 
 
 
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 MsgBox "The AdvancedSearchComplete Event fired." 
 
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


[Application Object](application-object-outlook.md)

