---
title: Search.Save Method (Outlook)
keywords: vbaol11.chm2260
f1_keywords:
- vbaol11.chm2260
ms.prod: outlook
api_name:
- Outlook.Search.Save
ms.assetid: a6dbec81-67fd-e337-b640-4f94ab36218f
ms.date: 06/08/2017
---


# Search.Save Method (Outlook)

Saves the search results to a Search Folder.


## Syntax

 _expression_ . **Save** **_SchFldrName_**

 _expression_ A variable that represents a **Search** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SchFldrName_|Required| **String**|A string that represents the Search Folder name.|

## Remarks

The  **Save** method displays an error if a Search Folder with the same name already exists.


## Example

The following Microsoft Visual Basic for Applications (VBA) example searches the Inbox for items with Subject line equal to 'Test' and saves the results in a Search Folder. The  `AdvanceSearchComplete` event procedure sets the **Boolean** `blnSearchComp` to **True** when the search is complete. This **Boolean** variable is used by the `TestAdvancedSearchComplete()` procedure to determine when the search is complete. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `TestAdvancedSearchComplete()` procedure must be called before the event procedure can be called by Outlook.


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
 
 sch.Save("Subject Test") 
 
End Sub
```


## See also


#### Concepts


[Search Object](search-object-outlook.md)

