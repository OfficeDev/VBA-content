---
title: Search.Tag Property (Outlook)
keywords: vbaol11.chm2258
f1_keywords:
- vbaol11.chm2258
ms.prod: outlook
api_name:
- Outlook.Search.Tag
ms.assetid: f0341885-ea75-2277-e55b-827f62165ab2
ms.date: 06/08/2017
---


# Search.Tag Property (Outlook)

Returns a  **String** specifying the name of the current search. The **Tag** property is used to identify a specific search. Read-only.


## Syntax

 _expression_ . **Tag**

 _expression_ A variable that represents a **Search** object.


## Remarks

The  **Tag** property is set by using the **[AdvancedSearch](application-advancedsearch-method-outlook.md)** method when the **[Search](search-object-outlook.md)** object is created.


## Example

The following Visual Basic for Applications (VBA) example searches through the user's  **Inbox** for all items that do not have a flag. The name "FlagSearch", specified by the **Tag** property, is given to the search. The `AdvanceSearchComplete` event procedure sets the boolean `blnSearchComp` to **True** when the search is complete. This boolean variable is used by the `TestAdvancedSearchComplete()` procedure to determine when the search is complete. The sample code must be placed in a class module such as **ThisOutlookSession**, and the  `TestAdvancedSearchComplete()` subroutine must be called before the event procedure can be called by Outlook. The `AdvanceSearchComplete` event procedure displays the tag to the user so the user can identify which search was completed because usually the search is asynchronous (use the **[IsSynchronous](search-issynchronous-property-outlook.md)** property to determine if the search will be synchronous or asynchronous), and you can execute multiple searches simultaneously.


```vb
Public blnSearchComp As Boolean 
 
 
 
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 MsgBox "The AdvancedSearchComplete Event fired for " &; _ 
 
 SearchObject.Tag &; " and the scope was " &; SearchObject.Scope 
 
 blnSearchComp = True 
 
End Sub 
 
 
 
Sub TestAdvancedSearch111Complete() 
 
 'List all items in the Inbox that do NOT have a flag: 
 
 Dim objSch As Outlook.Search 
 
 Const strF As String = "urn:schemas:httpmail:messageflag IS NULL" 
 
 Const strS As String = "Inbox" 
 
 Dim rsts As Outlook.Results 
 
 Dim i As Integer 
 
 blnSearchComp = False 
 
 Const strF1 As String = "urn:schemas:mailheader:subject = 'Test'" 
 
 Const strS1 As String = "Inbox" 
 
 Set objSch = _ 
 
 Application.AdvancedSearch(Scope:=strS1, Filter:=strF1, _ 
 
 Tag:="FlagSearch") 
 
 While blnSearchComp = False 
 
 DoEvents 
 
 Wend 
 
 Set rsts = objSch.Results 
 
 For i = 1 To rsts.Count 
 
 MsgBox rsts.Item(i).SenderName 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[Search Object](search-object-outlook.md)

