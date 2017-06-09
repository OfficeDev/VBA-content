---
title: Search.SearchSubFolders Property (Outlook)
keywords: vbaol11.chm2256
f1_keywords:
- vbaol11.chm2256
ms.prod: outlook
api_name:
- Outlook.Search.SearchSubFolders
ms.assetid: 26dd1970-ba59-9f6a-8cf6-3dba0f9668b2
ms.date: 06/08/2017
---


# Search.SearchSubFolders Property (Outlook)

Returns a  **Boolean** indicating whether the scope of the specified search included the subfolders of any folders searched. Read-only.


## Syntax

 _expression_ . **SearchSubFolders**

 _expression_ A variable that represents a **Search** object.


## Remarks

This property is determined by the  _SearchSubfolders_ argument of the **[AdvancedSearch](application-advancedsearch-method-outlook.md)** method and is specified when the search is initiated. If **True** , the **Search** object searches through any subfolders in the specified filter path.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a  **Search** object. The user's **Inbox** is specified as the scope of the search and the **SearchSubFolders** property is set to **True** . The event subroutine fires when the search has completed and displays the **Tag** and **Scope** properties for the new object as well as the results of the search.


```vb
Public blnSearchComp As Boolean 
 
 
 
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 MsgBox "The AdvancedSearchComplete Event fired for " &; SearchObject.Tag &; _ 
 
 " and the scope was " &; SearchObject.Scope 
 
 blnSearchComp = True 
 
End Sub 
 
 
 
Sub TestAdvancedSearchComplete() 
 
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
 
 SearchSubFolders:=True, Tag:="FlagSearch") 
 
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

