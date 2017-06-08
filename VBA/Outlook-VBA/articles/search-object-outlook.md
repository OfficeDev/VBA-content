---
title: Search Object (Outlook)
keywords: vbaol11.chm2248
f1_keywords:
- vbaol11.chm2248
ms.prod: outlook
api_name:
- Outlook.Search
ms.assetid: 226a5d49-3caf-90dd-725c-265404d1939f
ms.date: 06/08/2017
---


# Search Object (Outlook)

Contains information about individual searches performed against Outlook items.


## Remarks

The  **Search** object contains properties that define the type of search and the parameters of the search itself.

Use the  **[Application](http://msdn.microsoft.com/library/797003e7-ecd1-eccb-eaaf-32d6ddde8348%28Office.15%29.aspx)** object's **[AdvancedSearch](http://msdn.microsoft.com/library/7b433d8b-08b9-dff1-b854-287d76b47a90%28Office.15%29.aspx)** method to return a **Search** object.

Use the  **[AdvancedSearchComplete](http://msdn.microsoft.com/library/4f33ad44-20a3-62cd-aa1b-db74581ebb3c%28Office.15%29.aspx)** event to determine when a given search has completed.


## Example

The following Microsoft Visual Basic for Applications (VBA) example returns a search object named "SubjectSearch" and displays the object's  **[Tag](http://msdn.microsoft.com/library/f0341885-ea75-2277-e55b-827f62165ab2%28Office.15%29.aspx)** and **[Filter](http://msdn.microsoft.com/library/f6040465-da73-56f6-edb7-06d93bb8b531%28Office.15%29.aspx)** property values. The **Tag** property is used to identify a specific search once it has completed.


```
Sub SearchInboxFolder() 
 
'Searches the Inbox 
 
 
 
 Dim objSch As Search 
 
 Const strF As String = _ 
 
 "urn:schemas:mailheader:subject = 'Office Christmas Party'" 
 
 Const strS As String = "Inbox" 
 
 Const strTag As String = "SubjectSearch" 
 
 Set objSch = Application.AdvancedSearch(Scope:=strS, _ 
 
 Filter:=strF, SearchSubFolders:=True, Tag:=strTag) 
 
 
 
End Sub 
 

```

The following VBA example displays information about the search and the results of the search.




```
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
 
 
 
 Dim objRsts As Results 
 
 MsgBox "The search " &amp; SearchObject.Tag &amp; "has completed. 
 
 Set objRsts = SearchObject.Results 
 
 'Print out number in Results collection 
 
 Debug.Print objRsts.Count 
 
 'Print out each member of Results collection 
 
 For Each Item In objRsts 
 
 Debug.Print Item 
 
 Next 
 
 
 
End Sub 
 

```


## Methods



|**Name**|
|:-----|
|[GetTable](http://msdn.microsoft.com/library/3aba6b77-73a3-9620-9c18-b2e03c7b63bc%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/a6dbec81-67fd-e337-b640-4f94ab36218f%28Office.15%29.aspx)|
|[Stop](http://msdn.microsoft.com/library/c087e5aa-a846-56e1-a808-e8718096c3c9%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/9db2bcd4-d191-68c9-dd2a-f14a8372d541%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/178d0f62-75f9-20bd-d6dc-bcf04ae37422%28Office.15%29.aspx)|
|[Filter](http://msdn.microsoft.com/library/f6040465-da73-56f6-edb7-06d93bb8b531%28Office.15%29.aspx)|
|[IsSynchronous](http://msdn.microsoft.com/library/e240cc55-26c3-a560-4ee2-84b15da95e52%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/edd9777f-a764-8e35-4a66-05a0f838de0e%28Office.15%29.aspx)|
|[Results](http://msdn.microsoft.com/library/405166fa-d0bc-33d2-f4aa-908fb821edd6%28Office.15%29.aspx)|
|[Scope](http://msdn.microsoft.com/library/aa4b9aea-029f-6f80-87b1-b99c04ff9631%28Office.15%29.aspx)|
|[SearchSubFolders](http://msdn.microsoft.com/library/26dd1970-ba59-9f6a-8cf6-3dba0f9668b2%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/8d5a2300-dc21-0fbe-c7c0-17741caae30a%28Office.15%29.aspx)|
|[Tag](http://msdn.microsoft.com/library/f0341885-ea75-2277-e55b-827f62165ab2%28Office.15%29.aspx)|

## See also


#### Other resources


[Search Object Members](http://msdn.microsoft.com/library/543773b8-9f38-8d3e-2279-8f2a581ccd18%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
