---
title: Store.IsInstantSearchEnabled Property (Outlook)
keywords: vbaol11.chm3264
f1_keywords:
- vbaol11.chm3264
ms.prod: outlook
api_name:
- Outlook.Store.IsInstantSearchEnabled
ms.assetid: 0fba75cc-c506-157b-7dfa-ec438e932f5c
ms.date: 06/08/2017
---


# Store.IsInstantSearchEnabled Property (Outlook)

Returns a  **Boolean** that indicates whether Instant Search is enabled and operational on a store. Read-only.


## Syntax

 _expression_ . **IsInstantSearchEnabled**

 _expression_ A variable that represents a **Store** object.


## Remarks

Use  **IsInstantSearchEnabled** to evaluate whether you should use **ci_startswith** or **ci_phrasematch** operators in your query. If you use **ci_startswith** or **ci_phrasematch** in the query and Instant Search is not enabled, Outlook will return an error.


## Example

The following code sample accepts a matching string as an input parameter, constructs a DASL filter with the content indexing keyword  **ci_phrasematch** if Instant Search is enabled on the store, and returns the filter. Otherwise, if Instant Search is not operational, then the code sample returns a filter that uses the **like** keyword.

For more information on filtering with keywords, see [Filtering Items Using Query Keywords](http://msdn.microsoft.com/library/d7e6b169-c5fd-7acc-f077-658a153a921f%28Office.15%29.aspx).




```vb
Function CreateSubjectRestriction(criteria As String) As String 
 
 Dim result As String 
 
 If Application.Session.DefaultStore.IsInstantSearchEnabled Then 
 
 result = "@SQL=" &; Chr(34) &; "urn:schemas:httpmail:subject" _ 
 
 &; Chr(34) &; " ci_phrasematch '" &; criteria &; "'" 
 
 Else 
 
 result = "@SQL=" &; Chr(34) &; "urn:schemas:httpmail:subject" _ 
 
 &; Chr(34) &; " like '%" &; criteria &; "%'" 
 
 End If 
 
 CreateSubjectRestriction = result 
 
End Function
```


## See also


#### Concepts


[Store Object](store-object-outlook.md)

