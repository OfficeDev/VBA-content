---
title: LanguageSettings Object (Office)
keywords: vbaof11.chm231000
f1_keywords:
- vbaof11.chm231000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.LanguageSettings
ms.assetid: 936f7d61-87e5-e153-08d4-f8c5c8ef0710
---


# LanguageSettings Object (Office)

Returns information about the language settings in a Microsoft Office application.


## Remarks

Use Application.LanguageSettings.LanguageID( _MsoAppLanguageID_ ), where[MsoAppLanguageID](http://msdn.microsoft.com/library/78196ded-10d3-2088-f263-44a771ee78b4%28Office.15%29.aspx) is a constant used to return locale identifier (LCID) information to the specified application.


## Example

The following example returns the install language, user interface language, and Help language LCIDs in a message box.


```
MsgBox "The following locale IDs are registered " &amp; _ 
 "for this application: Install Language - " &amp; _ 
 Application.LanguageSettings.LanguageID(msoLanguageIDInstall) &amp; _ 
 " User Interface Language - " &amp; _ 
 Application.LanguageSettings.LanguageID(msoLanguageIDUI) &amp; _ 
 " Help Language - " &amp; _ 
 Application.LanguageSettings.LanguageID(msoLanguageIDHelp)
```

Use  **Application.LanguageSettings.LanguagePreferredForEditing** to determine which LCIDs are registered as preferred editing languages for the application, as in the following example.




```
If Application.LanguageSettings. _ 
 LanguagePreferredForEditing(msoLanguageIDEnglishUS) Then 
 MsgBox "U.S. English is one of the chosen editing languagess." 
End If
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/48bd707e-4dac-df46-fa5b-e8d1159aa19d%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/6c7f0a01-af17-c246-5b52-4c70d45568e7%28Office.15%29.aspx)|
|[LanguageID](http://msdn.microsoft.com/library/a1efbab6-000f-d87e-296b-b58be9ad5194%28Office.15%29.aspx)|
|[LanguagePreferredForEditing](http://msdn.microsoft.com/library/345e29df-6cb7-13cc-a8ec-22196f38fc62%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/5f10ab2b-bbab-7a91-a298-42f12e1c1b22%28Office.15%29.aspx)|

## See also


#### Other resources


[LanguageSettings Object Members](http://msdn.microsoft.com/library/068383c2-78f1-2299-2087-9eaa3409e6fe%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
