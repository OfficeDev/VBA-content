---
title: NameSpace.Logon Method (Outlook)
keywords: vbaol11.chm767
f1_keywords:
- vbaol11.chm767
ms.prod: outlook
api_name:
- Outlook.NameSpace.Logon
ms.assetid: 167c632b-0d52-a1e4-8dcd-57d301cde3c9
ms.date: 06/08/2017
---


# NameSpace.Logon Method (Outlook)

Logs the user on to MAPI, obtaining a MAPI session.


## Syntax

 _expression_ . **Logon**( **_Profile_** , **_Password_** , **_ShowDialog_** , **_NewSession_** )

 _expression_ A variable that represents a **NameSpace** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Profile_|Optional| **Variant**|The MAPI profile name, as a  **String** , to use for the session. Specify an empty string to use the default profile for the current session.|
| _Password_|Optional| **Variant**|The password (if any), as a  **String** , associated with the profile. This parameter exists only for backwards compatibility and for security reasons, it is not recommended for use. Microsoft Oultook will prompt the user to specify a password in most system configurations. This is your logon password and should not be confused with PST passwords.|
| _ShowDialog_|Optional| **Variant**| **True** to display the MAPI logon dialog box to allow the user to select a MAPI profile.|
| _NewSession_|Optional| **Variant**| **True** to create a new Microsoft Outlook session. Since multiple sessions cannot be created in Outlook, this parameter should be specified as True only if a session does not already exist.|

## Remarks

Use the  **Logon** method only to log on to a specific profile when Outlook is not already running. This is because only one Outlook process can run at a time, and that Outlook process uses only one profile and supports only one MAPI session. When users start Outlook a second time, that instance of Outlook runs within the same Outlook process, does not create a new process, and uses the same profile.

If Outlook is already running, using this method does not create a new Outlook session or change the current profile to a different one. 

If Outlook is not running and you only want to start Outlook with the default profile, do not use the  **Logon** method. A better alternative is shown in the following code example, `InitializeMAPI`: first, instantiate the Outlook [Application](application-object-outlook.md) object, then reference a default folder such as the Inbox. This has the side effect of initializing MAPI to use the default profile and to make the object model fully functional.




```vb
Sub InitializeMAPI ()

    ' Start Outlook.
    Dim olApp As Outlook.Application
    Set olApp = CreateObject("Outlook.Application")
    
    ' Get a session object. 
    Dim olNs As Outlook.NameSpace
    Set olNs = olApp.GetNamespace("MAPI")
    
    ' Create an instance of the Inbox folder. 
    ' If Outlook is not already running, this has the side
    ' effect of initializing MAPI.
    Dim mailFolder As Outlook.Folder
    Set mailFolder = olNs.GetDefaultFolder(olFolderInbox)

    ' Continue to use the object model to automate Outlook.
End Sub
```

Starting in Outlook 2010, if you have multiple profiles, you have configured Outlook to always use a default profile, and you use the  **Logon** method to log on to the default profile without prompting the user, the user will receive a prompt to choose a profile anyway. To avoid this behavior, do not use the **Logon** method; use the workaround suggested in the preceding `InitializeMAPI` example instead.


## Example

This Visual Basic for Applications example uses the  **Logon** method to log on to a new session, displaying the dialog box to verify the profile name and enter password.


```vb
Sub StartOutlook() 
    Dim myNameSpace As Outlook.NameSpace 
  
    Set myNameSpace = Application.GetNamespace("MAPI") 
    myNameSpace.Logon "LatestProfile", , True, True 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)
#### Other resources


[How to: Obtain and Log On to an Instance of Outlook](http://msdn.microsoft.com/library/ef369364-6500-2759-3ef4-ed4411112e96%28Office.15%29.aspx)


