---
title: Application.FollowHyperlink Method (Access)
keywords: vbaac10.chm12561
f1_keywords:
- vbaac10.chm12561
ms.prod: access
api_name:
- Access.Application.FollowHyperlink
ms.assetid: b5142ca6-8d67-c42b-81a4-5417265a50b0
ms.date: 06/08/2017
---


# Application.FollowHyperlink Method (Access)

The  **FollowHyperlink** method opens the document or Web page specified by a hyperlink address.


## Syntax

 _expression_. **FollowHyperlink**( ** _Address_**, ** _SubAddress_**, ** _NewWindow_**, ** _AddHistory_**, ** _ExtraInfo_**, ** _Method_**, ** _HeaderInfo_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Address_|Required|**String**|A string expression that evaluates to a valid hyperlink address.|
| _SubAddress_|Optional|**String**| A string expression that evaluates to a named location in the document specified by the _address_ argument. The default is a zero-length string (" ").|
| _NewWindow_|Optional|**Boolean**|A  **Boolean** value where **True** (?1) opens the document in a new window and **False** (0) opens the document in the current window. The default is **False**.|
| _AddHistory_|Optional|**Boolean**|A  **Boolean** value where **True** adds the hyperlink to the History folder and **False** doesn't add the hyperlink to the History folder. The default is **True**.|
| _ExtraInfo_|Optional|**Variant**|A string or an array of  **Byte** data that specifies additional information for navigating to a hyperlink. For example, this argument may be used to specify a search parameter for an .asp or .idc file. In your Web browser, the _extrainfo_ argument may appear after the hyperlink address, separated from the address by a question mark (?). You don't need to include the question mark when you specify the _extrainfo_ argument.|
| _Method_|Optional|**MsoExtraInfoMethod**|A  **[MsoExtraInfoMethod](http://msdn.microsoft.com/library/eb8edb9c-2a9a-62b5-f592-e40a2325a555%28Office.15%29.aspx)** constant that specifies how the _extrainfo_ argument is attached.|
| _HeaderInfo_|Optional|**String**|A string that specifies header information. By default the  _headerinfo_ argument is a zero-length string.|

## Remarks

By using the  **FollowHyperlink** method, you can follow a hyperlink that doesn't exist in a control. This hyperlink may be supplied by you or by the user. For example, you can prompt a user to enter a hyperlink address in a dialog box, then use the **FollowHyperlink** method to follow that hyperlink.

You can use the  _extrainfo_ and _method_ arguments to supply additional information when navigating to a hyperlink. For example, you can supply parameters to a search engine.

You can use the  **[Follow](hyperlink-follow-method-access.md)** method to follow a hyperlink associated with a control.


## Example

The following function prompts a user for a hyperlink address and then follows the hyperlink:


```vb
Function GetUserAddress() As Boolean 
    Dim strInput As String 
 
    On Error GoTo Error_GetUserAddress 
    strInput = InputBox("Enter a valid address") 
    Application.FollowHyperlink strInput, , True 
    GetUserAddress = True 
 
Exit_GetUserAddress: 
    Exit Function 
 
Error_GetUserAddress: 
    MsgBox Err &; ": " &; Err.Description 
    GetUserAddress = False 
    Resume Exit_GetUserAddress 
End Function
```

You could call this function with a procedure such as the following:




```vb
Sub CallGetUserAddress() 
    If GetUserAddress = True Then 
        MsgBox "Successfully followed hyperlink." 
    Else 
        MsgBox "Could not follow hyperlink." 
    End If 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-access.md)

