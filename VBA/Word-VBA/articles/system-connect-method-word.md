---
title: System.Connect Method (Word)
keywords: vbawd10.chm154468454
f1_keywords:
- vbawd10.chm154468454
ms.prod: word
api_name:
- Word.System.Connect
ms.assetid: c2f2bc89-89a7-8ca0-3e78-ea558068b044
ms.date: 06/08/2017
---


# System.Connect Method (Word)

Establishes a connection to a network drive.


## Syntax

 _expression_ . **Connect**( **_Path_** , **_Drive_** , **_Password_** )

 _expression_ Required. A variable that represents a **[System](system-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path for the network drive (for example, "\\Project\Info").|
| _Drive_|Optional| **Variant**|A number corresponding to the letter you want to assign to the network drive, where 0 (zero) corresponds to the first available drive letter, 1 corresponds to the second available drive letter, and so on. If this argument is omitted, the next available letter is used.|
| _Password_|Optional| **Variant**|The password, if the network drive is protected with a password.|

## Security

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security Notes for Microsoft Office Solution Developers](system-connect-method-word.md). 


## Remarks

Use the  **Dialogs** property with the **wdDialogConnect** constant to display the **Connect To Network Drive** dialog box. The following example displays the **Connect To Network Drive** dialog box, with a preset path shown.


```vb
With Dialogs(wdDialogConnect) 
 .Path = "\\Marketing\Public" 
 .Show 
End With
```


## Example

This example establishes a connection to a network drive (\\Project\Info) protected with the password contained in the String variable, and then assigns the network drive to the next available drive letter.


```
System.Connect Path:="\\Project\Info", Password:=strPassword
```

This example establishes a connection to a network drive (\\Team1\Public) and assigns the network drive to the third available drive letter.




```
System.Connect Path:="\\Team1\Public", Drive:=2
```


## See also


#### Concepts


[System Object](system-object-word.md)

