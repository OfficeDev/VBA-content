---
title: System.PrivateProfileString Property (Word)
keywords: vbawd10.chm154468362
f1_keywords:
- vbawd10.chm154468362
ms.prod: word
api_name:
- Word.System.PrivateProfileString
ms.assetid: 737fb157-4665-5e31-240a-347bd7334005
ms.date: 06/08/2017
---


# System.PrivateProfileString Property (Word)

Returns or sets a string in a settings file or the Microsoft Windows registry. Read/write  **String** .


## Syntax

 _expression_ . **PrivateProfileString**( **_FileName_** , **_Section_** , **_Key_** )

 _expression_ An expression that returns a **[System](system-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The file name for the settings file. If there is no path specified, the Windows folder is assumed.|
| _Section_|Required| **String**|The name of the section in the settings file that contains Key. In a Windows settings file, the section name appears between brackets before the associated keys (do not include the brackets with Section). If you are returning the value of an entry from the Windows registry, Section should be the complete path to the subkey, including the subtree (for example, "HKEY_CURRENT_USER\Software\Microsoft\Office\version\Word\Options").|
| _Key_|Required| **String**|The key setting or registry entry value you want to retrieve. In a Windows settings file, the key name is followed by an equal sign (=) and the setting. If you are returning the value of an entry from the Windows registry, Key should be the name of an entry in the subkey specified by Section (for example, "STARTUP-PATH").|

## Remarks

You can write macros that use a settings file to store and retrieve settings. For example, you can store the name of the active document when you exit Microsoft Word so that it can be reopened automatically the next time you start Word. A settings file is a text file with information arranged like the information in the Windows 3.x WIN.INI file.


## Example

This example sets the current document name as the LastFile setting under the MacroSettings heading in Settings.txt.


```
System.PrivateProfileString("C:\Settings.txt", "MacroSettings", _ 
 "LastFile") = ActiveDocument.FullName
```

This example returns the LastFile setting from Settings.txt and then opens the document stored in LastFile.




```vb
LastFile = System.PrivateProfileString("C:\Settings.Txt", _ 
 "MacroSettings", "LastFile") 
If LastFile <> "" Then Documents.Open FileName:=LastFile
```

This example displays the value of the EmailName entry from the Windows registry.




```vb
aName = System.PrivateProfileString("", _ 
 "HKEY_CURRENT_USER\Software\Microsoft\" _ 
 &; "Windows\CurrentVersion\Internet Settings", "EmailName") 
MsgBox aName
```


## See also


#### Concepts


[System Object](system-object-word.md)

