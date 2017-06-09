---
title: DeleteSetting Statement
keywords: vblr6.chm1020901
f1_keywords:
- vblr6.chm1020901
ms.prod: office
ms.assetid: e80dec3d-f3e3-a94f-69ae-930e62898ad6
ms.date: 06/08/2017
---


# DeleteSetting Statement

Deletes a section or key setting from an application's entry in the Windows [registry](vbe-glossary.md) or (on the Macintosh) information in the application's initialization file.

 **Syntax**

 **DeleteSetting  _appname_,** **_section_** [ **,** **_key_** ]

The  **DeleteSetting** statement syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_appname_**|Required. [String expression](vbe-glossary.md) containing the name of the application or[project](vbe-glossary.md) to which the section or key setting applies. On the Macintosh, this is the filename of the initialization file in the Preferences folder in the System folder.|
|**_section_**|Required. String expression containing the name of the section where the key setting is being deleted. If only  **_appname_** and **_section_** are provided, the specified section is deleted along with all related key settings.|
|**_key_**|Optional. String expression containing the name of the key setting being deleted.|
 **Remarks**
If all [arguments](vbe-glossary.md) are provided, the specified setting is deleted. A run-time error occurs if you attempt to use the **DeleteSetting** statement on a non-existent section or key setting.

## Example

The following example first uses the  **SaveSetting** statement to make entries in the Windows registry (or .ini file on 16-bit Windows platforms) for the application, and then uses the **DeleteSetting** statement to remove them. Because no **_key_** argument is specified, the whole section is deleted, including the section name and all its keys.


```vb
' Place some settings in the registry. 
SaveSetting appname := "MyApp", section := "Startup", _ 
 key := "Top", setting := 75 
SaveSetting "MyApp","Startup", "Left", 50 
' Remove section and all its settings from registry. 
DeleteSetting "MyApp", "Startup"
```


