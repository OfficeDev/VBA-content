---
title: SaveSetting Statement
keywords: vblr6.chm1020904
f1_keywords:
- vblr6.chm1020904
ms.prod: office
ms.assetid: f15549da-3c84-0991-592a-9d715fd488f3
ms.date: 06/08/2017
---


# SaveSetting Statement

Saves or creates an application entry in the application's entry in the Windows [registry](vbe-glossary.md) or (on the Macintosh) information in the application's initialization file.

 **Syntax**

 **SaveSetting  _appname_,** **_section,_**_key, setting_

The  **SaveSetting** statement syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_appname_**|Required. [String expression](vbe-glossary.md) containing the name of the application or[project](vbe-glossary.md) to which the setting applies. On the Macintosh, this is the filename of the initialization file in the Preferences folder in the System folder.|
|**_section_**|Required. String expression containing the name of the section where the key setting is being saved.|
|**_key_**|Required. String expression containing the name of the key setting being saved.|
|**_setting_**|Required. [Expression](vbe-glossary.md) containing the value that **_key_** is being set to.|
 **Remarks**
An error occurs if the key setting can't be saved for any reason.

## Example

The following example first uses the  **SaveSetting** statement to make entries in the Windows registry (or .ini file on 16-bit Windows platforms) for the application, and then uses the **DeleteSetting** statement to remove them.


```vb
' Place some settings in the registry. 
SaveSetting appname := "MyApp", section := "Startup", _ 
 key := "Top", setting := 75 
SaveSetting "MyApp","Startup", "Left", 50 
' Remove section and all its settings from registry. 
DeleteSetting "MyApp", "Startup" 

```


