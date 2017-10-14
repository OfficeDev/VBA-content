---
title: Application.SetDefaultTheme Method (Word)
keywords: vbawd10.chm158335390
f1_keywords:
- vbawd10.chm158335390
ms.prod: word
api_name:
- Word.Application.SetDefaultTheme
ms.assetid: 7c51ff47-92d7-724f-0334-b789d2441313
ms.date: 06/08/2017
---


# Application.SetDefaultTheme Method (Word)

Sets a default theme for Word to use with new documents, e-mail messages, or web pages.

## Syntax

_expression_. **SetDefaultTheme** (**_Name_**, **_DocumentType_**)

_expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the theme you want to assign as the default theme plus any theme formatting options you want to apply. The format of this string is "themennn" where _theme_ and _nnn_ are defined in the [Themes](#themes) table.|
| _DocumentType_|Required|**WdDocumentMedium**|The type of new document to which you are assigning a default theme.|

<br/>

#### Themes

|**String**|**Description**|
|:-----|:-----|
|theme|The name of the folder that contains the data for the requested theme. (The default location for theme data folders is C:\Program Files\Common Files\Microsoft Shared\Themes.) You must use the folder name for the theme rather than the display name that appears in the  **Theme** dialog box (**Theme** command, **Format** menu).|
|nnn|A three-digit string that indicates which theme formatting options to activate (1 to activate, 0 to deactivate). The digits correspond to the **Vivid Colors**, **Active Graphics**, and **Background Image** check boxes in the **Theme** dialog box (**Theme** command, **Format** menu). If this string is omitted, the default value for _nnn_ is "011" (Active Graphics and Background Image are activated).|

<br/>

## Remarks

Setting a default theme will not apply that theme to the blank document automatically created when you start Word. Any new documents you create after that will have the default theme.

You can also use the **ThemeName** property to return and set the default theme for new e-mail messages.


## Example

This example specifies that Word use the Blueprint theme for all new e-mail messages.

```vb
Application.SetDefaultTheme "blueprnt", wdEmailMessage
```

This example specifies that Word use the Expedition theme with Active Graphics for all new web pages.

```vb
Application.SetDefaultTheme "expeditn 010", wdWebPage
```


## See also

#### Concepts

- [Application Object](application-object-word.md)

