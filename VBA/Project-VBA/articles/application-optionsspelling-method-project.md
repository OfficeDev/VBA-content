---
title: Application.OptionsSpelling Method (Project)
keywords: vbapj.chm614
f1_keywords:
- vbapj.chm614
ms.prod: project-server
api_name:
- Project.Application.OptionsSpelling
ms.assetid: e0085f68-a57d-c117-cc81-ad11f363c5f4
ms.date: 06/08/2017
---


# Application.OptionsSpelling Method (Project)

Sets options for the spelling checker.


## Syntax

 _expression_. **OptionsSpelling**( ** _TaskName_**, ** _TaskNotes_**, ** _TaskText1_**, ** _TaskText2_**, ** _TaskText3_**, ** _TaskText4_**, ** _TaskText5_**, ** _TaskText6_**, ** _TaskText7_**, ** _TaskText8_**, ** _TaskText9_**, ** _TaskText10_**, ** _ResourceCode_**, ** _ResourceName_**, ** _ResourceNotes_**, ** _ResourceGroup_**, ** _ResourceText1_**, ** _ResourceText2_**, ** _ResourceText3_**, ** _ResourceText4_**, ** _ResourceText5_**, ** _AssignNotes_**, ** _IgnoreUppercase_**, ** _IgnoreNumberWords_**, ** _AlwaysSuggest_**, ** _UseCustomDictionary_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TaskName_|Optional|**Boolean**|**True** if task names are checked.|
| _TaskNotes_|Optional|**Boolean**|**True** if task notes are checked.|
| _TaskText1_|Optional|**Boolean**|**True** if the **Text1** field of tasks is checked.|
| _TaskText2_|Optional|**Boolean**|**True** if the **Text2** field of tasks is checked.|
| _TaskText3_|Optional|**Boolean**|**True** if the **Text3** field of tasks is checked.|
| _TaskText4_|Optional|**Boolean**|**True** if the **Text4** field of tasks is checked.|
| _TaskText5_|Optional|**Boolean**|**True** if the **Text5** field of tasks is checked.|
| _TaskText6_|Optional|**Boolean**|**True** if the **Text6** field of tasks is checked.|
| _TaskText7_|Optional|**Boolean**|**True** if the **Text7** field of tasks is checked.|
| _TaskText8_|Optional|**Boolean**|**True** if the **Text8** field of tasks is checked.|
| _TaskText9_|Optional|**Boolean**|**True** if the **Text9** field of tasks is checked.|
| _TaskText10_|Optional|**Boolean**|**True** if the **Text10** field of tasks is checked.|
| _ResourceCode_|Optional|**Boolean**|**True** if resource codes are checked.|
| _ResourceName_|Optional|**Boolean**|**True** if resource names are checked.|
| _ResourceNotes_|Optional|**Boolean**|**True** if resource notes are checked.|
| _ResourceGroup_|Optional|**Boolean**|**True** if resource groups are checked.|
| _ResourceText1_|Optional|**Boolean**|**True** if the **Text1** field of resources is checked.|
| _ResourceText2_|Optional|**Boolean**|**True** if the **Text2** field of resources is checked.|
| _ResourceText3_|Optional|**Boolean**|**True** if the **Text3** field of resources is checked.|
| _ResourceText4_|Optional|**Boolean**|**True** if the **Text4** field of resources is checked.|
| _ResourceText5_|Optional|**Boolean**|**True** if the **Text5** field of resources is checked.|
| _AssignNotes_|Optional|**Boolean**|**True** if assignment notes are checked.|
| _IgnoreUppercase_|Optional|**Boolean**|**True** if words consisting entirely of uppercase letters are ignored.|
| _IgnoreNumberWords_|Optional|**Boolean**|**True** if words that contain numbers are ignored.|
| _AlwaysSuggest_|Optional|**Boolean**|**True** if Project always suggests alternate spellings to misspelled words.|
| _UseCustomDictionary_|Optional|**Boolean**|**True** if the custom dictionary is used.|

### Return Value

 **Boolean**


## Remarks

If an argument is omitted, the default value is specified by the setting on the  **Proofing** tab of the **Project Options** dialog box.


 **Note**  The list of fields to check for spelling on the  **Proofing** tab includes fields up to **Text30** for task, resource, and assignment custom fields.

Using the  **OptionsSpelling** method without specifying any arguments displays the **Project Options** dialog box with the **General** tab selected.

You can also use the  **[SpellCheckField](application-spellcheckfield-method-project.md)** method to change the state of a spell check field.


