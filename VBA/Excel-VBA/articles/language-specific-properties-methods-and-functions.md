---
title: Language-specific Properties, Methods, and Functions
keywords: vbaxl10.chm5278881
f1_keywords:
- vbaxl10.chm5278881
ms.prod: excel
ms.assetid: abf2101c-93ee-352b-6a67-478b9eb09003
ms.date: 06/08/2017
---


# Language-specific Properties, Methods, and Functions

The Excel Visual Basic for Applications (VBA) object model has language-specific elements for use with Asian and right-to-left languages.

The following table lists methods that have language-specific arguments. Methods that have new arguments or fewer arguments than in earlier versions of Excel are noted.


|**Method**|**Objects**|**Comments**|
|:-----|:-----|:-----|
| **[Add](phonetics-add-method-excel.md)**| **Phonetics**||
| **[AddLabel](shapes-addlabel-method-excel.md)**| **Shapes**||
| **[AddTextbox](shapes-addtextbox-method-excel.md)**| **Shapes**||
| **AutoFormat**| **Range**||
| **[CheckSpelling](application-checkspelling-method-excel.md)**| **Application**,  **Chart**,  **Range**,  **Worksheet**|Added  **_SpellLang_** and removed **_IgnoreInitialAlefHamza_** and **_IgnoreFinalYaa_**|
| **[Find](range-find-method-excel.md)**| **Application**,  **Range**|Removed  **_MatchControlCharacters_**,  **_MatchDiacritics_**,  **_MatchKashida_**, and  **_MatchAlefHamza_**|
| **[GetPhonetic](application-getphonetic-method-excel.md)**| **Application**||
| **[Replace](range-replace-method-excel.md)**| **Range**|Removed  **_MatchControlCharacters_**,  **_MatchDiacritics_**,  **_MatchKashida_**, and  **_MatchAlefHamza_**|
| **[SetPhonetic](range-setphonetic-method-excel.md)**| **Range**||
| **[Sort](range-sort-method-excel.md)**| **Range**|Removed  **_IgnoreControlCharacters_**,  **_IgnoreDiacritics_**, and  **_IgnoreKashida_**|
| **[SortSpecial](range-sortspecial-method-excel.md)**| **Range**||
Properties that return or set language-specific attributes are listed in the following table.


|**Property**|**Objects**|
|:-----|:-----|
| **[AddIndent](range-addindent-property-excel.md)**| **Range**,  **Style**|
| **[AddressLocal](range-addresslocal-property-excel.md)**| **Range**|
| **[Alignment](phonetic-alignment-property-excel.md)**| **Phonetic**,  **Phonetics**,  **TextEffectFormat**,  **TickLabels**|
| **[CharacterType](phonetic-charactertype-property-excel.md)**| **Phonetic**,  **Phonetics**|
| **[ControlCharacters](application-controlcharacters-property-excel.md)**| **Application**|
| **[CursorMovement](application-cursormovement-property-excel.md)**| **Application**|
| **[DefaultSheetDirection](application-defaultsheetdirection-property-excel.md)**| **Application**|
| **[DisplayRightToLeft](worksheet-displayrighttoleft-property-excel.md)**| **Window**,  **Worksheet**|
| **[FileFormat](workbook-fileformat-property-excel.md)**| **Workbook**|
| **[HorizontalAlignment](axistitle-horizontalalignment-property-excel.md)**| **AxisTitle**,  **ChartTitle**,  **DataLabel**,  **DataLabels**,  **DisplayUnitLabel**,  **Range**,  **Style**,  **TextFrame**|
| **[IMEMode](validation-imemode-property-excel.md)**| **Validation**|
| **[International](application-international-property-excel.md)**| **Application**|
| **[Item](phonetics-item-property-excel.md)**| **Phonetics**|
| **[Length](phonetics-length-property-excel.md)**| **Phonetics**|
| **[Phonetic](range-phonetic-property-excel.md)**| **Range**|
| **[PhoneticCharacters](characters-phoneticcharacters-property-excel.md)**| **Characters**|
| **[Phonetics](range-phonetics-property-excel.md)**| **Range**|
| **[ReadingOrder](axistitle-readingorder-property-excel.md)**| **AxisTitle**,  **ChartTitle**,  **DataLabel**,  **DataLabels**,  **DisplayUnitLabel**,  **Range**,  **Style**,  **TextFrame**,  **TickLabels**|
| **[Start](phonetics-start-property-excel.md)**| **Phonetics**|
| **[VerticalAlignment](axistitle-verticalalignment-property-excel.md)**| **AxisTitle**,  **ChartTitle**,  **DataLabel**,  **DataLabels**,  **DisplayUnitLabels**,  **Range**,  **Style**,  **TextFrame**|
The following are language-specific worksheet functions:

-  **[FindB](list-of-worksheet-functions-available-to-visual-basic.md)**
    
-  **[ReplaceB](list-of-worksheet-functions-available-to-visual-basic.md)**
    
-  **[SearchB](list-of-worksheet-functions-available-to-visual-basic.md)**
    
-  **[USDollar](list-of-worksheet-functions-available-to-visual-basic.md)**
    

