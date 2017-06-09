---
title: SpellingOptions Object (Excel)
keywords: vbaxl10.chm717072
f1_keywords:
- vbaxl10.chm717072
ms.prod: excel
api_name:
- Excel.SpellingOptions
ms.assetid: 3ba7d0b4-bebb-0cc9-cb50-066d1c19d876
ms.date: 06/08/2017
---


# SpellingOptions Object (Excel)

Represents the various spell checking options for a worksheet.


## Remarks

Use the  **[SpellingOptions](application-spellingoptions-property-excel.md)** property of the **[Application](application-object-excel.md)** object to return a **SpellingOptions** object.

Once a  **SpellingOptions** object is returned, you can use its following properties to set or return various spell checking options.


-  **[ArabicModes](spellingoptions-arabicmodes-property-excel.md)**
    
-  **[DictLang](spellingoptions-dictlang-property-excel.md)**
    
-  **[GermanPostReform](spellingoptions-germanpostreform-property-excel.md)**
    
-  **[HebrewModes](spellingoptions-hebrewmodes-property-excel.md)**
    
-  **[IgnoreCaps](spellingoptions-ignorecaps-property-excel.md)**
    
-  **[IgnoreFileNames](spellingoptions-ignorefilenames-property-excel.md)**
    
-  **[IgnoreMixedDigits](spellingoptions-ignoremixeddigits-property-excel.md)**
    
-  **[KoreanCombineAux](spellingoptions-koreancombineaux-property-excel.md)**
    
-  **[KoreanProcessCompound](spellingoptions-koreanprocesscompound-property-excel.md)**
    
-  **[KoreanUseAutoChangeList](spellingoptions-koreanuseautochangelist-property-excel.md)**
    
-  **[SuggestMainOnly](spellingoptions-suggestmainonly-property-excel.md)**
    
-  **[UserDict](spellingoptions-userdict-property-excel.md)**
    

## Example

The following example uses the  **[IgnoreCaps](spellingoptions-ignorecaps-property-excel.md)** property to disable spell checking for words that have all capitalized letters. In this example, "Testt", but not "TESTT", is identified by the spell checker.


```vb
Sub IgnoreAllCAPS() 
 
 ' Place mispelled versions of the same word in all caps and mixed case. 
 Range("A1").Formula = "Testt" 
 Range("A2").Formula = "TESTT" 
 
 With Application.SpellingOptions 
 .SuggestMainOnly = True 
 .IgnoreCaps = True 
 End With 
 
 ' Run a spell check. 
 Cells.CheckSpelling 
 
End Sub
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


