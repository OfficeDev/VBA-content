---
title: Global Object (Word)
keywords: vbawd10.chm2489
f1_keywords:
- vbawd10.chm2489
ms.prod: word
api_name:
- Word.Global
ms.assetid: b91e7459-08d5-ea8c-42e0-f7b9bfd1a72c
ms.date: 06/08/2017
---


# Global Object (Word)

Contains top-level properties and methods that do not need to be preceded by the  **Application** property.


## Remarks

The following two statements have the same result. One statement uses the  **Application** property to access the **Documents** collection, and one does not. Both statements are equal and achieve the same result.


```
Documents(1).Content.Bold = True 
Application.Documents(1).Content.Bold = True
```


## Methods



|**Name**|
|:-----|
|[BuildKeyCode](global-buildkeycode-method-word.md)|
|[CentimetersToPoints](global-centimeterstopoints-method-word.md)|
|[ChangeFileOpenDirectory](global-changefileopendirectory-method-word.md)|
|[CheckSpelling](global-checkspelling-method-word.md)|
|[CleanString](global-cleanstring-method-word.md)|
|[DDEExecute](global-ddeexecute-method-word.md)|
|[DDEInitiate](global-ddeinitiate-method-word.md)|
|[DDEPoke](global-ddepoke-method-word.md)|
|[DDERequest](global-dderequest-method-word.md)|
|[DDETerminate](global-ddeterminate-method-word.md)|
|[DDETerminateAll](global-ddeterminateall-method-word.md)|
|[GetSpellingSuggestions](global-getspellingsuggestions-method-word.md)|
|[Help](global-help-method-word.md)|
|[InchesToPoints](global-inchestopoints-method-word.md)|
|[KeyString](global-keystring-method-word.md)|
|[LinesToPoints](global-linestopoints-method-word.md)|
|[MillimetersToPoints](global-millimeterstopoints-method-word.md)|
|[NewWindow](global-newwindow-method-word.md)|
|[PicasToPoints](global-picastopoints-method-word.md)|
|[PixelsToPoints](global-pixelstopoints-method-word.md)|
|[PointsToCentimeters](global-pointstocentimeters-method-word.md)|
|[PointsToInches](global-pointstoinches-method-word.md)|
|[PointsToLines](global-pointstolines-method-word.md)|
|[PointsToMillimeters](global-pointstomillimeters-method-word.md)|
|[PointsToPicas](global-pointstopicas-method-word.md)|
|[PointsToPixels](global-pointstopixels-method-word.md)|
|[Repeat](global-repeat-method-word.md)|

## Properties



|**Name**|
|:-----|
|[ActiveDocument](global-activedocument-property-word.md)|
|[ActivePrinter](global-activeprinter-property-word.md)|
|[ActiveProtectedViewWindow](global-activeprotectedviewwindow-property-word.md)|
|[ActiveWindow](global-activewindow-property-word.md)|
|[AddIns](global-addins-property-word.md)|
|[Application](global-application-property-word.md)|
|[AutoCaptions](global-autocaptions-property-word.md)|
|[AutoCorrect](global-autocorrect-property-word.md)|
|[AutoCorrectEmail](global-autocorrectemail-property-word.md)|
|[CaptionLabels](global-captionlabels-property-word.md)|
|[CommandBars](global-commandbars-property-word.md)|
|[Creator](global-creator-property-word.md)|
|[CustomDictionaries](global-customdictionaries-property-word.md)|
|[CustomizationContext](global-customizationcontext-property-word.md)|
|[Dialogs](global-dialogs-property-word.md)|
|[Documents](global-documents-property-word.md)|
|[FileConverters](global-fileconverters-property-word.md)|
|[FindKey](global-findkey-property-word.md)|
|[FontNames](global-fontnames-property-word.md)|
|[HangulHanjaDictionaries](global-hangulhanjadictionaries-property-word.md)|
|[IsObjectValid](global-isobjectvalid-property-word.md)|
|[IsSandboxed](global-issandboxed-property-word.md)|
|[KeyBindings](global-keybindings-property-word.md)|
|[KeysBoundTo](global-keysboundto-property-word.md)|
|[LandscapeFontNames](global-landscapefontnames-property-word.md)|
|[Languages](global-languages-property-word.md)|
|[LanguageSettings](global-languagesettings-property-word.md)|
|[ListGalleries](global-listgalleries-property-word.md)|
|[MacroContainer](global-macrocontainer-property-word.md)|
|[Name](global-name-property-word.md)|
|[NormalTemplate](global-normaltemplate-property-word.md)|
|[Options](global-options-property-word.md)|
|[Parent](global-parent-property-word.md)|
|[PortraitFontNames](global-portraitfontnames-property-word.md)|
|[PrintPreview](global-printpreview-property-word.md)|
|[ProtectedViewWindows](global-protectedviewwindows-property-word.md)|
|[RecentFiles](global-recentfiles-property-word.md)|
|[Selection](global-selection-property-word.md)|
|[ShowVisualBasicEditor](global-showvisualbasiceditor-property-word.md)|
|[StatusBar](global-statusbar-property-word.md)|
|[SynonymInfo](global-synonyminfo-property-word.md)|
|[System](global-system-property-word.md)|
|[Tasks](global-tasks-property-word.md)|
|[Templates](global-templates-property-word.md)|
|[VBE](global-vbe-property-word.md)|
|[Windows](global-windows-property-word.md)|
|[WordBasic](global-wordbasic-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
