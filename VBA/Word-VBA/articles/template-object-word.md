---
title: Template Object (Word)
keywords: vbawd10.chm2410
f1_keywords:
- vbawd10.chm2410
ms.prod: word
api_name:
- Word.Template
ms.assetid: 47d1d92d-bba9-3f2a-9c71-22ac43159bd3
ms.date: 06/08/2017
---


# Template Object (Word)

Represents a document template. The  **Template** object is a member of the **[Templates](templates-object-word.md)** collection. The **Templates** collection includes all the available **Template** objects.


## Remarks

Use  **Templates** (Index), where Index is the template name or the index number, to return a single **Template** object. The following example saves the Memo2.dot template if it is in the **Templates** collection.


```
For Each aTemp In Templates 
 If LCase(aTemp.Name) = "memo2.dot" Then aTemp.Save 
Next aTemp
```

The index number represents the position of the template in the  **Templates** collection. The following example opens the first template in the **Templates** collection.




```
Templates(1).OpenAsDocument
```

The  **Add** method is not available for the **Templates** collection. Instead, you can add a template to the **Templates** collection by doing any of the following:


- Using the  **Open** method with the **Documents** collection to open a document based on a template or a template
    
- Using the  **Add** method with the **Documents** collection to open a new document based on a template
    
- Using the  **Add** method with the **Addins** collection to load a global template
    
- Using the  **AttachedTemplate** property with the **Document** object to attach a template to a document
    
Use the  **NormalTemplate** property to return a template object that refers to the Normal template. Use the **AttachedTemplate** property to return the template attached to the specified document.

Use the  **DefaultFilePath** property to return or set the location of user or workgroup templates (that is, the folder where you want to store these templates). The following example displays the user template folder from the **File Locations** tab in the **Options** dialog box ( **Tools** menu).




```
MsgBox Options.DefaultFilePath(wdUserTemplatesPath)
```


## Methods



|**Name**|
|:-----|
|[OpenAsDocument](template-openasdocument-method-word.md)|
|[Save](template-save-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](template-application-property-word.md)|
|[BuildingBlockEntries](template-buildingblockentries-property-word.md)|
|[BuildingBlockTypes](template-buildingblocktypes-property-word.md)|
|[BuiltInDocumentProperties](template-builtindocumentproperties-property-word.md)|
|[Creator](template-creator-property-word.md)|
|[CustomDocumentProperties](template-customdocumentproperties-property-word.md)|
|[FarEastLineBreakLanguage](template-fareastlinebreaklanguage-property-word.md)|
|[FarEastLineBreakLevel](template-fareastlinebreaklevel-property-word.md)|
|[FullName](template-fullname-property-word.md)|
|[JustificationMode](template-justificationmode-property-word.md)|
|[KerningByAlgorithm](template-kerningbyalgorithm-property-word.md)|
|[LanguageID](template-languageid-property-word.md)|
|[LanguageIDFarEast](template-languageidfareast-property-word.md)|
|[ListTemplates](template-listtemplates-property-word.md)|
|[Name](template-name-property-word.md)|
|[NoLineBreakAfter](template-nolinebreakafter-property-word.md)|
|[NoLineBreakBefore](template-nolinebreakbefore-property-word.md)|
|[NoProofing](template-noproofing-property-word.md)|
|[Parent](template-parent-property-word.md)|
|[Path](template-path-property-word.md)|
|[Saved](template-saved-property-word.md)|
|[Type](template-type-property-word.md)|
|[VBProject](template-vbproject-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
