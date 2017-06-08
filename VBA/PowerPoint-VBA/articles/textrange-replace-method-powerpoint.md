---
title: TextRange.Replace Method (PowerPoint)
keywords: vbapp10.chm569035
f1_keywords:
- vbapp10.chm569035
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.Replace
ms.assetid: 046d1c3d-fd3e-7871-e31e-6529b77fcd60
ms.date: 06/08/2017
---


# TextRange.Replace Method (PowerPoint)

Finds specific text in a text range, replaces the found text with a specified string, and returns a  **TextRange** object that represents the first occurrence of the found text. Returns **Nothing** if no match is found.


## Syntax

 _expression_. **Replace**( **_FindWhat_**, **_ReplaceWhat_**, **_After_**, **_MatchCase_**, **_WholeWords_** )

 _expression_ A variable that represents a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FindWhat_|Required|**String**|The text to search for.|
| _ReplaceWhat_|Required|**String**|The text you want to replace the found text with.|
| _After_|Optional|**Integer**|The position of the character (in the specified text range) after which you want to search for the next occurrence of FindWhat. For example, if you want to search from the fifth character of the text range, specify 4 for After. If this argument is omitted, the first character of the text range is used as the starting point for the search.|
| _MatchCase_|Optional|**MsoTriState**|Determines whether a distinction is made on the basis of case.|
| _WholeWords_|Optional|**MsoTriState**|Determines whether only whole words are found.|

### Return Value

TextRange


## Remarks

The  _MatchCase_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. Does not distinguish between uppercase and lowercase characters.|
|**msoTrue**|Distinguish between uppercase and lowercase characters.|
The  _WholeWords_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. Does not find only entire words.|
|**msoTrue**|Finds only entire words.|

## Example

This example replaces every whole-word occurrence of "like" in all of the shapes in the active presentation with "NOT LIKE".


```vb
Sub ReplaceText()

    

    Dim oSld As Slide

    Dim oShp As Shape

    Dim oTxtRng As TextRange

    Dim oTmpRng As TextRange

     

    Set oSld = Application.ActivePresentation.Slides(1)

    

    For Each oShp In oSld.Shapes

        Set oTxtRng = oShp.TextFrame.TextRange

        Set oTmpRng = oTxtRng.Replace(FindWhat:="like", _
            Replacewhat:="NOT LIKE", WholeWords:=True)

        Do While Not oTmpRng Is Nothing
            Set oTxtRng = oTxtRng.Characters(oTmpRng.Start + oTmpRng.Length, _
                oTxtRng.Length)

            Set oTmpRng = oTxtRng.Replace(FindWhat:="like", _
                Replacewhat:="NOT LIKE", WholeWords:=True)
        Loop

    Next oShp



End Sub
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

