---
title: Find Method (VBA Add-In Object Model)
keywords: vbob6.chm1098972
f1_keywords:
- vbob6.chm1098972
ms.prod: office
ms.assetid: cf7a4b4e-89e7-91ea-2f9b-880384cd3339
ms.date: 06/08/2017
---


# Find Method (VBA Add-In Object Model)



Searches the active [module](vbe-glossary.md) for a specified string.
 **Syntax**
 _object_**.Find(**_target_, _startline_, _startcol_, _endline_, _endcol_ [, _wholeword_ ] [, _matchcase_ ] [, _patternsearch_ ] **) As Boolean**
The  **Find** syntax has these parts:


| <strong>Part</strong>  | <strong>Description</strong>                                                                                                                                                                                                                 |
|:-----------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>        | Required. An [object expression](vbe-glossary.md) that evaluates to an object in the Applies To list.                                                                                                                                        |
| <em>target</em>        | Required. A [String](vbe-glossary.md) containing the text or pattern you want to find.                                                                                                                                                       |
| <em>startline</em>     | Required. A [Long](vbe-glossary.md) specifying the line at which you want to start the search; will be set to the line of the match if one is found. The first line is number 1.                                                             |
| <em>startcol</em>      | Required. A  <strong>Long</strong> specifying the column at which you want to start the search; will be set to the column containing the match if one is found. The first column is 1.                                                       |
| <em>endline</em>       | Required. A  <strong>Long</strong> specifying the last line of the match if one is found. The last line may be specified as -1.                                                                                                              |
| <em>endcol</em>        | Required. A  <strong>Long</strong> specifying the last line of the match if one is found. The last column may be designated as -1.                                                                                                           |
| <em>wholeword</em>     | Optional. A [Boolean](vbe-glossary.md) value specifying whether to only match whole words. If <strong>True</strong>, only matches whole words. <strong>False</strong> is the default.                                                        |
| <em>matchcase</em>     | Optional. A  <strong>Boolean</strong> value specifying whether to match case. If <strong>True</strong>, the search is case sensitive. <strong>False</strong> is the default.                                                                 |
| <em>patternsearch</em> | Optional. A  <strong>Boolean</strong> value specifying whether or not the target string is a regular expression pattern. If <strong>True</strong>, the target string is a regular expression pattern. <strong>False</strong> is the default. |

 **Remarks**
 **Find** returns **True** if a match is found and **False** if a match isn't found.
The  _matchcase_ and _patternsearch_[arguments](vbe-glossary.md) are mutually exclusive; if both arguments are passed as **True**, an error occurs.
The content of the  **Find** dialog box isn't affected by the **Find** method.
The specified range of lines and columns is inclusive, so a search can find the pattern on the specified last line if  _endcol_ is supplied as either -1 or the length of the line.

