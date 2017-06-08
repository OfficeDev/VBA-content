---
title: Module.Find Method (Access)
keywords: vbaac10.chm12286
f1_keywords:
- vbaac10.chm12286
ms.prod: access
api_name:
- Access.Module.Find
ms.assetid: 6b8fcd1a-a490-19a0-1692-fb01f213c639
ms.date: 06/08/2017
---


# Module.Find Method (Access)

Finds specified text in a standard module or class module.


## Syntax

 _expression_. **Find**( ** _Target_**, ** _StartLine_**, ** _StartColumn_**, ** _EndLine_**, ** _EndColumn_**, ** _WholeWord_**, ** _MatchCase_**, ** _PatternSearch_** )

 _expression_ A variable that represents a **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Target_|Required|**String**|The text that you want to find.|
| _StartLine_|Required|**Long**|The line on which to begin searching. If a match is found, the value of the  _StartLine_ argument is set to the line on which the beginning character of the matching text is found.|
| _StartColumn_|Required|**Long**|The column on which to begin searching. Each character in a line is in a separate column, beginning with zero on the left side of the module. If a match is found, the value of the  _StartColumn_ argument is set to the column on which the beginning character of the matching text is found.|
| _EndLine_|Required|**Long**|The line on which to stop searching. If a match is found, the value of the  _EndLine_ argument is set to the line on which the ending character of the matching text is found.|
| _EndColumn_|Required|**Long**|The column on which to stop searching. If a match is found, the value of the  _EndColumn_ argument is set to the column on which the beginning character of the matching text is found.|
| _WholeWord_|Optional|**Boolean**|**True** results in a search for whole words only. The default is **False**.|
| _MatchCase_|Optional|**Boolean**|**True** results in a search for words with case matching the _Target_ argument. The default is **False**.|
| _PatternSearch_|Optional|**Boolean**|**True** results in a search in which the _Target_ argument may contain wildcard characters such as an asterisk (*) or a question mark (?). The default is **False**.|

### Return Value

Boolean


## Remarks

The  **Find** method searches for the specified text string in a **Module** object. If the string is found, the **Find** method returns **True**.

To determine the position in the module at which the search text was found, pass empty variables to the  **Find** method for the _StartLine_,  _StartColumn_,  _EndLine_, and  _EndColumn_ arguments. If a match is found, these arguments will contain the line number and column position at which the search text begins ( _StartLine_,  _StartColumn_) and ends ( _EndLine_,  _EndColumn_).

For example, if the search text is found on line 5, begins at column 10, and ends at column 20, the values of these arguments will be:  _StartLine_ = 5, _StartColumn_ = 10, _EndLine_ = 5, _EndColumn_ = 20.


## Example

The following function finds a specified string in a module and replaces the line that contains that string with a new specified line.


```vb
Function FindAndReplace(strModuleName As String, _ 
 strSearchText As String, _ 
 strNewText As String) As Boolean 
 Dim mdl As Module 
 Dim lngSLine As Long, lngSCol As Long 
 Dim lngELine As Long, lngECol As Long 
 Dim strLine As String, strNewLine As String 
 Dim intChr As Integer, intBefore As Integer, _ 
 intAfter As Integer 
 Dim strLeft As String, strRight As String 
 
 ' Open module. 
 DoCmd.OpenModule strModuleName 
 ' Return reference to Module object. 
 Set mdl = Modules(strModuleName) 
 
 ' Search for string. 
 If mdl.Find(strSearchText, lngSLine, lngSCol, lngELine, _ 
 lngECol) Then 
 ' Store text of line containing string. 
 strLine = mdl.Lines(lngSLine, Abs(lngELine - lngSLine) + 1) 
 ' Determine length of line. 
 intChr = Len(strLine) 
 ' Determine number of characters preceding search text. 
 intBefore = lngSCol - 1 
 ' Determine number of characters following search text. 
 intAfter = intChr - CInt(lngECol - 1) 
 ' Store characters to left of search text. 
 strLeft = Left$(strLine, intBefore) 
 ' Store characters to right of search text. 
 strRight = Right$(strLine, intAfter) 
 ' Construct string with replacement text. 
 strNewLine = strLeft &; strNewText &; strRight 
 ' Replace original line. 
 mdl.ReplaceLine lngSLine, strNewLine 
 FindAndReplace = True 
 Else 
 MsgBox "Text not found." 
 FindAndReplace = False 
 End If 
 
Exit_FindAndReplace: 
 Exit Function 
 
Error_FindAndReplace: 
 
MsgBox Err &; ": " &; Err.Description 
 FindAndReplace = False 
 Resume Exit_FindAndReplace 
End Function
```


## See also


#### Concepts


[Module Object](module-object-access.md)

