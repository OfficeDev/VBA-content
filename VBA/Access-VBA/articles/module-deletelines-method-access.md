---
title: Module.DeleteLines Method (Access)
keywords: vbaac10.chm12278
f1_keywords:
- vbaac10.chm12278
ms.prod: access
api_name:
- Access.Module.DeleteLines
ms.assetid: 57f65c6c-4d9c-3abd-065b-b75d1ada06cb
ms.date: 06/08/2017
---


# Module.DeleteLines Method (Access)

The  **DeleteLines** method deletes lines from a standard module or a class module.


## Syntax

 _expression_. **DeleteLines**( ** _StartLine_**, ** _Count_** )

 _expression_ A variable that represents a **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _StartLine_|Required|**Long**| The number of the line from which to begin deleting.|
| _Count_|Required|**Long**|The number of lines to delete.|

### Return Value

Nothing


## Remarks

Lines in a module are numbered beginning with one. To determine the number of lines in a module, use the  **[CountOfLines](module-countoflines-property-access.md)** property.

To replace one line with another line, use the  **[ReplaceLine](module-replaceline-method-access.md)** method.


## Example

The following example deletes a specified line from a module.


```vb
Function DeleteWholeLine(strModuleName, strText As String) _ 
 As Boolean 
 Dim mdl As Module, lngNumLines As Long 
 Dim lngSLine As Long, lngSCol As Long 
 Dim lngELine As Long, lngECol As Long 
 Dim strTemp As String 
 
 On Error GoTo Error_DeleteWholeLine 
 DoCmd.OpenModule strModuleName 
 Set mdl = Modules(strModuleName) 
 
 If mdl.Find(strText, lngSLine, lngSCol, lngELine, lngECol) Then 
 lngNumLines = Abs(lngELine - lngSLine) + 1 
 strTemp = LTrim$(mdl.Lines(lngSLine, lngNumLines)) 
 strTemp = RTrim$(strTemp) 
 If strTemp = strText Then 
 mdl.DeleteLines lngSLine, lngNumLines 
 Else 
 MsgBox "Line contains text in addition to '" _ 
 &; strText &; "'." 
 End If 
 Else 
 MsgBox "Text '" &; strText &; "' not found." 
 End If 
 DeleteWholeLine = True 
 
Exit_DeleteWholeLine: 
 Exit Function 
 
Error_DeleteWholeLine: 
 MsgBox Err &; " :" &; Err.Description 
 DeleteWholeLine = False 
 Resume Exit_DeleteWholeLine 
End Function
```

You could call this function from a procedure such as the following, which searches the module Module1 for a constant declaration and deletes it.




```vb
Sub DeletePiConst() 
 If DeleteWholeLine("Module1", "Const conPi = 3.14") Then 
 Debug.Print "Constant declaration deleted successfully." 
 Else 
 Debug.Print "Constant declaration not deleted." 
 End If 
End Sub
```


## See also


#### Concepts


[Module Object](module-object-access.md)

