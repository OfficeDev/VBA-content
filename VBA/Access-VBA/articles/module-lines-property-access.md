---
title: Module.Lines Property (Access)
keywords: vbaac10.chm12275
f1_keywords:
- vbaac10.chm12275
ms.prod: access
api_name:
- Access.Module.Lines
ms.assetid: a230ffef-6640-178f-b3a5-edd1e171a8f6
ms.date: 06/08/2017
---


# Module.Lines Property (Access)

The  **Lines** property returns a string containing the contents of a specified line or lines in a standard module or a class module. Read-only **String**.


## Syntax

 _expression_. **Lines**( ** _Line_**, ** _NumLines_** )

 _expression_ A variable that represents a **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Line_|Required|**Long**|The number of the first line to return.|
| _NumLines_|Required|**Long**|The number of lines to return.|

## Remarks

Lines in a module are numbered beginning with 1. For example, if you read the  **Lines** property with a value of 1 for the _line_ argument and 1 for the _numlines_ argument, the **Lines** property returns a string containing the text of the first line in the module.

To insert a line of text into a module, use the  **[InsertLines](module-insertlines-method-access.md)** method.


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

