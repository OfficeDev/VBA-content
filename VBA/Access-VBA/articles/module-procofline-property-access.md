---
title: Module.ProcOfLine Property (Access)
keywords: vbaac10.chm12283
f1_keywords:
- vbaac10.chm12283
ms.prod: access
api_name:
- Access.Module.ProcOfLine
ms.assetid: 64a21820-923d-a816-6b6e-2a679d0e09ac
ms.date: 06/08/2017
---


# Module.ProcOfLine Property (Access)

The  **ProcOfLine** property returns the name of the procedure that contains a specified line in a standard module or a class module. Read-only string.


## Syntax

 _expression_. **ProcOfLine**( ** _Line_**, ** _pprockind_** )

 _expression_ A variable that represents a **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Line_|Required|**Long**|The number of a line in the module.|
| _pprockind_|Required|**vbext_ProcKind**|The type of procedure. See the Remarks section for the possible settings.|

## Remarks

The  _ProcKind_ argument can be one of the following **vbext_ProcKind** constants:



|**Constant**|**Description**|
|:-----|:-----|
|**vbext_pk_Get**|A  **Property Get** procedure.|
|**vbext_pk_Let**|A  **Property Let** procedure.|
|**vbext_pk_Proc**|A  **Sub** or **Function** procedure.|
|**vbext_pk_Set**|A  **Property Set** procedure.|
For any given line number, the  **ProcOfLine** property returns the name of the procedure that contains that line. Since comments and compilation constants immediately preceding a procedure definition are considered part of that procedure, the **ProcOfLine** property may return the name of a procedure for a line that isn't within the body of the procedure. The **[ProcStartLine](module-procstartline-property-access.md)** property indicates the line on which a procedure begins; the **[ProcBodyLine](module-procbodyline-property-access.md)** property indicates the line on which the procedure definition begins (the body of the procedure).

Note that the  _pprockind_ argument indicates whether the line belongs to a **Sub** or **Function** procedure, a **Property Get** procedure, a **Property Let** procedure, or a **Property Set** procedure. To determine what type of procedure a line is in, pass a variable of type **Long** to the **ProcOfLine** property, then check the value of that variable.


 **Note**  The  **ProcBodyLine** property treats **Sub** and **Function** procedures similarly, but distinguishes between each type of **Property** procedure.


## Example

The following function procedure lists the names of all procedures in a specified module:


```vb
Public Function AllProcs(ByVal strModuleName As String) 
 
 Dim mdl As Module 
 Dim lngCount As Long 
 Dim lngCountDecl As Long 
 Dim lngI As Long 
 Dim strProcName As String 
 Dim astrProcNames() As String 
 Dim intI As Integer 
 Dim strMsg As String 
 Dim lngR As Long 
 
 ' Open specified Module object. 
 DoCmd.OpenModule strModuleName 
 
 ' Return reference to Module object. 
 Set mdl = Modules(strModuleName) 
 
 ' Count lines in module. 
 lngCount = mdl.CountOfLines 
 
 ' Count lines in Declaration section in module. 
 lngCountDecl = mdl.CountOfDeclarationLines 
 
 ' Determine name of first procedure. 
 strProcName = mdl.ProcOfLine(lngCountDecl + 1, lngR) 
 
 ' Initialize counter variable. 
 intI = 0 
 
 ' Redimension array. 
 ReDim Preserve astrProcNames(intI) 
 
 ' Store name of first procedure in array. 
 astrProcNames(intI) = strProcName 
 
 ' Determine procedure name for each line after declarations. 
 For lngI = lngCountDecl + 1 To lngCount 
 ' Compare procedure name with ProcOfLine property value. 
 If strProcName <> mdl.ProcOfLine(lngI, lngR) Then 
 ' Increment counter. 
 intI = intI + 1 
 strProcName = mdl.ProcOfLine(lngI, lngR) 
 ReDim Preserve astrProcNames(intI) 
 ' Assign unique procedure names to array. 
 astrProcNames(intI) = strProcName 
 End If 
 Next lngI 
 
 strMsg = "Procedures in module '" &; strModuleName &; "': " &; vbCrLf &; vbCrLf 
 For intI = 0 To UBound(astrProcNames) 
 strMsg = strMsg &; astrProcNames(intI) &; vbCrLf 
 Next intI 
 
 ' Message box listing all procedures in module. 
 MsgBox strMsg 
End Function
```


## See also


#### Concepts


[Module Object](module-object-access.md)

