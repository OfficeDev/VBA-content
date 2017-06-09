---
title: Module.ProcStartLine Property (Access)
keywords: vbaac10.chm12280
f1_keywords:
- vbaac10.chm12280
ms.prod: access
api_name:
- Access.Module.ProcStartLine
ms.assetid: ef9a1ab4-f992-5077-b52b-d16cba10f697
ms.date: 06/08/2017
---


# Module.ProcStartLine Property (Access)

The  **ProcStartLine** property returns avalue identifying the line at which a specified procedure begins in a standard module or a class module. Read-only **Long**.


## Syntax

 _expression_. **ProcStartLine**( ** _ProcName_**, ** _ProcKind_** )

 _expression_ A variable that represents a **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ProcName_|Required|**String**|The name of a procedure in the module.|
| _ProcKind_|Required|**vbext_ProcKind**|The type of procedure. See the Remarks section for the possible settings.|

## Remarks

The  _ProcKind_ argument can be one of the following **vbext_ProcKind** constants:



|**Constant**|**Description**|
|:-----|:-----|
|**vbext_pk_Get**|A  **Property Get** procedure.|
|**vbext_pk_Let**|A  **Property Let** procedure.|
|**vbext_pk_Proc**|A  **Sub** or **Function** procedure.|
|**vbext_pk_Set**|A  **Property Se** t procedure.|
A procedure begins with any comments and compilation constants that immediately precede the procedure definition, denoted by one of the following:


- A  **Sub** statement.
    
- A  **Function** statement.
    
- A  **Property Get** statement.
    
- A  **Property Let** statement.
    
- A  **Property Set** statement.
    
The  **ProcStartLine** property returns the number of the line on which the specified procedure begins. The beginning of the procedure may include comments or compilation constants that precede the procedure definition.

To determine the line on which the procedure definition begins, use the  **[ProcBodyLine](module-procbodyline-property-access.md)** property. This property returns the number of the line that begins with a **Sub**, **Function**, **Property Get**, **Property Let**, or **Property Set** statement.

The  **ProcStartLine** and **ProcBodyLine** properties can have the same value, if the procedure definition is the first line of the procedure. If the procedure definition isn't the first line of the procedure, the **ProcBodyLine** property will have a greater value than the **ProcStartLine** property.

It may be easier to determine where a procedure begins if you have the  **Procedure Separator** option selected. With this option selected, there is a line between the end of a procedure and the beginning of the next procedure. The first line of code (or blank line) below the procedure separator is the first line of the following procedure, which is the line returned by the **ProcStartLine** property. The **Procedure Separator** option is located on the **Editor** tab of the **Options** dialog box, available by clicking **Options** on the **Tools** menu.


 **Note**  The  **ProcCountLines** property treats **Sub** and **Function** procedures similarly, but distinguishes between each type of Property procedure.


## Example

The following example displays a message indicating where a particular procedure starts in a particular form module.


```vb
Dim strForm As String 
Dim strProc As String 
 
strForm = "Products" 
strProc = "Form_Activate" 
 
MsgBox "The procedure " &; strProc &; " starts on line " &; _ 
 Forms(strForm).Module.ProcStartLine(strProc, vbext_pk_Proc) &; "."
```


## See also


#### Concepts


[Module Object](module-object-access.md)

