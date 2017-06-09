---
title: Module.ProcBodyLine Property (Access)
keywords: vbaac10.chm12282
f1_keywords:
- vbaac10.chm12282
ms.prod: access
api_name:
- Access.Module.ProcBodyLine
ms.assetid: b81affb6-a3ca-3bda-59f0-9fb809b34d2d
ms.date: 06/08/2017
---


# Module.ProcBodyLine Property (Access)

The  **ProcBodyLine** property returns the number of the line at which the body of a specified procedure begins in a standard module or a class module. Read-only **Long**.


## Syntax

 _expression_. **ProcBodyLine**( ** _ProcName_**, ** _ProcKind_** )

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
|**vbext_pk_Set**|A  **Property Set** procedure.|
The body of a procedure begins with the procedure definition, denoted by one of the following:


- A  **Sub** statement.
    
- A  **Function** statement.
    
- A  **Property Get** statement.
    
- A  **Property Let** statement.
    
- A  **Property Set** statement.
    
The  **ProcBodyLine** property returns a number that identifies the line on which the procedure definition begins. In contrast, the **[ProcStartLine](module-procstartline-property-access.md)** property returns a number that identifies the line at which a procedure is separated from the preceding procedure in a module. Any comments or compilation constants that precede the procedure definition (the body of a procedure) are considered part of the procedure, but the **ProcBodyLine** property ignores them.


 **Note**  The  **ProcBodyLine** property treats **Sub** and **Function** procedures similarly, but distinguishes between each type of Property procedure.


## Example

The following example displays a message indicating on which line the procedure definition begins.


```vb
Dim strForm As String 
Dim strProc As String 
 
strForm = "Products" 
strProc = "Products_Subform_Enter" 
 
MsgBox "The definition of the " &; strProc &; " procedure begins on line " &; _ 
 Forms(strForm).Module.ProcStartLine(strProc, vbext_pk_Proc) &; "."
```


## See also


#### Concepts


[Module Object](module-object-access.md)

