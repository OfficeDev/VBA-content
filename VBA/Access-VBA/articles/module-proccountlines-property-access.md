---
title: Module.ProcCountLines Property (Access)
keywords: vbaac10.chm12281
f1_keywords:
- vbaac10.chm12281
ms.prod: access
api_name:
- Access.Module.ProcCountLines
ms.assetid: d85cacb5-127a-68a1-3bff-cc13a8a7e9ed
ms.date: 06/08/2017
---


# Module.ProcCountLines Property (Access)

The  **ProcCountLines** property returns the number of lines in a specified procedure in a standard module or a class module. Read-only **Long**.


## Syntax

 _expression_. **ProcCountLines**( ** _ProcName_**, ** _ProcKind_** )

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
The procedure begins with any comments and compilation constants that immediately precede the procedure definition, denoted by one of the following:


- A  **Sub** statement.
    
- A  **Function** statement.
    
- A  **Property Get** statement.
    
- A  **Property Let** statement.
    
- A  **Property Set** statement.
    
The  **ProcCountLines** property returns the number of lines in a procedure, beginning with the line returned by the **[ProcStartLine](module-procstartline-property-access.md)** property and ending with the line that ends the procedure. The procedure may be ended with **End Sub**, **End Function**, or **End Property**.


 **Note**  The  **ProcCountLines** property treats **Sub** and **Function** procedures similarly, but distinguishes between each type of Property procedure.


## Example

The following example displays a message indicating the number of lines in a given procedure.


```vb
Dim strForm As String 
Dim strProc As String 
 
strForm = "Products" 
strProc = "Form_Activate" 
 
MsgBox "There are " &; Forms(strForm).Module.ProcCountLines(strProc, vbext_pk_Proc) &; _ 
 " lines in the " &; strProc &; " procedure."
```


## See also


#### Concepts


[Module Object](module-object-access.md)

