
# Module.ProcCountLines Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


The  **ProcCountLines** property returns the number of lines in a specified procedure in a standard module or a class module. Read-only **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ProcCountLines**( **_ProcName_**,  **_ProcKind_**)

 _expression_A variable that represents a  **Module** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ProcName|Required| **String**|The name of a procedure in the module.|
|ProcKind|Required| **vbext_ProcKind**|The type of procedure. See the Remarks section for the possible settings.|

## Remarks
<a name="sectionSection1"> </a>

The ProcKind argument can be one of the following **vbext_ProcKind** constants:



|**Constant**|**Description**|
|:-----|:-----|
| **vbext_pk_Get**|A  **Property Get** procedure.|
| **vbext_pk_Let**|A  **Property Let** procedure.|
| **vbext_pk_Proc**|A  **Sub** or **Function** procedure.|
| **vbext_pk_Set**|A  **Property Set** procedure.|
The procedure begins with any comments and compilation constants that immediately precede the procedure definition, denoted by one of the following:


- A  **Sub** statement.
    
- A  **Function** statement.
    
- A  **Property Get** statement.
    
- A  **Property Let** statement.
    
- A  **Property Set** statement.
    
The  **ProcCountLines** property returns the number of lines in a procedure, beginning with the line returned by the ** [ProcStartLine](ef9a1ab4-f992-5077-b52b-d16cba10f697.md)** property and ending with the line that ends the procedure. The procedure may be ended with **End Sub**,  **End Function**, or  **End Property**.


 **Note**  The  **ProcCountLines** property treats **Sub** and **Function** procedures similarly, but distinguishes between each type of Property procedure.


## Example
<a name="sectionSection2"> </a>

The following example displays a message indicating the number of lines in a given procedure.


```
Dim strForm As String 
Dim strProc As String 
 
strForm = "Products" 
strProc = "Form_Activate" 
 
MsgBox "There are " &amp; Forms(strForm).Module.ProcCountLines(strProc, vbext_pk_Proc) &amp; _ 
 " lines in the " &amp; strProc &amp; " procedure."
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Module Object](e04272fa-9c29-2567-bd15-1cea38906894.md)
#### Other resources


 [Module Object Members](c2e71012-645e-b818-1247-9775f221619e.md)
