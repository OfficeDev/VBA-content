
# Module.ProcBodyLine Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


The  **ProcBodyLine** property returns the number of the line at which the body of a specified procedure begins in a standard module or a class module. Read-only **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **ProcBodyLine**( **_ProcName_**,  **_ProcKind_**)

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
The body of a procedure begins with the procedure definition, denoted by one of the following:


- A  **Sub** statement.
    
- A  **Function** statement.
    
- A  **Property Get** statement.
    
- A  **Property Let** statement.
    
- A  **Property Set** statement.
    
The  **ProcBodyLine** property returns a number that identifies the line on which the procedure definition begins. In contrast, the ** [ProcStartLine](ef9a1ab4-f992-5077-b52b-d16cba10f697.md)**property returns a number that identifies the line at which a procedure is separated from the preceding procedure in a module. Any comments or compilation constants that precede the procedure definition (the body of a procedure) are considered part of the procedure, but the  **ProcBodyLine** property ignores them.


 **Note**  The  **ProcBodyLine** property treats **Sub** and **Function** procedures similarly, but distinguishes between each type of Property procedure.


## Example
<a name="sectionSection2"> </a>

The following example displays a message indicating on which line the procedure definition begins.


```
Dim strForm As String 
Dim strProc As String 
 
strForm = "Products" 
strProc = "Products_Subform_Enter" 
 
MsgBox "The definition of the " &amp; strProc &amp; " procedure begins on line " &amp; _ 
 Forms(strForm).Module.ProcStartLine(strProc, vbext_pk_Proc) &amp; "."
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Module Object](e04272fa-9c29-2567-bd15-1cea38906894.md)
#### Other resources


 [Module Object Members](c2e71012-645e-b818-1247-9775f221619e.md)
