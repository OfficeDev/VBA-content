
# CodeProject.Path Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


You can use the  **Path** property to determine the location where data is stored for a Microsoft Access project (.adp) or Microsoft Access database. Read-only **String**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Path**

 _expression_A variable that represents a  **CodeProject** object.


## Remarks
<a name="sectionSection1"> </a>

You can use the  **Path** property to determine the location of information stored through the ** [CurrentProject](e6baae73-1eeb-b48f-d35e-b3e921378561.md)**or  ** [CodeProject](70b71f57-df23-2cf7-23f5-147053a8ec26.md)**objects of a project or database.


## Example
<a name="sectionSection2"> </a>

The following example displays a message indicating the disk location of the current Access project or database.


```
MsgBox "The current database is located at " &amp; Application.CurrentProject.Path &amp; "." 
 

```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [CodeProject Object](70b71f57-df23-2cf7-23f5-147053a8ec26.md)
#### Other resources


 [CodeProject Object Members](cd3b6b70-8312-2f2f-0f4d-7679d8bea9f5.md)
