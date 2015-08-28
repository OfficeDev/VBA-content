
# Form.AfterInsert Property (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the ** [AfterInsert](07140c13-ce7c-91f2-7451-d7f834653ef2.md)**event occurs. Read/write.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AfterInsert**

 _expression_A variable that represents a  **Form** object.


## Remarks
<a name="sectionSection1"> </a>

Valid values for this property are "macroname" where macroname is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the **BeforeInsert** event for the specified object, or "=functionname()" wherefunctionname is the name of a user-defined function.


## Example
<a name="sectionSection2"> </a>

The following example specifies that when the  **AfterInsert** event occurs on the first form of the current project, the associated event procedure should run.


```
Forms(0).After Insert = "[Event Procedure]" 

```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Form Object](72ef9219-142b-b690-b696-3eba9a5d4522.md)
#### Other resources


 [Form Object Members](e1976b58-28ca-8f76-cdf3-6732cb06ce6c.md)
