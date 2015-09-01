
# ChartBorder.Application Property (Word)

 **Last modified:** July 28, 2015

When used without an object qualifier, returns an  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)**object that represents the Microsoft Word application. When used with an object qualifier, returns an  **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.

## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  ** [ChartBorder](eea90670-c599-2ec8-5b7b-c946a4bcd638.md)** object.


## Example

The following example displays a message about the application that created  `myObject`.


```
Set myObject = ActiveDocument 
If myObject.Application.Value = "Microsoft Word" Then 
 MsgBox "This is a Word Application object." 
Else 
 MsgBox "This is not a Word Application object." 
End If
```


## See also


#### Concepts


 [ChartBorder Object](eea90670-c599-2ec8-5b7b-c946a4bcd638.md)
#### Other resources


 [ChartBorder Object Members](208fbc56-c413-c830-c010-00f7851b297a.md)
