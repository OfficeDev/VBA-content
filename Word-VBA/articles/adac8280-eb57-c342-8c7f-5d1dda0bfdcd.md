
# SeriesCollection.Application Property (Word)

 **Last modified:** July 28, 2015

When used without an object qualifier, returns an  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)**object that represents the Microsoft Word application. When used with an object qualifier, returns an  **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.

## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  ** [SeriesCollection](785d61ff-96c9-b9b0-ed98-e992d9adeda6.md)** object.


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


 [SeriesCollection Object](785d61ff-96c9-b9b0-ed98-e992d9adeda6.md)
#### Other resources


 [SeriesCollection Object Members](310e4bfe-0132-ad36-7a72-f37afaba7983.md)
