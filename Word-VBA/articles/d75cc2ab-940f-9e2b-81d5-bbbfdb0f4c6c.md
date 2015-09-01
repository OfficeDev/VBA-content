
# Window.Panes Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Returns a  ** [Panes](6ed6353c-9134-f47d-a108-13e84eced8ff.md)**collection that represents all the window panes for the specified window.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Panes**

 _expression_An expression that returns a  ** [Window](d92f83f9-ae44-56c0-4584-7a9359253c6d.md)**object.


## Remarks
<a name="sectionSection1"> </a>

For information about returning a single member of a collection, see  [Returning an Object from a Collection](28f76384-f495-9640-a7c8-10ada3fac727.md).


## Example
<a name="sectionSection2"> </a>

This example splits the active window in half.


```
If ActiveDocument.ActiveWindow.Panes.Count = 1 Then _ 
 ActiveDocument.ActiveWindow.Panes.Add
```

This example activates the first pane in the window for Document2.




```
Windows("Document2").Panes(1).Activate
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Window Object](d92f83f9-ae44-56c0-4584-7a9359253c6d.md)
#### Other resources


 [Window Object Members](c0dec747-3695-4f96-ea25-05b6494aad7e.md)
