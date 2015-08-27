
# WebNavigationBarSet.DeleteSetAndInstances Method (Publisher)

 **Last modified:** July 28, 2015

Deletes a Web navigation bar set and all instances of it in the current document.

## Syntax

 _expression_. **DeleteSetAndInstances**

 _expression_A variable that represents a  **WebNavigationBarSet** object.


## Example

The following example iterates through the  **WebNavigationBarSets** collection and deletes each set from the active document.


```
Dim objWebNavBarSet As WebNavigationBarSet 
For Each objWebNavBarSet In ActiveDocument.WebNavigationBarSets 
 objWebNavBarSet.DeleteSetAndInstances 
Next objWebNavBarSet
```

