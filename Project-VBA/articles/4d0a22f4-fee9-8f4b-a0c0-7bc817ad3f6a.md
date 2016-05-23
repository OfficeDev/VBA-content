
# CodeMask Object (Project)

The  **CodeMask** object is a collection of **[CodeMaskLevel](cef1b15f-c7f1-3b95-49a1-00854a74d9da.md)** objects that define the code mask for an outline code in Project.


## Example

The following example adds three levels to a code mask.


```vb
Sub DefineLocationCodeMask(objCodeMask As CodeMask) 
 
    objCodeMask.Add _ 
        Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 
        Length:=2, Separator:="." 
 
    objCodeMask.Add _ 
        Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 
        Separator:="." 
 
    objCodeMask.Add _ 
        Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 
        Length:=3, Separator:="." 
End Sub
```

