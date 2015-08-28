
# LCase Function

 **Last modified:** July 28, 2015


Returns a  [String](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) that has been converted to lowercase.
 **Syntax**
 **LCase**( _string_)
The required  _string_ [argument](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) is any valid [string expression](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md). If  _string_ contains [Null](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), Null is returned.
 **Remarks**
Only uppercase letters are converted to lowercase; all lowercase letters and nonletter characters remain unchanged.

## Example

This example uses the  **LCase** function to return a lowercase version of a string.


```
Dim UpperCase, LowerCase
Uppercase = "Hello World 1234"    ' String to convert.
Lowercase = Lcase(UpperCase)    ' Returns "hello world 1234".


```

