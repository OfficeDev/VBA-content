
# Overflow (Error 6)

 **Last modified:** July 28, 2015

An overflow results when you try to make an assignment that exceeds the limitations of the target of the assignment. This error has the following causes and solutions:




- The result of an assignment, calculation, or  [data type](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) conversion is too large to be represented within the range of values allowed for that type of [variable](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md).
    
    Assign the value to a variable of a type that can hold a larger range of values.
    
- An assignment to a  [property](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) exceeds the maximum value the property can accept.
    
    Make sure your assignment fits the range for the property to which it is made.
    
- You attempt to use a number in a calculation, and that number is coerced into an integer, but the result is larger than an integer. For example:
    
  ```
      Dim x As Long 
    x = 2000 * 365   ' Error: Overflow
  ```


    To work around this situation, type the number, like this:
    


  ```
      Dim x As Long 
    x = 2000 * 365   ' Error: Overflow
  ```




  ```
      Dim x As Long 
    x = CLng(2000) * 365
  ```


    To work around this situation, type the number, like this:
    


  ```
      Dim x As Long 
    x = CLng(2000) * 365
  ```


For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).
