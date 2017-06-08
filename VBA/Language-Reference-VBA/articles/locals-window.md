---
title: Locals Window
ms.prod: office
ms.assetid: 32e88a9a-853c-e0ec-37ba-364706cf2958
ms.date: 06/08/2017
---


# Locals Window


![Locals window](images/local_ZA01201622.gif)



Automatically displays all of the declared variables in the current procedure and their values.

When the  **Locals** window is visible, it is automatically updated every time there is a change from Run to[Break mode](vbe-glossary.md) or you navigate in the stack display.

You can:


- Resize the column headers by dragging the border to the right or the left.
    
- Close the window by clicking the Close box. If the Close box is not visible, double-click the  **Title** bar to make the Close box visible, then click it.
    


## Window Elements

 **Calls Stack Button**

Opens the  **Call** **Stack** dialog box which lists the procedures in the call stack.

 **Expression**

Lists the name of the variables.

The first variable in the list is a special module variable and can be expanded to display all module level variables in the current module. For a class module, the system variable  `<Me>` is defined. For standard modules, the first variable is the is defined. For standard modules, the first variable is the `<name of the current module>`. Global variables and variables in other projects are not accessible from the Locals window.

You cannot edit data in this column.

 **Value**

List the value of the variable.

When you click on a value in the Value column, the cursor changes to an I-beam. You can edit a value and then press ENTER, the UP ARROW key, the DOWN ARROW key, TAB, SHIFT+TAB, or click on the screen to validate the change. If the value is illegal, the Edit field remains active and the value is highlighted. A message box describing the error also appears. Cancel a change by pressing ESC.

All numeric variables must have a value listed. String variables can have an empty Value list.

Variables that contain subvariables can be expanded and collapsed. Collapsed variables do not display a value but each subvariable does. The expand icon, 
![Expand icon](images/expand_ZA01201606.gif) and the collapse icon,
![Collapse icon](images/collapse_ZA01201589.gif) appear to the left of the variable.

 **Type**

Lists the variable type. You cannot edit data in this column.


