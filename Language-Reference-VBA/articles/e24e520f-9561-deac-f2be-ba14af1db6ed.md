
# This interaction between compiled and design environment components is not supported

 **Last modified:** July 28, 2015

This error has the following causes and solutions:




- This occurs when two components are running together, where one component (such as a form or a UserControl) was previously compiled and is now running using the runtime (msvbvm60.dll), and the other component is being run in the IDE. For example, a compiled UserControl running on a form in the IDE. The problem occurs because the internal memory structure between an item running in the IDE and a compiled object is slightly different and not always compatible. In general, though, you shouldn't encounter a problem with this unless you are passing an instance of a UserControl (Me) to a host form or other component through a  **Property** or **Sub** procedure.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).
