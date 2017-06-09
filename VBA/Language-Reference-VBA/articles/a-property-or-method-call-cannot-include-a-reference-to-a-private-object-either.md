---
title: A property or method call cannot include a reference to a private object, either as an argument or as a return value (Error 98)
ms.prod: office
ms.assetid: 1f4e72f6-1972-4337-a56a-adc366264954
ms.date: 06/08/2017
---


# A property or method call cannot include a reference to a private object, either as an argument or as a return value (Error 98)
Private objects should never be passed outside a project. The following, all of which are prohibited, are possible causes for the error:


- A client invoked a property or method of an out-of-process component and attempted to pass a reference to a private object as one of the arguments. A client invoked a property or method of an out-of-process component and the component attempted to return a reference to a private object, or to assign such a reference to a  **ByRef** argument.
    
- An out-of-process component has invoked a call-back method on its client and attempted to pass a reference to a private object.
    
- An out-of-process component attempted to pass a reference to a private object as an argument of an event it was raising.
    
- A client attempted to assign a private object reference to a  **ByRef** argument of an event it was handling.
    

Note that although Visual Basic prevents you from passing references to nonvisual private objects across processes, there are some cases in which Visual Basic can't detect this error and thus can't prevent it. Private objects are not designed to be used outside your project. If you pass them to a client, you may jeopardize program stability and cause incompatibility with future versions of Visual Basic. If you need to pass a private class of your own to a client, set the  **Instancing** property to a value other than **Private**.
For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

