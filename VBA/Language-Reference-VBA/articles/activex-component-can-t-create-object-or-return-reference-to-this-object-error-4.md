---
title: ActiveX component can't create object or return reference to this object (Error 429)
keywords: vblr6.chm1016046
f1_keywords:
- vblr6.chm1016046
ms.prod: office
ms.assetid: b2eb3773-bc6e-4291-8c17-19f4038fe01b
ms.date: 06/08/2017
---


# ActiveX component can't create object or return reference to this object (Error 429)

Creating objects requires that the object's [class](vbe-glossary.md) be registered in the system[registry](vbe-glossary.md) and that any associated[dynamic-link libraries (DLL)](vbe-glossary.md) be available. This error has the following causes and solutions:



- The class isn't registered. For example, the system registry has no mention of the class, or the class is mentioned, but specifies either a file of the wrong type or a file that can't be found. If possible, try to start the object's application. If the registry information is out of date or wrong, the application should check the registry and correct the information. If starting the application doesn't fix the problem, rerun the application's setup program.
    
- A DLL required by the object can't be used, either because it can't be found, or it was found but was corrupted. Make sure all associated DLLs are available. For example, the Data Access Object (DAO) requires supporting DLLs that vary among platforms. You may have to rerun the setup program for such an object if that is what is causing this error.
    
- The object is available on the machine, but it is a licensed [Automation object](vbe-glossary.md), and can't verify the availability of the license necessary to instantiate it.
    
    Some objects can be instantiated only after the component finds a license key, which verifies that the object is registered for instantiation on the current machine. When a reference is made to an object through a properly installed [type library](vbe-glossary.md) or[object library](vbe-glossary.md), the correct key is supplied automatically.
    
    If the attempt to instantiate is the result of a  **CreateObject** or **GetObject** call, the object must find the key. In this case, it may search the system registry or look for a special file that it creates when it is installed, for example, one with the extension .lic. If the key can't be found, the object can't be instantiated. If an end user has improperly set up the object's application, inadvertently deleted a necessary file, or changed the system registry, the object may not be able to find its key. If the key can't be found, the object can't be instantiated. In this case, the instantiation may work on the developer's system, but not on the user's system. It may be necessary for the user to reinstall the licensed object.
    
- You are trying to use the  **GetObject** function to retrieve a reference to class created with Visual Basic. **GetObject** can't be used to obtain a reference to a class created with Visual Basic.
    
- Access to the object has explicitly been denied. For example, you may be trying to access a data object that's currently being used and is locked to prevent deadlock situations. If that's the case, you may be able to access the object at another time.
    

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

