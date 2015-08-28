
# Module not found

 **Last modified:** July 28, 2015

 [Modules](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) aren't loaded from a code reference â€” they must be part of the [project](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md). This error has the following cause and solution:




- The requested module doesn't exist in the specified project. For example, the statement  `MyModule.SomeVar = 5` generates this error when `MyModule` isn't visible in the project `MyProject`. See your  [host application](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) documentation for information on including the module in the project.
    

