
# Understanding the Customization File

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_

Each section header in the customization file consists of square brackets ( **[]** ) containing a type and parameter. The four section types are indicated by the literal strings **connect**, **sql**, **userlist**, or **logs**. The parameter is the literal string, the default, a user-specified identifier, or nothing.

Therefore, each section is marked with one of the following section headers:



```
 
[ connect   default     ]
[ connect   identifier  ]
[ sql       default     ]
[ sql       identifier  ]
[ userlist  identifier  ]
[ logs                  ]
```

The section headers have the following parts.


|**Part**|**Description**|
|:-----|:-----|
|**connect**|A literal string that modifies a connection string.|
|**sql**|A literal string that modifies a command string.|
|**userlist**|A literal string that modifies the access rights of a specific user.|
|**logs**|A literal string that specifies a log file recording operational errors.|
|**default**|A literal string that is used if no identifier is specified or found.|
| _identifier_|A string that matches a string in the  **connect** or **command** string.
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Use this section if the section header contains <b>connect</b>  and the identifier string is found in the connection string.</p></li><li><p>Use this section if the section header contains <b>sql</b>  and the identifier string is found in the command string.</p></li><li><p>Use this section if the section header contains <b>userlist</b>  and the identifier string matches a <b>connect</b>  section identifier.</p></li></ul>|
The  **DataFactory** calls the handler, passing client parameters. The handler searches for whole strings in the client parameters that match identifiers in the appropriate section headers. If a match is found, the contents of that section are applied to the client parameter.
A particular section is used under the following circumstances:

- A  **connect** section is used if the value part of the client connect string keyword, " **Data Source=** _value_ ", matches a **connect** section identifier _._
    
- An  **sql** section is used if the client command string contains a string that matches an **sql** section identifier.
    
- A  **connect** or **sql** section with a default parameter is used if there is no matching identifier.
    
- A  **userlist** section is used if the **userlist** section identifier matches a **connect** section identifier. If there is a match, the contents of the **userlist** section are applied to the connection governed by the **connect** section.
    
- If the string in a connection or command string does not match the identifier in any  **connect** or **sql** section header, and there is no **connect** or **sql** section header with a default parameter, then the client string is used without modification.
    
- The  **logs** section is used whenever the **DataFactory** is in operation.
    
