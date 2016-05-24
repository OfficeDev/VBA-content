
# Syntax error in WITH OWNERACCESS OPTION declaration. (Error 3257)

 **Last modified:** December 30, 2015

 _ **Applies to:** Access 2013 | Access 2016_

Possible causes:



- The WITH OWNERACCESS OPTION declaration is incomplete or includes a space between OWNER and ACCESS.
    
- The declaration appears in an unexpected and disallowed position in the SQL statement. For example:
    
  ```
  SELECT * WITH OWNERACCESS OPTION FROM [My Table]; 

  ```


    The WITH OWNERACCESS OPTION declaration should appear at the end of the SQL statement, usually after the ORDER BY clause, if present:
    


  ```
  SELECT * FROM [My Table] WITH OWNERACCESS OPTION;
  ```

