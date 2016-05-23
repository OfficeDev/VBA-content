
# Syntax error in WITH OWNERACCESS OPTION declaration. (Error 3257)

 **Last modified:** December 30, 2015

 _ **Applies to:** Access 2013 | Access 2016_

Possible causes:



- The WITH OWNERACCESS OPTION declaration is incomplete or includes a space between OWNER and ACCESS.
    
- The declaration appears in an unexpected and disallowed position in the SQL statement. For example:
    
```sql
SELECT * WITH OWNERACCESS OPTION FROM [My Table]; 
```


    The WITH OWNERACCESS OPTION declaration should appear at the end of the SQL statement, usually after the ORDER BY clause, if present:
    


```sql
SELECT * FROM [My Table] WITH OWNERACCESS OPTION;
```

 **ACCESS SUPPORT RESOURCES**[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)[Access help on support.office.com](https://support.office.com/search/results?query=Access)[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&amp;tab=question&amp;status=all&amp;auth=1)[Search for specific Access error codes on Bing](http://www.bing.com/)[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)
