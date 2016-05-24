
# Circular reference caused by <query reference>. (Error 3102)

 **Last modified:** December 30, 2015

 _ **Applies to:** Access 2013 | Access 2016_

You tried to execute a query that depends on itself for data. For example, this error occurs if you execute either of the following queries:

Query1



```
SELECT * FROM Employees, Query2;

```

Query2



```
SELECT * FROM Query1;


```

Redesign the queries to eliminate the dependency.
