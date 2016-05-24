
# The cascading options for the new reference conflict with existing reference <name>. (Error 3707)

 **Last modified:** December 30, 2015

 _ **Applies to:** Access 2013 | Access 2016_

This error occurs if a CASCADE action is defined on a column that already has another type of CASCADE action. For example, if CASCADE DELETE is already specified, the user will be prevented from trying to add CASCADE UPDATE. To apply the desired CASCADE action, the original CONSTRAINT must be dropped. This can be done with the ALTER TABLE ALTER COLUMN syntax or with the DROP CONSTRAINT syntax.

