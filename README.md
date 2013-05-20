ADO.vbs
====
a lib for QTP/UFT to connect MSSQL in ADODB.    
You can also use it to connect other DB system by ODBC

## execute SQL
call function **SQL\_OutputFullTable** to get a full result from a *Dbstr* by *strSQL*

    res = SQL\_OutputFullTable(Dbstr,strSQL)    

call function **SQL\_Output1Line** to get the first line from a _Dbstr_ by *strSQL*

    res = SQL\_Output1Line(Dbstr,strSQL)    

call function **SQL\_Output1Field** to get the first Field from a _Dbstr_ by *strSQL*

    res = SQL\_Output1Field(Dbstr,strSQL)    

Functions above all returns a Dictionary Object to store the result of *strSQL* in some simple rules(read the code, you'll know the rule), so that the variable *res* is a Dictionary Object.

## Get values from RES
There are functions to get the values from *res* easier:
**SQL\_GetValue1Line**

arguments:

> *res*,the result of SQL\_GetValue1Line
> *Field*,the Field Name of the table created by *strSQL*
**SQL\_GetValue1Field**

arguments:

> *res*,the result of SQL\_GetValue1Line
> *rowNum*,the row Number of table created by *strSQL*
**SQL\_GetValue1FullTable**

arguments:

> *res*,the result of SQL\_GetValue1Line
> *rowNum*,the row Number of table created by *strSQL*
> *Field*,the Field Name of the table created by *strSQL*
