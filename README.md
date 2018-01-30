# EpSqlGenCore
Excel / Json generator  for SQL databases based on EpPlus library

Supported databases ( tested mainly with Oracle ) :  

 - Oracle 
 - MSSQL 
 - MySQL 
 - PgSql
 - SqLite 
 
## Usage - simple
**XLSX output** 
  

     dotnet EpSqlGenCore.dll AuthorFilms.sql

**JSON output** 
  

     dotnet EpSqlGenCore.dll AuthorFilms.sql -j

   
**JSON output to console** 
 

     dotnet EpSqlGenCore.dll AuthorFilms.sql -jc

**Help** 

      dotnet EpSqlGenCore.dll  -h
   
## Usage -advanced
**sql definition file for simple one tab XLSX output**

    dotnet EpSqlGenCore.dll MySqlQuery.sql -oMyOutputFileName -a:MyArgument1:Argument1Type:Argumet1value -a:MyArgument2:Argument2Type:Argumet2value

 **json definition file for complex XLSX output with more tabs**

    dotnet EpSqlGenCore.dll MyJsonDefinition.json -oMyOutputFileName -a:MyArgument1:Argument1Type:Argumet1value -aMyArgument2:Argument2Type:Argumet2value

**Sample-sql def to XLSX** 
 
    dotnet EpSqlGenCore.dll MySqlQuerry.sql -oMyOutputfile  -dt -a:Statuses:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Zmluva:number:505115

**Sample-json def to  XLSX**  

    dotnet EpSqlGenCore.dll Test.json -oMyOutputfile -do -a:Stavy:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Contract:integer :505115

**Sample-sql def to JSON**  XLSX 

    dotnet EpSqlGenCore.dll MySqlQuerry.sql -oMyOutputfile -j -a:Stavy:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Contract:integer:505115

**Supported arguments types**  
string || char || varchar2 || varchar || date || integer || decimal || number || array

