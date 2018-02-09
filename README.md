# EpSqlGenCore
Simple Excel / Json generator  for SQL databases, based on EpPlus library - https://github.com/JanKallman/EPPlus 

For Core version I using now Vahid fork https://github.com/VahidN/EPPlus.Core 


Supported databases ( tested mainly with Oracle and PgSql ) :  

 - Oracle    providerName="System.Data.OracleClient"     https://github.com/ericmend/oracleClientCore-2.0
			 providerName="Mono.Data.OracleClientCore"
 - MSSQL	 providerName="System.Data.SqlClient"
 - MySQL	 providerName="MySql.Data"
 - PgSql	 providerName="Npgsql"
 
## Usage - simple
**XLSX output** 
  
	 // compiled to portable format
     dotnet EpSqlGenCore.dll AuthorFilms.sql	
	 // compiled to native format ( e.g. exe in Windows)
	 EpSqlGenCore.exe  AuthorFilms.sql			

**JSON output** 
  
     dotnet EpSqlGenCore.dll AuthorFilms.sql -j
	 EpSqlGenCore.exe AuthorFilms.sql -j

   
**JSON output to console** 
 
     dotnet EpSqlGenCore.dll AuthorFilms.sql -jc
	 EpSqlGenCore.exe AuthorFilms.sql -j

**Help** 

      dotnet EpSqlGenCore.dll -h
	  EpSqlGenCore.exe -h
      
**Directory settings for simple usage** 

Don't forget set definitions/outputs directory in app config to your path. But you can set full path to config or output file in cmd line too

    <configuration>
      <appSettings>
        <add key="DefinitionsDir" value="c:\work\EpSqlGenCore\definitions"/>
        <add key="OutputsDir" value="c:\work\EpSqlGenCore\outputs"/>
       ...
      </appSettings>

   
## Usage -advanced
**sql definition file for simple one tab XLSX output (portable/Win exe sample)**

    dotnet EpSqlGenCore.dll MySqlQuery.sql -oMyOutputFileName -a:MyArgument1:Argument1Type:Argumet1value -a:MyArgument2:Argument2Type:Argumet2value
	EpSqlGenCore.exe MySqlQuery.sql -oMyOutputFileName -a:MyArgument1:Argument1Type:Argumet1value -a:MyArgument2:Argument2Type:Argumet2value

 **json definition file for complex XLSX output with more tabs**

    dotnet EpSqlGenCore.dll MyJsonDefinition.json -oMyOutputFileName -a:MyArgument1:Argument1Type:Argumet1value -aMyArgument2:Argument2Type:Argumet2value

**Sample-sql def to XLSX** 
 
    dotnet EpSqlGenCore.dll MySqlQuerry.sql -oMyOutputfile  -dt -a:Statuses:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Zmluva:number:505115

**Sample-json def to  XLSX**  

    dotnet EpSqlGenCore.dll Test.json -oMyOutputfile -do -a:Stavy:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Contract:integer :505115

**Sample-sql def to JSON**

    dotnet EpSqlGenCore.dll MySqlQuerry.sql -oMyOutputfile -j -a:Stavy:array:'P9','K9','O9' -a:Ids:array:31,32,3 -a:Produkt:string:UO -a:Od:date:4.2.2015 -a:Contract:integer:505115

**Supported arguments types**  
string || char || varchar2 || varchar || date || integer || decimal || number || array

