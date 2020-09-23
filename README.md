<div align="center">

## SQL\_Fix


</div>

### Description

Corrects for reserved SQL characters in SQL queries. This will correct your SQL statement if an apostrophe or 'pipe' character is in the SQL query. It is a better fix than the SQL functions which replace ' with '' because those will actually still fail in certain situations (such as in FindFirst commands).
 
### More Info
 
Bad SQL statement as string

Proper SQL Statement as string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeremy H](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeremy-h.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeremy-h-sql-fix__1-5695/archive/master.zip)





### Source Code

```
Public Function SQL_Fix(ByVal sSQL as string) as string
  Dim sTempSQL as string
  'replace apostrophes
  sTempSQL = Replace(sSQL, "'", "' & Chr(39) & '")
  'replace pipe symbols
  SQL_Fix = Replace(sTempSQL, "|", "' & Chr(124) & '")
End Function
```

