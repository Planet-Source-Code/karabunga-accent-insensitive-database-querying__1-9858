<div align="center">

## Accent Insensitive database querying


</div>

### Description

Since MS-Access doesn't support accent insensitive queries by itself (MS SQL Server does as far as I know), I had to create a function that would fix the problem. With this function, it is possible to turn any SQL query into an accent insensitive query. With a few little modifications, it works great with ASP too!
 
### More Info
 
Example: STRSQL = "SELECT * FROM MyTable WHERE animal LIKE '%" & AccIns("ELEPHANT") & "%'"

This will return any record where animal = Élephant, Elephant, éléphant, eléphant, etc. You get the picture. Now have fun! :)

You need to know how SQL queries work.

An accent insensitive string to use against a database.

My friend's computer exploded last time he used this function, so watch out!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Karabunga](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/karabunga.md)
**Level**          |Advanced
**User Rating**    |5.0 (40 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/karabunga-accent-insensitive-database-querying__1-9858/archive/master.zip)





### Source Code

```
Function AccIns(Str As String) As String
  Dim CurLtr As String * 1
  For x = 1 To Len(Str)
    CurLtr = Mid(Str, x, 1)
    Select Case CurLtr
      Case "e", "é", "è", "ê", "ë", "E", "É", "È", "Ê", "Ë"
        AccIns = AccIns & "[eéèêëEÉÈÊË]"
      Case "a", "à", "â", "ä", "A", "À", "Â", "Ä"
        AccIns = AccIns & "[aàâäAÀÂÄ]"
      Case "i", "ì", "ï", "î", "I", "Ì", "Ï", "Î"
        AccIns = AccIns & "[iïîìIÏÎÌ]"
      Case "o", "ô", "ö", "ò", "O", "Ô", "Ö", "Ò"
        AccIns = AccIns & "[oôöòOÔÖÒ]"
      Case "u", "ù", "û", "ü", "U", "Ù", "Û", "Ü"
        AccIns = AccIns & "[uûüùUÛÜÙ]"
      Case "c", "ç", "C", "Ç"
        AccIns = AccIns & "[cCçÇ]"
      Case Else
        AccIns = AccIns & CurLtr
    End Select
  Next
End Function
```

