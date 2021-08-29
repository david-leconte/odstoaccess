# ODSToAccess

This class allows the user to select an ODS file and to send the info contained into it into a Microsoft Access file. The header (first line) of the ODS file must contain the same fields as the selected table.

Arguments : 
1. the ODS file 
2. the MS Access file
3. the table name

**Example**

```sh
    java -jar odstoaccess c:/pathto/thesheet.ods c:/pathto/thedatabase.accdb tablename