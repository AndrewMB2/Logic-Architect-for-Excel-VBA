Logic Architect for Excel/VBA is a toolset for manipulating tabular data efficiently in Excel VBA which greatly simplifies project development. It can be used with:

- Data on worksheets, whether organised as a table or not.
- Databases.
- Virtual tables, which store data within VBA.

The main features are:

- There are extensive facilities for manipulating data on worksheets
- Databases can be created, maintained (including adding new tables) and queried with SQL with simple programming.
- There is a wide range of data transformations.
- There is a full object model for use in VBA where parameters are easy to change.
- Methods can be called individually so can be placed where required in VBA code.
- In addition to high level procedures (such as left join, unpivot, etc.) access is also provided to individual records and fields.
- Calculation is done with Excel formulae. Regular Expressions are supported for data extraction.
- User defined functions in VBA are supported.
- It provides an easy way to use SQL on data in the same workbook, to maintain workbook data in an Access database or save it in other formats.
- It is straightforward to create reports and ensure that formulae are protected and formatting is applied.
- It runs on all versions of Excel from 2007 onwards.

Get and Transform in Excel is frequently used to transform data and provides a user interface. If Get and Transform needs to form part of a VBA application, queries can be run from VBA, but if they need to be altered this requires editing the M code created by the Get and Transform user interface. Also it does not use Excel formulae for calculation or allow VBA access to individual records. Logic Architect provides most of the transformations which are required in practice in a way which is easier to use in VBA, more flexible and frequently faster.

Logic Architect is written as a set of class modules:

- TableData manipulates data on a worksheet.
- RsetData manipulates data in a database using ADODB.
- ArrData manipulates data in virtual tables.
