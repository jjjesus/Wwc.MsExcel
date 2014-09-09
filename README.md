Wwc.MsExcel
===========

Parsing an MS-Excel spreadsheet in C#.

The MS-Excel spreadsheet is in the SourceData directory.  As part of the build process, the spreadsheet is embededed in the DLL.

The spreadsheet contains a Header Row containing field names.

I believe that using the System.Data.DataSet means that this solution is based on ADO.NET and an OLE DB Adapter or Provider for Microsoft Excel.
