# ReadXLSX
Read a XLSX file without Excel installed.

This tool will accept a pipe delimited command line argument for reading and 
exporting an .XLSX file format to a pipe delimited flat file. 

Arguments : 

1. Spreadsheet to be read. Full path and filename. 

2. Rows to be read. Start row hyphen End row. Example: 2-155. 
   If the end row is omitted it will read to the last used row. 

3. Columns to be exported. Comma delimited. 

4. Export file. Full path and filename. 

Example arguments: C:\temp\temp.xlsx|12-155|A,C,D|C:\temp\output.txt

This is an example of how it would be called with C#. This assumes it is in the same directory as the app calling it.

    ProcessStartInfo startInfo = new ProcessStartInfo();
    startInfo.FileName = Path.GetDirectoryName(Application.ExecutablePath) + "\\ReadXLSX.exe";
    startInfo.Arguments = @"C:\temp\temp.xlsx|12-155|A,C,D|C:\temp\output.txt";
    Process.Start(startInfo);
