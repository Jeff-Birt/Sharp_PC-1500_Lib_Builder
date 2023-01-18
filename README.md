# Sharp_PC-1500_Lib_Builder
Excel file to create Sharp PC-1500 Lib files and complete memory map

I was wishing to create a more complete Sharp PC-1500 memory map by combining date from Many sources (credit to come.) It occured to me it woudl be nice to be able to automatically generate LIBrary files from this memory map, one for each device. For fun this was all done in Excel and even though I'm no VBA wizard it all works pretty well. 

There is a worksheet at the very right side called 'Path'. Cell A2 in 'Path' contains the path you wish the generated libray files to be saved to. If this cell is blank you will be prompted with a save dialog to select the desired destiantion folder. The path is then saved to Cell A2 in 'Path' If you choose to save the Excel file after selecting the destination folder you will not be prompted for it again. If you want to use a different folder then simpley delete 'Path' Cell A2.

You will not be prompted to overwrite any existing lib files of the name.

I'm sure there are lots of fun and interesting ways to break this VBA script.

Use: ---->>>>> You must have the Developer tab showing in Excel
Developer Tab->Macros->GenerateLibs: Generates library files and saves them to path. Only 'Short' comments are inluded in library files.
Developer Tab->Macros->GenerateMemoryMap: Generates a complete PC-1500 memory map with all devices and both 'Short' and 'Long' comments.
