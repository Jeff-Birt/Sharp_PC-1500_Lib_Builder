Excel file to create Sharp PC-1500 Lib files and complete memory map

I was wishing to create a more complete Sharp PC-1500 memory map by combining data from many sources (credit to come.) It occurred to me it would be nice to be able to automatically generate Library files from this memory map, one for each device. For fun this was all done in Excel and even though I'm no VBA wizard it all works pretty well.

There is a worksheet at the very right side called 'Path'. Cell A2 in 'Path' contains the path you wish the generated library files to be saved to. If this cell is blank, you will be prompted with a save dialog to select the desired destination folder. The path is then saved to Cell A2 in 'Path' If you choose to save the Excel file after selecting the destination folder you will not be prompted for it again. If you want to use a different folder, then simply delete 'Path' Cell A2.
You will not be prompted to overwrite any existing lib files of the name.

I'm sure there are lots of fun and interesting ways to break this VBA script.

Use: ---->>>>> You must have the Developer tab showing in Excel 

Developer Tab->Macros->GenerateLibs: Generates library files and saves them to path. Only 'Short' comments are included in library files. 

Developer Tab->Macros->GenerateMemoryMap: Generates a complete PC-1500 memory map with all devices and both 'Short' and 'Long' comments.

Tabs: 
PC-1500_Mem_Map – This is the unified PC-1500 memory map which includes all related devices. It is created by the macro “GenerateMemoryMap”

PC-1500_Lib_Builider – This is where new data is entered. Short comment used for library files. Short and Long comments concatenated for make tab “PC-1500_Mem_Map”

CE-150, CE-151, CE-155, CE-158, CE-159, CE-161, ROM4, ROM8, ROM16 – Running macro “GenerateLibs” populates these tabs and then creates/saves library files from them.

Path - contains the path you wish the generated library files to be saved to. If this cell is blank, you will be prompted with a save dialog to select the desired destination folder. The path is then saved to Cell A2 in 'Path' If you choose to save the Excel file after selecting the destination folder you will not be prompted for it again. If you want to use a different folder, then simply delete 'Path' Cell A2.


Other files:
PC-1500_Memory_Map_Norlin Rober_Pocket_Computer_Newsletter.docx - Memmory map by Norline Rober
PC-1500_Memory_Map_ROM1500.xlsx - Memory map from ROM1500 program
Sharp lh5801 Opcode Tables.xlsx - Opcodes by number and mnumonic in both Sharp and Z80 formats
Sharp PC-1500 Development Notes.docx - Random PC-1500 develepment notes of mine
Sharp PC-1500 Error Codes.docx - List of all error codes for PC-1500 and accesories
Sharp_PC-1500_LCD_Mapping.png - Graphic explaining odd LCD mapping
