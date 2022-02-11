# Excel-VBA-Tools
Useful libraries for Excel VBA:
 - [LibExcelTables.bas](https://github.com/cristianbuse/Excel-VBA-Tools/blob/master/src/LibExcelTables.bas)  
   Useful ```ListObject``` related methods:
     - ```AddListRows```, ```DeleteListRows```: adds/deletes a variable number of ```ListRows``` to/from a ```ListObject``` and works even if the parent ```Worksheet``` is protected with the ```UserInterfaceOnly``` flag set to ```True``` without the need to unprotect. See related [SO answer](https://stackoverflow.com/a/70832694/8488913)
   - ```GetListObject```: retrieve table by name without the need to know the parent ```Worksheet```
   - ```IsListObjectFiltered```: check if table is filtered without the need for error handling
 - [LibExcelBookItems.bas](https://github.com/cristianbuse/Excel-VBA-Tools/blob/master/src/LibExcelBookItems.bas)  
   Store/retrieve```String``` items in a ```Workbook``` using encapsulated custom XML functionality. No need to write any XML.
     - ```BookItem```: parametric property Get/Let. To delete a property simply set the value to a null string e.g. BookItem(ThisWorkbook, "itemName") = vbNullString
     - ```GetBookItemNames```: retrieve a collection of all item names
	 
## Submodules
Some of the modules in this repository require some additional library code modules which are available in the [submodules folder](https://github.com/cristianbuse/Excel-VBA-Tools/tree/master/submodules) or you can get their latest version here:  
* [LibMemory](https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/LibMemory.bas)

Note that submodules are not available in the Zip download. If cloning via GitHub Desktop the submodules will be pulled automatically by default. If cloning via Git Bash then use something like:
```
$ git clone https://github.com/cristianbuse/Excel-VBA-Tools
$ git submodule init
$ git submodule update
```
or:
```
$ git clone --recurse-submodules https://github.com/cristianbuse/Excel-VBA-Tools
```	 
	 
## License
MIT License

Copyright (c) 2022 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
