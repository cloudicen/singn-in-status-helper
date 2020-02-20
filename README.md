# singn-in-status-helper

## A Qt project that can process Excel workbook：

### 1、Native function：
You have a excel workbook that contains students' information: [real name] [id] [others....] and [net name]. 
Also you have a txt file that contains the name list of the students who have signed in, but the name list only have studets's net name.

This tiny tool can help you to process the excel file, sort the record by [id], and mark the record with [signed in] and [Not signed].
The result can be saved to source file or save as other .xslx file.

### 2、About code:
This project is for practicing Qt programming skills. This project covers these points:
- signal and slot
- close event
- auto layout
- QAxObject -- handel COM object
- excel's VBA objects and functions
- VBA with Qt
- Qt multi thread -- future,futureWatcher,QtConcurrent  

this project can be a reference to these point... maybe.

### 3、About next
native function may not update in the future.  

the QAxObject part may be packed to a library to handel excel file. Its in a very primary stage right now.

So, this project is basically an example， as a note. 
