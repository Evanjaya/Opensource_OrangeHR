# Opensource_OrangeHR
EDMI Entrance testing project.

This is entrance test project to PT EDMI Indonesia. this project is for Katalon Studio and developed using version 7.7.1/

How to Use :
1. Download this project. Can be from Katolon build in Git (choose master branch) or by donloading te source code.
2. In katalon Studio, at right corner go to profiles -> default, change value of :
   - screenshotDir : Screenshot directory.
   - xlsDir : xls data file(s) directory.
   - attachDir : Attachment files directory.
  To the directories in your local computer.
3. Yaou can edit xls datafile or begin to explore the project.

XLS Datafile :
Located in the project, folder XLS Data Files (Default file name : EDMI_Entrance_Test_Datafile.xls). Consist of 5 sheets :
1. StartEnd : for determining which record to start and which record to end the test.
2. Personal : Contains value of editable field at Personal Details section.
3. Custom : Contains value of editable field at Custom Fields section.
4. Attachment : Contains value of editable field at Attachments section. You can add more than one attachment in one record by write same no in "No Personal" column, which is no of record in Personal sheet. The script will attach files with same No Personal in the same run.
