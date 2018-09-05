# ContentsVBA
VBA to create table of contens and hyperlink back to the contents page

A useful trick to automatically create table of contens in Excel learnt from Chris Newman's blog.
https://www.thespreadsheetguru.com/blog/2015/3/28/automate-building-table-of-contents-spreadsheet-with-vba#disqus_thread
Save it for future.

Click the File tab.
Under Help, click Options.
Click Customize Ribbon.
Under Customize the Ribbon, select the Developer check box.

On the Developer tab, in the Code group, click Visual Basic.
In the Visual Basic Editor, on the Insert menu, click Module.
In the code window of the module, copy the macro code: CreateTOC
Insert as many module as you want...

In the Visual Basic Editor, on the File menu, click Close and Return to Microsoft Excel.
On the File menu, click Save As, and then save the file as an Excel Macro-Enabled Workbook (.xlsm).

On the Developer tab, in the Code group, click Macro Security.
Under Macro Settings, click Enable all macros (not recommended, potentially dangerous code can run), and then click OK.

On the Developer tab, in the Code group, click Macros.
In the Macro window, select the Create_TOC macro, and click Run.
