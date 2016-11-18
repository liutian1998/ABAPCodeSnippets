## Use VBA Recorded Code in ABAP  
With a tiny trick is it easy possible to use recorded Microsoft Office VBA (Visual Basic for Applications) code directly in ABAP. It is necessary to use only the Microsoft Script Control, with a little preparation for Excel (Project001.zExcelViaVBScript.FunctionModule.abap). The recorded VBA code is stored in an ABAP include (Project001.zExcelTest.include.abap), which is read via a function module (Project001.zReadVBAInclAsString.abap). Last but not least the VBScript and VBA code must be concatenate and then is it read to execute - that's all.  
 
You can find the corresponding post in the SCN [here](http://scn.sap.com/community/abap/blog/2016/08/21/how-to-use-vba-recorded-code-in-abap).  
 
**Hint:** The function module zReadVBAInclAsString deletes leading and trailing spaces and adds now the leading point in each line automatically. So be careful if you use underscores to break VBA code lines.
