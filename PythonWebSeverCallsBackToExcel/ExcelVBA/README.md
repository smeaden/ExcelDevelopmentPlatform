I choose not to check in a whole workbook into source control because of (a) size, (b) macro security.  

I prefer to check in the VBA code modules separately so they are viewable like any other source code.  

But this means I need to give instructions as how to build a workbook containing the code.  

I have given a module called devBuild which automates importing the code modules.  So the process should be quick/terse.

- Open Excel and create a new workbook
- Go to the VBA IDE and then for the new workbook import the code module named 'devBuild.bas'
- Once imported, double-click on devBuild to open the code module in the IDE editor
- Navigate to the top of the code where you should find 'Private Sub BuildWorkbook()', click in the subroutine to place the cursor there 
- With the cursor placed in the BuildWorkbook subroutine press F5 to import the other code modules

Now you should have the other code modules imported.  
In the future, I expect the build process to take care of Tools->References but for this project there are none.

If the process fails then simply import the modules manually.  This build process is just a quick time saver over manual imports (there is no extra logic here).
