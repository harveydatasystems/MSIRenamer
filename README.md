# MSIRenamer
used to rename an MSI to include the MSI version number

The idea and bulk of the code came from:
https://stackoverflow.com/questions/3359974/how-to-include-version-number-in-vs-setup-project-output-filename/9891044

with additional credit given to. 
https://alteridem.net/2008/05/20/read-properties-from-an-msi-file/

though, i have made additional changes to make it a bit more robust. Including preserving the original msi name and file. 

to use the code, I copied the exe file to the "source" directory that i keep all of my repositories in. 
%userprofile%\source\MSIVerRenamer.exe

then inside visual studio, bring up the properties of the installer, locate the "PostBuildEvent" and execute the program. 
%userprofile%\source\MSIVerRenamer.exe "$(BuiltOuputPath)"

