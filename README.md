TabGroupJumper
==============
Add-in exposing keyboard shortcuts for jumping between left/right tabgroups in VS11

VS2015 Support
==============
I've ported the addin to VSIX format to make it compatible with VS2015. The new repo is here: https://github.com/mrdooz/TabGroupJumperVSIX.


Building/Installing
===================
Load the solution in Visual Studio, and build (in Release). This will produce output in TabGroupJumper\bin\Release.

To register the plugin with Visual Studio, you will need to copy 3 files to Visual Studio's plugin directory, located in 
$(USER_HOME)/Documents\Visual Studio 2013\Addins (create the folder if it doesn't exist).

Copy the output from the Release folder, along with TabGroupJumper.AddIn to the Addins root (all files should be put directly in the 
Addins root).

Start Visual Studio, and you should see a TabGroupJumper item under the Tools menu.
