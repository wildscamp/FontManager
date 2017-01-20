# Commmand-line Font Manager for Windows

This utility makes it very easy and fast to add and remove fonts from Windows from the
command line. Traditionally this has been cumbersome and at times required user interaction
for it to work.

At this point, this utility is very basic and still considered beta as it has not been
battle tested. However, for the normal use case it works very well.

Usage
----

`FontMtr.exe [<font files>* [-i|-install <font file>] [-u|-uninstall <font file>]]`

_Examples_

* `FontMgr.exe -i fontfile.ttf`
* `FontMgr.exe -u fontb.ttf`
* `FontMgr.exe fontfile.ttf fontb.ttf hello.ttf -q` - Installs all three fonts
  and supresses the output.
* `FontMgr.exe -?` - Prints help
* `FontMgr.exe -i C:\Users\user\Desktop\fontname.ttf otherfont.ttf` - Installs
  both fonts onto the system.

Additional Notes
----

* If a font file of the same name already exists, FontMgr.exe will just skip it. At present
  There is no functionality to force it to overwrite or install.