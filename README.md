# VBA-Module-Import-Export
This is [vbac.wsf](https://github.com/vbaidiot/ariawase) like compose/decompose script for modules of VBA ,but specialized for my use.


* works for excel macro file only.
* before compose,lfs are converted to crlfs in the source file.
* there is a constant isFixedMode in scripts 
    * When isFixedMode is true, suppose location of files and folders are definite and works silently,
    * When isFixedMode  is false ,the dialog open and ask the location of target macrofile or source folder.
* Scripts Fix_xxxx and xxxx are almost all same but constant isFixedMode.
* Scripts Fix_xxxx_.vbs are for xlam scripts.
* In case targetFiles is not determined implicitly ,Its extension should be written as targetExt in the script .

## When isFixedMode  is True 

Fixed mode suposes like below folder location 
(supose parent folder and macro file names are same)

+ xxxx
    + Fix_compose.vbs
    + Fix_decompose.vbs
    + xxxx.xlsm
    + src
        + aaaa.bas
        + bbbb.cls
    
## When isFixedMode is False

###  decompose

    when select macro file xxxx.xlsm ,works below
    (make same name folder xxxx and subfolder src, and decompose)

+ yyyy
    + xxxx.xlsm
    + xxxx
        + src
            + aaaa.bas
            + bbbb.cls

### compose

    when select source folder xxxx ,works below
    (if not exists, make macro file yyyy.xlsm.
    (yyyy is parent folder name)
    if yyyy.xlsm exists,recompose it.)

 + yyyy
    + xxxx
        + aaaa.bas
        + bbbb.cls
    
    + yyyy.xlsm     
