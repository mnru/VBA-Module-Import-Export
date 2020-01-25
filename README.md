# VBA-Module-Import-Export
This is [vbac.wsf](https://github.com/vbaidiot/ariawase) like compose/decompose script for modules of VBA ,but specialized for my use.

* works for excel macro file only.
* before compose,in the source file lf is converted to crlf 
* there is a constant isFixedMode in scripts 
    * When isFixedMode is true, suppose location of files and folder definite and works silently,
    * When isFixedMode  is false ,the dialog open and ask files and folder location.
    * Scripts Fix_xxxx and xxxx are almost all same but constant isFixedMode.

## When isFixedMode  is True 

Fixed mode suposes like below folder location 
(supose parent folder and macro file names are same)

+ xxxx
    + bin
        + xxxx.xlsm
    + src
        + aaaa.bas
        + bbbb.cls
    + Fix_compose.vbs
    + Fix_decompose.vbs

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
    (if not exists, make folder bin and macro file yyyy.xlsm.
    (yyyy is parent folder name)
    if yyyy.xlsm exists,recompose it.)

 + yyyy
    + xxxx
        + aaaa.bas
        + bbbb.cls
    + bin
        + yyyy.xlsm     
