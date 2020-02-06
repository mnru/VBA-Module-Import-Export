# VBA-Module-Import-Export
This is [vbac.wsf](https://github.com/vbaidiot/ariawase) like compose/decompose script for modules of VBA ,but specialized for my use.

Install script for xlam is modified version of [VBAFormatter's](https://github.com/fuku2014/VBAFormatter) one.

* works for excel macro file only.
* before compose,lfs are converted to crlfs in the source file.
* there is a constant isFixedMode in scripts 
    * When isFixedMode is true, suppose location of files and folders are definite and works silently,
    * When isFixedMode  is false ,the dialog open and ask the location of target macrofile or source folder.
* Scripts compose and decompose are pretty different ,but only prefix different scripts are almost same but paremeters defined in the head of scripts.
* In case targetFiles is not determined implicitly ,Its extension should be written as targetExt in the script .
* Install.vbs and UnInstall.vbs can be used for xlam file.If folder composition is same as the fixed mode,they can be used as is. 

## When isFixedMode  is True 

Fixed mode suposes like below folder location 
(supose parent folder and macro file names are same)

+ xxxx
    + Fix_compose.vbs
    + Fix_decompose.vbs
    + (Install.vbs)
    + (UnInstall.vbs)
    + xxxx.xlsm
    + (xxxx.xlam)
    + src
        + aaaa.bas
        + bbbb.cls
    
## When isFixedMode is False(Dialog mode)

###  decompose

    when you select excel file xxxx.xl* ,this script works below
    (make same name folder xxxx and subfolder src, and decompose)

+ yyyy
    + xxxx.xl*(various type of excel file)
    + xxxx
        + src
            + aaaa.bas
            + bbbb.cls

### compose

    when you select source folder xxxx ,this script works below

    (if not exists, make macro file yyyy.xlsm.
    (yyyy is parent folder name)
    if yyyy.xlsm exists,recompose it.
    extension is determined by parameter in the script.)

 + yyyy
    + xxxx
        + aaaa.bas
        + bbbb.cls
    
    + yyyy.xlsm     
