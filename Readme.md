#Yealink USB Phone helper for VB

## What
YLHelper is a C++ DLL project to make a wrapper for YLTELBOX.DLL. 
I built it because it is not possible to use the YLTELBOX.DLL directly in VB.
The reason is used in YLTELBOX.DLL _CDecl_ calling convention.

## Some details
YLHelper.dll exports a number of functions:

	  YL_DeviceIoControl_
	  YL_Open
	  YL_Close
	  YL_BeReady
	  YL_BeNotReady
	  YL_Calling
	  YL_GetRinging
	  YL_Talking
	  YL_SwitchToUsbChannels
	  YL_SwitchToPstnChannels
	  YL_DefaultInUsbChannels
	  YL_DefaultInPstnChannels
	  YL_UseUsbChannelsOnly

## How
To call the functions in VB you can use *Yealink.bas* module in your project or as example.
To handle callbacks as you wish you can change the `YL_CallBack` function in CallBacks module.

## Limitations
Unfortunately, some variables mentioned in the docs are not defined in the 
the *YLTELBOX.h*. For example, `YL_IOCTL_GEN_GETROUTECODE` or `KEY_CENTER`.
That is why some features was not implemented. 

## Contents

* Lib 		
Files from Yealink SDK (first version I think) 
* Source		
Source files of the helper project (Visual Studio 6 C++  project)
* YealinkTest	
Test/Example project (VB 6)

## License 
MIT