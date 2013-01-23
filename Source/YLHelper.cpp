// YLHelper.cpp : Defines the entry point for the DLL application.
//

#include "Afx.h"
#include "stdafx.h"
#include "YLTELBOX.h"

#pragma comment(lib, "YLTELBOX.lib")


BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
	return true;
}

extern DWORD WINAPI YL_Open(void*	lpInBuffer)
{
	return YL_DeviceIoControl(YL_IOCTL_OPEN_DEVICE, lpInBuffer, 0, 0, 0);
}

extern DWORD WINAPI YL_DeviceIoControl_(DWORD dwIoControlCode, 
		void*	lpInBuffer = 0,		DWORD nInBufferSize = 0, 
		void*	lpOutBuffer = 0,	DWORD nOutBufferSize = 0)
{
	return YL_DeviceIoControl(dwIoControlCode, lpInBuffer, nInBufferSize, lpOutBuffer, nOutBufferSize);
}

extern DWORD WINAPI YL_Close()
{
	return YL_DeviceIoControl(YL_IOCTL_CLOSE_DEVICE, 0, 0, 0, 0);
}

extern DWORD WINAPI YL_BeReady()
{
	return YL_DeviceIoControl(YL_IOCTL_GEN_READY, 0, 0, 0, 0);
}

extern DWORD WINAPI YL_BeNotReady()
{
	return YL_DeviceIoControl(YL_IOCTL_GEN_UNREADY, 0, 0, 0, 0);
}

extern DWORD WINAPI YL_Calling()
{
	return YL_DeviceIoControl(YL_IOCTL_GEN_CALLOUT, 0, 0, 0, 0);
}

extern DWORD WINAPI YL_GetRinging()
{
	return YL_DeviceIoControl(YL_IOCTL_GEN_RINGBACK, 0, 0, 0, 0);
}

extern DWORD WINAPI YL_Talking()
{
	return YL_DeviceIoControl(YL_IOCTL_GEN_TALKING, 0, 0, 0, 0);
}

extern DWORD WINAPI YL_SwitchToUsbChannels()
{
	return YL_DeviceIoControl(YL_IOCTL_GEN_GOTOUSB, 0, 0, 0, 0);
}

extern DWORD WINAPI YL_SwitchToPstnChannels()
{
	return YL_DeviceIoControl(YL_IOCTL_GEN_GOTOPSTN, 0, 0, 0, 0);
}

extern DWORD WINAPI YL_DefaultInUsbChannels()
{
	return YL_DeviceIoControl(YL_IOCTL_GEN_DEFAULTUSB, 0, 0, 0, 0);
}

extern DWORD WINAPI YL_DefaultInPstnChannels()
{
	return YL_DeviceIoControl(YL_IOCTL_GEN_DEFAULTPSTN, 0, 0, 0, 0);
}

extern DWORD WINAPI YL_UseUsbChannelsOnly()
{
	return YL_DeviceIoControl(YL_IOCTL_GEN_ONLYUSB, 0, 0, 0, 0);
}

