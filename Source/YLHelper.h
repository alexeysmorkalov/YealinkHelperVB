
// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the YLHELPER_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// YLHELPER_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef YLHELPER_EXPORTS
#define YLHELPER_API __declspec(dllexport)
#else
#define YLHELPER_API __declspec(dllimport)
#endif

#define WINAPI __stdcall


// This class is exported from the YLHelper.dll
class YLHELPER_API CYLHelper {
public:
	CYLHelper(void);
	// TODO: add your methods here.
};

extern YLHELPER_API int nYLHelper;

YLHELPER_API int fnYLHelper(void);

