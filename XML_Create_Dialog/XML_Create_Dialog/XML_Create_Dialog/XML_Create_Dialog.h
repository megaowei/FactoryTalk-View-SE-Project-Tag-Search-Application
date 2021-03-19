
// XML_Create_Dialog.h : main header file for the PROJECT_NAME application
//

#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif

#include "resource.h"		// main symbols

#include <iostream>
#include <tchar.h>
#include <io.h>
#include <comdef.h> 

// CXML_Create_DialogApp:
// See XML_Create_Dialog.cpp for the implementation of this class
//

class CXML_Create_DialogApp : public CWinApp
{
public:
	CXML_Create_DialogApp();

// Overrides
public:
	virtual BOOL InitInstance();

// Implementation

	DECLARE_MESSAGE_MAP()
};

extern CXML_Create_DialogApp theApp;