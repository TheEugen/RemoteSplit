#pragma once


#include <windows.h>
#include <ShObjIdl.h>
#include <cguid.h>
#include <atlbase.h>
#include <rdpencomapi.h>

#include <initguid.h>

// c6bfcd38-8ce9-404d-8ae8-f31d00c65cb5"
//DEFINE_GUID(IID_IRDPSRAPIViewer, 0xc6bfcd38, 0x8ce9, 0x404d, 0x8ae8, 0xf3, 0x1d, 0x00, 0xc6, 0x5c, 0xb5);

// 32be5ed2-5c86-480f-a914-0ff8885a1b3f
//DEFINE_GUID(CLSID_RDPViewer, 0x32be5ed2, 0x5c86, 0x480f, 0xa914, 0x0f, 0xf8, 0x88, 0x5a, 0x1b, 0x3f);

#include "Utils.h"

#pragma comment(lib, "atls.lib")

class Viewer
{
	CComPtr<IRDPSRAPIViewer> m_viewer;
	CComPtr<IUnknown> m_adivsor;

public:

	void start(HWND hWnd);
	IRDPSRAPIViewer* getAPIViewer();
	void putStreamOnWindow(HINSTANCE hInst, HWND hWnd);
};

