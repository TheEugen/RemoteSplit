#include "Viewer.h"



void Viewer::putStreamOnWindow(HINSTANCE hInst, HWND hWnd)
{
	/*
	// While looking at the calls made I have never seen this create window called it is always returns 0 and is recreated
	// NOTE: because OleCreate is always failing with E_NOSUCHINTERFACE
	HWND ACTIVEX_WINDOW = CreateWindowEx(WS_EX_APPWINDOW, L"AX", L"{32be5ed2-5c86-480f-a914-0ff8885a1b3f}", WS_CHILD | WS_VISIBLE, 0, 20, 1200, 600, hWnd, 0, hInst, 0);
	if (!ACTIVEX_WINDOW)
	{
		outputLog(GetDlgItem(GetParent(hWnd), IDC_OUTPUTLOG), L"ActiveX window creating failed with error code: ");
		outputLog(GetDlgItem(GetParent(hWnd), IDC_OUTPUTLOG), std::to_wstring(GetLastError()).c_str());
	}

	if (ACTIVEX_WINDOW)
	{
		IUnknown* a = 0;
		SendMessageA(ACTIVEX_WINDOW, AX_QUERYINTERFACE, (WPARAM) & __uuidof(IUnknown*), (LPARAM)&a);
	}
	else
	{
		//ACTIVEX_WINDOW = CreateWindowA("AX", "}32BE5ED2-5C86-480F-A914-0FF8885A1B3F}", WS_CHILD | WS_VISIBLE, 0, 0, 1200, 600, hWnd, 0, hInst, 0);
		ACTIVEX_WINDOW = CreateWindowA("AX", "}32BE5ED2-5C86-480F-A914-0FF8885A1B3F}", WS_CHILD | WS_VISIBLE, 0, 0, 1200, 600, hWnd, 0, hInst, 0);
		if (!ACTIVEX_WINDOW)
		{
			outputLog(GetDlgItem(GetParent(hWnd), IDC_OUTPUTLOG), L"ActiveX window creating failed with error code: ");
			outputLog(GetDlgItem(GetParent(hWnd), IDC_OUTPUTLOG), std::to_wstring(GetLastError()).c_str());
		}
		SendMessageA(ACTIVEX_WINDOW, AX_RECREATE, 0, (LPARAM)getAPIViewer());	
	}
	
	SendMessageA(ACTIVEX_WINDOW, AX_INPLACE, 1, 0);
	ShowWindow(ACTIVEX_WINDOW, SW_MAXIMIZE);
	*/

	
	HWND ACTIVEX_WINDOW_T = CreateWindowEx(WS_EX_APPWINDOW, L"AXClient", L"{32be5ed2-5c86-480f-a914-0ff8885a1b3f}", WS_CHILD | WS_VISIBLE, 0, 0, 1200, 600, hWnd, 0, hInst, 0);
	if (!ACTIVEX_WINDOW_T)
	{
		outputLog(GetDlgItem(GetParent(hWnd), IDC_OUTPUTLOG), L"ActiveX window creating failed with error code: ");
		outputLog(GetDlgItem(GetParent(hWnd), IDC_OUTPUTLOG), std::to_wstring(GetLastError()).c_str());

		ACTIVEX_WINDOW_T = CreateWindowA("AXClient", "}32BE5ED2-5C86-480F-A914-0FF8885A1B3F}", WS_CHILD | WS_VISIBLE, 0, 0, 1200, 600, hWnd, 0, hInst, 0);
		if (!ACTIVEX_WINDOW_T)
		{
			outputLog(GetDlgItem(GetParent(hWnd), IDC_OUTPUTLOG), L"ActiveX window creating failed with error code: ");
			outputLog(GetDlgItem(GetParent(hWnd), IDC_OUTPUTLOG), std::to_wstring(GetLastError()).c_str());
		}
	}
	
	
	SendMessageA(ACTIVEX_WINDOW_T, AX_RECREATE, 0, (LPARAM)getAPIViewer());
	
	
	//HWND ACTIVEX_WINDOW_T = CreateWindowEx(WS_EX_APPWINDOW, L"AXClient", L"{32be5ed2-5c86-480f-a914-0ff8885a1b3f}", WS_CHILD | WS_VISIBLE, 0, 0, 1200, 600, hWnd, nullptr, hInst, (LPVOID)getAPIViewer());


	SendMessageA(ACTIVEX_WINDOW_T, AX_INPLACE, 1, 0);
	ShowWindow(ACTIVEX_WINDOW_T, SW_MAXIMIZE);
	
	
}

void Viewer::start(HWND hWnd)
{
	// COM init
	HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);

	// create Viewer object
	hr = CoCreateInstance(__uuidof(RDPViewer), nullptr, CLSCTX_INPROC_SERVER, __uuidof(IRDPSRAPIViewer), (void**)&m_viewer);

	// container for connection points
	CComPtr<IConnectionPointContainer> cpc;
	m_viewer->QueryInterface(&cpc);

	// find connection point
	CComPtr<IConnectionPoint> cp;
	hr = cpc->FindConnectionPoint(__uuidof(_IRDPSessionEvents), &cp);

	DWORD connectionCookie;
	cp->Advise(m_adivsor, &connectionCookie);

	// create file window object
	IFileOpenDialog* pFileOpen;
	hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_ALL,
		IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));
	
	
	COMDLG_FILTERSPEC filterspec[1] = { {L"XML Files", L"*.xml"} };
	pFileOpen->SetFileTypes(1, filterspec);

	// show the file window and get user input
	hr = pFileOpen->Show(nullptr);
	IShellItem* pItem;
	hr = pFileOpen->GetResult(&pItem);
	PWSTR pszFilePath;
	hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);
	if (hr == S_OK)
	{
		wchar_t connectionStr[1680] = { 0 };
		DWORD bytesRead = 0;
		// open the file containing connection str
		HANDLE file = CreateFile(pszFilePath, GENERIC_READ, 0, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
		if (ReadFile(file, connectionStr, 1680 * 2, &bytesRead, nullptr))
		{
			outputLog(GetDlgItem(hWnd, IDC_OUTPUTLOG), L"Successfully read connection file!\n");
			outputLog(GetDlgItem(hWnd, IDC_OUTPUTLOG), std::to_wstring(bytesRead).c_str());

			hr = m_viewer->Connect(BStr(connectionStr), BStr("User123"), BStr(""));
			if (hr == S_OK)
			{
				outputLog(GetDlgItem(hWnd, IDC_OUTPUTLOG), L"Successfully connected!\n");
			}
			else
			{
				outputLog(GetDlgItem(hWnd, IDC_OUTPUTLOG), L"Establishing connection failed!\n");
			}
		}
	}

	

	
}

IRDPSRAPIViewer* Viewer::getAPIViewer()
{
	return m_viewer;
}