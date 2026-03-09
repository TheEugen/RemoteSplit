
#include "Sharer.h"

void Sharer::start(HWND hWnd)
{
	//IID_IRDPSRAPISharingSession
	HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);

	// create SharingSession object
	hr = CoCreateInstance(CLSID_RDPSession, nullptr, CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER, IID_IRDPSRAPISharingSession, (void**)&m_session);

	// container for connection points
	CComPtr<IConnectionPointContainer> cpc;
	m_session->QueryInterface(&cpc);

	// find connection point
	CComPtr<IConnectionPoint> cp;
	hr = cpc->FindConnectionPoint(DIID__IRDPSessionEvents, &cp);

	// link our connection point to our sink to receive events
	DWORD connectionCookie = 0;
	m_events = CComPtr<RDPSessionEvents>(new RDPSessionEvents());
	cp->Advise(m_events, &connectionCookie);

	m_events->setOutputLog(GetDlgItem(hWnd, IDC_OUTPUTLOG));

	//p->mAttendees = -1;
	//p->mConnectionString.clear();

	//VARIANT variant = { variant.vt = VT_BOOL, variant.boolVal = VARIANT_TRUE };

	// get session properties and alter them
	CComPtr<IRDPSRAPISessionProperties> sessionProps;
	hr = m_session->get_Properties(&sessionProps);
	hr = sessionProps->put_Property(BStr(L"DrvConAttach"), Variant(true));
	hr = sessionProps->put_Property(BStr(L"PortProtocol"), Variant(AF_INET));
	hr = sessionProps->put_Property(BStr(L"PortId"), Variant(m_port));

	// 24 bit = true colour
	//hr = m_session->put_ColorDepth(24);

	///////////////
	/*
	IRDPSRAPIApplication* app = nullptr;
	hr = CoCreateInstance(CLSID_RDPSRAPIApplication, nullptr, CLSCTX_INPROC_SERVER, __uuidof(IRDPSRAPIApplication), (void**)&app);
	
	//m_session->QueryInterface(__uuidof(IRDPSRAPIApplication), (void**)&app);
	IRDPSRAPIWindowList* windowList2 = nullptr;
	app->get_Windows(&windowList2);
	for(long i = 0; i < 100; ++i)
	{
		IRDPSRAPIWindow* window2 = nullptr;
		windowList2->get_Item(i, &window2);
	}

	CComPtr<IRDPSRAPIApplicationFilter> appFilter;
	m_session->get_ApplicationFilter(&appFilter);

	CComPtr<IRDPSRAPIApplicationList> appList;
	CComPtr<IRDPSRAPIWindowList> windowList;
	appFilter->get_Applications(&appList);
	appFilter->get_Windows(&windowList);
	std::vector<IRDPSRAPIApplication*> apps;
	std::vector<IRDPSRAPIWindow*> windows;
	for(long i = 0; i < 100; ++i)
	{
		apps.push_back(0);
		windows.push_back(0);
		IRDPSRAPIApplication* y = nullptr;
		IRDPSRAPIWindow* x = nullptr;
		hr = appList->get_Item(i, &y);
		hr = windowList->get_Item(i, &x);
	}
	*/
	/*
	CComPtr<IRDPSRAPIApplicationFilter> appFilter;
	m_session->get_ApplicationFilter(&appFilter);
	CComPtr<IRDPSRAPIWindowList> windowList;
	appFilter->get_Windows(&windowList);

	std::vector<BSTR> windowNames;

	IEnumVARIANT* pEnum = nullptr;
	if (SUCCEEDED(windowList->get__NewEnum((IUnknown**)&pEnum)))
	{
		int count = 0;
		VARIANT var;
		VariantInit(&var);
		
		while (pEnum->Next(1, &var, nullptr) == S_OK)
		{
			if (var.vt == VT_DISPATCH && var.pdispVal != nullptr)
			{
				IRDPSRAPIWindow* pWindow = nullptr;
				if (SUCCEEDED(var.pdispVal->QueryInterface(__uuidof(IRDPSRAPIWindow), (void**)&pWindow)))
				{
					++count;
					//BStr* a = new BStr(L"");
					BSTR b;
					windowNames.push_back(b);
					pWindow->get_Name(&windowNames.back());
					pWindow->put_Shared(VARIANT_FALSE);
					//if(windowNames.back() != BStr(L"FreundesListe"))
						//pWindow->put_Shared(VARIANT_FALSE);
					//else
						//pWindow->put_Shared(VARIANT_TRUE);
					// Do something with pWindow
					pWindow->Release();
					pWindow->ge
				}
			}
			VariantClear(&var);
		}
		pEnum->Release();
	}
	*/
	////////////


	m_rect = { m_rect.left = ::GetSystemMetrics(SM_XVIRTUALSCREEN),
				m_rect.top = ::GetSystemMetrics(SM_YVIRTUALSCREEN),
				m_rect.right = ::GetSystemMetrics(SM_CXVIRTUALSCREEN),
				m_rect.bottom = ::GetSystemMetrics(SM_CYVIRTUALSCREEN) };

	m_rect = { 0, 0, 1920, 1080 };

	// set the region to be shared
	hr = m_session->SetDesktopSharedRect(m_rect.left, m_rect.top, m_rect.right, m_rect.bottom);

	// open the session
	hr = m_session->Open();

	// create invitations
	CComPtr<IRDPSRAPIInvitationManager> invitationManager;
	hr = m_session->get_Invitations(&invitationManager);
	CComPtr<IRDPSRAPIInvitation> invitation;
	hr = invitationManager->CreateInvitation(nullptr, BStr(L"Server"), BStr(L""), m_maxAttendes, &invitation);
	if (hr == S_OK)
	{
		BSTR bConnectionStr;
		hr = invitation->get_ConnectionString(&bConnectionStr);

		if (hr == S_OK)
		{
			BStr connectionStr = bConnectionStr;

			// create file window object
			IFileSaveDialog* pFileOpen;
			hr = CoCreateInstance(CLSID_FileSaveDialog, NULL, CLSCTX_ALL,
				IID_IFileSaveDialog, reinterpret_cast<void**>(&pFileOpen));

			// TODO: if user doesnt add extension, automatically add it
			COMDLG_FILTERSPEC filterspec[1] = { {L"XML Files", L"*.xml"} };
			pFileOpen->SetFileTypes(1, filterspec);

			// show the file window and get user input
			hr = pFileOpen->Show(nullptr);
			IShellItem* pItem;
			hr = pFileOpen->GetResult(&pItem);
			PWSTR pszFilePath;
			hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);

			// create the file and write our connection str into it
			HANDLE file = CreateFile(pszFilePath, GENERIC_WRITE, 0, 0, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, 0);
			WriteFile(file, (void*)connectionStr, connectionStr.strSize() * 2, 0, 0);

			CloseHandle(file);
		}
	}
	

}