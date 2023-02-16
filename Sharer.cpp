
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

	m_rect = { m_rect.left = ::GetSystemMetrics(SM_XVIRTUALSCREEN),
				m_rect.top = ::GetSystemMetrics(SM_YVIRTUALSCREEN),
				m_rect.right = ::GetSystemMetrics(SM_CXVIRTUALSCREEN),
				m_rect.bottom = ::GetSystemMetrics(SM_CYVIRTUALSCREEN) };

	m_rect = { 0, 0, 1000, 1000 };

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