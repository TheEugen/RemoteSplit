#pragma once

#include <Windows.h>

//#include "associated.h"




class AXClient : public IUnknown, public IOleClientSite, public IAdviseSink, public IDispatch, public IOleInPlaceSite
{
	HWND m_window = nullptr;
	HWND m_parentWindow = nullptr;
	IStorage* m_storage = nullptr;
	IOleObject* m_oleObj = nullptr;
	const IID* m_IID = &IID_IOleObject;
	ULONG m_refCount = 0;
	IViewObject* m_viewObj = nullptr;
	IDataObject* m_dataObj = nullptr;
	DWORD m_adviseToken = 0;

public:

	bool InPlace = false;
	bool CalledCanInPlace = false;
	bool ExternalPlace = false;

	// IUnknown
	HRESULT __stdcall QueryInterface(REFIID iid, void** ppvObject) override;
	ULONG __stdcall AddRef() override;
	ULONG __stdcall Release() override;

	// IOleClientSite
	HRESULT __stdcall SaveObject() override;
	HRESULT __stdcall GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker **ppmk) override;
	HRESULT __stdcall GetContainer(IOleContainer** ppContainer) override;
	HRESULT __stdcall ShowObject() override;
	HRESULT __stdcall OnShowWindow(BOOL fShow) override;
	HRESULT __stdcall RequestNewObjectLayout() override;

	// IAdviseSink
	void __stdcall OnDataChange(FORMATETC* pFormatEtc, STGMEDIUM* pStgmed) override;
	void __stdcall OnViewChange(DWORD dwAspect, LONG lIndex) override;
	void __stdcall OnRename(IMoniker* pmk) override;
	void __stdcall OnSave() override;
	void __stdcall OnClose() override;

	// IDispatch
	HRESULT _stdcall GetTypeInfoCount(unsigned int* pctinfo) override;
	HRESULT _stdcall GetTypeInfo(unsigned int iTInfo, LCID lcid, ITypeInfo FAR* FAR* ppTInfo) override;
	HRESULT _stdcall GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR*, unsigned int cNames, LCID lcid, DISPID FAR*) override;
	HRESULT _stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS FAR* pDispParams,
		VARIANT FAR* pVarResult, EXCEPINFO FAR* pExcepInfo, unsigned int FAR* puArgErr) override;

	// IOleInPlaceSite
	HRESULT __stdcall GetWindow(HWND* p) override;
	HRESULT __stdcall ContextSensitiveHelp(BOOL) override;
	HRESULT __stdcall CanInPlaceActivate() override;
	HRESULT __stdcall OnInPlaceActivate() override;
	HRESULT __stdcall OnUIActivate() override;
	HRESULT __stdcall GetWindowContext(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT r1, LPRECT r2, LPOLEINPLACEFRAMEINFO o) override;
	HRESULT __stdcall Scroll(SIZE s) override;
	HRESULT __stdcall OnUIDeactivate(int) override;
	HRESULT __stdcall OnInPlaceDeactivate() override;
	HRESULT __stdcall DiscardUndoState() override;
	HRESULT __stdcall DeactivateAndUndo() override;
	HRESULT __stdcall OnPosRectChange(LPCRECT) override;


	// custom
	void setWindowAndParent(HWND window);
	HWND customGetWindow();
	HWND customGetParentWindow();
	IStorage* getStorageP();
	IStorage** getStoragePP();
	IOleObject* getOleObjectP();
	IOleObject** getOleObjectPP();
	const IID* getIIDP() const;
	DWORD* getAdviseTokenP();
	IViewObject* getViewObjP();
	IViewObject** getViewObjPP();
	IDataObject** getDataObjPP();

	virtual ~AXClient();
};

// custom

AXClient::~AXClient() = default;

void AXClient::setWindowAndParent(HWND window)
{
	m_window = window;
	m_parentWindow = GetParent(window);
}

IStorage* AXClient::getStorageP() { return m_storage; }
IStorage** AXClient::getStoragePP() { return &m_storage; }
IOleObject* AXClient::getOleObjectP() { return m_oleObj; }
IOleObject** AXClient::getOleObjectPP() { return &m_oleObj; }
DWORD* AXClient::getAdviseTokenP() { return &m_adviseToken; }
IViewObject* AXClient::getViewObjP() { return m_viewObj; }
IViewObject** AXClient::getViewObjPP() { return &m_viewObj; }
IDataObject** AXClient::getDataObjPP() { return &m_dataObj; }
const IID* AXClient::getIIDP() const { return m_IID; }
HWND AXClient::customGetWindow() { return m_window; }
HWND AXClient::customGetParentWindow() { return m_parentWindow; }

// IUnknown

HRESULT AXClient::QueryInterface(REFIID iid, void** ppvObject)
{
	if(!ppvObject)
		return E_POINTER;

	if(iid == IID_IUnknown)
		*ppvObject = (IUnknown*)this;
	if(iid == IID_IOleClientSite)
		*ppvObject = (IOleClientSite*)this;
	if(iid == IID_IAdviseSink)
		*ppvObject = (IAdviseSink*)this;
	if(iid == IID_IDispatch)
		*ppvObject = (IDispatch*)this;
	if(iid == IID_IOleInPlaceSite)
		*ppvObject = (IOleInPlaceSite*)this;

	if(!ppvObject)
		return E_NOINTERFACE;

	this->AddRef();
	return S_OK;
}

ULONG AXClient::AddRef()
{
	return ++m_refCount;
}

ULONG AXClient::Release()
{
	if(--m_refCount == 0)
	{
		delete this;
		return 0;
	}

	return m_refCount;
}

// IOleClientSite

HRESULT __stdcall AXClient::SaveObject()
{
	return S_OK;
}
HRESULT __stdcall AXClient::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk)
{
	*ppmk = 0;
	return E_NOTIMPL;
}
HRESULT __stdcall AXClient::GetContainer(IOleContainer** ppContainer)
{
	*ppContainer = 0;
	return E_FAIL;
}
HRESULT __stdcall AXClient::ShowObject()
{
	return S_OK;
}
HRESULT __stdcall AXClient::OnShowWindow(BOOL fShow)
{
	return S_OK;
}
HRESULT __stdcall AXClient::RequestNewObjectLayout()
{
	return S_OK;
}

// IAdviseSink

void __stdcall AXClient::OnDataChange(FORMATETC* pFormatEtc, STGMEDIUM* pStgmed) {}
void __stdcall AXClient::OnViewChange(DWORD dwAspect, LONG lIndex) {}
void __stdcall AXClient::OnRename(IMoniker* pmk) {}
void __stdcall AXClient::OnSave() {}
void __stdcall AXClient::OnClose() {}


// IDispatch

HRESULT _stdcall AXClient::GetTypeInfoCount( unsigned int* pctinfo)
{
	return E_NOTIMPL;
}

HRESULT _stdcall AXClient::GetTypeInfo(unsigned int iTInfo, LCID lcid, ITypeInfo FAR* FAR* ppTInfo) 
{
	return E_NOTIMPL;
}

HRESULT _stdcall AXClient::GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR*, unsigned int cNames, LCID lcid, DISPID FAR*)
{
	return E_NOTIMPL;
}

HRESULT _stdcall AXClient::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS FAR* pDispParams,
										VARIANT FAR* pVarResult,EXCEPINFO FAR* pExcepInfo,unsigned int FAR* puArgErr)
{
	/*AXDISPATCHNOTIFICATION axd = { ax, dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr };
	if (ax->DispatchNotificationWindow)
	{
		SendMessage(ax->DispatchNotificationWindow, ax->DispatchNotificationMessage, 0, (LPARAM)&axd);
	}
	if (ax->DispatchNotificationFunction)
	{
		ax->DispatchNotificationFunction(&axd);
		return S_OK;
	}
	*/
	return S_OK;
}

// IOleInPlaceSite methods
HRESULT __stdcall AXClient::GetWindow(HWND* p)
{
	*p = customGetWindow();
	return S_OK;
}

HRESULT __stdcall AXClient::ContextSensitiveHelp(BOOL)
{
	return E_NOTIMPL;
}


HRESULT __stdcall AXClient::CanInPlaceActivate()
{
	if (InPlace)
	{
		CalledCanInPlace = true;
		return S_OK;
	}
	return S_FALSE;
}

HRESULT __stdcall AXClient::OnInPlaceActivate()
{
	return S_OK;
}

HRESULT __stdcall AXClient::OnUIActivate()
{
	return S_OK;
}

HRESULT __stdcall AXClient::GetWindowContext(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT r1, LPRECT r2, LPOLEINPLACEFRAMEINFO o)
{
	*ppFrame = (IOleInPlaceFrame*)this;
	AddRef();

	*ppDoc = nullptr;
	GetClientRect(customGetWindow(), r1);
	GetClientRect(customGetWindow(), r2);
	o->cb = sizeof(OLEINPLACEFRAMEINFO);
	o->fMDIApp = false;
	o->hwndFrame = customGetParentWindow();
	o->haccel = 0;
	o->cAccelEntries = 0;

	return S_OK;
}

HRESULT __stdcall AXClient::Scroll(SIZE s)
{
	return E_NOTIMPL;
}

HRESULT __stdcall AXClient::OnUIDeactivate(int)
{
	return S_OK;
}

HRESULT __stdcall AXClient::OnInPlaceDeactivate()
{
	return S_OK;
}

HRESULT __stdcall AXClient::DiscardUndoState()
{
	return S_OK;
}

HRESULT __stdcall AXClient::DeactivateAndUndo()
{
	return S_OK;
}

HRESULT __stdcall AXClient::OnPosRectChange(LPCRECT)
{
	return S_OK;
}



LRESULT createClientAXWindow(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, AXClient* axClient)
{
	/*
	// get class id from window msg
	//AXClient* axClient;
	char tit[1000] = { 0 };
	HRESULT hr;
	wchar_t wtit[1000] = { 0 };
	GetWindowTextW(hWnd, wtit, 1000);
	WideCharToMultiByte(CP_ACP, 0, wtit, -1, tit, 1000, nullptr, nullptr);
	
	
	if (message == WM_CREATE)
	{
		//if (tit[0] == '}')
			//return 0; // No Creation at the momemt
	}
	if (message == AX_RECREATE) // lParam is the IUnknown
	{
		tit[0] = '{';
	}
	*/

	// create AXClient obj
	axClient = new AXClient();

	// giving window a ref to our obj
	SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)axClient);
	
	// giving our obj ref to the windows
	axClient->setWindowAndParent(hWnd);

	// 
	HRESULT hr = StgCreateDocfile(0, STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_DIRECT | STGM_CREATE, 0, axClient->getStoragePP());
	if(hr == S_OK)
	{
		/*
		REFIID rid = *axClient->getIIDP();
		if (message == WM_CREATE)
		{
			CLSID cid = { 0 };
			CLSIDFromString(wtit, &cid);
			
			// create ole object
			// since RDPViewer doesnt support OleCache OLERENDER_DRAW should always fail
			hr = OleCreate(cid, IID_IOleObject, OLERENDER_DRAW, 0, axClient, axClient->getStorageP(), (void**)axClient->getOleObjectPP());
			if (hr != S_OK)
			{
				hr = OleCreate(cid, IID_IOleObject, OLERENDER_NONE, 0, axClient, axClient->getStorageP(), (void**)axClient->getOleObjectPP());
				//hr = OleCreate(__uuidof(RDPViewer), IID_IOleObject, OLERENDER_NONE, 0, axClient, axClient->getStorageP(), (void**)axClient->getOleObjectPP());
				if(hr != S_OK)
				{
					// critical error
					delete axClient;
					SetWindowLongPtr(hWnd, GWLP_USERDATA, 0);
					return E_ABORT;
				}
			}			
		}
		else
		{
			auto u = (IUnknown*)lParam;
			if (u)
				u->QueryInterface(rid, (void**)axClient->getOleObjectPP());
		}
		
		REFIID rid = *axClient->getIIDP();
		CLSID cid = { 0 };
		CLSIDFromString(wtit, &cid);
		hr = OleCreate(cid, IID_IOleObject, OLERENDER_NONE, 0, axClient, axClient->getStorageP(), (void**)axClient->getOleObjectPP());
		*/
		auto viewer = (IUnknown*)lParam;
		if (viewer)
			viewer->QueryInterface(IID_IOleObject, (void**)axClient->getOleObjectPP());
		
		// give ole obj ref to our client site obj (axclient)
		axClient->getOleObjectP()->SetClientSite(axClient);

		// TODO: error catching
		hr = OleSetContainedObject(axClient->getOleObjectP(), true);
		hr = axClient->getOleObjectP()->Advise(axClient, axClient->getAdviseTokenP());
		hr = axClient->getOleObjectP()->QueryInterface(IID_IViewObject, (void**)axClient->getViewObjPP());
		hr = axClient->getOleObjectP()->QueryInterface(IID_IDataObject, (void**)axClient->getDataObjPP());

		if (axClient->getViewObjPP())
			axClient->getViewObjP()->SetAdvise(DVASPECT_CONTENT, 0, (IAdviseSink*)axClient);

		return 0;
	}

	return E_FAIL;
}



LRESULT CALLBACK AXWndProcT(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	AXClient* axClient = 0;
	switch(message)
	{
	case WM_COMMAND:
		return 0;
	case WM_DESTROY:
		return true;
	case WM_PAINT:
		return 0;
	case WM_CREATE:
		return 0;
	case AX_RECREATE:
		createClientAXWindow(hWnd, message, wParam, lParam, axClient);
		return 0;
	case WM_GETMINMAXINFO:
		return 0;
	case WM_NCCALCSIZE:
		return 0;
	case WM_SIZE:
		return 0;
	case AX_INPLACE:
	{

		auto* ax = (AXClient*)GetWindowLongPtr(hWnd, GWLP_USERDATA);
		if (!ax)
			return 0;
		if (!ax->getOleObjectP())
			return 0;
		RECT rect;
		HRESULT hr;
		::GetClientRect(hWnd, &rect);
		
		if (ax->InPlace == false && wParam == 1) // Activate In Place
		{
			ax->InPlace = true;
			ax->ExternalPlace = lParam;
			hr = ax->getOleObjectP()->DoVerb(OLEIVERB_INPLACEACTIVATE, nullptr, ax, 0, hWnd, &rect);
			InvalidateRect(hWnd, nullptr, true);
			return 1;
		}
		
		if (ax->InPlace == true && wParam == 0) // Deactivate
		{
			ax->InPlace = false;
		
			IOleInPlaceObject* iib;
			ax->getOleObjectP()->QueryInterface(IID_IOleInPlaceObject, (void**)&iib);
			if (iib)
			{
				iib->UIDeactivate();
				iib->InPlaceDeactivate();
				iib->Release();
				InvalidateRect(hWnd, nullptr, true);
				return 1;
			}
		}
		return 0;

	}
	case AX_QUERYINTERFACE:
	{
		auto const* p = (char*)wParam;
		auto* ax = (AXClient*)GetWindowLongPtr(hWnd, GWLP_USERDATA);
		if (!ax)
			return 0;
		return ax->getOleObjectP()->QueryInterface((REFIID)*p, (void**)lParam);
	}
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}

	return DefWindowProc(hWnd, message, wParam, lParam);
}


static ATOM AXClientRegister(HINSTANCE hInstance)
{
	WNDCLASSEXW wcex = { 0 };

	wcex.cbSize = sizeof(WNDCLASSEX);

	// wC.style = CS_GLOBALCLASS | CS_DBLCLKS;
	wcex.style = CS_GLOBALCLASS | CS_DBLCLKS; //CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc = AXWndProcT;
	//wcex.cbClsExtra = 0;
	// ?
	wcex.cbWndExtra = 4;
	wcex.hInstance = hInstance;
	//wcex.hInstance = GetModuleHandle(0);
	wcex.lpszClassName = L"AXClient";


	//wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_REMOTESPLIT));
	//wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
	//wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	//wcex.lpszMenuName = MAKEINTRESOURCEW(IDC_REMOTESPLIT);
	//wcex.lpszClassName = szWindowClass;
	//wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));


	return RegisterClassExW(&wcex);
}