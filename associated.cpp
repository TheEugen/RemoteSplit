// AX.CPP
#pragma once

#include <windows.h>
#include <comdef.h>

#include "associated.h"

//#include <oleidl.h>

//#pragma comment (lib, "Ole32.lib")


#pragma warning (disable: 4244)
#pragma warning (disable: 4800)

// AXClientSite class
// ------- Implement member functions
AXClientSite::AXClientSite()
{
	refNum = 0;
	CalledCanInPlace = 0;
	InPlace = 0;
}

AXClientSite :: ~AXClientSite()
{
}


// IUnknown methods
STDMETHODIMP AXClientSite::QueryInterface(REFIID iid, void** ppvObject)
{
	*ppvObject = 0;
	if (iid == IID_IOleClientSite)
		*ppvObject = (IOleClientSite*)this;
	if (iid == IID_IUnknown)
		*ppvObject = this;
	if (iid == IID_IAdviseSink)
		*ppvObject = (IAdviseSink*)this;
	if (iid == IID_IDispatch)
		*ppvObject = (IDispatch*)this;
	if (iid == SpecialIIDForceOK)
		*ppvObject = (IDispatch*)this;
	if (iid == IID_IServiceProvider)
		*ppvObject = (IServiceProvider*)this;
	if (ExternalPlace == false)
	{
		if (iid == IID_IOleInPlaceSite)
			*ppvObject = (IOleInPlaceSite*)this;
		if (iid == IID_IOleInPlaceFrame)
			*ppvObject = (IOleInPlaceFrame*)this;
		if (iid == IID_IOleInPlaceUIWindow)
			*ppvObject = (IOleInPlaceUIWindow*)this;

		
		//if (iid == IID_IOleInPlaceSiteEx)
			//*ppvObject = (IOleInPlaceSiteEx*)this;
	}

	if (iid == IID_IOleCache)
		*ppvObject = (IOleCache*)this;

	//* Log Call
	if (*ppvObject)
	{
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG)  AXClientSite::AddRef()
{
	refNum++;
	return refNum;
}

STDMETHODIMP_(ULONG)  AXClientSite::Release()
{
	refNum--;
	return refNum;
}


// IServiceManager 	
STDMETHODIMP AXClientSite::QueryService(REFGUID guid, REFIID iid, void** ppvObject)
{
	if (!ax->CustomServiceProvider)
		return E_NOINTERFACE;
	return ax->CustomServiceProvider->QueryService(guid, iid, ppvObject);
}



// IOleClientSite methods
STDMETHODIMP AXClientSite::SaveObject()
{
	return S_OK;
}

STDMETHODIMP AXClientSite::GetMoniker(DWORD dwA, DWORD dwW, IMoniker** pm)
{
	*pm = 0;
	return E_NOTIMPL;
}

STDMETHODIMP AXClientSite::GetContainer(IOleContainer** pc)
{
	*pc = 0;
	return E_FAIL;
}

STDMETHODIMP AXClientSite::ShowObject()
{
	return S_OK;
}

STDMETHODIMP AXClientSite::OnShowWindow(BOOL f)
{
	InvalidateRect(Window, 0, TRUE);
	InvalidateRect(Parent, 0, TRUE);
	return S_OK;
}

STDMETHODIMP AXClientSite::RequestNewObjectLayout()
{
	return S_OK;
}

STDMETHODIMP_(void) AXClientSite::OnViewChange(DWORD dwAspect, LONG lIndex)
{
}

STDMETHODIMP_(void) AXClientSite::OnRename(IMoniker* pmk)
{
}

STDMETHODIMP_(void) AXClientSite::OnSave()
{
}

STDMETHODIMP_(void) AXClientSite::OnClose()
{
}


// IOleInPlaceSite methods
STDMETHODIMP AXClientSite::GetWindow(HWND* p)
{
	*p = Window;
	return S_OK;
}

STDMETHODIMP AXClientSite::ContextSensitiveHelp(BOOL)
{
	return E_NOTIMPL;
}


STDMETHODIMP AXClientSite::CanInPlaceActivate()
{
	if (InPlace)
	{
		CalledCanInPlace = true;
		return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP AXClientSite::OnInPlaceActivate()
{
	return S_OK;
}

STDMETHODIMP AXClientSite::OnUIActivate()
{
	return S_OK;
}

STDMETHODIMP AXClientSite::GetWindowContext(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT r1, LPRECT r2, LPOLEINPLACEFRAMEINFO o)
{
	*ppFrame = (IOleInPlaceFrame*)this;
	AddRef();

	*ppDoc = NULL;
	GetClientRect(Window, r1);
	GetClientRect(Window, r2);
	o->cb = sizeof(OLEINPLACEFRAMEINFO);
	o->fMDIApp = false;
	o->hwndFrame = Parent;
	o->haccel = 0;
	o->cAccelEntries = 0;

	return S_OK;
}

STDMETHODIMP AXClientSite::Scroll(SIZE s)
{
	return E_NOTIMPL;
}

STDMETHODIMP AXClientSite::OnUIDeactivate(int)
{
	return S_OK;
}

STDMETHODIMP AXClientSite::OnInPlaceDeactivate()
{
	return S_OK;
}

STDMETHODIMP AXClientSite::DiscardUndoState()
{
	return S_OK;
}

STDMETHODIMP AXClientSite::DeactivateAndUndo()
{
	return S_OK;
}

STDMETHODIMP AXClientSite::OnPosRectChange(LPCRECT)
{
	return S_OK;
}


// IOleInPlaceFrame methods
STDMETHODIMP AXClientSite::GetBorder(LPRECT l)
{
	GetClientRect(Window, l);
	return S_OK;
}

STDMETHODIMP AXClientSite::RequestBorderSpace(LPCBORDERWIDTHS b)
{
	//return S_OK;
	return E_NOTIMPL;
}

STDMETHODIMP AXClientSite::SetBorderSpace(LPCBORDERWIDTHS b)
{
	return S_OK;
}

STDMETHODIMP AXClientSite::SetActiveObject(IOleInPlaceActiveObject* pV, LPCOLESTR s)
{
	ax->Pao = pV;
	return S_OK;
}


STDMETHODIMP AXClientSite::SetStatusText(LPCOLESTR t)
{
	return E_NOTIMPL;
}

STDMETHODIMP AXClientSite::EnableModeless(BOOL f)
{
	return E_NOTIMPL;
}

STDMETHODIMP AXClientSite::TranslateAccelerator(LPMSG, WORD)
{
	return E_NOTIMPL;
}


// IDispatch Methods
HRESULT _stdcall AXClientSite::GetTypeInfoCount(
	unsigned int* pctinfo) {
	return E_NOTIMPL;
}

HRESULT _stdcall AXClientSite::GetTypeInfo(
	unsigned int iTInfo,
	LCID lcid,
	ITypeInfo FAR* FAR* ppTInfo) {
	return E_NOTIMPL;
}

HRESULT _stdcall AXClientSite::GetIDsOfNames(
	REFIID riid,
	OLECHAR FAR* FAR*,
	unsigned int cNames,
	LCID lcid,
	DISPID FAR*) {
	return E_NOTIMPL;
}

HRESULT _stdcall AXClientSite::Cache(
	/* [unique][in] */ __RPC__in_opt FORMATETC* pformatetc,
	/* [in] */ DWORD advf,
	/* [out] */ __RPC__out DWORD* pdwConnection) {
	return E_NOTIMPL;
}
HRESULT _stdcall AXClientSite::Uncache(
	/* [in] */ DWORD dwConnection) {
	return E_NOTIMPL;
}
HRESULT _stdcall AXClientSite::EnumCache(
	/* [out] */ __RPC__deref_out_opt IEnumSTATDATA** ppenumSTATDATA) {
	return E_NOTIMPL;
}
HRESULT _stdcall AXClientSite::InitCache(
	/* [unique][in] */ __RPC__in_opt IDataObject* pDataObject) {
	return E_NOTIMPL;
}
HRESULT _stdcall AXClientSite::SetData(
	/* [unique][in] */ __RPC__in_opt FORMATETC* pformatetc,
	/* [unique][in] */ __RPC__in_opt STGMEDIUM* pmedium,
	/* [in] */ BOOL fRelease) {
	return E_NOTIMPL;
}


// Other Methods
void AX::Init(char* cls)
{
	wchar_t x[1000] = { 0 };
	MultiByteToWideChar(CP_ACP, 0, cls, -1, x, 1000);
	CLSIDFromString(x, &clsid);
	iid = (IID*)&IID_IOleObject;
	OleObject = 0;
	Storage = 0;
	View = 0;
	Data = 0;
	Unk = 0;
	Pao = 0;
	AdviseToken = 0;
	CustomServiceProvider = 0;
	memset(DAdviseToken, 0, sizeof(DAdviseToken));
	Site.ax = this;
	DispatchNotificationFunction = 0;
	DispatchNotificationWindow = 0;
	DispatchNotificationMessage = 0;
}

AX::AX(char* cls)
{
	Init(cls);
}



void AX::Clean()
{
	if (Site.InPlace == true)
	{
		Site.InPlace = false;
		IOleInPlaceObject* iib = 0;
		if (OleObject)
			OleObject->QueryInterface(IID_IOleInPlaceObject, (void**)&iib);
		if (iib)
		{
			iib->UIDeactivate();
			iib->InPlaceDeactivate();
			iib->Release();
		}
	}

	if (AdviseToken && OleObject)
	{
		OleObject->Unadvise(AdviseToken);
		AdviseToken = 0;
	}
	if (Data)
	{
		for (int i = 0; i < 100; i++)
			if (DAdviseToken[i])
				Data->DUnadvise(DAdviseToken[i]);
		memset(DAdviseToken, 0, sizeof(DAdviseToken));
	}


	if (Pao) Pao->Release();
	if (Unk) Unk->Release();
	if (Data) Data->Release();
	if (View) View->Release();
	if (Storage) Storage->Release();
	if (OleObject) OleObject->Release();
	Unk = 0;
	Data = 0;
	View = 0;
	Storage = 0;
	OleObject = 0;

}

AX :: ~AX()
{
	Clean();
}

CLSID AX::GetCLSID()
{
	return clsid;
}

HRESULT _stdcall AXClientSite::InsertMenus(HMENU h, LPOLEMENUGROUPWIDTHS x)
{

	return E_NOTIMPL;
}
HRESULT _stdcall AXClientSite::SetMenu(HMENU h, HOLEMENU hO, HWND hw)
{
	AX* t = (AX*)ax;

	return E_NOTIMPL;
}

HRESULT _stdcall AXClientSite::RemoveMenus(HMENU h)
{
#ifdef WINCE
	return E_NOTIMPL;
#else
	AX* t = (AX*)ax;
	if (t->AddMenu)
	{
		if (!h)
			return S_OK;

		int c = GetMenuItemCount(h);
		for (int i = c; i >= 0; i--)
		{
			HMENU hh = GetSubMenu(h, i);
			if (hh == Menu)
				RemoveMenu(h, i, MF_BYPOSITION);
		}
		if (h == Menu)
			DestroyMenu(h);

		//DrawMenuBar(Window);
		DrawMenuBar(Parent);
		return S_OK;
	}
	return E_NOTIMPL;
#endif
}

//extern HRESULT ProcessCmd(char*);
HRESULT _stdcall AXClientSite::Invoke(
	DISPID dispIdMember,
	REFIID riid,
	LCID lcid,
	WORD wFlags,
	DISPPARAMS FAR* pDispParams,
	VARIANT FAR* pVarResult,
	EXCEPINFO FAR* pExcepInfo,
	unsigned int FAR* puArgErr)
{
	AXDISPATCHNOTIFICATION axd = { ax, dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr };
	if (ax->DispatchNotificationWindow)
	{
		SendMessage(ax->DispatchNotificationWindow, ax->DispatchNotificationMessage, 0, (LPARAM)&axd);
	}
	if (ax->DispatchNotificationFunction)
	{
		ax->DispatchNotificationFunction(&axd);
		return S_OK;
	}

	return S_OK;
}

void _stdcall AXClientSite::OnDataChange(FORMATETC* pFormatEtc, STGMEDIUM* pStgmed)
{
	// Notify our app that a change is being requested
	return;
}

// Window Procedure for AX control
LRESULT CALLBACK AXWndProc(HWND hh, UINT mm, WPARAM ww, LPARAM ll)
{
	if (mm == WM_CREATE || mm == AX_RECREATE)
	{
		AX* ax;
		char tit[1000] = { 0 };
		HRESULT hr;
		wchar_t wtit[1000] = { 0 };
		GetWindowTextW(hh, wtit, 1000);
		WideCharToMultiByte(CP_ACP, 0, wtit, -1, tit, 1000, 0, 0);
		if (mm == WM_CREATE)
		{
			if (tit[0] == '}')
				return 0; // No Creation at the momemt
		}
		if (mm == AX_RECREATE) // ll is the IUnknown
		{
			tit[0] = '{';
		}


		ax = new AX(tit);
		SetWindowLongPtr(hh, GWLP_USERDATA, (LONG_PTR)ax);
		ax->Site.Window = hh;
		ax->Site.Parent = GetParent(hh);

		hr = StgCreateDocfile(0, STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_DIRECT | STGM_CREATE, 0, &ax->Storage);
		ax->Site.Window = hh;

		REFIID rid = *ax->iid;
		if (mm == WM_CREATE)
		{
			CLSID cid = ax->GetCLSID();

			hr = OleCreate(cid, rid, OLERENDER_DRAW, 0, &ax->Site, ax->Storage, (void**)&ax->OleObject);
			if (FAILED(hr))
			{
				hr = OleCreate(cid, rid, OLERENDER_NONE, 0, &ax->Site, ax->Storage, (void**)&ax->OleObject);
			}
		}
		else
			if (mm == AX_RECREATE)
			{
				IUnknown* u = (IUnknown*)ll;
				if (u)
					u->QueryInterface(rid, (void**)&ax->OleObject);
			}

		if (!ax->OleObject)
		{
			delete ax;
			SetWindowLongPtr(hh, GWLP_USERDATA, 0);
			return -1;
		}
		ax->OleObject->SetClientSite(&ax->Site);

		hr = OleSetContainedObject(ax->OleObject, TRUE);
		hr = ax->OleObject->Advise(&ax->Site, &ax->AdviseToken);
		hr = ax->OleObject->QueryInterface(IID_IViewObject, (void**)&ax->View);
		hr = ax->OleObject->QueryInterface(IID_IDataObject, (void**)&ax->Data);
		if (ax->View)
			hr = ax->View->SetAdvise(DVASPECT_CONTENT, 0, &ax->Site);

		return 0;
	}

	if (mm == WM_DESTROY)
	{
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;
		ax->Clean();
		return true;
	}

	if (mm == WM_COMMAND)
	{
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;

		return SendMessage(ax->CommandCallbackWindow, mm, ww, ll);
	}

	if (mm == AX_SETSERVICEPROVIDER)
	{
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;
		ax->CustomServiceProvider = (IServiceProvider*)ll;
		return 1;
	}

	if (mm == AX_CONNECTOBJECT)
	{
		// ww = (IID*)  Interface ID to which we request connection
		// ll = (void*) Pointer to a AX_CONNECTSTRUCT to fill in
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;

		int* wx = (int*)ww;
		AX_CONNECTSTRUCT* tcs = (AX_CONNECTSTRUCT*)ll;

		if (tcs->Advisor == 0)
		{
			tcs->Advisor = (IUnknown*)&ax->Site;
			ax->Site.SpecialIIDForceOK = tcs->SpecialIIDForceOK;
		}
		tcs->id = AXConnectObject((IUnknown*)ax->OleObject, (REFIID)*wx, tcs->Advisor, &tcs->icpc, &tcs->icp);
		ax->tcs = new AX_CONNECTSTRUCT;
		memcpy(ax->tcs, tcs, sizeof(AX_CONNECTSTRUCT));
		return 0;
	}

	if (mm == AX_DISCONNECTOBJECT)
	{
		// Use of ConnectObject
		// ww = (IID*)  Interface ID to which we request connection
		char* p = (char*)ww;
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;

		AX_CONNECTSTRUCT* tcs = ax->tcs;
		AXDisconnectObject(tcs->icpc, tcs->icp, tcs->id);
		delete ax->tcs;
		ax->tcs = 0;
		return 0;
	}


	if (mm == AX_SETCOMMANDCALLBACK)
	{
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;

		ax->CommandCallbackWindow = (HWND)ww;
		return 1;
	}

	if (mm == AX_SETDISPATCHNOTIFICATION)
	{
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;

		// ww = Window
		// ll = Message
		ax->DispatchNotificationWindow = (HWND)ww;
		ax->DispatchNotificationMessage = ll;
		return 1;
	}

	if (mm == AX_SETDISPATCHNOTIFICATIONFUNC)
	{
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;

		ax->DispatchNotificationFunction = (void(__stdcall*)(AXDISPATCHNOTIFICATION*))ww;
		return 1;
	}

	if (mm == AX_SETDATAADVISE)
	{
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;

		// 1.Enum available FORMATETC structures
		// 2.Set Data Advise specified to index ww
		if (!ax->Data)
			return 0;

		IEnumFORMATETC* ief = 0;
		int i = 0;

		FORMATETC fe;
		ax->Data->EnumFormatEtc((LPARAM)ll, &ief);
		if (!ief)
			return 0;
		for (;;)
		{
			HRESULT hr = ief->Next(1, &fe, 0);
			if (hr != S_OK)
				break;
			if (ww == i)
				break;
			i++;
		}
		ief->Release();
		if (ww == -1)
			return i;

		if (ax->Data)
			ax->Data->DAdvise(&fe, 0, &ax->Site, &ax->DAdviseToken[ww]);

		return true;
	}

	if (mm == AX_GETAXINTERFACE)
	{
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		return (LONG_PTR)ax;
	}

	if (mm == AX_QUERYINTERFACE)
	{
		char* p = (char*)ww;
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;
		return ax->OleObject->QueryInterface((REFIID)*p, (void**)ll);
	}

	if (mm == WM_LBUTTONDBLCLK)
	{
		PostMessage(hh, AX_INPLACE, 1, 0);
		return 0;
	}


	if (mm == AX_INPLACE)
	{
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;
		if (!ax->OleObject)
			return 0;
		RECT rect;
		HRESULT hr;
		::GetClientRect(hh, &rect);

		if (ax->Site.InPlace == false && ww == 1) // Activate In Place
		{
			ax->Site.InPlace = true;
			ax->Site.ExternalPlace = ll;
			hr = ax->OleObject->DoVerb(OLEIVERB_INPLACEACTIVATE, 0, &ax->Site, 0, hh, &rect);
			InvalidateRect(hh, 0, true);
			return 1;
		}

		if (ax->Site.InPlace == true && ww == 0) // Deactivate
		{
			ax->Site.InPlace = false;

			IOleInPlaceObject* iib;
			ax->OleObject->QueryInterface(IID_IOleInPlaceObject, (void**)&iib);
			if (iib)
			{
				iib->UIDeactivate();
				iib->InPlaceDeactivate();
				iib->Release();
				InvalidateRect(hh, 0, true);
				return 1;
			}
		}
		return 0;
	}

	if (mm == WM_SIZE)
	{
		AX* ax = (AX*)GetWindowLongPtr(hh, GWLP_USERDATA);
		if (!ax)
			return 0;
		if (!ax->OleObject)
			return 0;

		// Send Notification to parent
		NMHDR nh = { 0 };
		nh.code = AXN_SIZE;
		nh.hwndFrom = hh;
		nh.idFrom = GetWindowLong(hh, GWL_ID);
		SendMessage(ax->Site.Parent, WM_NOTIFY, 0, (LPARAM)&nh);
		DefWindowProc(hh, mm, ww, ll);
		/*
		if (ax->Site.InPlace == true)
		{
		SendMessage(hh,AX_INPLACE,0,0);
		InvalidateRect(hh,0,true);
		SendMessage(hh,AX_INPLACE,1,0);
		}
		*/

		IOleInPlaceObject* pl;
		ax->OleObject->QueryInterface(IID_IOleInPlaceObject, (void**)&pl);
		if (!pl)
			return 0;
		RECT r;
		GetClientRect(ax->Site.Window, &r);
		pl->SetObjectRects(&r, &r);
		pl->Release();
		return 0;
	}

	if (mm == WM_GETMINMAXINFO)
		return 0;

	if(mm == WM_NCCALCSIZE)
		return 0;
	
	/*
	if(mm == WM_NCCREATE)
	{
		CREATESTRUCT* createStruct = (CREATESTRUCT*)ll;
		return true;
	}
	*/

	return DefWindowProc(hh, mm, ww, ll);
}


int AXConnectObject(IUnknown* Container, REFIID riid, IUnknown* Advisor, IConnectionPointContainer** picpc, IConnectionPoint** picp)
{
	HRESULT hr = 0;
	unsigned long tid = 0;
	IConnectionPointContainer* icpc = 0;
	IConnectionPoint* icp = 0;
	*picpc = 0;
	*picp = 0;

	Container->QueryInterface(IID_IConnectionPointContainer, (void**)&icpc);
	if (icpc)
	{
		*picpc = icpc;
		icpc->FindConnectionPoint(riid, &icp);
		if (icp)
		{
			*picp = icp;
			hr = icp->Advise(Advisor, &tid);
			icp->Release();
		}
		icpc->Release();
	}
	return tid;
}

void AXDisconnectObject(IConnectionPointContainer* icpc, IConnectionPoint* icp, unsigned int Cookie)
{
	unsigned long hr = 0;
	icp->Unadvise(Cookie);
	icp->Release();
	icpc->Release();
}



// Registration function
ATOM AXRegister()
{
	WNDCLASSEXW wC = { 0 };
	wC.cbSize = sizeof(wC);


	wC.style = CS_GLOBALCLASS | CS_DBLCLKS;
	wC.lpfnWndProc = AXWndProc;
	wC.cbWndExtra = 4;
	wC.hInstance = GetModuleHandle(0);
	wC.lpszClassName = L"AX";

	return RegisterClassExW(&wC);
}

