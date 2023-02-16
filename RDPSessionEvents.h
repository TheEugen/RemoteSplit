#pragma once

#include <Windows.h>
#include <atlbase.h>
#include <rdpencomapi.h>

static void OnAttendeeConnected(IDispatch* pAttendee) {
	IRDPSRAPIAttendee* pRDPAtendee;
	pAttendee->QueryInterface(__uuidof(IRDPSRAPIAttendee), (void**)&pRDPAtendee);
	pRDPAtendee->put_ControlLevel(CTRL_LEVEL::CTRL_LEVEL_VIEW);
	//AddText(infoLog, L"An attendee connected!\r\n");
}

static void OnAttendeeDisconnected(IDispatch* pAttendee) {
	IRDPSRAPIAttendeeDisconnectInfo* info;
	ATTENDEE_DISCONNECT_REASON reason;
	pAttendee->QueryInterface(__uuidof(IRDPSRAPIAttendeeDisconnectInfo), (void**)&info);
	if (info->get_Reason(&reason) == S_OK) {
		//char *textReason = NULL;
		const char* textReason;
		switch (reason) {
		case ATTENDEE_DISCONNECT_REASON_APP:
			textReason = "Viewer terminated session!";
			break;
		case ATTENDEE_DISCONNECT_REASON_ERR:
			textReason = "Internal Error!";
			break;
		case ATTENDEE_DISCONNECT_REASON_CLI:
			textReason = "Attendee requested termination!";
			break;
		default:
			textReason = "Unknown reason!";
		}
		//AddText(infoLog, dupcat("Attendee disconnected\r\n   Reason: ", textReason, "\r\n", 0));

	}
	pAttendee->Release();
	//picp = 0;
	//picpc = 0;
	//disconnect();
}

static void OnControlLevelChangeRequest(IDispatch* pAttendee, CTRL_LEVEL RequestedLevel)
{	
	IRDPSRAPIAttendee* pRDPAtendee;
	pAttendee->QueryInterface(__uuidof(IRDPSRAPIAttendee), (void**)&pRDPAtendee);
	if (pRDPAtendee->put_ControlLevel(RequestedLevel) == S_OK)
	{
		/*
		switch (RequestedLevel)
		{
		case CTRL_LEVEL_NONE:
			AddText(infoLog, L"Level changed to CTRL_LEVEL_NONE!\r\n");
			break;
		case CTRL_LEVEL_VIEW:
			AddText(infoLog, L"Level changed to CTRL_LEVEL_VIEW!\r\n");
			break;
		case CTRL_LEVEL_INTERACTIVE:
			AddText(infoLog, L"Level changed to CTRL_LEVEL_INTERACTIVE!\r\n");
			break;
		}
		*/
	}
}

class RDPSessionEvents : public _IRDPSessionEvents
{
private:

	int m_refCount = 0;
protected:

	virtual void OnViewerConnected() {}

public:

	HRESULT __stdcall QueryInterface(REFIID riid, void** pp) override
	{
		// COPIED
		if (!pp)
			return E_POINTER;
		*pp = nullptr;
		if (IsEqualIID(riid, DIID__IRDPSessionEvents))
			*pp = static_cast<IDispatch*>(this);
		if (IsEqualIID(riid, IID_IDispatch))
			*pp = static_cast<IDispatch*>(this);
		if (IsEqualIID(riid, IID_IUnknown))
			*pp = static_cast<IUnknown*>(this);
		if (!(*pp))
			return E_NOINTERFACE;
		this->AddRef();
		return S_OK;
	}

	ULONG __stdcall AddRef() override { ++m_refCount; return 0; }
	ULONG __stdcall Release() override { if (--m_refCount == 0) delete this; return 0; }
	HRESULT __stdcall GetTypeInfoCount(UINT*) override { return E_NOTIMPL; }
	HRESULT __stdcall GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
	HRESULT __stdcall GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override { return E_NOTIMPL; }
	HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID, LCID, WORD, DISPPARAMS* dispParams, VARIANT*, EXCEPINFO*, UINT*)
	{
		// COPIED
		switch (dispIdMember)
		{
		case DISPID_RDPSRAPI_EVENT_ON_ATTENDEE_CONNECTED:
			OnAttendeeConnected(dispParams->rgvarg[0].pdispVal);
			break;
		case DISPID_RDPSRAPI_EVENT_ON_ATTENDEE_DISCONNECTED:
			OnAttendeeDisconnected(dispParams->rgvarg[0].pdispVal);
			break;
		case DISPID_RDPSRAPI_EVENT_ON_CTRLLEVEL_CHANGE_REQUEST:
			OnControlLevelChangeRequest(dispParams->rgvarg[1].pdispVal, (CTRL_LEVEL)dispParams->rgvarg[0].intVal);
			break;
		}
		return S_OK;
	}
	
};