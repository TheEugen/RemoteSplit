#pragma once

#include <Windows.h>
#include <atlbase.h>
#include <rdpencomapi.h>

#include "Utils.h"



static void onViewerConnected(IDispatch* viewer)
{

}

class RDPSessionEvents : public _IRDPSessionEvents
{
private:

	int m_refCount = 0;

	HWND m_outputLog;

protected:

	virtual void OnViewerConnected() {}


public:

	void OnAttendeeConnected(IDispatch* pAttendee)
	{
		IRDPSRAPIAttendee* pRDPAtendee;
		pAttendee->QueryInterface(__uuidof(IRDPSRAPIAttendee), (void**)&pRDPAtendee);
		pRDPAtendee->put_ControlLevel(CTRL_LEVEL::CTRL_LEVEL_INTERACTIVE);
		//AddText(infoLog, L"An attendee connected!\r\n");
		outputLog(m_outputLog, L"A viewer connected!\n");
	}

	void OnAttendeeDisconnected(IDispatch* pAttendee)
	{
		IRDPSRAPIAttendeeDisconnectInfo* info;
		ATTENDEE_DISCONNECT_REASON reason;
		pAttendee->QueryInterface(__uuidof(IRDPSRAPIAttendeeDisconnectInfo), (void**)&info);
		if (info->get_Reason(&reason) == S_OK) {
			//char *textReason = NULL;
			const wchar_t* textReason;
			switch (reason) {
			case ATTENDEE_DISCONNECT_REASON_APP:
				textReason = L"Viewer terminated session!\n";
				break;
			case ATTENDEE_DISCONNECT_REASON_ERR:
				textReason = L"Internal Error!\n";
				break;
			case ATTENDEE_DISCONNECT_REASON_CLI:
				textReason = L"Attendee requested termination!\n";
				break;
			default:
				textReason = L"Unknown reason!\n";
			}
			//AddText(infoLog, dupcat("Attendee disconnected\r\n   Reason: ", textReason, "\r\n", 0));
			outputLog(m_outputLog, L"Viewer disconnected | Reason: ");
			outputLog(m_outputLog, textReason);

		}
		pAttendee->Release();
		//picp = 0;
		//picpc = 0;
		//disconnect();
	}

	void OnControlLevelChangeRequest(IDispatch* pAttendee, CTRL_LEVEL RequestedLevel)
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


	void setOutputLog(HWND outputLog) { m_outputLog = outputLog; }

	void OnViewerConnectFailed()
	{
		outputLog(m_outputLog, L"Viewer tried to connect and failed!\n");
	}

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
		case DISPID_RDPSRAPI_EVENT_ON_VIEWER_CONNECTFAILED:
			OnViewerConnectFailed();
		}
		return S_OK;
	}
	
};