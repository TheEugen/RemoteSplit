#pragma once

#include <Windows.h>

#include <vector>
#include <string>

#include "Resource.h"

#include "associated.h"

static void outputLog(HWND hWndLog, LPCWSTR text)
{
	int len = GetWindowTextLength(hWndLog);
	SendMessage(hWndLog, EM_SETSEL, (WPARAM)len, (LPARAM)len);
	SendMessage(hWndLog, EM_REPLACESEL, 0, (LPARAM)text);
}

static bool changeToServerStyle(HINSTANCE hInstance, HWND hWnd)
{
	// hide server and client buttons
	ShowWindow(GetDlgItem(hWnd, IDC_BTN_SERVER), SW_HIDE);
	ShowWindow(GetDlgItem(hWnd, IDC_BTN_CLIENT), SW_HIDE);

	// create new buttons
	CreateWindow(L"button", L"Start Sharingsession", WS_VISIBLE | WS_CHILD | WS_BORDER, 50, 100, 100, 50, hWnd, (HMENU)IDC_BTN_SERVER_STARTSHARING, hInstance, nullptr);
	CreateWindow(L"button", L"Stop Sharingsession", WS_VISIBLE | WS_CHILD | WS_BORDER, 200, 100, 100, 50, hWnd, (HMENU)IDC_BTN_SERVER_STOPSHARING, hInstance, nullptr);

	// create info log box
	CreateWindow(L"edit", 0, WS_CHILD | WS_VISIBLE | ES_MULTILINE | WS_VSCROLL | ES_AUTOVSCROLL, 20, 150, 350, 150, hWnd, (HMENU)IDC_OUTPUTLOG, hInstance, 0);

	int len = GetWindowTextLength(GetDlgItem(hWnd, IDC_OUTPUTLOG));
	SendMessage(GetDlgItem(hWnd, IDC_OUTPUTLOG), EM_SETSEL, (WPARAM)len, (LPARAM)len);
	SendMessage(GetDlgItem(hWnd, IDC_OUTPUTLOG), EM_REPLACESEL, 0, (LPARAM)L"Test123");

	return true;
}

static bool changeToClientStyle(HINSTANCE hInstance, HWND hWnd)
{
	// resize mainwindow
	// TODO: define window size in Resources.h
	SetWindowPos(hWnd, HWND_TOP, 100, 100, 1200, 800, SWP_SHOWWINDOW);

	// hide server and client buttons
	ShowWindow(GetDlgItem(hWnd, IDC_BTN_SERVER), SW_HIDE);
	ShowWindow(GetDlgItem(hWnd, IDC_BTN_CLIENT), SW_HIDE);

	// move close button
	SetWindowPos(GetDlgItem(hWnd, IDC_BTN_CLOSE), HWND_TOP, 1050, 700, 100, 50, SWP_SHOWWINDOW);

	// create new buttons
	CreateWindow(L"button", L"Connect", WS_VISIBLE | WS_CHILD | WS_BORDER, 50, 700, 100, 50, hWnd, (HMENU)IDC_BTN_CLIENT_CONNECT, hInstance, nullptr);
	CreateWindow(L"button", L"Disconnect", WS_VISIBLE | WS_CHILD | WS_BORDER, 200, 700, 100, 50, hWnd, (HMENU)IDC_BTN_CLIENT_DISCONNECT, hInstance, nullptr);

	// create info log box
	CreateWindow(L"edit", 0, WS_CHILD | WS_VISIBLE | ES_MULTILINE | WS_VSCROLL | ES_AUTOVSCROLL, 350, 650, 350, 150, hWnd, (HMENU)IDC_OUTPUTLOG, hInstance, 0);

	// create stream window
	CreateWindow(L"edit", nullptr, WS_CLIPCHILDREN | WS_CHILD | WS_VISIBLE, 0, 0, 1200, 600, hWnd, (HMENU)IDC_VIEW_CLIENT_STREAM, hInstance, nullptr);
	SendMessageA(GetDlgItem(hWnd, IDC_VIEW_CLIENT_STREAM), EM_SETREADONLY, 1, 0);

	return true;
}

static void PutStreamOnWindow(HINSTANCE hInstance, HWND hWnd, IRDPSRAPIViewer* viewer) {
	// While looking at the calls made I have never seen this create window called it is always returns 0 and is recreated
	HWND ACTIVEX_WINDOW = CreateWindowEx(0, L"AX", L"{32be5ed2-5c86-480f-a914-0ff8885a1b3f}", WS_CHILD | WS_VISIBLE, 0, 20, 1200, 600, hWnd, 0, hInstance, 0);
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
		ACTIVEX_WINDOW = CreateWindowA("AX", "}32BE5ED2-5C86-480F-A914-0FF8885A1B3F}", WS_CHILD | WS_VISIBLE, 0, 0, 1200, 600, hWnd, 0, hInstance, 0);
		if (!ACTIVEX_WINDOW)
		{
			outputLog(GetDlgItem(GetParent(hWnd), IDC_OUTPUTLOG), L"ActiveX window creating failed with error code: ");
			outputLog(GetDlgItem(GetParent(hWnd), IDC_OUTPUTLOG), std::to_wstring(GetLastError()).c_str());
		}
		SendMessageA(ACTIVEX_WINDOW, AX_RECREATE, 0, (LPARAM)viewer);
	}

	SendMessageA(ACTIVEX_WINDOW, AX_INPLACE, 1, 0);
	ShowWindow(ACTIVEX_WINDOW, SW_MAXIMIZE);
}

struct Variant : VARIANT
{
	Variant(bool b) { VARIANT::boolVal = b ? VARIANT_TRUE : VARIANT_FALSE; VARIANT::vt = VT_BOOL; }
	Variant(int i) { VARIANT::intVal = i; VARIANT::vt = VT_I4; }
};


class BStr
{
public:
	BStr() : data(nullptr) {}
	BStr(const wchar_t* str) : data(::SysAllocString(str)) {}
	BStr(const std::string& s) :data(nullptr) { FromUtf8(s); }
	BStr& operator=(const std::string& s) { return FromUtf8(s); }
	~BStr() { ::SysFreeString(data); }
	operator BSTR() { return data; }
	BSTR* operator&() { ::SysFreeString(data); data = nullptr; return &data; }
	bool operator==(const wchar_t* p) const { return std::wstring(p) == data; }
	std::string ToUtf8() const
	{
		std::string s;
		if (data)
		{
			s.resize(::WideCharToMultiByte(CP_UTF8, 0, data, -1, 0, 0, 0, 0));
			::WideCharToMultiByte(CP_UTF8, 0, data, -1, const_cast<char*>(s.data()), s.size(), 0, 0);
			if (s.size() > 0)
				s.resize(s.size() - 1);
		}
		return s;
	}
	BStr& FromUtf8(const std::string& s)
	{
		::SysFreeString(data);
		std::vector<wchar_t> buf(::MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, 0, 0));
		::MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, buf.data(), buf.size());
		data = ::SysAllocString(buf.data());
		return *this;
	}

	unsigned int strSize()
	{
		return this->ToUtf8().size();
	}

private:
	BSTR data;
};