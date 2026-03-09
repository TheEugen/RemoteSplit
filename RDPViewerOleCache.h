#pragma once

#include <Windows.h>
#include <rdpencomapi.h>

class RDPViewerOleCache : public IRDPSRAPIViewer, IOleCache
{


public:

	// IOleCache Methods
	HRESULT STDMETHODCALLTYPE Cache(
		/* [unique][in] */ __RPC__in_opt FORMATETC* pformatetc,
		/* [in] */ DWORD advf,
		/* [out] */ __RPC__out DWORD* pdwConnection);

	HRESULT STDMETHODCALLTYPE Uncache(
		/* [in] */ DWORD dwConnection);

	HRESULT STDMETHODCALLTYPE EnumCache(
		/* [out] */ __RPC__deref_out_opt IEnumSTATDATA** ppenumSTATDATA);

	HRESULT STDMETHODCALLTYPE InitCache(
		/* [unique][in] */ __RPC__in_opt IDataObject* pDataObject);

	HRESULT STDMETHODCALLTYPE SetData(
		/* [unique][in] */ __RPC__in_opt FORMATETC* pformatetc,
		/* [unique][in] */ __RPC__in_opt STGMEDIUM* pmedium,
		/* [in] */ BOOL fRelease);
};

// IOleCache Methods
HRESULT _stdcall RDPViewerOleCache::Cache(
	/* [unique][in] */ __RPC__in_opt FORMATETC* pformatetc,
	/* [in] */ DWORD advf,
	/* [out] */ __RPC__out DWORD* pdwConnection) {
	return E_NOTIMPL;
}
HRESULT _stdcall RDPViewerOleCache::Uncache(
	/* [in] */ DWORD dwConnection) {
	return E_NOTIMPL;
}
HRESULT _stdcall RDPViewerOleCache::EnumCache(
	/* [out] */ __RPC__deref_out_opt IEnumSTATDATA** ppenumSTATDATA) {
	return E_NOTIMPL;
}
HRESULT _stdcall RDPViewerOleCache::InitCache(
	/* [unique][in] */ __RPC__in_opt IDataObject* pDataObject) {
	return E_NOTIMPL;
}
HRESULT _stdcall RDPViewerOleCache::SetData(
	/* [unique][in] */ __RPC__in_opt FORMATETC* pformatetc,
	/* [unique][in] */ __RPC__in_opt STGMEDIUM* pmedium,
	/* [in] */ BOOL fRelease) {
	return E_NOTIMPL;
}