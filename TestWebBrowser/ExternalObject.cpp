#include "StdAfx.h"
#include "ExternalObject.h"

CExternalObject::CExternalObject(void)
{
	m_ExternalList[L"SayHello"]		= EXTFUNC_ID_HELLO;
	m_ExternalList[L"SayGoodBye"]	= EXTFUNC_ID_GOODBYE;
}

CExternalObject::~CExternalObject(void)
{
}

//IUnknown
HRESULT STDMETHODCALLTYPE CExternalObject::QueryInterface(/* [in] */ REFIID riid, 
														  /* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject)
{
	if (IsEqualIID(riid, IID_IUnknown))
	{
		*ppvObject = this;
	}
	else if (IsEqualIID(riid, IID_IDispatch))
	{
		*ppvObject = (IDispatch*)this;
	}
	else
	{
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}

	return S_OK;
}

ULONG STDMETHODCALLTYPE CExternalObject::AddRef(void)
{
	return 1;
}

ULONG STDMETHODCALLTYPE CExternalObject::Release(void)
{
	return 1;
}

//IDispatch
HRESULT STDMETHODCALLTYPE CExternalObject::GetTypeInfoCount(/* [out] */ __RPC__out UINT *pctinfo)
{
	*pctinfo = 0;   //没有类型库  
	return S_OK;  
}

HRESULT STDMETHODCALLTYPE CExternalObject::GetTypeInfo(/* [in] */ UINT iTInfo, 
													   /* [in] */ LCID lcid, 
													   /* [out] */ __RPC__deref_out_opt ITypeInfo **ppTInfo)
{
	*ppTInfo = NULL;  
	return E_NOTIMPL;  
}

HRESULT STDMETHODCALLTYPE CExternalObject::GetIDsOfNames(/* [in] */ __RPC__in REFIID riid, 
														 /* [size_is][in] */ __RPC__in_ecount_full(cNames) LPOLESTR *rgszNames, 
														 /* [range][in] */ UINT cNames, 
														 /* [in] */ LCID lcid, 
														 /* [size_is][out] */ __RPC__out_ecount_full(cNames) DISPID *rgDispId)
{
	for (UINT i=0; i<cNames; i++)  
	{  
		map<CString, UINT>::iterator iter = m_ExternalList.find(rgszNames[i]);  
		if ( m_ExternalList.end() != iter )  
		{  
			rgDispId[i] = iter->second;  
		}  
		else  
		{  
			rgDispId[i] = DISPID_UNKNOWN;  
		}  
	}  

	return S_OK;
}

HRESULT STDMETHODCALLTYPE CExternalObject::Invoke(	/* [in] */ DISPID dispIdMember, 
													/* [in] */ REFIID riid, 
													/* [in] */ LCID lcid, 
													/* [in] */ WORD wFlags, 
													/* [out][in] */ DISPPARAMS *pDispParams, 
													/* [out] */ VARIANT *pVarResult, 
													/* [out] */ EXCEPINFO *pExcepInfo, 
													/* [out] */ UINT *puArgErr)
{
	if (0==dispIdMember ||  
		(dispIdMember!=EXTFUNC_ID_HELLO && dispIdMember!=EXTFUNC_ID_GOODBYE) ||  
		0==(DISPATCH_METHOD&wFlags || DISPATCH_PROPERTYGET&wFlags))  
	{  
		return E_NOTIMPL;  
	}

	if (pVarResult)  
	{  
		CComVariant var(true);  
		*pVarResult = var;  
	}  

	//判断属性
	if (DISPATCH_PROPERTYGET&wFlags)
	{
		pVarResult->vt = VT_DISPATCH;//object
		pVarResult->pdispVal = this;
		return S_OK;
	}

	USES_CONVERSION;  

	//调用本地方法
	switch (dispIdMember)  
	{  
	case EXTFUNC_ID_HELLO:  
		if (pDispParams &&								//参数数组有效  
			pDispParams->cArgs==1 &&					//参数个数为1  
			pDispParams->rgvarg[0].vt==VT_BSTR &&		//参数类型满足  
			pDispParams->rgvarg[0].bstrVal)				//参数值有效  
		{  
			CString strVal(OLE2T(pDispParams->rgvarg[0].bstrVal));  
			AtlMessageBox(NULL, strVal.GetBuffer(0), L"SayHello");
		}  
		break;  

	case EXTFUNC_ID_GOODBYE:  
		if (pDispParams &&								//参数数组有效  
			pDispParams->cArgs==2 &&					//参数个数为2 
			pDispParams->rgvarg[1].vt==VT_I4 &&	
			pDispParams->rgvarg[0].vt==VT_BSTR &&		//参数类型满足  
			pDispParams->rgvarg[1].bstrVal &&
			pDispParams->rgvarg[0].bstrVal)				//参数值有效  
		{  
			CString strVal;
			strVal.Format(L"%d %s", pDispParams->rgvarg[1].bstrVal, 
									pDispParams->rgvarg[0].bstrVal);
			AtlMessageBox(NULL, strVal.GetBuffer(0), L"SayGoodBye");
		}  
		break;  
	}  

	return S_OK;  
}