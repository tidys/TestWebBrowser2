// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include "AboutDlg.h"

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// center the dialog on the screen
	CenterWindow();

	// set icons
	HICON hIcon = (HICON)::LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), 
		IMAGE_ICON, ::GetSystemMetrics(SM_CXICON), ::GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR);
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = (HICON)::LoadImage(_Module.GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), 
		IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR);
	SetIcon(hIconSmall, FALSE);

	InitIE();
	NavigateTo(L"http://www.baidu.com");
	return TRUE;
}

void CMainDlg::OnClose()
{
	EndDialog(0);
	return;
}

void CMainDlg::OnGo(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	WCHAR szUrl[MAX_PATH] = {0};
	if (GetDlgItemText(IDE_URL, szUrl, MAX_PATH))
	{
		NavigateTo(szUrl);
	}
}

void CMainDlg::OnBack(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	if (m_spWebBrowser)
	{
		m_spWebBrowser->GoBack();
	}
}

void CMainDlg::OnForward(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	if (m_spWebBrowser)
	{
		m_spWebBrowser->GoForward();
	}
}

void CMainDlg::OnStop(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	if (m_spWebBrowser)
	{
		m_spWebBrowser->Stop();
	}
}

void CMainDlg::OnRefresh(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	CComVariant vLevel;

	if (m_spWebBrowser)
	{
		vLevel = REFRESH_COMPLETELY;
		m_spWebBrowser->Refresh2(&vLevel);
	}
}

void CMainDlg::OnAbout(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	CAboutDlg dlg;
	dlg.DoModal();
}

void CMainDlg::OnScript1(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	ExecScript1(L"window.alert(\'ExecScript1\')");
}

void CMainDlg::OnScript2(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	ExecScript2();
}

void CMainDlg::NavigateTo(LPCWSTR pszUrl)
{
	if (m_spWebBrowser)
	{
		CComVariant v;
		m_spWebBrowser->Navigate(CComBSTR(pszUrl), &v, &v, &v, &v);
	}
}
void __stdcall CMainDlg::OnDownloadBegin()
{
	OutputDebugString(L"OnDownloadBegin");
}

void __stdcall CMainDlg::OnDownloadComplete()
{
	OutputDebugString(L"OnDownloadComplete");
}

void __stdcall CMainDlg::OnBeforeNavigate(LPDISPATCH pDisp, VARIANT* URL, VARIANT* Flags, VARIANT* TargetFrameName, VARIANT* PostData, VARIANT* Headers, BOOL* Cancel)
{
	if (URL)
	{
		CString strUrl;
		strUrl.Format(L"OnBeforeNavigate %s", URL->bstrVal);
		OutputDebugString(strUrl);

		//可指定Cancel来阻止访问
		if (strUrl.CompareNoCase(L"OnBeforeNavigate http://www.sina.com/")==0)
		{
			*Cancel = TRUE;
			OutputDebugString(L"http://www.sina.com/ is Canceled!");

		}
	}
}

void __stdcall CMainDlg::OnNavigateComplete(LPDISPATCH pDisp, VARIANT* URL)
{
	if (URL)
	{
		CString strUrl;
		strUrl.Format(L"OnNavigateComplete %s", URL->bstrVal);
		OutputDebugString(strUrl);
	}
}

void __stdcall CMainDlg::OnDocumentComplete(LPDISPATCH pDisp, VARIANT* URL)
{
	if (URL)
	{
		CString strUrl;
		strUrl.Format(L"OnDocumentCompleteIe %s", URL->bstrVal);
		OutputDebugString(strUrl);
	}
}

void CMainDlg::InitIE()
{
	//查询得到IWebBrowser2接口
	CAxWindow wndIE = GetDlgItem(IDC_IE);
	if (wndIE.IsWindow())
	{
		wndIE.QueryControl(&m_spWebBrowser);
	}

	//设置浏览器内容回调接口
	wndIE.SetExternalDispatch((IDispatch*)(&m_ExternalObject));
}

void CMainDlg::ExecScript1(LPCWSTR pszJs)
{
	if (!m_spWebBrowser)
	{
		return;
	}

	CComPtr<IDispatch> spDispDoc;
	m_spWebBrowser->get_Document(&spDispDoc);

	if(!spDispDoc) 
	{
		return;
	}

	CComPtr<IHTMLDocument2> spHtmlDoc;
	CComPtr<IHTMLWindow2> spHtmlWindow;
	HRESULT hr = spDispDoc->QueryInterface(IID_IHTMLDocument2,(void**)&spHtmlDoc);
	if (spHtmlDoc)
	{
		if (SUCCEEDED(spHtmlDoc->get_parentWindow(&spHtmlWindow)) && spHtmlWindow)
		{
			CComBSTR bstrJs  = pszJs;
			CComBSTR bstrlan = L"javascript";
			VARIANT varRet;
			varRet.vt = VT_EMPTY;

			spHtmlWindow->execScript(bstrJs, bstrlan, &varRet);
		}
	}
}

void CMainDlg::ExecScript2()
{
	CComPtr<IDispatch> spDisp;
	HRESULT hr = m_spWebBrowser->get_Document(&spDisp);
	if (SUCCEEDED(hr))
	{
		CComQIPtr<IHTMLDocument2> spDoc2 = spDisp;
		if (spDoc2)
		{
			CComDispatchDriver spScript;
			hr = spDoc2->get_Script(&spScript);
			if (SUCCEEDED(hr))
			{
// 				{
// 					CComVariant varRet;  
// 					spScript.Invoke0(L"test1", &varRet);  
// 					int a = 10;
// 				}

				//--Add1
				{
					CComVariant var1 = 10, var2 = 20, varRet;  
					spScript.Invoke2(L"Add1", &var1, &var2, &varRet);  

					CString strVal;
					strVal.Format(L"%d", varRet.intVal);
					OutputDebugString(strVal.GetBuffer(0));
				}

				//--Add2
				{
					CComVariant var1 = 10, var2 = 20, varRet;  
					spScript.Invoke2(L"Add2", &var1, &var2, &varRet);  

					CComDispatchDriver spArray = varRet.pdispVal;  

					//获取数组中元素个数，这个length在JS中是Array对象的属性
					CComVariant varArrayLen;  
					spArray.GetPropertyByName(L"length", &varArrayLen); 

					//获取数组中第0,1,2个元素的值：  
					CComVariant varValue[3];  
					spArray.GetPropertyByName(L"0", &varValue[0]);  
					spArray.GetPropertyByName(L"1", &varValue[1]);  
					spArray.GetPropertyByName(L"2", &varValue[2]);  

					CString strVal;
					strVal.Format(L"%d %d %d", varValue[0].intVal,
											   varValue[1].intVal,
											   varValue[2].intVal);
					OutputDebugString(strVal.GetBuffer(0));
				}

				//--Add3
				{
					CComVariant var1 = 10, var2 = 20, varRet;  
					spScript.Invoke2(L"Add3", &var1, &var2, &varRet);  
					CComDispatchDriver spData = varRet.pdispVal;  

					CComVariant varValue1, varValue2;  
					spData.GetPropertyByName(L"result", &varValue1);  
					spData.GetPropertyByName(L"str", &varValue2); 

					CString strVal;
					strVal.Format(L"%d %s", varValue1.intVal, varValue2.bstrVal);
					OutputDebugString(strVal.GetBuffer(0));
				}
			}
		}
	}
}
