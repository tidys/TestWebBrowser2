#include "StdAfx.h"

#include "resource.h"
#include "AboutDlg.h"

BOOL CAboutDlg::OnInitDialog(CWindow wndFocus, LPARAM lInitParam)
{
	CRect rcClient;
	GetClientRect(&rcClient);
	rcClient.bottom -= 50;

	//创建包容窗口
	m_wndIE.Create(m_hWnd, rcClient, L"", 
				   WS_CHILD|WS_VISIBLE|WS_CLIPCHILDREN|WS_CLIPSIBLINGS);

	//创建ActiveX控件
	CComPtr<IUnknown> punkCtrl;
	m_wndIE.CreateControlEx(L"Shell.Explorer", NULL, NULL, &punkCtrl);

	//查询接口导航到指定url
	if (punkCtrl)
	{
		CComQIPtr<IWebBrowser2> pWB2 = punkCtrl;
		CComVariant v;

		if (pWB2)
		{
			this->DispEventAdvise(punkCtrl, &__uuidof(DWebBrowserEvents2));
			pWB2->Navigate(CComBSTR(L"http://jimwen.net"), &v, &v, &v, &v);
		}
	}

	return 0;
}

void CAboutDlg::OnDestroy()
{
	CComPtr<IUnknown> punkCtrl;
	m_wndIE.QueryControl(&punkCtrl);
	this->DispEventUnadvise(punkCtrl, &__uuidof(DWebBrowserEvents2));
}

void CAboutDlg::OnClose()
{
	EndDialog(0);
}

void CAboutDlg::OnOK(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	EndDialog(0);
}

void __stdcall CAboutDlg::OnDownloadBegin()
{
	OutputDebugString(L"OnDownloadBegin ABout");
}

void __stdcall CAboutDlg::OnDownloadComplete()
{
	OutputDebugString(L"OnDownloadComplete ABout");
}
