// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#include "ExternalObject.h"

class CMainDlg : public CAxDialogImpl<CMainDlg>,
				 public IDispEventImpl<IDC_IE,CMainDlg>
{
public:
	enum { IDD = IDD_MAINDLG };

	BEGIN_MSG_MAP(CMainDlg)
		CHAIN_MSG_MAP(CAxDialogImpl<CMainDlg>)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
		MSG_WM_CLOSE(OnClose)
		COMMAND_ID_HANDLER_EX(IDB_GO,		OnGo)
		COMMAND_ID_HANDLER_EX(IDB_BACK,		OnBack)
		COMMAND_ID_HANDLER_EX(IDB_FORWARD,	OnForward)
		COMMAND_ID_HANDLER_EX(IDB_STOP,		OnStop)
		COMMAND_ID_HANDLER_EX(IDB_REFRESH,	OnRefresh)
		COMMAND_ID_HANDLER_EX(IDB_ABOUT,	OnAbout)
		COMMAND_ID_HANDLER_EX(IDB_SCRIPT1,	OnScript1)
		COMMAND_ID_HANDLER_EX(IDB_SCRIPT2,	OnScript2)
	END_MSG_MAP()

	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	void OnClose();

	void OnGo(UINT uNotifyCode, int nID, CWindow wndCtl);
	void OnBack(UINT uNotifyCode, int nID, CWindow wndCtl);
	void OnForward(UINT uNotifyCode, int nID, CWindow wndCtl);
	void OnStop(UINT uNotifyCode, int nID, CWindow wndCtl);
	void OnRefresh(UINT uNotifyCode, int nID, CWindow wndCtl);
	void OnAbout(UINT uNotifyCode, int nID, CWindow wndCtl);
	void OnScript1(UINT uNotifyCode, int nID, CWindow wndCtl);
	void OnScript2(UINT uNotifyCode, int nID, CWindow wndCtl);

private:
	CComPtr<IWebBrowser2> m_spWebBrowser;
	CExternalObject	m_ExternalObject;

	void InitIE();
	void NavigateTo(LPCWSTR pszUrl);
	void ExecScript1(LPCWSTR pszJs);
	void ExecScript2();

public:
	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY(IDC_IE, DISPID_DOWNLOADBEGIN, OnDownloadBegin)
		SINK_ENTRY(IDC_IE, DISPID_DOWNLOADCOMPLETE, OnDownloadComplete)
		SINK_ENTRY(IDC_IE, DISPID_BEFORENAVIGATE2, OnBeforeNavigate)
		SINK_ENTRY(IDC_IE, DISPID_NAVIGATECOMPLETE2, OnNavigateComplete)
		SINK_ENTRY(IDC_IE, DISPID_DOCUMENTCOMPLETE, OnDocumentComplete)
	END_SINK_MAP()
	void __stdcall OnDownloadBegin();
	void __stdcall OnDownloadComplete();
	void __stdcall OnBeforeNavigate(LPDISPATCH pDisp, VARIANT* URL, VARIANT* Flags, VARIANT* TargetFrameName, VARIANT* PostData, VARIANT* Headers, BOOL* Cancel);
	void __stdcall OnNavigateComplete(LPDISPATCH pDisp, VARIANT* URL);
	void __stdcall OnDocumentComplete(LPDISPATCH pDisp, VARIANT* URL);
};
