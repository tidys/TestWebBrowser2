#pragma once

static _ATL_FUNC_INFO Info_DownloadBegin		= {CC_STDCALL, VT_EMPTY, 0};
static _ATL_FUNC_INFO Info_OnDownloadComplete	= {CC_STDCALL, VT_EMPTY, 0};

class CAboutDlg : public CDialogImpl<CAboutDlg>,
				  public IDispEventSimpleImpl<1, CAboutDlg, &DIID_DWebBrowserEvents2>
{
public:
	enum {IDD=IDD_ABOUT};

	BEGIN_MSG_MAP_EX(CAboutDlg)
		MSG_WM_INITDIALOG(OnInitDialog)
		MSG_WM_DESTROY(OnDestroy)
		MSG_WM_CLOSE(OnClose)
		COMMAND_ID_HANDLER_EX(IDOK, OnOK)
	END_MSG_MAP()

	BOOL OnInitDialog(CWindow wndFocus, LPARAM lInitParam);
	void OnDestroy();
	void OnClose();
	void OnOK(UINT uNotifyCode, int nID, CWindow wndCtl);

private:
	CAxWindow		m_wndIE;

public:
	BEGIN_SINK_MAP(CAboutDlg)
		SINK_ENTRY_INFO(1, DIID_DWebBrowserEvents2, DISPID_DOWNLOADBEGIN, OnDownloadBegin, &Info_DownloadBegin)
		SINK_ENTRY_INFO(1, DIID_DWebBrowserEvents2, DISPID_DOWNLOADCOMPLETE, OnDownloadComplete, &Info_OnDownloadComplete)
	END_SINK_MAP()
	void __stdcall OnDownloadBegin();
	void __stdcall OnDownloadComplete();
};
