
// RelaxTimeDlg.h: 头文件
//

#pragma once


// CRelaxTimeDlg 对话框
class CRelaxTimeDlg : public CDialogEx
{
// 构造
public:
	CRelaxTimeDlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_RELAXTIME_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CString E_Day;
	CListCtrl m_list;
	afx_msg void OnBnClickedAdday();
	afx_msg void OnBnClickedSbday();
//	afx_msg void OnEnChangeDay();
	afx_msg void OnEnSetfocusDay();

	bool CheckTextIsNotNull();
	

public:
	
	
	 const CString TextValue1 = _T("请输入要修改的天数");
	const  CString TextValue2 = _T("请选择要修改的序号");
	void InitListCtrl();
	void closedb();
	CString GetTime(CString CBText);
	void ConnectDB();
	void chushihua();
	BOOL GetcbCurSel();
	BOOL ChangeDataTime(CString NewData,CString Times);

	CString GetSystemTime();
	


	_ConnectionPtr m_pConnection; //连接access数据库的链接对象 


	CString strCBText;

	CComboBox m_cb1;
	afx_msg void OnTimer(UINT_PTR nIDEvent);
};
