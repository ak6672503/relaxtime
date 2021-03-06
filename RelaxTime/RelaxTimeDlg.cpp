﻿
// RelaxTimeDlg.cpp: 实现文件
//

#include "pch.h"
#include "framework.h"
#include "RelaxTime.h"
#include "RelaxTimeDlg.h"
#include "afxdialogex.h"
//#include "x64/Debug/msado15.tlh"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CRelaxTimeDlg 对话框



CRelaxTimeDlg::CRelaxTimeDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_RELAXTIME_DIALOG, pParent)
	, E_Day(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CRelaxTimeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_DAY, E_Day);
	DDX_Control(pDX, IDC_LIST1, m_list);
	DDX_Control(pDX, IDC_COMBO1, m_cb1);
}

BEGIN_MESSAGE_MAP(CRelaxTimeDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_ADday, &CRelaxTimeDlg::OnBnClickedAdday)
	ON_BN_CLICKED(IDC_SBday, &CRelaxTimeDlg::OnBnClickedSbday)
//	ON_EN_CHANGE(IDC_DAY, &CRelaxTimeDlg::OnEnChangeDay)
	ON_EN_SETFOCUS(IDC_DAY, &CRelaxTimeDlg::OnEnSetfocusDay)
	ON_WM_TIMER()
END_MESSAGE_MAP()


// CRelaxTimeDlg 消息处理程序

BOOL CRelaxTimeDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	SetTimer(1, 10000, 0);


	chushihua();
	InitListCtrl();
	



	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CRelaxTimeDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

void CRelaxTimeDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

HCURSOR CRelaxTimeDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


//增加
void CRelaxTimeDlg::OnBnClickedAdday()
{
	UpdateData();
	
	if (CheckTextIsNotNull()) {
		
		CString NeedData;
		if (GetcbCurSel()){
			//把提取出来的数值，加上要修改的数值，然后用取小数点后1位的方式变成cstring
			NeedData.Format(_T("%.1lf"), _ttof(E_Day) + _ttof(GetTime(strCBText)));
			CString NowTime = GetSystemTime();
			if(ChangeDataTime(NeedData,NowTime)){
			MessageBox(_T("数据更改成功"));
			//MessageBox(GetSystemTime());
			m_list.DeleteAllItems();
			InitListCtrl();
			}
			else {
				MessageBox(_T("数据更改失败"));
			}
		}
	}
	
	
}

//减少
void CRelaxTimeDlg::OnBnClickedSbday()
{
	UpdateData();

	if (CheckTextIsNotNull()) {

		CString NeedData;
		if (GetcbCurSel()) {
			//把提取出来的数值，加上要修改的数值，然后用取小数点后1位的方式变成cstring
			double linshi;
			linshi = _ttof(GetTime(strCBText)) - _ttof(E_Day);
			if (linshi >= 0.0) {
				NeedData.Format(_T("%.1lf"),linshi);
				CString NowTime = GetSystemTime();
				if (ChangeDataTime(NeedData, NowTime)) {
					MessageBox(_T("数据更改成功"));

					m_list.DeleteAllItems();
					InitListCtrl();
				}
				else {
					MessageBox(_T("数据更改失败"));
				}
			}
			else
			{
				MessageBox(_T("假期不够抵扣"));
			}
			
			
		}
	}
	
}
	

//文本输入框获取焦点
void CRelaxTimeDlg::OnEnSetfocusDay()
{
	GetDlgItem(IDC_DAY)->SetWindowText(_T(""));
}


//判断文本框是否输入内容
bool CRelaxTimeDlg::CheckTextIsNotNull() {

	if (E_Day == TextValue1 || E_Day == "") {
		MessageBox(_T("请在文本框中输入要修改的天数"));
		return false;
	}
	return true;
}

//listctrl框初始化
void CRelaxTimeDlg::InitListCtrl()
{
	m_cb1.ResetContent();
	m_cb1.AddString(TextValue2);

 	ConnectDB();
	HRESULT hr;
	_RecordsetPtr pRentRecordset;
	hr = pRentRecordset.CreateInstance(__uuidof(Recordset));
	if (FAILED(hr)) {
		MessageBox(_T("创建记录集对象失败."));
		return;
	}

	CString strSQL;
	strSQL = _T("select * from sPeople");
	try {

		hr = pRentRecordset->Open(_variant_t(strSQL), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		if (SUCCEEDED(hr)) {
			_variant_t var;
			CString strValue;
			int curItem = 0;
			while (!pRentRecordset->GetadoEOF()) {
				var = pRentRecordset->GetCollect(_T("ID"));
				if (var.vt != NULL)
					strValue = (LPCTSTR)_bstr_t(var);
				m_list.InsertItem(curItem, strValue);

				m_cb1.AddString(strValue);

				var = pRentRecordset->GetCollect(_T("sName"));
				if (var.vt != NULL)
					strValue = (LPCTSTR)_bstr_t(var);
				m_list.SetItemText(curItem, 1, strValue);

				var = pRentRecordset->GetCollect(_T("sTime"));
				if (var.vt != NULL)
					strValue = (LPCTSTR)_bstr_t(var);
				m_list.SetItemText(curItem, 2, strValue);

				var = pRentRecordset->GetCollect(_T("sData"));
				if (var.vt != NULL)
					strValue = (LPCTSTR)_bstr_t(var);
				m_list.SetItemText(curItem, 3, strValue);

				pRentRecordset->MoveNext();
				curItem++;
			}
			UpdateData();
		}
		else {
			MessageBox(_T("打开结果记录集失败."));
			return;
		}
	}
	catch (_com_error* e) {
		MessageBox(e->ErrorMessage());

	}
	m_cb1.SetCurSel(0);

	pRentRecordset->Close();
	pRentRecordset = NULL;
	closedb();
}

//通过选择序号提取出某个人还剩下的假期
CString CRelaxTimeDlg::GetTime(CString CBText)
{
	ConnectDB();
	_RecordsetPtr pDVDIDRecordset;
	pDVDIDRecordset.CreateInstance(__uuidof(Recordset));

	CString value = (_T(""));
	CString strSQL, strValue;
	strSQL.Format(_T("select sTime from sPeople where id = %d"), _ttoi(CBText));

	try
	{
		HRESULT hr = pDVDIDRecordset->Open(_variant_t(strSQL), m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
		if (FAILED(hr)) {
			return value;
		}
	}
	catch (_com_error* e)
	{
		MessageBox(e->ErrorMessage());
		return value;
	}

	_variant_t var;
	var = pDVDIDRecordset->GetCollect(_T("sTime"));
	if (var.vt != VT_NULL) {
		strValue = (LPCTSTR)_bstr_t(var);
		value = strValue;
	}
	else {
		value = "";
	}

	pDVDIDRecordset->Close();
	pDVDIDRecordset = NULL;
	closedb();
	return value;
}

//连接数据库
void CRelaxTimeDlg::ConnectDB()
{
	CoInitialize(NULL);  //初始化 
	m_pConnection.CreateInstance(__uuidof(Connection)); //实例化对象  

	 //连到具体某个mdb ，此处的的Provider语句因Access版本的不同而有所不同。
	try {
		m_pConnection->
			Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=MyAccess.accdb;Persist Security Info = False;", "", "", adModeUnknown);
	}
	catch (_com_error e)
	{
		CString errormessage;
		errormessage.Format(_T("连接数据库失败!\n错误信息:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);//显示错误信息

	}
}


//关闭数据库
void CRelaxTimeDlg::closedb()
{
	//退出程序时的处理  ，关闭数据库的相关操作
	if (m_pConnection->State)
	{
		m_pConnection->Close();
	}
}


//初始化函数 所有文本的东西都丢在这里
void CRelaxTimeDlg::chushihua() {
	m_list.InsertColumn(0, _T("序号"), LVCFMT_LEFT, 100);
	m_list.InsertColumn(1, _T("名字"), LVCFMT_LEFT, 150);
	m_list.InsertColumn(2, _T("剩余假期"), LVCFMT_LEFT, 150);
	m_list.InsertColumn(3, _T("修改时间"), LVCFMT_LEFT, 300);


	

}
//计时器，拿来设定文本框提示符
void CRelaxTimeDlg::OnTimer(UINT_PTR nIDEvent)
{
	UpdateData();
	if (E_Day == "" || E_Day == TextValue1) {
		GetDlgItem(IDC_DAY)->SetWindowText(TextValue1);
		SetFocus();
	}

	CDialogEx::OnTimer(nIDEvent);
}

//判断是否有选择序号
BOOL CRelaxTimeDlg::GetcbCurSel(){

int nIndex = m_cb1.GetCurSel();
m_cb1.GetLBText(nIndex, strCBText); //获取到选择的下拉框的具体文本

if (strCBText == TextValue2)
{
	MessageBox(_T("请选择了要修改的序号再使用此功能"));
	return false;
}
return true;

 }

//修改数据库的时间
BOOL CRelaxTimeDlg::ChangeDataTime(CString NewData,CString Times)
{
	ConnectDB();

	HRESULT hr;

	_CommandPtr pCommand; //申请cmd对象
	hr = pCommand.CreateInstance(__uuidof(Command));
	if (FAILED(hr)) {
		MessageBox(_T("创建cmd对象失败"));
		return false;
	}
	pCommand->ActiveConnection = m_pConnection;
	CString strSql;
	
			strSql.Format(_T("update sPeople set sTime = \'%s\',sData = \'%s\' where ID = %d"), NewData,Times, _ttoi(strCBText));
			pCommand->CommandText = _bstr_t(strSql);
			try {
				hr = pCommand->Execute(NULL, NULL, adCmdText);
				if (FAILED(hr)) {
					return false;
				}
			}
			catch (_com_error* e) {
				MessageBox(e->ErrorMessage());
			}
		

	pCommand = NULL;
	closedb();
	return TRUE;
}

//获取系统时间
CString CRelaxTimeDlg::GetSystemTime() {

	CString m_strDateTime;
	CTime m_time;
	m_time = CTime::GetCurrentTime();             //获取当前时间日期  
	
	m_strDateTime = m_time.Format(_T("%Y-%m-%d %H:%M:%S"));   //格式化日期时间  
	//UpdateData(false);

	return  m_strDateTime;
}