// NetWorkDlg.h : 头文件
//

#pragma once

#include <mshtml.h>

// CNetWorkDlg 对话框
class CNetWorkDlg : public CDialog
{
// 构造
public:
	CNetWorkDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_NETWORK_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持
	BOOL bReady;				 //判断yahoo邮箱页面是否完全打开


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

	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButtonSumbit();
	void LoginYahoo(void);
	void SumbitPage(void);
	void EnumIE(void);
	// void EnumFrame(IHTMLDocument2 * pIHTMLDocument2)
	void EnumFrame(IHTMLDocument2 * pIHTMLDocument2);
	// void EnumForm(IHTMLDocument2 * pIHTMLDocument2)  
	void EnumForm(IHTMLDocument2 * pIHTMLDocument2,IDispatch*spDispDoc);
	// void EnumField(CComDispatchDriver spInputElement,CString ComType,CString ComVal,CString ComName)
	void EnumField(CComDispatchDriver spInputElement , CString ComType , CString ComVal , CString ComName);
	afx_msg void OnBnClickedButton2();
	CString m_MailName;
	CString m_MailPsw;
};
