// NetWorkDlg.h : ͷ�ļ�
//

#pragma once

#include <mshtml.h>

// CNetWorkDlg �Ի���
class CNetWorkDlg : public CDialog
{
// ����
public:
	CNetWorkDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_NETWORK_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��
	BOOL bReady;				 //�ж�yahoo����ҳ���Ƿ���ȫ��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
