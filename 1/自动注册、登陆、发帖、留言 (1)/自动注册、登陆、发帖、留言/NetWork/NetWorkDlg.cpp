// NetWorkDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "NetWork.h"
#include "NetWorkDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

int	RegOrComment;//ע��򷢱���־

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CNetWorkDlg �Ի���




CNetWorkDlg::CNetWorkDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CNetWorkDlg::IDD, pParent)
	, m_MailName(_T(""))
	, m_MailPsw(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CNetWorkDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT_MAILNAME, m_MailName);
	DDX_Text(pDX, IDC_EDIT_MAILPSW, m_MailPsw);
}

BEGIN_MESSAGE_MAP(CNetWorkDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON1, &CNetWorkDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON_SUMBIT, &CNetWorkDlg::OnBnClickedButtonSumbit)
	ON_BN_CLICKED(IDC_BUTTON2, &CNetWorkDlg::OnBnClickedButton2)
END_MESSAGE_MAP()


// CNetWorkDlg ��Ϣ�������

BOOL CNetWorkDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	RegOrComment=0;

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CNetWorkDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CNetWorkDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CNetWorkDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CNetWorkDlg::OnBnClickedButton1()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	RegOrComment=0;//ע�������Ա
	SumbitPage();//�Զ��ύע����Ϣ

//	����ٶ���ҳ��IWebBrowser2�ؼ�	
//	CComVariant vtUrl(TEXT("http://www.baidu.com"));
//	CComVariant	vtEmpty;
//	explorer1.h��
//	explorer1	m_MyWeb;
//	����IWebBrowser2*	m_MyWeb;
//	m_MyWeb.Navigate2(&vtUrl, &vtEmpty, &vtEmpty, &vtEmpty, &vtEmpty);
}

void CNetWorkDlg::OnBnClickedButtonSumbit()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	UpdateData(TRUE);

	if(m_MailPsw.IsEmpty()||	m_MailName.IsEmpty())	
	{
		MessageBox(TEXT("�������û���������"));
		return;
	}

	LoginYahoo();/*�Զ���¼�Ż�����*/
}

void CNetWorkDlg::LoginYahoo(void)
{
	/*�Զ���¼�Ż�����*/
	//#include <mshtml.h>
	//CString m_strURL = TEXT("http://cn.mail.yahoo.com/");  //URL��ַ
	CString m_strURL = TEXT("http://mail.126.com/");
	
	VARIANT id, index;
	
	bReady=0;
	BSTR bsStatus;
	BSTR bsPW = m_MailPsw.AllocSysString();
	BSTR bsUser=m_MailName.AllocSysString();

	CString mStr;
	HRESULT hr;	
	hr = CoInitialize(NULL);
	if(!SUCCEEDED(hr))		return;

	IWebBrowser2* pBrowser;
	hr = CoCreateInstance (CLSID_InternetExplorer, NULL, CLSCTX_LOCAL_SERVER, IID_IWebBrowser2, (LPVOID *)&pBrowser); //Create an Instance of web browser,�������

	if(hr!=S_OK)
	{
		char a[1024] = {0};
		DWORD b = GetLastError();
		sprintf(a,"CoCreateInstance fail a=%d hr = 0x%x",b,hr);
		return;
	}
	COleVariant vaURL(m_strURL) ;
	COleVariant null;
	
	VARIANT_BOOL pBool = TRUE;
	pBrowser->put_Visible( pBool );//Comment out this line if you dont want the browser to be displayed,FALSE���������

	pBrowser->Navigate2(vaURL,null,null,null,null) ; //Open the URL page,����ҳ
	//	m_MyWeb.Navigate2(vaURL,null,null,null,null) ; ������ҳ

	
	while(!bReady)  //This while loop mask sure that the page is fully loaded before we go to the next page
	{
		//����û��ֶ��ر�IE����,�˳�ѭ��
		SHANDLE_PTR hHwnd;
		pBrowser->get_HWND(&hHwnd);
		if (NULL == hHwnd)
		{
			bReady=1;
			return;
		}
		
		//�ȴ���ҳ��ȫ��,�˳�ѭ��
		pBrowser->get_StatusText(&bsStatus);
		mStr=bsStatus;
		if(mStr==TEXT("���") || mStr==TEXT("���" )|| mStr==TEXT("Done") )
		{
			bReady=1;
			break;
		}
		
		Sleep(200);
	}		

	
	IDispatch* pDisp;
	hr=pBrowser->get_Document(&pDisp);//Get the underlying document object of the browser
	
	if (pDisp == NULL )
	{
		CoUninitialize();
		return;
	}
	IHTMLDocument2* pHTMLDocument2;
	
	hr = pDisp->QueryInterface( IID_IHTMLDocument2,
		(void**)&pHTMLDocument2 );//Ask for an HTMLDocument2 interface
	
	if (hr != S_OK)
	{
		CoUninitialize();
		return ;
	}
	IHTMLElementCollection* pColl = NULL;//Enumerate the HTML elements
	
	hr = pHTMLDocument2->get_all( &pColl );
	if (hr != S_OK || pColl == NULL)
	{
		pHTMLDocument2->Release();
		CoUninitialize();
		return;
	}
	LONG celem;
	hr = pColl->get_length( &celem );//Find the count of the elements
	
	if ( hr != S_OK )
	{
		pHTMLDocument2->Release();
		pColl->Release();
		CoUninitialize();
		return;
	}

	for ( int i=0; i< celem; i++ )//Loop through each elment
	{
		IDispatch* pDisp2;
		
		V_VT(&id) = VT_I4;
		V_I4(&id) = i;
		V_VT(&index) = VT_I4;
		V_I4(&index) = 0;
		hr = pColl->item( id,index, &pDisp2 );//Get an element
		
		if ( hr == S_OK )
		{
			IHTMLElement* pElem;
			
			//Ask for an HTMLElemnt interface
			hr = pDisp2->QueryInterface(IID_IHTMLElement,(void **)&pElem);
			if ( hr == S_OK )
			{
				BSTR bstr;
				
				IHTMLInputTextElement* pUser;//We need to check for input elemnt on login screen
				hr = pDisp2->QueryInterface(IID_IHTMLInputTextElement,(void **)&pUser );
				if ( hr == S_OK )
				{
					pUser->get_name(&bstr);
					mStr=bstr;
					//::MessageBox(0, mStr, 0, 0);
					//pUser->get_type(&bstr);
					//CString cstrType = bstr;
					//::MessageBox(0, cstrType, 0, 0);
					//if(mStr==TEXT("login"))
					//if (mStr == TEXT("idInput"))
					if (mStr == TEXT(""))				
					{//Is this a User Id field
						MessageBox(TEXT("5"));
						pUser->put_value(bsUser);//Paste the User Id
					}
					//else if(mStr==TEXT("passwd"))
					else if(mStr == TEXT("password"))
					{//Or, is this a password field
						pUser->put_value(bsPW);// Paste your password into the field                 
					}
					pUser->Release();					
				}
				else
				{
					IHTMLInputButtonElement* pButton;
					//hr = pDisp2->QueryInterface(IID_IHTMLInputButtonElement,(void **)&pButton);
					hr = pDisp2->QueryInterface(IID_IHTMLButtonElement,(void **)&pButton);
					if ( hr == S_OK )
					{
						//pButton->get_type(&bstr);//This will send the all the information in correct format         
						pButton->get_value(&bstr);
						mStr=bstr;
						pButton->get_type(&bstr);
						CString cstrType = bstr;
						if (cstrType == TEXT("submit"))
						{
							//if (mStr==TEXT("�� ¼"))
							{
								pElem->click();							
								i=celem;
							}
						}
						pButton->Release();
					}
				}
				pElem->Release();
			}
			pDisp2->Release();
		}
	}	
	pColl->Release();	
	pHTMLDocument2->Release();//For the next page open a fresh document	
	pDisp->Release();	
	CoUninitialize();
}

void CNetWorkDlg::SumbitPage(void)
{
	CoInitialize(NULL); //��ʼ��COM
	EnumIE();             //ö�������       
	CoUninitialize();   //�ͷ�COM
}


void CNetWorkDlg::EnumIE(void)
{
	
	CString sURL;	//sURLΪҪע�����ַ����http://www.Ice.com

	if(!RegOrComment)	sURL=TEXT("file://D:/������Աע��.htm");
	else				sURL=TEXT("file://D:/�����Ż�.htm");	
	/*���ص���ַ�����ַ�ʽ
		1��file://D:/������Աע��.htm			
		2��http://localhost:1909/Default.aspx			��Ϊ�������Ǳ��ز��ԣ����������������˿�����VS2008����ʱ����ַ�õ���
			���Խ�����ĸ���ַ�����ĸ���ַ��	�����ҵ�����Щ����IIS�ģ���Ĭ����80�˿ڣ�Ҳ������http://localhost:80/Default.aspxΪ��ַ
			������������ʹ򲻿�
	*/


	CComQIPtr<IWebBrowser2>spBrowser;
	HRESULT	hr = CoCreateInstance (CLSID_InternetExplorer, NULL, CLSCTX_LOCAL_SERVER, IID_IWebBrowser2, (LPVOID *)&spBrowser);	//�������
	if(hr!=S_OK)		return;

	COleVariant vaURL(sURL) ;	
	COleVariant null;	

	VARIANT_BOOL pBool=TRUE;
	spBrowser->put_Visible( pBool );//TRUE��ʾ�������FALSE���������
	spBrowser->Navigate2(vaURL,null,null,null,null) ; //Open the URL page,����ҳ


	Sleep(2000);//�ȴ���ҳ�������
	if(!RegOrComment)	sURL=TEXT("file:///D:/������Աע��.htm");	//ע�⣺�����file����������б�ܡ�	�����http��ͷ�ģ�����Ҫ���if���
	else				sURL=TEXT("file:///D:/�����Ż�.htm");	

	CComPtr<IShellWindows> spShellWin;   
	 hr=spShellWin.CoCreateInstance(CLSID_ShellWindows);   
	if (FAILED(hr))   		return;   

	long nCount=0;    //ȡ�������ʵ������(Explorer��IExplorer)   
	spShellWin->get_Count(&nCount); 
	if (0==nCount)   		return; 

	for(int i=0; i<nCount; i++)   
	{  
		CComPtr<IDispatch> spDispIE;   
		hr=spShellWin->Item(CComVariant((long)i), &spDispIE);   
		if (FAILED(hr)) continue; 

		spBrowser=spDispIE; 
		if (!spBrowser) continue; 

		IDispatch*spDispDoc;   
		hr=spBrowser->get_Document(&spDispDoc);   
		if (FAILED(hr)) continue; 

		CComQIPtr<IHTMLDocument2>spDocument2 =spDispDoc;   
		if (!spDocument2) continue;       

		CString cIEUrl_Filter;  //����URL(�����Ǵ�URL����վ����Ч);
		cIEUrl_Filter=sURL; 
	
		CComBSTR IEUrl;
		spBrowser->get_LocationURL(&IEUrl);    
		CString cIEUrl_Get;     //�ӻ�����ȡ�õ�HTTP��������URL;
		cIEUrl_Get=IEUrl;
		cIEUrl_Get=cIEUrl_Get.Left(cIEUrl_Filter.GetLength()); //��ȡǰ��Nλ

//MessageBox(cIEUrl_Get+TEXT("2")+cIEUrl_Filter); 

		if (cIEUrl_Get==cIEUrl_Filter)
		{
			// �������е��ˣ��Ѿ��ҵ���IHTMLDocument2�Ľӿ�ָ��   

//	CComPtr<IHTMLFormElement> pForm;
//	if (SUCCEEDED(spBrowser->get_Document( &spDispDoc)))	spDocument2 = spDispDoc;
//	if(SUCCEEDED(spDispDoc->QueryInterface(IID_IHTMLFormElement,(void**)&pForm)))
//	if(SUCCEEDED(pFormElement->item(id,index, &spDispDoc)))
//	spDispDoc->QueryInterface(IID_IHTMLInputTextElement,(void**)&pInputElement);
										
			EnumForm(spDocument2,spDispDoc); //ö�����еı� 
		}  				
	}   
}


void CNetWorkDlg::EnumFrame(IHTMLDocument2 * pIHTMLDocument2)
{
	if (!pIHTMLDocument2) return;       
	HRESULT   hr;   

	CComPtr<IHTMLFramesCollection2> spFramesCollection2;   
	pIHTMLDocument2->get_frames(&spFramesCollection2); //ȡ�ÿ��frame�ļ���   

	long nFrameCount=0;        //ȡ���ӿ�ܸ���   
	hr=spFramesCollection2->get_length(&nFrameCount);   
	if (FAILED(hr)|| 0==nFrameCount) return;   

	for(long i=0; i<nFrameCount; i++)   
	{   
		CComVariant vDispWin2; //ȡ���ӿ�ܵ��Զ����ӿ�   
		hr = spFramesCollection2->item(&CComVariant(i), &vDispWin2);   
		if (FAILED(hr)) continue;       
		CComQIPtr<IHTMLWindow2>spWin2 = vDispWin2.pdispVal;   
		if (!spWin2) continue; //ȡ���ӿ�ܵ�   IHTMLWindow2   �ӿ�       
		CComPtr <IHTMLDocument2> spDoc2;   
		spWin2->get_document(&spDoc2); //ȡ���ӿ�ܵ�   IHTMLDocument2   �ӿ�

		EnumForm(spDoc2,spDoc2);      //�ݹ�ö�ٵ�ǰ�ӿ��   IHTMLDocument2   �ϵı�form   
	}   
}

void CNetWorkDlg::EnumForm(IHTMLDocument2 * pIHTMLDocument2,IDispatch*spDispDoc)
{
	CString sDbg;

	if (!pIHTMLDocument2) return; 
	EnumFrame(pIHTMLDocument2);   //�ݹ�ö�ٵ�ǰIHTMLDocument2�ϵ��ӿ��frame   

	HRESULT hr;
	USES_CONVERSION;       

	CComQIPtr<IHTMLElementCollection> spElementCollection;   
	hr=pIHTMLDocument2->get_forms(&spElementCollection); //ȡ�ñ�����   
	if (FAILED(hr))   	return;   
 
	long nFormCount=0;           //ȡ�ñ���Ŀ   
	hr=spElementCollection->get_length(&nFormCount);   
	if (FAILED(hr))   	return;   

	for(long i=0; i<nFormCount; i++)   
	{   
		IDispatch *pDisp = NULL;   //ȡ�õ�i���   
		hr=spElementCollection->item(CComVariant(i),CComVariant(),&pDisp);   
		if(FAILED(hr)) continue;   

		CComQIPtr<IHTMLFormElement> spFormElement= pDisp;   
		pDisp->Release();   

		long nElemCount=0;         //ȡ�ñ��������Ŀ   
		hr=spFormElement->get_length(&nElemCount);   
		if(FAILED(hr)) continue;   

		for(long j=0; j<nElemCount; j++)   
		{  			
			CComDispatchDriver spInputElement; //ȡ�õ�j�����   
			hr=spFormElement->item(CComVariant(j), CComVariant(), &spInputElement);   
			if(FAILED(hr)) continue;   

			CComVariant vName,vVal,vType;     //ȡ�ñ�������ƣ���ֵ������ 
			hr=spInputElement.GetPropertyByName(TEXT("name"), &vName);   
			if(FAILED(hr)) continue;   
			hr=spInputElement.GetPropertyByName(TEXT("value"), &vVal);   
			if(FAILED(hr)) continue;   
			hr=spInputElement.GetPropertyByName(TEXT("type"), &vType);   
			if(FAILED(hr)) continue;   

			LPCTSTR lpName= vName.bstrVal ? OLE2CT(vName.bstrVal) : TEXT(""); //δ֪����   
			LPCTSTR lpVal=  vVal.bstrVal  ? OLE2CT(vVal.bstrVal)  : TEXT(""); //��ֵ��δ����   
			LPCTSTR lpType= vType.bstrVal ? OLE2CT(vType.bstrVal) : TEXT(""); //δ֪����  

		//	sDbg.Format(TEXT("%s,%s,%s"),lpType,lpVal,lpName);//ͨ�������Ի��򣬿��Ի��ÿһ���ؼ�������ָ�꣬���Ը�����Ҫѡ��һ��ָ��
		//	MessageBox(sDbg);

			EnumField(spInputElement,lpType,lpVal,lpName);					  //���ݲ������������͡�ֵ����


			if((CString)lpVal==TEXT("�ύ"))			break;
			if((CString)lpVal==TEXT("��������"))		break;

		}//����ѭ������     
	}//��ѭ������      
}

void CNetWorkDlg::EnumField(CComDispatchDriver spInputElement , CString ComType , CString ComVal , CString ComName)
{
	CString csTemp;
	bool	IsWant=TRUE;//�Ƿ���Ҫ��ı���
/*********************************************ע�������Ա*********************************************/		
	if		   (	ComName=="TextBox1"	)								csTemp=TEXT("����bing��");				//�û���
	else	if (	ComName=="TextBox2"||	ComName=="TextBox3"	)		csTemp=TEXT("zc2011");					//����	
	else	if (	ComName=="TextBox4"	)								csTemp=TEXT("2011��2��10�� ������");	//"��������"
	else	if (	ComName=="TextBox5"	)								csTemp=TEXT("lxz@qq.com");				//"Email:"
	else	if (	ComName=="DropDownList1"	)						csTemp=TEXT("�������ѧ�뼼��");		//"רҵ:"
	else	if (	ComName=="TextBox7"	)								csTemp=TEXT("����bing��   1662888517@qq.com");				//ComVal=="���˼��"
	else	IsWant=FALSE;

	if(IsWant)//�������Ҫ��ı�������д
	{
		CComVariant vSetStatus(csTemp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}

	if(ComVal=="�ύ")
	{
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		MessageBox(TEXT("���Կ������Ƿ��Ѿ�������"));
		pHElement->click();   
	}
/*********************************************�ڱ�������������*********************************************/		
	//lpName	TextBox_Daily	Button_Submit0	TextBox_Comment		Button_Submit
	//lpVal		������д���������·�̰�......		������־	''	��������

	IsWant=TRUE;
	if			(	ComName=="TextBox_Daily")			csTemp=TEXT("Hello!	��������(*^__^*)");			//д��־		
	else	if  (	ComName=="TextBox_Comment")			csTemp=TEXT("Thanks!	��������(*^__^*)");		//д����
	else	IsWant=FALSE;
		
	if(IsWant)//�������Ҫ��ı�������д
	{
		CComVariant vSetStatus(csTemp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}

	if(ComName=="Button_Submit0")	//������־
	{ 
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		MessageBox(TEXT("���Կ�����־�Ƿ��Ѿ�д����"));
	//	pHElement->click();		//�����Ҫ������־����ȡ������ע�ͣ�
								//����EnumForm()��		��if((CString)lpVal==TEXT("��������"))break;����һ�������ϡ�if((CString)lpVal==TEXT("������־"))break;��

	}
	else	if (ComName=="Button_Submit")	//��������
	{ 
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		MessageBox(TEXT("���Կ��������Ƿ��Ѿ�������"));
		pHElement->click();  
	}	

/*********************************************��дQQ��ҳ*********************************************/		
/*	���´�����������������ع����Ĵ����ʽд�ģ�����Ⱥܸߣ��Ͳ����ˣ�����Ū�����Ժ󷢲��򵥵��Դ��
���ߵķ��׾����Ǻ�ֵ�ñ����ѧϰ�ģ�	�����ٶȺ͹ȸ��ϣ���������ע�͵�Ҳ�������������߾���֮һ��
��ַ���ǵ��ˣ���ֻ�����˱�����ҳ�����Դ򿪿��������߲��͵�����Ϊ��fjssharpsword��ר�������ٴθ�л���ߵ���˽����Ŷ


	//�����˺�
	if ((ComType.Find(TEXT("text"))>=0)  &&  ComVal=="����дһ������Ҫ������ʺ�" && ComName==TEXT(""))
	{
		Tmp=TEXT("qq");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}
	//1����@qq.com,3����@foxmail.com
	else	if ((ComType.Find(TEXT("select-one"))>=0)  &&ComName==TEXT("emailselect"))
	{ 
		Tmp=TEXT("3");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}
	//����˺��Ƿ����
	else	if ((ComType.Find(TEXT("button"))>=0)   &&ComName==TEXT(""))
	{ 
		//button�����ָ�Ϊ123456
		Tmp=TEXT("123456");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
		//�����ť
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		pHElement->click();  
	}	
	//�ǳ�
	else	if ((ComType.Find(TEXT("text"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{
		Tmp=TEXT("�ǳ�");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}
	//��
	else	if ((ComType.Find(TEXT("select-one"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("2011");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//��
	else	if ((ComType.Find(TEXT("select-one"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("1");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}
	//��
	else	if ((ComType.Find(TEXT("select-one"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("1");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//ComVal==TEXT("1")1�����У�2����Ů
	else	if ((ComType.Find(TEXT("radio"))>=0)  &&ComVal==TEXT("2")&& ComName==TEXT("gender"))
	{
		//���ѡ��
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		pHElement->click();  
	}	
	//password
	else	if ((ComType.Find(TEXT("password"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("password");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//repassword
	else	if ((ComType.Find(TEXT("password"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("password");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//ʡ
	else	if ((ComType.Find(TEXT("select-one"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("ʡ");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//��
	else	if ((ComType.Find(TEXT("select-one"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("��");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//��
	else	if ((ComType.Find(TEXT("��"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("password");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//��֤��
	else	if ((ComType.Find(TEXT("text"))>=0) &&  ComVal==TEXT("������������ͼ�п������ַ��������ִ�Сд")  &&ComName==TEXT("verifycode"))
	{ 
		Tmp=TEXT("cfsb");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	

	else	if(ComType.Find(TEXT("submit"))!=-1)
	{
//		MessageBox(TEXT("SUBMIT:"));
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		pHElement->click();                
	}*/
/*


IHTMLElement* pElem;
//��ó�����ҳԪ��HTMLElement�ӿ�
hr = spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pElem);
//��ͼ��ó�����Ԫ�ؽӿ�

IHTMLAnchorElement* pSubAnchor;
hr = spInputElement->QueryInterface(IID_IHTMLAnchorElement, (void **)&pSubAnchor);
if(hr == S_OK) {
BSTR bstr;
pElem->get_id(&bstr);
CString strID;
strID = bstr;
if(strID == "a_submit"){
MessageBox(TEXT("���Կ�����Ϣ�Ƿ��Ѿ�������"));
pElem->click();
}
pSubAnchor->Release();
}


MessageBox(ComName);
IHTMLFormElement*pForm; 


	spInputElement->QueryInterface(IID_IHTMLFormElement,(void**)&pForm);
	pForm->submit();

MessageBox(ComVal);
*/ 

//MessageBox(ComName);
//MessageBox(ComVal);
//MessageBox(ComType);

}

void CNetWorkDlg::OnBnClickedButton2()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	RegOrComment=1;//���������־
	SumbitPage();//�Զ��ύע����Ϣ

}
