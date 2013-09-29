// NetWorkDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "NetWork.h"
#include "NetWorkDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

int	RegOrComment;//注册或发表日志

// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CNetWorkDlg 对话框




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


// CNetWorkDlg 消息处理程序

BOOL CNetWorkDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	RegOrComment=0;

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CNetWorkDlg::OnPaint()
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
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CNetWorkDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CNetWorkDlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	RegOrComment=0;//注册冰网会员
	SumbitPage();//自动提交注册信息

//	载入百度网页至IWebBrowser2控件	
//	CComVariant vtUrl(TEXT("http://www.baidu.com"));
//	CComVariant	vtEmpty;
//	explorer1.h类
//	explorer1	m_MyWeb;
//	或者IWebBrowser2*	m_MyWeb;
//	m_MyWeb.Navigate2(&vtUrl, &vtEmpty, &vtEmpty, &vtEmpty, &vtEmpty);
}

void CNetWorkDlg::OnBnClickedButtonSumbit()
{
	// TODO: 在此添加控件通知处理程序代码
	UpdateData(TRUE);

	if(m_MailPsw.IsEmpty()||	m_MailName.IsEmpty())	
	{
		MessageBox(TEXT("请输入用户名和密码"));
		return;
	}

	LoginYahoo();/*自动登录雅虎邮箱*/
}

void CNetWorkDlg::LoginYahoo(void)
{
	/*自动登录雅虎邮箱*/
	//#include <mshtml.h>
	//CString m_strURL = TEXT("http://cn.mail.yahoo.com/");  //URL地址
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
	hr = CoCreateInstance (CLSID_InternetExplorer, NULL, CLSCTX_LOCAL_SERVER, IID_IWebBrowser2, (LPVOID *)&pBrowser); //Create an Instance of web browser,打开浏览器

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
	pBrowser->put_Visible( pBool );//Comment out this line if you dont want the browser to be displayed,FALSE隐藏浏览器

	pBrowser->Navigate2(vaURL,null,null,null,null) ; //Open the URL page,打开网页
	//	m_MyWeb.Navigate2(vaURL,null,null,null,null) ; 载入网页

	
	while(!bReady)  //This while loop mask sure that the page is fully loaded before we go to the next page
	{
		//如果用户手动关闭IE窗口,退出循环
		SHANDLE_PTR hHwnd;
		pBrowser->get_HWND(&hHwnd);
		if (NULL == hHwnd)
		{
			bReady=1;
			return;
		}
		
		//等待网页完全打开,退出循环
		pBrowser->get_StatusText(&bsStatus);
		mStr=bsStatus;
		if(mStr==TEXT("完毕") || mStr==TEXT("完成" )|| mStr==TEXT("Done") )
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
							//if (mStr==TEXT("登 录"))
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
	CoInitialize(NULL); //初始化COM
	EnumIE();             //枚举浏览器       
	CoUninitialize();   //释放COM
}


void CNetWorkDlg::EnumIE(void)
{
	
	CString sURL;	//sURL为要注册的网址，如http://www.Ice.com

	if(!RegOrComment)	sURL=TEXT("file://D:/冰网会员注册.htm");
	else				sURL=TEXT("file://D:/冰网门户.htm");	
	/*本地的网址有两种方式
		1】file://D:/冰网会员注册.htm			
		2】http://localhost:1909/Default.aspx			因为我做的是本地测试，所有用这个，这个端口是由VS2008调试时的网址得到的
			调试结果是哪个网址就填哪个网址，	网上找到的那些配置IIS的，都默认是80端口，也就是以http://localhost:80/Default.aspx为网址
			很明显我这里就打不开
	*/


	CComQIPtr<IWebBrowser2>spBrowser;
	HRESULT	hr = CoCreateInstance (CLSID_InternetExplorer, NULL, CLSCTX_LOCAL_SERVER, IID_IWebBrowser2, (LPVOID *)&spBrowser);	//打开浏览器
	if(hr!=S_OK)		return;

	COleVariant vaURL(sURL) ;	
	COleVariant null;	

	VARIANT_BOOL pBool=TRUE;
	spBrowser->put_Visible( pBool );//TRUE显示浏览器，FALSE隐藏浏览器
	spBrowser->Navigate2(vaURL,null,null,null,null) ; //Open the URL page,打开网页


	Sleep(2000);//等待网页加载完毕
	if(!RegOrComment)	sURL=TEXT("file:///D:/冰网会员注册.htm");	//注意：这里的file后面有三个斜杠。	如果是http开头的，则不需要这个if语句
	else				sURL=TEXT("file:///D:/冰网门户.htm");	

	CComPtr<IShellWindows> spShellWin;   
	 hr=spShellWin.CoCreateInstance(CLSID_ShellWindows);   
	if (FAILED(hr))   		return;   

	long nCount=0;    //取得浏览器实例个数(Explorer和IExplorer)   
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

		CString cIEUrl_Filter;  //设置URL(必须是此URL的网站才有效);
		cIEUrl_Filter=sURL; 
	
		CComBSTR IEUrl;
		spBrowser->get_LocationURL(&IEUrl);    
		CString cIEUrl_Get;     //从机器上取得的HTTP的完整的URL;
		cIEUrl_Get=IEUrl;
		cIEUrl_Get=cIEUrl_Get.Left(cIEUrl_Filter.GetLength()); //截取前面N位

//MessageBox(cIEUrl_Get+TEXT("2")+cIEUrl_Filter); 

		if (cIEUrl_Get==cIEUrl_Filter)
		{
			// 程序运行到此，已经找到了IHTMLDocument2的接口指针   

//	CComPtr<IHTMLFormElement> pForm;
//	if (SUCCEEDED(spBrowser->get_Document( &spDispDoc)))	spDocument2 = spDispDoc;
//	if(SUCCEEDED(spDispDoc->QueryInterface(IID_IHTMLFormElement,(void**)&pForm)))
//	if(SUCCEEDED(pFormElement->item(id,index, &spDispDoc)))
//	spDispDoc->QueryInterface(IID_IHTMLInputTextElement,(void**)&pInputElement);
										
			EnumForm(spDocument2,spDispDoc); //枚举所有的表单 
		}  				
	}   
}


void CNetWorkDlg::EnumFrame(IHTMLDocument2 * pIHTMLDocument2)
{
	if (!pIHTMLDocument2) return;       
	HRESULT   hr;   

	CComPtr<IHTMLFramesCollection2> spFramesCollection2;   
	pIHTMLDocument2->get_frames(&spFramesCollection2); //取得框架frame的集合   

	long nFrameCount=0;        //取得子框架个数   
	hr=spFramesCollection2->get_length(&nFrameCount);   
	if (FAILED(hr)|| 0==nFrameCount) return;   

	for(long i=0; i<nFrameCount; i++)   
	{   
		CComVariant vDispWin2; //取得子框架的自动化接口   
		hr = spFramesCollection2->item(&CComVariant(i), &vDispWin2);   
		if (FAILED(hr)) continue;       
		CComQIPtr<IHTMLWindow2>spWin2 = vDispWin2.pdispVal;   
		if (!spWin2) continue; //取得子框架的   IHTMLWindow2   接口       
		CComPtr <IHTMLDocument2> spDoc2;   
		spWin2->get_document(&spDoc2); //取得子框架的   IHTMLDocument2   接口

		EnumForm(spDoc2,spDoc2);      //递归枚举当前子框架   IHTMLDocument2   上的表单form   
	}   
}

void CNetWorkDlg::EnumForm(IHTMLDocument2 * pIHTMLDocument2,IDispatch*spDispDoc)
{
	CString sDbg;

	if (!pIHTMLDocument2) return; 
	EnumFrame(pIHTMLDocument2);   //递归枚举当前IHTMLDocument2上的子框架frame   

	HRESULT hr;
	USES_CONVERSION;       

	CComQIPtr<IHTMLElementCollection> spElementCollection;   
	hr=pIHTMLDocument2->get_forms(&spElementCollection); //取得表单集合   
	if (FAILED(hr))   	return;   
 
	long nFormCount=0;           //取得表单数目   
	hr=spElementCollection->get_length(&nFormCount);   
	if (FAILED(hr))   	return;   

	for(long i=0; i<nFormCount; i++)   
	{   
		IDispatch *pDisp = NULL;   //取得第i项表单   
		hr=spElementCollection->item(CComVariant(i),CComVariant(),&pDisp);   
		if(FAILED(hr)) continue;   

		CComQIPtr<IHTMLFormElement> spFormElement= pDisp;   
		pDisp->Release();   

		long nElemCount=0;         //取得表单中域的数目   
		hr=spFormElement->get_length(&nElemCount);   
		if(FAILED(hr)) continue;   

		for(long j=0; j<nElemCount; j++)   
		{  			
			CComDispatchDriver spInputElement; //取得第j项表单域   
			hr=spFormElement->item(CComVariant(j), CComVariant(), &spInputElement);   
			if(FAILED(hr)) continue;   

			CComVariant vName,vVal,vType;     //取得表单域的名称，数值，类型 
			hr=spInputElement.GetPropertyByName(TEXT("name"), &vName);   
			if(FAILED(hr)) continue;   
			hr=spInputElement.GetPropertyByName(TEXT("value"), &vVal);   
			if(FAILED(hr)) continue;   
			hr=spInputElement.GetPropertyByName(TEXT("type"), &vType);   
			if(FAILED(hr)) continue;   

			LPCTSTR lpName= vName.bstrVal ? OLE2CT(vName.bstrVal) : TEXT(""); //未知域名   
			LPCTSTR lpVal=  vVal.bstrVal  ? OLE2CT(vVal.bstrVal)  : TEXT(""); //空值，未输入   
			LPCTSTR lpType= vType.bstrVal ? OLE2CT(vType.bstrVal) : TEXT(""); //未知类型  

		//	sDbg.Format(TEXT("%s,%s,%s"),lpType,lpVal,lpName);//通过弹出对话框，可以获得每一个控件的三项指标，可以根据需要选择一个指标
		//	MessageBox(sDbg);

			EnumField(spInputElement,lpType,lpVal,lpName);					  //传递并处理表单域的类型、值、名


			if((CString)lpVal==TEXT("提交"))			break;
			if((CString)lpVal==TEXT("发表评论"))		break;

		}//表单域循环结束     
	}//表单循环结束      
}

void CNetWorkDlg::EnumField(CComDispatchDriver spInputElement , CString ComType , CString ComVal , CString ComName)
{
	CString csTemp;
	bool	IsWant=TRUE;//是否想要填的表单？
/*********************************************注册冰网会员*********************************************/		
	if		   (	ComName=="TextBox1"	)								csTemp=TEXT("冰灵bing魂");				//用户名
	else	if (	ComName=="TextBox2"||	ComName=="TextBox3"	)		csTemp=TEXT("zc2011");					//密码	
	else	if (	ComName=="TextBox4"	)								csTemp=TEXT("2011年2月10日 星期四");	//"出生日期"
	else	if (	ComName=="TextBox5"	)								csTemp=TEXT("lxz@qq.com");				//"Email:"
	else	if (	ComName=="DropDownList1"	)						csTemp=TEXT("计算机科学与技术");		//"专业:"
	else	if (	ComName=="TextBox7"	)								csTemp=TEXT("冰灵bing魂   1662888517@qq.com");				//ComVal=="个人简介"
	else	IsWant=FALSE;

	if(IsWant)//如果是想要填的表单，则填写
	{
		CComVariant vSetStatus(csTemp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}

	if(ComVal=="提交")
	{
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		MessageBox(TEXT("可以看看表单是否已经填上啦"));
		pHElement->click();   
	}
/*********************************************在冰网发帖并留言*********************************************/		
	//lpName	TextBox_Daily	Button_Submit0	TextBox_Comment		Button_Submit
	//lpVal		在这里写下你的心历路程吧......		发表日志	''	发表评论

	IsWant=TRUE;
	if			(	ComName=="TextBox_Daily")			csTemp=TEXT("Hello!	快乐兔有(*^__^*)");			//写日志		
	else	if  (	ComName=="TextBox_Comment")			csTemp=TEXT("Thanks!	开心兔有(*^__^*)");		//写留言
	else	IsWant=FALSE;
		
	if(IsWant)//如果是想要填的表单，则填写
	{
		CComVariant vSetStatus(csTemp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}

	if(ComName=="Button_Submit0")	//发表日志
	{ 
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		MessageBox(TEXT("可以看看日志是否已经写上啦"));
	//	pHElement->click();		//如果需要发表日志，则取消这行注释，
								//并在EnumForm()的		【if((CString)lpVal==TEXT("发表评论"))break;】这一句后面加上【if((CString)lpVal==TEXT("发表日志"))break;】

	}
	else	if (ComName=="Button_Submit")	//发表评论
	{ 
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		MessageBox(TEXT("可以看看留言是否已经填上啦"));
		pHElement->click();  
	}	

/*********************************************填写QQ网页*********************************************/		
/*	以下代码基本上是依照下载过来的代码格式写的，冗余度很高，就不改了，等我弄懂了以后发布简单点的源码
作者的奉献精神是很值得表扬和学习的，	整个百度和谷歌上，能用且有注释的也就两三个，作者就是之一啊
网址不记得了，我只保存了本地网页，可以打开看看。作者博客的名字为【fjssharpsword的专栏】，再次感谢作者的无私奉献哦


	//邮箱账号
	if ((ComType.Find(TEXT("text"))>=0)  &&  ComVal=="请填写一个您想要申请的帐号" && ComName==TEXT(""))
	{
		Tmp=TEXT("qq");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}
	//1代表@qq.com,3代表@foxmail.com
	else	if ((ComType.Find(TEXT("select-one"))>=0)  &&ComName==TEXT("emailselect"))
	{ 
		Tmp=TEXT("3");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}
	//检测账号是否可用
	else	if ((ComType.Find(TEXT("button"))>=0)   &&ComName==TEXT(""))
	{ 
		//button的名字改为123456
		Tmp=TEXT("123456");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
		//点击按钮
		IHTMLElement*  pHElement;
		spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pHElement);
		pHElement->click();  
	}	
	//昵称
	else	if ((ComType.Find(TEXT("text"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{
		Tmp=TEXT("昵称");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}
	//年
	else	if ((ComType.Find(TEXT("select-one"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("2011");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//月
	else	if ((ComType.Find(TEXT("select-one"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("1");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}
	//日
	else	if ((ComType.Find(TEXT("select-one"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("1");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//ComVal==TEXT("1")1代表男，2代表女
	else	if ((ComType.Find(TEXT("radio"))>=0)  &&ComVal==TEXT("2")&& ComName==TEXT("gender"))
	{
		//点击选择
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
	//省
	else	if ((ComType.Find(TEXT("select-one"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("省");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//市
	else	if ((ComType.Find(TEXT("select-one"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("市");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//县
	else	if ((ComType.Find(TEXT("县"))>=0) &&  ComVal==TEXT("")  &&ComName==TEXT(""))
	{ 
		Tmp=TEXT("password");
		CComVariant vSetStatus(Tmp);
		spInputElement.PutPropertyByName(TEXT("value"),&vSetStatus);
	}	
	//验证码
	else	if ((ComType.Find(TEXT("text"))>=0) &&  ComVal==TEXT("请输入您在下图中看到的字符，不区分大小写")  &&ComName==TEXT("verifycode"))
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
//获得常规网页元素HTMLElement接口
hr = spInputElement->QueryInterface(IID_IHTMLElement,(void **)&pElem);
//试图获得超链接元素接口

IHTMLAnchorElement* pSubAnchor;
hr = spInputElement->QueryInterface(IID_IHTMLAnchorElement, (void **)&pSubAnchor);
if(hr == S_OK) {
BSTR bstr;
pElem->get_id(&bstr);
CString strID;
strID = bstr;
if(strID == "a_submit"){
MessageBox(TEXT("可以看看信息是否已经填上啦"));
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
	// TODO: 在此添加控件通知处理程序代码
	RegOrComment=1;//发表冰网日志
	SumbitPage();//自动提交注册信息

}
