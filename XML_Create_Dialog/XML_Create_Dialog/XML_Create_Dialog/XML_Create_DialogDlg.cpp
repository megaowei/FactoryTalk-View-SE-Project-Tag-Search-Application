
// XML_Create_DialogDlg.cpp : implementation file
//

#include "stdafx.h"
#include "XML_Create_Dialog.h"
#include "XML_Create_DialogDlg.h"
#include "afxdialogex.h"
#include "XML_Help.h"
#include "ParameterSelection.h"

using namespace std;
#import <msxml3.dll>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

char XMLFileName[500]="\0";
char szXmlFile[500]="\0";
char FilePath[500]="\0";
int m_parameter = -1;

// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CXML_Create_DialogDlg dialog




CXML_Create_DialogDlg::CXML_Create_DialogDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CXML_Create_DialogDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CXML_Create_DialogDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_EDIT1, m_XML_Files);
	DDX_Control(pDX, IDC_EDIT2, m_Merged_XML);
}

BEGIN_MESSAGE_MAP(CXML_Create_DialogDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CXML_Create_DialogDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CXML_Create_DialogDlg::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON3, &CXML_Create_DialogDlg::OnBnClickedButton3)
	ON_BN_CLICKED(IDC_BUTTON4, &CXML_Create_DialogDlg::OnBnClickedButton4)
	ON_BN_CLICKED(IDC_BUTTON5, &CXML_Create_DialogDlg::OnBnClickedButton5)
	ON_BN_CLICKED(IDC_BUTTON6, &CXML_Create_DialogDlg::OnBnClickedButton6)
END_MESSAGE_MAP()


// CXML_Create_DialogDlg message handlers

BOOL CXML_Create_DialogDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
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

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon



	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CXML_Create_DialogDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CXML_Create_DialogDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CXML_Create_DialogDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CXML_Create_DialogDlg::OnBnClickedButton1()
{
	TCHAR pszPath[MAX_PATH];
	CString test;
	int strL;
    BROWSEINFO bi;
	CString m_XML_Files_Path;

    bi.hwndOwner      = this->GetSafeHwnd();  
    bi.pidlRoot       = NULL;  
    bi.pszDisplayName = NULL;   
    bi.lpszTitle      = TEXT("Please select folder");   
    bi.ulFlags        = BIF_NEWDIALOGSTYLE | BIF_DONTGOBELOWDOMAIN | BIF_BROWSEFORCOMPUTER | BIF_RETURNONLYFSDIRS | BIF_RETURNFSANCESTORS;  
    bi.lpfn           = NULL;   
    bi.lParam         = 0;  
    bi.iImage         = 0;   
  
    LPITEMIDLIST pidl = SHBrowseForFolder(&bi);  
    if (pidl == NULL)  
    {  
        return;  
    }  
  
	else   
    {  
       memset(pszPath, 0, sizeof(pszPath));
       SHGetPathFromIDList(pidl, pszPath);
	   m_XML_Files_Path = pszPath;

	   test = m_XML_Files_Path;
	   strL=WideCharToMultiByte(CP_ACP,0,m_XML_Files_Path,test.GetLength(),NULL,0,NULL,NULL);
	   WideCharToMultiByte(CP_ACP,0,m_XML_Files_Path,test.GetLength(),FilePath,strL,NULL,NULL);

	   SetDlgItemText(IDC_EDIT1,m_XML_Files_Path);

	   strcpy_s(XMLFileName,FilePath);

	   strcat_s(XMLFileName,"\\*.xml\0");

    }
	
}


void CXML_Create_DialogDlg::OnBnClickedButton2()
{
	BROWSEINFO broInfo = {0};
    TCHAR szDisName[MAX_PATH] = {0};
	CString test;
	CString m_XML_File_Path;
	int strL;


    broInfo.hwndOwner = this->m_hWnd;
    broInfo.pidlRoot  = NULL;
    broInfo.pszDisplayName = szDisName;
    broInfo.lpszTitle = _T("Please select folder");
    broInfo.ulFlags   = BIF_NEWDIALOGSTYLE | BIF_DONTGOBELOWDOMAIN | BIF_BROWSEFORCOMPUTER | BIF_RETURNONLYFSDIRS | BIF_RETURNFSANCESTORS;
    broInfo.lpfn      = NULL;
    broInfo.lParam    = NULL;
    broInfo.iImage    = IDR_MAINFRAME;


    LPITEMIDLIST pIDList = SHBrowseForFolder(&broInfo);
    if (pIDList != NULL)
    {
       memset(szDisName, 0, sizeof(szDisName));
       SHGetPathFromIDList(pIDList, szDisName);
	   m_XML_File_Path = szDisName;

	   test = m_XML_File_Path;
	   strL=WideCharToMultiByte(CP_ACP,0,m_XML_File_Path,test.GetLength(),NULL,0,NULL,NULL);
	   WideCharToMultiByte(CP_ACP,0,m_XML_File_Path,test.GetLength(),szXmlFile,strL,NULL,NULL);
	   strcat(szXmlFile,"\\Search_XML_File.xml\0");

	   SetDlgItemText(IDC_EDIT2,m_XML_File_Path);

    }
	
}


void CXML_Create_DialogDlg::OnBnClickedButton3()
{
	ProgressBar_XML dlg;

	dlg.DoModal();

	

	return;
}




void CXML_Create_DialogDlg::OnBnClickedButton4()
{
	XML_Help dlg;
	dlg.DoModal();

}


void CXML_Create_DialogDlg::OnBnClickedButton5()
{
	CAboutDlg dlgAbout;
	dlgAbout.DoModal();
}


void CXML_Create_DialogDlg::OnBnClickedButton6()
{
	ParameterSelection dlgParameter;
	dlgParameter.DoModal();
}
