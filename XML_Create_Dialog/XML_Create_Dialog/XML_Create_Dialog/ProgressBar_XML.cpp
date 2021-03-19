// ProgressBar_XML.cpp : 实现文件
//

#include "stdafx.h"
#include "XML_Create_Dialog.h"
#include "ProgressBar_XML.h"
#include "afxdialogex.h"


using namespace std;
#import <msxml3.dll>

extern char XMLFileName[500];
extern char szXmlFile[500];
extern char FilePath[500];
extern int m_parameter;

int num_Process =1;

#define CHECK_AND_RELEASE(pInterface)  \
    if(pInterface) \
    {\
        pInterface->Release();\
        pInterface = NULL;\
    }\


// ProgressBar_XML 对话框

IMPLEMENT_DYNAMIC(ProgressBar_XML, CDialog)

ProgressBar_XML::ProgressBar_XML(CWnd* pParent /*=NULL*/)
	: CDialog(ProgressBar_XML::IDD, pParent)
{

}

ProgressBar_XML::~ProgressBar_XML()
{
}

void ProgressBar_XML::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_PROGRESS1, m_ProgressBar);
}


BEGIN_MESSAGE_MAP(ProgressBar_XML, CDialog)
	ON_WM_CREATE()
	ON_WM_SHOWWINDOW()
	ON_BN_CLICKED(IDCANCEL, &ProgressBar_XML::OnBnClickedCancel)
END_MESSAGE_MAP()


// ProgressBar_XML 消息处理程序


int ProgressBar_XML::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CDialog::OnCreate(lpCreateStruct) == -1)
		return -1;

	Num_XMLFile=0;
	num_XML_Files = 1;

	return 0;
}


void ProgressBar_XML::OnShowWindow(BOOL bShow, UINT nStatus)
{
	CDialog::OnShowWindow(bShow, nStatus);

	// TODO: ÔÚ´Ë´¦Ìí¼ÓÏûÏ¢´¦Àí³ÌÐò´úÂë
}

BOOL ProgressBar_XML::OnInitDialog()
{
	CDialog::OnInitDialog();

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
	//AfxMessageBox(_T("Please check the folder path of XML files based on Help page!"),MB_OK|MB_ICONERROR);

    try
   {
    long handle; //用于查找的句柄
    struct _finddata_t fileinfo;  //文件信息的结构体

    handle=_findfirst(XMLFileName,&fileinfo);         //第一次查找
    if(-1==handle)  
	{
		AfxMessageBox(_T("Please check the folder path of XML files !"),MB_OK|MB_ICONERROR);
	    this->EndDialog(true);
		return -1;
	}
	else
	{
		
	    Num_XMLFile=Num_XMLFile+1;

	    while(!_findnext(handle,&fileinfo))
	    {
		   Num_XMLFile=Num_XMLFile+1;
	    }

	    m_ProgressBar.SetRange(1,Num_XMLFile);

	    m_pThread_XML = AfxBeginThread(ThreadForCreateXML,this);
	    m_pThread_ProgressBar = AfxBeginThread(ThreadForProgressBar,this);


	    if ( m_pThread_XML ==NULL || m_pThread_ProgressBar==NULL)
	    {
		     AfxMessageBox(_T("Please check the compatibility of software!"),MB_OK|MB_ICONERROR);
	    }
   }
   }

   catch(...)
  {
	AfxMessageBox(_T("Please check the problem based on Help page!"),MB_OK|MB_ICONERROR);
	this->EndDialog(true);
   }

	return TRUE;
}



void ProgressBar_XML::OnBnClickedCancel()
{
	num_Process = 0;
	Sleep(300);
	CDialog::OnCancel();

	return;
}
UINT ThreadForCreateXML(LPVOID pParm)
{
	CoInitialize(NULL);
	
	ProgressBar_XML *pXML = (ProgressBar_XML *)pParm;

	BOOL bResult = FALSE;
    IXMLDOMDocument *pIXMLDOMDocument=NULL;//,*spXMLDoc=NULL;
    IXMLDOMElement *pIXMLDOMElement=NULL;
	IXMLDOMElement *pElement=NULL;
    IXMLDOMProcessingInstruction *pIXMLDOMProcessingInstruction=NULL;
    IXMLDOMNode *pIXMLDOMNode = NULL, *TagNode=NULL, *AttributeNode=NULL;
	//IXMLDOMNodeListPtr spNodeList=NULL;
    HRESULT hr ;
    BSTR Tagname=(_bstr_t)"parameter";//bstrValue,
	_variant_t TagName, ScreenName;
	long numtag=0, numfile=0;
	char xmlFilepath[1000];
	char *parameterSele=NULL;

	wstring strFindText (_T("parameter"));

    MSXML2::IXMLDOMDocumentPtr spXMLDoc;
    spXMLDoc.CreateInstance(__uuidof(MSXML2::DOMDocument30));

    MSXML2::IXMLDOMElementPtr spRoot = NULL;
    MSXML2::IXMLDOMNodeListPtr spNodeList = NULL;
    MSXML2::IXMLDOMNamedNodeMapPtr spNameNodeMap=NULL; 

    long handle; //用于查找的句柄
    struct _finddata_t fileinfo;//文件信息的结构体
try
    
{    
	switch (m_parameter)
	{
	case 0:
		parameterSele = "#1";
		break;
	case 1:
		parameterSele = "#2";
		break;
	case 2:
		parameterSele = "#3";
		break;
	case 3:
		parameterSele = "#4";
		break;
	case 4:
		parameterSele = "#5";
		break;
	case 5:
		parameterSele = "#6";
		break;
	case 6:
		parameterSele = "#100";
		break;
	case 7:
		parameterSele = "#101";
		break;
	case 8:
		parameterSele = "#102";
		break;
	case 9:
		parameterSele = "#103";
		break;
	case 10:
		parameterSele = "#200";
		break;
	case 11:
		parameterSele = "#201";
		break;
	case 12:
		parameterSele = "#202";
		break;
	case 13:
		parameterSele = "#104";
		break;
	case 14:
		parameterSele = "#105";
		break;
	default:
		parameterSele = "#106";
		break;
	}
	
	
	hr=CoCreateInstance(CLSID_DOMDocument, NULL, CLSCTX_SERVER, IID_IXMLDOMDocument, (LPVOID*)(&pIXMLDOMDocument));
        SUCCEEDED(hr) ? 0 : throw hr;

        if(pIXMLDOMDocument)
        {
            hr=pIXMLDOMDocument->createElement((_bstr_t)(char*)"TagGroup", &pIXMLDOMElement);
            if(SUCCEEDED(hr) && pIXMLDOMElement)
            {

                if(SUCCEEDED(hr))
                {
                    hr=pIXMLDOMDocument->createProcessingInstruction(_T("xml"), _T("version='1.0'"), &pIXMLDOMProcessingInstruction);
                    if(SUCCEEDED(hr) && pIXMLDOMProcessingInstruction)
                    {
                      pIXMLDOMDocument->appendChild(pIXMLDOMProcessingInstruction, &pIXMLDOMNode);
                      pIXMLDOMDocument->putref_documentElement(pIXMLDOMElement);

//////////For scanning the first XML files//////////////////////////////////////////////////////////////////////////////////////////////////////
					  handle=_findfirst(XMLFileName,&fileinfo);         //Find the first xml file and get the info about that
		              memset(xmlFilepath,0,1000);

		              numfile=numfile+1;
				      num_XML_Files = numfile;
					  Sleep(100);

		              strcpy(xmlFilepath, FilePath);
				      strcat(xmlFilepath,"\\\0");
		              strcat(xmlFilepath,fileinfo.name);
		              strcat(xmlFilepath,"\0");

                      spXMLDoc->load((_bstr_t)xmlFilepath);

		              spRoot = spXMLDoc->documentElement;

		              spNodeList = spRoot->selectNodes("//parameter");

					  for (long i = 0; i != spNodeList->length; ++i)
                      {//search all named tags

                         spNameNodeMap=spNodeList->item[i]->attributes;

						 for (long j=0;j!=spNameNodeMap->length;++j)
						 {

						 if (strcmp((char*)(_bstr_t)spNameNodeMap->item[j]->nodeName , "name")==0 && strcmp((char*)(_bstr_t)spNameNodeMap->item[j]->nodeValue,parameterSele)==0)
	                     {
							 TagName = spNameNodeMap->item[2]->nodeValue;
		                     ScreenName.SetString(fileinfo.name);
							 hr=pIXMLDOMDocument->createElement((_bstr_t)(char*)"Tag", &pElement);
	                         pElement->setAttribute((_bstr_t)(char*)"Name", TagName);
	                         pElement->setAttribute((_bstr_t)(char*)"ScreenName", ScreenName);

					         pIXMLDOMElement->appendChild((IXMLDOMNode*)pElement, &pIXMLDOMNode);

							 break;
	                     }
						 
						 }
	
                       }

					   pIXMLDOMDocument->save((_variant_t)szXmlFile);

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


					      while(!_findnext(handle,&fileinfo)&&num_Process)               //循环查找其他符合的文件，知道找不到其他的为止
                          {
		                      memset(xmlFilepath,0,1000);

		                      numfile=numfile+1;
							  num_XML_Files = numfile;


							  pXML->ProgressStr.Format(_T("Merging XML files %d of %d : %s"),num_XML_Files,pXML->Num_XMLFile,CStringW(fileinfo.name));

							   Sleep(100);

		                       strcpy(xmlFilepath, FilePath);
							   strcat(xmlFilepath,"\\\0");
		                       strcat(xmlFilepath,fileinfo.name);
		                       strcat(xmlFilepath,"\0");

                               spXMLDoc->load((_bstr_t)xmlFilepath);

		                       spRoot = spXMLDoc->documentElement;

		                       spNodeList = spRoot->selectNodes("//parameter");


					  for (long i = 0; i != spNodeList->length; ++i)
                      {

                         spNameNodeMap=spNodeList->item[i]->attributes;

						 for (long j=0;j!=spNameNodeMap->length;++j)
						 {

						 if (strcmp((char*)(_bstr_t)spNameNodeMap->item[j]->nodeName , "name")==0 && strcmp((char*)(_bstr_t)spNameNodeMap->item[j]->nodeValue,parameterSele)==0)
	                     {
							 TagName = spNameNodeMap->item[2]->nodeValue;
		                     ScreenName.SetString(fileinfo.name);
							 hr=pIXMLDOMDocument->createElement((_bstr_t)(char*)"Tag", &pElement);
	                         pElement->setAttribute((_bstr_t)(char*)"Name", TagName);
	                         pElement->setAttribute((_bstr_t)(char*)"ScreenName", ScreenName);

					         pIXMLDOMElement->appendChild((IXMLDOMNode*)pElement, &pIXMLDOMNode);

							 break;
	                     }
						 
						 }
	
                       }

					   pIXMLDOMDocument->save((_variant_t)szXmlFile);

                    }
                }
			    CHECK_AND_RELEASE(pIXMLDOMNode);
                CHECK_AND_RELEASE(pIXMLDOMProcessingInstruction);
                
            }

		}
		CHECK_AND_RELEASE(pIXMLDOMElement);
		}

	    CHECK_AND_RELEASE(pIXMLDOMDocument);
    
	    CHECK_AND_RELEASE(pElement);
	    CHECK_AND_RELEASE(TagNode);
	    CHECK_AND_RELEASE(AttributeNode);

		if (pXML->Num_XMLFile <= (numfile+1))
		{
			AfxMessageBox(_T("Successfully!"),MB_OK);
			SetDlgItemText(pXML->m_hWnd,IDCANCEL,_T("Finished"));
		}

}
	
catch(...)
{
        CHECK_AND_RELEASE(pIXMLDOMElement);
        CHECK_AND_RELEASE(pIXMLDOMDocument);
        CHECK_AND_RELEASE(pIXMLDOMNode);
        CHECK_AND_RELEASE(pIXMLDOMProcessingInstruction);
	    AfxMessageBox(_T("Please check the folder path of XML files based on Help page!"),MB_OK|MB_ICONERROR);
        //DisplayErrorToUser();
}

    _findclose(handle); 

	CoInitialize(NULL);

	return 0;
}
UINT ThreadForProgressBar(LPVOID pParm)
{ 
	ProgressBar_XML *pXML = (ProgressBar_XML *)pParm;
	
	while(num_Process)
	{
	   pXML ->m_ProgressBar.SetPos(num_XML_Files);
	   pXML ->SetDlgItemText(IDC_STATIC,pXML ->ProgressStr);
	   Sleep(100);
	}
	return 0;
}