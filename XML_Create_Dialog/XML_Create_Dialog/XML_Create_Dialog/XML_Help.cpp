// XML_Help.cpp : 实现文件
//

#include "stdafx.h"
#include "XML_Create_Dialog.h"
#include "XML_Help.h"
#include "afxdialogex.h"


// XML_Help 对话框

IMPLEMENT_DYNAMIC(XML_Help, CDialog)

XML_Help::XML_Help(CWnd* pParent /*=NULL*/)
	: CDialog(XML_Help::IDD, pParent)
{

}

XML_Help::~XML_Help()
{
}

void XML_Help::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(XML_Help, CDialog)
END_MESSAGE_MAP()


// XML_Help 消息处理程序
