// ParameterSelection.cpp : 实现文件
//

#include "stdafx.h"
#include "XML_Create_Dialog.h"
#include "ParameterSelection.h"
#include "afxdialogex.h"

extern int m_parameter;

// ParameterSelection 对话框

IMPLEMENT_DYNAMIC(ParameterSelection, CDialog)

ParameterSelection::ParameterSelection(CWnd* pParent /*=NULL*/)
	: CDialog(ParameterSelection::IDD, pParent)
{

	//  m_parameterselection = 0;
	m_parameterSelection = 0;
}

ParameterSelection::~ParameterSelection()
{
}

void ParameterSelection::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Radio(pDX, IDC_RADIO1, m_parameterSelection);
	DDV_MinMaxInt(pDX, m_parameterSelection, -1, 15);
}

BEGIN_MESSAGE_MAP(ParameterSelection, CDialog)
	ON_BN_CLICKED(IDOK, &ParameterSelection::OnBnClickedOk)
END_MESSAGE_MAP()


// ParameterSelection 消息处理程序


void ParameterSelection::OnBnClickedOk()
{
	UpdateData(TRUE);
	m_parameter = m_parameterSelection;

	CDialog::OnOK();
}


