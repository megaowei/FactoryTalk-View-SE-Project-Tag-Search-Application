
// XML_Create_DialogDlg.h : header file
//

#pragma once
#include "ProgressBar_XML.h"

extern char XMLFileName[500];
extern char szXmlFile[500];
extern char FilePath[500];
extern int m_parameter;

// CXML_Create_DialogDlg dialog
class CXML_Create_DialogDlg : public CDialogEx
{
// Construction
public:
	CXML_Create_DialogDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_XML_CREATE_DIALOG_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	CEdit m_XML_Files;
	CEdit m_Merged_XML;
	
	afx_msg void OnBnClickedButton1();

public:
	afx_msg void OnBnClickedButton2();
public:
	afx_msg void OnBnClickedButton3();
	afx_msg void OnBnClickedButton4();
	afx_msg void OnBnClickedButton5();
	afx_msg void OnBnClickedButton6();
};
