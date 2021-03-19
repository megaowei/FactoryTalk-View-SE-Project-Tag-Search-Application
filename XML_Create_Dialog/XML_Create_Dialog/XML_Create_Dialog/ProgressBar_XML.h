#pragma once


// ProgressBar_XML �Ի���
static int num_XML_Files = 0;
extern int num_Process ;



UINT ThreadForCreateXML(LPVOID pParm);
UINT ThreadForProgressBar(LPVOID pParm);

class ProgressBar_XML : public CDialog
{
	DECLARE_DYNAMIC(ProgressBar_XML)

public:
	ProgressBar_XML(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~ProgressBar_XML();

// �Ի�������
	enum { IDD = IDD_DIALOG1 };

protected:
	HICON m_hIcon;
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��
	virtual BOOL OnInitDialog();

	DECLARE_MESSAGE_MAP()
public:
	CProgressCtrl m_ProgressBar;
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
	short Num_XMLFile;
	CWinThread *m_pThread_XML;
	CWinThread *m_pThread_ProgressBar;
	afx_msg void OnBnClickedCancel();
	char XMLFileName_P[500];
	char szXmlFile_P[500];
	char FilePath_P[500];
	CString ProgressStr;
};
