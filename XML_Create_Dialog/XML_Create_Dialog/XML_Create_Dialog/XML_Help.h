#pragma once


// XML_Help �Ի���

class XML_Help : public CDialog
{
	DECLARE_DYNAMIC(XML_Help)

public:
	XML_Help(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~XML_Help();

// �Ի�������
	enum { IDD = IDD_DIALOG2 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
};
