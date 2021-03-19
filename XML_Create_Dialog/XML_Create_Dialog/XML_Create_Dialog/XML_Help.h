#pragma once


// XML_Help 对话框

class XML_Help : public CDialog
{
	DECLARE_DYNAMIC(XML_Help)

public:
	XML_Help(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~XML_Help();

// 对话框数据
	enum { IDD = IDD_DIALOG2 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
};
