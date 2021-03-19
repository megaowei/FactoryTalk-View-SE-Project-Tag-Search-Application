#pragma once


// ParameterSelection 对话框

class ParameterSelection : public CDialog
{
	DECLARE_DYNAMIC(ParameterSelection)

public:
	ParameterSelection(CWnd* pParent = NULL);   // 标准构造函数
	virtual ~ParameterSelection();

// 对话框数据
	enum { IDD = IDD_DIALOG3 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

	DECLARE_MESSAGE_MAP()
public:
//	int m_parameterselection;
	int m_parameterSelection;
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedRadio14();
};
