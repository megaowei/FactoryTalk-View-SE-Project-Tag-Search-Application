#pragma once


// ParameterSelection �Ի���

class ParameterSelection : public CDialog
{
	DECLARE_DYNAMIC(ParameterSelection)

public:
	ParameterSelection(CWnd* pParent = NULL);   // ��׼���캯��
	virtual ~ParameterSelection();

// �Ի�������
	enum { IDD = IDD_DIALOG3 };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

	DECLARE_MESSAGE_MAP()
public:
//	int m_parameterselection;
	int m_parameterSelection;
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedRadio14();
};
