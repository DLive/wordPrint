// wordPrintDlg.h : ͷ�ļ�
//

#pragma once


// CwordPrintDlg �Ի���
#define	WM_MY_TIMEER	WM_USER+100
class CwordPrintDlg : public CDialog
{
// ����
public:
	CwordPrintDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_WORDPRINT_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��

	BOOL printWord(CString Path,int count);


// ʵ��
public:
	void getExeUrl(CString *url);
	void readConfigure();
	void writeConfigure();
protected:
	HICON m_hIcon;
	
	void getPrinter();
	
	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnClose();
	afx_msg void OnTimer(UINT_PTR nIDEvent);
//	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
};
