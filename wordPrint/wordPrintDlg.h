// wordPrintDlg.h : 头文件
//

#pragma once


// CwordPrintDlg 对话框
#define	WM_MY_TIMEER	WM_USER+100
class CwordPrintDlg : public CDialog
{
// 构造
public:
	CwordPrintDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_WORDPRINT_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持

	BOOL printWord(CString Path,int count);


// 实现
public:
	void getExeUrl(CString *url);
	void readConfigure();
	void writeConfigure();
protected:
	HICON m_hIcon;
	
	void getPrinter();
	
	// 生成的消息映射函数
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
