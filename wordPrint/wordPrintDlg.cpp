// wordPrintDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "wordPrint.h"
#include "wordPrintDlg.h"


#include "CApplication.h"   
#include "CSelection.h"
#include "CDocuments.h"
#include "CDocument0.h"

#include "WINSPOOL.H"

#include <AtlBase.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CwordPrintDlg �Ի���




CwordPrintDlg::CwordPrintDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CwordPrintDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CwordPrintDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CwordPrintDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BEGIN, &CwordPrintDlg::OnBnClickedButton1)
	ON_WM_CLOSE()
	ON_WM_TIMER()
//	ON_WM_HSCROLL()
END_MESSAGE_MAP()


// CwordPrintDlg ��Ϣ�������

BOOL CwordPrintDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	
	getPrinter();//��ʼ����ӡ��
	readConfigure();

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CwordPrintDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CwordPrintDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CwordPrintDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}
BOOL CwordPrintDlg::printWord(CString Path,int count)//count��>��ӡ�ķ���
{
	CApplication	m_app;
	COleVariant covTrue((short)TRUE),
            covFalse((short)FALSE),
            covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	
	if (!m_app.CreateDispatch(_T("Word.Application"),NULL)) //����ʵ��
	{
		AfxMessageBox(_T("Couldn''t get Word object."));
		return FALSE;
	}
	m_app.put_Visible(TRUE);
	CDocuments m_docs(m_app.get_Documents());	//�½�һ���ĵ�
	
	CDocument0 doc;
	doc.AttachDispatch(m_docs.Open(COleVariant( Path ,VT_BSTR),
                    covFalse,    // Confirm Conversion.
                    covFalse,    // ReadOnly.
                    covFalse,    // AddToRecentFiles.
                    covOptional, // PasswordDocument.
                    covOptional, // PasswordTemplate.
                    covFalse,    // Revert.
                    covOptional, // WritePasswordDocument.
                    covOptional, // WritePasswordTemplate.
                    covOptional, // Format. // Last argument for Word 97
                    covOptional, // Encoding // New for Word 2000/2002
                    covTrue,     // Visible
                    covOptional, // OpenConflictDocument
                    covOptional, // OpenAndRepair
                    COleVariant((long)0),     // DocumentDirection wdDocumentDirection LeftToRight
                    covOptional  // NoEncodingDialog
                    )  // Close Open parameters
       ); // Close AttachDispatch(��)

	doc.PrintOut(covFalse,              // Background.
                        covOptional,           // Append.
                        covOptional,           // Range.
                        covOptional,           // OutputFileName.
                        covOptional,           // From.
                        covOptional,           // To.
                        covOptional,           // Item.
                        COleVariant((long)1),  // Copies.
                        COleVariant((long)2),//covOptional,           // Pages.
                        covOptional,           // PageType.
                        covOptional,           // PrintToFile.
                        covOptional,           // Collate.
                        covOptional,           // ActivePrinterMacGX.
                        covTrue,            // ManualDuplexPrint.
                        covOptional,           // PrintZoomColumn  New with Word 2002
                        covOptional,           // PrintZoomRow          ditto
                        covOptional,           // PrintZoomPaperWidth   ditto
                        covOptional);          // PrintZoomPaperHeight  ditto
	
	m_app.Quit(covFalse,  // SaveChanges.
                   covTrue,   // OriginalFormat.
                   covFalse   // RouteDocument.
                   );
	//CDocuments m_doc(m_app
	return FALSE;
}
void CwordPrintDlg::OnBnClickedButton1()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	
	printWord(_T("F:\\123.doc"),1);
	SetTimer(WM_MY_TIMEER,10000,NULL);

}

void CwordPrintDlg::getPrinter()
{

	CComboBox* pbox=((CComboBox*)GetDlgItem(IDC_PRINTER));
	pbox->ResetContent();

	DWORD Flags =  PRINTER_ENUM_FAVORITE |PRINTER_ENUM_LOCAL;   //local   printers  
    PRINTER_INFO_2* pPrinterEnum=new PRINTER_INFO_2;
    DWORD dwCount=0,dwBytes=0;
	TCHAR Name[500];  
	memset(Name,   0,   sizeof(TCHAR) * 500)   ;   
    if (!EnumPrinters(Flags,Name,2,(LPBYTE)pPrinterEnum,sizeof(PRINTER_INFO_2),&dwBytes,&dwCount))
    {
        if (pPrinterEnum) delete pPrinterEnum;
        pPrinterEnum=(PRINTER_INFO_2*)(new BYTE[dwBytes]);
        EnumPrinters(Flags,Name,2,(LPBYTE)pPrinterEnum,dwBytes,&dwBytes,&dwCount);
    }
	
	DWORD i = 0 ;
	for( i = 0 ; i < dwCount ; i++ ) {
		CString printerName = pPrinterEnum[ i ].pPrinterName ;
		pbox->AddString( printerName ) ;
	}

    if (pPrinterEnum) {delete []pPrinterEnum;}
}

void CwordPrintDlg::getExeUrl(CString *url)
{
	TCHAR exefileurl[MAX_PATH];
	GetModuleFileName(NULL,exefileurl,MAX_PATH);
	(_tcsrchr(exefileurl, _T('\\')))[1] = 0;
	*url=exefileurl;
}
void CwordPrintDlg::readConfigure()
{
	CString path;
	getExeUrl(&path);
	path+=_T("printer.ini");

	wchar_t name[256];
	wchar_t time[256];
	memset(name,0,256);
	memset(time,0,256);
	GetPrivateProfileString(_T("configure"),_T("printername"),_T(""),name,256,path);
	GetPrivateProfileString(_T("configure"),_T("timespace"),_T("60"),time,256,path);

	CComboBox* pbox=((CComboBox*)GetDlgItem(IDC_PRINTER));
	int oldSelect=pbox->FindString(0,name);
	if(oldSelect>=0)
	{
		pbox->SetCurSel(oldSelect);
	}
	

	CEdit* pedit=(CEdit*)GetDlgItem(IDC_TIMESPACE);
	pedit->SetWindowTextW((LPCTSTR)time);

}
void CwordPrintDlg::writeConfigure()
{
	CString path;
	getExeUrl(&path);
	path+=_T("printer.ini");

	CString printerName;
	CString	timespace;

	((CComboBox*)GetDlgItem(IDC_PRINTER))->GetWindowTextW(printerName);
	
	((CEdit*)GetDlgItem(IDC_TIMESPACE))->GetWindowTextW(timespace);
	WritePrivateProfileString(_T("configure"),_T("printername"),printerName,path);
	WritePrivateProfileString(_T("configure"),_T("timespace"),timespace,path);

}
void CwordPrintDlg::OnClose()
{
	// TODO: �ڴ������Ϣ�����������/�����Ĭ��ֵ
	writeConfigure();
	CDialog::OnClose();
	
}

void CwordPrintDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: �ڴ������Ϣ�����������/�����Ĭ��ֵ
	if(nIDEvent==WM_MY_TIMEER )
		TRACE0("qq\n");
	CDialog::OnTimer(nIDEvent);
}

//void CwordPrintDlg::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
//{
//	// TODO: �ڴ������Ϣ�����������/�����Ĭ��ֵ
//
//	CDialog::OnHScroll(nSBCode, nPos, pScrollBar);
//}
