// wordPrint.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CwordPrintApp:
// �йش����ʵ�֣������ wordPrint.cpp
//

class CwordPrintApp : public CWinApp
{
public:
	CwordPrintApp();

// ��д
	public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CwordPrintApp theApp;