// �����Ϳ������á������ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\OFFICE11\\MSWORD.OLB" no_namespace
// CCustomProperties ��װ��

class CCustomProperties : public COleDispatchDriver
{
public:
	CCustomProperties(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CCustomProperties(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCustomProperties(const CCustomProperties& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// CustomProperties ����
public:
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(LPCTSTR Name, LPCTSTR Value)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, Value);
		return result;
	}

	// CustomProperties ����
public:

};