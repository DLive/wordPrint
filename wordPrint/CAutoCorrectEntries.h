// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

#import "C:\\Program Files\\Microsoft Office\\OFFICE11\\MSWORD.OLB" no_namespace
// CAutoCorrectEntries 包装类

class CAutoCorrectEntries : public COleDispatchDriver
{
public:
	CAutoCorrectEntries(){} // 调用 COleDispatchDriver 默认构造函数
	CAutoCorrectEntries(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAutoCorrectEntries(const CAutoCorrectEntries& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// AutoCorrectEntries 方法
public:
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
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
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
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, Value);
		return result;
	}
	LPDISPATCH AddRichText(LPCTSTR Name, LPDISPATCH Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_DISPATCH ;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, Range);
		return result;
	}

	// AutoCorrectEntries 属性
public:

};
