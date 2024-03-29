// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

#import "C:\\Program Files\\Microsoft Office\\OFFICE11\\MSWORD.OLB" no_namespace
// CDocumentEvents2 包装类

class CDocumentEvents2 : public COleDispatchDriver
{
public:
	CDocumentEvents2(){} // 调用 COleDispatchDriver 默认构造函数
	CDocumentEvents2(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDocumentEvents2(const CDocumentEvents2& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// DocumentEvents2 方法
public:
	void New()
	{
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Open()
	{
		InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Close()
	{
		InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Sync(long SyncEventType)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SyncEventType);
	}
	void XMLAfterInsert(LPDISPATCH NewXMLNode, BOOL InUndoRedo)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BOOL ;
		InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NewXMLNode, InUndoRedo);
	}
	void XMLBeforeDelete(LPDISPATCH DeletedRange, LPDISPATCH OldXMLNode, BOOL InUndoRedo)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_BOOL ;
		InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DeletedRange, OldXMLNode, InUndoRedo);
	}

	// DocumentEvents2 属性
public:

};
