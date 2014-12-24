#include "OleDbSQL.h"
#include "stdafx.h"

OleDbSQL::OleDbSQL()
{}

OleDbSQL::~OleDbSQL()
{}

HRESULT OleDbSQL::Initialize()
{
	// Initialize COM Environment
	HRESULT hResult;
	hResult = CoInitializeEx( NULL, COINIT_MULTITHREADED );
	return hResult;
}

HRESULT OleDbSQL::DataConnect()
{
	TCHAR szInitStr[1024];// Initialize connection string
	TCHAR pszServerName[64] = _T("localhost\\SQLEXPRESS");//Data source
	TCHAR pszDataSource[64] = _T("master");//Database name

	// Assemble connection string
	// I've created a database named master in SQLEXPRESS service.
	swprintf( szInitStr, 1024, _T("Provider=SQLNCLI11;Data Source=%s;Initial Catalog=%s;Integrated Security=SSPI"), 
		pszServerName, pszDataSource);

	CComPtr<IDataInitialize> m_pIDataInitialize;// OleDb data COM Interface
	// Call CoCreateInstance to create IDataInitialize instance
	HRESULT hResult = ::CoCreateInstance( CLSID_MSDAINITIALIZE, NULL,
		CLSCTX_INPROC_SERVER, IID_IDataInitialize, ( void** )&m_pIDataInitialize );
	if( FAILED( hResult ) )
        return E_FAIL;

	// Call IDataInitialize GetDataSource method to acquire IDBInitialize instance
    hResult = m_pIDataInitialize->GetDataSource( NULL, CLSCTX_INPROC_SERVER, 
		( LPCOLESTR )szInitStr, IID_IDBInitialize, ( IUnknown** )&pIDBInitialize);
    if( FAILED( hResult ) )
        return E_FAIL;
	
	// Call IDBInitialize Initialize method to initialize this instance
    hResult = pIDBInitialize->Initialize();
    if( FAILED( hResult ) )
        return E_FAIL;

	// Use interface IDBCreateSession to create session instance
	hResult = pIDBInitialize->QueryInterface( IID_IDBCreateSession, 
		( void** )&pSession );
    if( FAILED( hResult ) )
        return E_FAIL;

	// Call CreateSession method to create createcommand instance 
	CComPtr<IDBCreateCommand> m_pIDBCreateCommand;
	hResult = pSession->CreateSession( NULL, IID_IDBCreateCommand,
		(IUnknown**)&m_pIDBCreateCommand );
	if ( FAILED( hResult ) )
		return E_FAIL;

	// Commandtext interface 
	hResult = m_pIDBCreateCommand->CreateCommand(NULL,IID_ICommandText,
		(IUnknown**)&pICommandText);
	if ( FAILED( hResult ) )
		return E_FAIL;

    return S_OK;
}


HRESULT OleDbSQL::QueryExcute(TCHAR* pSQL)
{
	HRESULT hResult = S_OK;

	// Construct sql query and execute
	//TCHAR* pSQL = _T("create table airport(code varchar(10), name varchar(45))");
	hResult = pICommandText->SetCommandText(DBGUID_DEFAULT,pSQL);
	if ( FAILED( hResult ) )
		return E_FAIL;
							  
	hResult = pICommandText->Execute(NULL,IID_IRowset,NULL,NULL,(IUnknown**)&pIRowset);
	if ( FAILED( hResult ) )
		return E_FAIL;

	return S_OK;
}

HRESULT OleDbSQL::ReadData()
{
	HRESULT hResult = S_OK;

	// Create data bindings
	DBCOLUMNINFO* rgColumnInfo = NULL;
	LPWSTR pStringBuffer = NULL;
	CComPtr<IColumnsInfo> pIColumnsInfo;
	ULONG iCol;
	ULONG cColumns = 0;
	ULONG dwOffset = 0;
	DBBINDING* rgBindings = NULL;
	
	// Get column information
	// If succeeded, rgColumnInfo would contain an array with column data constructing information
	// Then cColumns is column number
	pIRowset->QueryInterface(IID_IColumnsInfo,(void**)&pIColumnsInfo);
	pIColumnsInfo->GetColumnInfo(&cColumns, &rgColumnInfo,&pStringBuffer);
	rgBindings = (DBBINDING*)HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY,cColumns * sizeof(DBBINDING));

 	// Create bindings
	for( iCol = 0; iCol < cColumns; iCol++ ) {

		rgBindings[iCol].iOrdinal	= rgColumnInfo[iCol].iOrdinal;
		rgBindings[iCol].dwPart		= DBPART_VALUE|DBPART_LENGTH|DBPART_STATUS;
		rgBindings[iCol].obStatus   = dwOffset;
		rgBindings[iCol].obLength   = dwOffset + sizeof(DBSTATUS);
		rgBindings[iCol].obValue    = dwOffset+sizeof(DBSTATUS)+sizeof(ULONG);
		rgBindings[iCol].dwMemOwner = DBMEMOWNER_CLIENTOWNED;
		rgBindings[iCol].eParamIO   = DBPARAMIO_NOTPARAM;
		rgBindings[iCol].bPrecision = rgColumnInfo[iCol].bPrecision;
		rgBindings[iCol].bScale     = rgColumnInfo[iCol].bScale;
		rgBindings[iCol].wType      = rgColumnInfo[iCol].wType;
		rgBindings[iCol].cbMaxLen   = rgColumnInfo[iCol].ulColumnSize;
		dwOffset = rgBindings[iCol].cbMaxLen + rgBindings[iCol].obValue;
		dwOffset = ROUNDUP(dwOffset);
	}

	// Release all temporary interfaces
	CoTaskMemFree(rgColumnInfo);
	CoTaskMemFree(pStringBuffer);
	pIColumnsInfo = NULL;

	// Accessor
	HACCESSOR phAccessor;
	CComPtr<IAccessor> pIAccessor;
	pIRowset->QueryInterface(IID_IAccessor,(void**)&pIAccessor);
	pIAccessor->CreateAccessor( DBACCESSOR_ROWDATA,cColumns,rgBindings,0,&phAccessor,NULL);
	pIAccessor = NULL;

	// Reading data
	void* pData = NULL;
	ULONG cRowsObtained;
	HROW* rghRows = NULL;
	ULONG iRow;
	LONG cRows = 1;// one row each time
	void* pCurData;
			
	// Create column data buffer and copy cRows rows here
	pData = HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY, dwOffset * cRows);
	// Create Data names
	// You can create a dynamic naming array for it, which means you can provide bindings
	// automatically.
	char* code = new char((int)rgBindings[0].cbMaxLen);
	char* name = new char((int)rgBindings[1].cbMaxLen);

	while( S_OK == pIRowset->GetNextRows(DB_NULL_HCHAPTER,0,cRows,&cRowsObtained, &rghRows) ) {

		// actually every time we operate on cRowsObtained rows, cRows total.
		for( iRow = 0; iRow < cRowsObtained; iRow++ ) {
			pCurData = (char*)pData + (dwOffset * iRow);
			pIRowset->GetData(rghRows[iRow],phAccessor,pCurData);

			// get code	value
			memcpy(code, (char*)pCurData+(int)rgBindings[0].obValue, (int)rgBindings[0].cbMaxLen);
			for(int i=0;i<(int)rgBindings[0].cbMaxLen;i++)
				printf("%c", *(code+i));
			printf("\t");

			// get name	value
			memcpy(name, (char*)pCurData+(int)rgBindings[1].obValue, (int)rgBindings[1].cbMaxLen);
			for(int i=0;i<(int)rgBindings[1].cbMaxLen;i++)
				printf("%c", *(name+i));
			printf("\n");
		}

		if( cRowsObtained ) {// Release column handle array
			pIRowset->ReleaseRows(cRowsObtained,rghRows,NULL,NULL,NULL);
		}
		CoTaskMemFree(rghRows);
		rghRows = NULL;
	}
			
	HeapFree(GetProcessHeap(),0,pData);

	return S_OK;
}


HRESULT OleDbSQL::DataDisconnect()
{
	HRESULT hResult = S_OK;

	this->pIRowset = NULL;
	this->pICommandText = NULL;
	this->pSession = NULL;
	this->pIDBInitialize = NULL;
	return S_OK;
}