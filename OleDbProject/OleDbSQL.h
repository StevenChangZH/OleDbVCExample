#pragma once

// Ole Db Operation class
class OleDbSQL
{
public:
	OleDbSQL();
	~OleDbSQL();

	HRESULT Initialize();			// Initialize COM environment
	HRESULT DataConnect();			// Connect database( SQL Server 12 Express)
	HRESULT QueryExcute(TCHAR* pSQL);// Execute SQL query
	HRESULT ReadData();				// Transform data and print out
	HRESULT DataDisconnect();		// Close database connection

	CComPtr<IRowset> pIRowset;		// rowset

private:
	CComPtr<IDBInitialize>     pIDBInitialize;// DB COM Interface
	CComPtr<IDBCreateSession>  pSession;// session object
	CComPtr<ICommandText>      pICommandText;//commandtext object
};





