// OleDbProject.cpp : console interface
//

#include "stdafx.h"
#include "OleDbProject.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// Application object

CWinApp theApp;

using namespace std;

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	int nRetCode = 0;

	HMODULE hModule = ::GetModuleHandle(NULL);

	if (hModule != NULL)
	{
		// Initialize MFC failed
		if (!AfxWinInit(hModule, NULL, ::GetCommandLine(), 0))
		{
			_tprintf(_T("Error: Initialize MFC Failed\n"));
			nRetCode = 1;
		}
		else
		{
			// Call OLE DB methods
			HRESULT hResult = S_OK;
			OleDbSQL* dbConn = new OleDbSQL();
			hResult = dbConn->Initialize();
			hResult = dbConn->DataConnect();
			if (hResult==S_OK)
				printf("Connection success\n");

			//sql
			TCHAR* pSQL = _T("select * from airport");
			hResult = dbConn->QueryExcute(pSQL);
			if (hResult==S_OK)
				printf("Execution success\n");
			dbConn->ReadData();
			dbConn->DataDisconnect();
		}
	}
	else
	{
		_tprintf(_T("Error: GetModuleHandle failed\n"));
		nRetCode = 1;
	}

	
	return nRetCode;
}
