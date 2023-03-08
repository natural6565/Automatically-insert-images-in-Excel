// pch.cpp: 미리 컴파일된 헤더에 해당하는 소스 파일

#include "pch.h"

CCriticalSection g_criticalExe;

CString INIReadString(CString strAppName, CString strKeyName, CString strFilePath)
{
	TCHAR szReturnString[1024] = { 0, };

	memset(szReturnString, NULL, 1024);
	GetPrivateProfileString(strAppName, strKeyName, _T(""), szReturnString, 1024, strFilePath);
	CString str;
	str.Format(_T("%s"), szReturnString);

	return str;
}

void INIWriteString(CString strAppName, CString strKeyName, CString strValue, CString strFilePath)
{
	WritePrivateProfileString(strAppName, strKeyName, strValue, strFilePath);
}

CString GetExePath()
{
	g_criticalExe.Lock();

	static TCHAR pBuf[256] = { 0, };
	memset(pBuf, NULL, sizeof(pBuf));
	GetModuleFileName(NULL, pBuf, sizeof(pBuf));

	CString strFilePath;
	
	strFilePath.Format(_T("%s"), pBuf);
	strFilePath = strFilePath.Left(strFilePath.ReverseFind(_T('\\')));
	g_criticalExe.Unlock();

	return strFilePath;
}


bool gv_bDropFile = true;
bool gv_bImageFolder = false, gv_bExcelFolder = false;
CString gv_InputFilePath;
CString gv_InputExcelFilePath;
CString gv_InIFilePath;
// 미리 컴파일된 헤더를 사용하는 경우 컴파일이 성공하려면 이 소스 파일이 필요합니다.
