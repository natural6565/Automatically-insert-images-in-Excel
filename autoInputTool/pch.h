// pch.h: 미리 컴파일된 헤더 파일입니다.
// 아래 나열된 파일은 한 번만 컴파일되었으며, 향후 빌드에 대한 빌드 성능을 향상합니다.
// 코드 컴파일 및 여러 코드 검색 기능을 포함하여 IntelliSense 성능에도 영향을 미칩니다.
// 그러나 여기에 나열된 파일은 빌드 간 업데이트되는 경우 모두 다시 컴파일됩니다.
// 여기에 자주 업데이트할 파일을 추가하지 마세요. 그러면 성능이 저하됩니다.
#pragma once

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN
#endif

#include "targetver.h"

#define _ATL__CSTRING_EXPLICIT_CONSTRUCTORS

#define _AFX_ALL_WARNINGS

#ifndef PCH_H
#define PCH_H

// 여기에 미리 컴파일하려는 헤더 추가
#include "framework.h"
#include <afxcmn.h>

#endif //PCH_H
#include <afxwin.h>
#include <afxext.h>

#include <afxdisp.h>

void INIWriteString(CString strAppName, CString strKeyName, CString strValue, CString strFilePath);
CString INIReadString(CString strAppName, CString strKeyName, CString strFilePath);

CString GetExePath();

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>
#endif
#ifndef _AFX_NO_AFXCWN_SUPPORT
#include <afxcmn.h>
#endif

#include <afxcontrolbars.h>

extern bool gv_bDropFile;
extern CString gv_InputFilePath;
extern CString gv_InputExcelFilePath;
extern CString gv_InIFilePath;
extern bool gv_bImageFolder;
extern bool gv_bExcelFolder;
