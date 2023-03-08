// CListBoxFileDrop.cpp: 구현 파일
//

#include "pch.h"
#include "autoInputTool.h"
#include "autoInputToolDlg.h"
#include "CListBoxFileDrop.h"


// CListBoxFileDrop

IMPLEMENT_DYNAMIC(CListBoxFileDrop, CListBox)

CListBoxFileDrop::CListBoxFileDrop()
{

}

CListBoxFileDrop::~CListBoxFileDrop()
{
}


BEGIN_MESSAGE_MAP(CListBoxFileDrop, CListBox)
	ON_WM_DROPFILES()
END_MESSAGE_MAP()



/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/19
함수:			CListBoxFileDrop::OnDropFiles
세부사항:		static, public
설명:			드래그한 파일의 주소를 리스트박스에 그림
매개변수:		HDROP hDropInfo
리턴값:			없음
노트:			CListBoxFileDrop 메시지 처리기
*****************************************************************************************/
void CListBoxFileDrop::OnDropFiles(HDROP hDropInfo)
{
	// TODO: 여기에 메시지 처리기 코드를 추가 및/또는 기본값을 호출합니다.
	UINT icheckItemCount = GetCount();
	
	if (icheckItemCount > 0)		//드래그 앤 드롭한 리스트박스에 1개라도 값이 들어가있으면
	{
		DeleteString(0);			//이전값을 지웁니다.
	}

	TCHAR FileName[MAX_PATH] = { 0, };
	UINT count = DragQueryFile(hDropInfo, 0xFFFFFFFF, FileName, MAX_PATH);

	for (UINT i = 0; i < count; i++)
	{
		DragQueryFile(hDropInfo, i, FileName, MAX_PATH);

		CString fileExtension;

		fileExtension = ExtractFileExtension(FileName);

		if (fileExtension == "" || fileExtension == "xlsx" || fileExtension == "xls") {
			AddString(FileName);
		}
	}
	CListBox::OnDropFiles(hDropInfo);
}


/*****************************************************************************************
최종수정일:		2022/06/30
함수:			CListBoxFileDrop::ExtractFileExtension
세부사항:		static, public
설명:			파일확장자를 추출함
매개변수:		TCHAR *pathName
리턴값:			CString : 파일확장자 가장앞의 값을 리턴
노트:
*****************************************************************************************/
CString CListBoxFileDrop::ExtractFileExtension(TCHAR *pathName)
{
	CString szFull = (LPCTSTR)pathName;
	CString fileExtension;
	AfxExtractSubString(fileExtension, szFull, 1, '.');

	return fileExtension;
}