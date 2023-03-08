
// autoInputToolDlg.h: 헤더 파일
//

#pragma once

#include <iostream>
#include <stdlib.h>
#include <fstream>
#include <string.h>
#include <vector>


#include "CListBoxFileDrop.h"
#include "CExcel.h"
#include <opencv2/opencv.hpp>
#include <opencv2/core.hpp>
#include <opencv2/highgui.hpp>

using namespace cv;
using namespace std;

#define MAX_WIDECOUNRY 24
#define MAX_PRMCOUNRY 17

class CautoInputToolDlgAutoProxy;

//extern char gv_strFormat1;
//extern char gv_strFormat2;
//extern bool gv_bDropFile;

typedef struct gv_save {
	std::vector<std::vector<CString>> v;
	CString folder_name;
}GV;

extern GV inputFile;
extern GV outputFIle;


// CautoInputToolDlg 대화 상자
class CautoInputToolDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CautoInputToolDlg);
	friend class CautoInputToolDlgAutoProxy;

// 생성입니다.
public:
	CautoInputToolDlg(CWnd* pParent = nullptr);	// 표준 생성자입니다.
	virtual ~CautoInputToolDlg();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_AUTOINPUTTOOL_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 지원입니다.
private:
	CFont m_font;

// 구현입니다.
protected:
	CautoInputToolDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// 생성된 메시지 맵 함수
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()
public:
	//시스템 자동작성함수////////////////////////
	afx_msg void OnLbnSelchangeInputadd();
	afx_msg void OnBnClickedOk();
	afx_msg void OnLvnItemchangedInputlist(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedLoadfolder();
	afx_msg void OnBnClickedAddbutton();
	afx_msg void OnBnClickedRemovebutton();
	afx_msg void OnLvnItemchangedAddedfile(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedResizecheckbutt();
	afx_msg void OnCbnSelchangeLanguageselectbox();
	afx_msg void OnLbnSelchangeOutputlist();
	afx_msg void SelectedImgPreview();
	afx_msg void OnLbnDblclkInputlist();
	afx_msg void OnCbnSelchangePlatform();
	afx_msg void OnBnClickedLoadfolderexcel();
	afx_msg void OnCbnSelchangeExtensionselect();
	afx_msg void OnBnClickedOrimagecheck();
	afx_msg void OnBnClickedDeletelog();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	/////////////////////////////////////////////
	//autoInputToolDlg 전역변수//////////////////
	CString strSelectTxt = L"";
	/////////////////////////////////////////////
	//다이얼로그 변수////////////////////////////
	CListBoxFileDrop m_ctrlListView;	//이미지폴더 파일드롭
	CListBoxFileDrop m_InputAddExcel;	//엑셀파일 파일드롭
	CListBox m_ListBox;					//왼쪽 이미지목록 리스트
	CListBox m_ListBox_Out;				//오른쪽 이미지목록 리스트
	CComboBox m_LanguageCtrlBox;		//언어설정 콤보박스
	CComboBox m_PlatformListBox;		//플랫폼설정 콤보박스
	CComboBox m_ExtensionSelect;		//확장자설정 콤보박스
	CComboBox m_ExcelFileList;			//엑셀파일목록 콤보박스
	CEdit m_FolderNumber;				//삽입될 폴더번호 에디트 박스
	CEdit m_EditLog;					//Log 출력 에디트 박스
	CEdit m_ReSizeLeftV;				//리사이즈 이미지 너비
	CEdit m_ReSizeRightV;				//리사이즈 이미지 높이
	CButton m_OriImageUse;				//원본이미지사용 체크박스
	CButton m_OKButton;					//이미지삽입시작버튼
	/////////////////////////////////////////////
	//OpenCv 변수////////////////////////////////
	Mat m_matImage, m_reImage;
	int mRadio;
	BITMAPINFO *m_pBitmapInfo;
	/////////////////////////////////////////////
	//언어 선택 Static 변수//////////////////////
	CStatic m_Language1;
	CStatic m_Language2;
	CStatic m_Language3;
	CStatic m_Language4;
	CStatic m_Language5;
	CStatic m_Language6;
	/////////////////////////////////////////////
	//검증 결과값 변수///////////////////////////
	CComboBox m_TxtOverBox1;
	CComboBox m_TxtOverBox2;
	CComboBox m_TxtOverBox3;
	CComboBox m_TxtOverBox4;
	CComboBox m_TxtOverBox5;
	CComboBox m_TxtOverBox6;
	
	CComboBox m_WrapErrBox1;
	CComboBox m_WrapErrBox2;
	CComboBox m_WrapErrBox3;
	CComboBox m_WrapErrBox4;
	CComboBox m_WrapErrBox5;
	CComboBox m_WrapErrBox6;
	
	CComboBox m_NotMatchingBox1;
	CComboBox m_NotMatchingBox2;
	CComboBox m_NotMatchingBox3;
	CComboBox m_NotMatchingBox4;
	CComboBox m_NotMatchingBox5;
	CComboBox m_NotMatchingBox6;

	CComboBox m_AbbreviationErr1;
	CComboBox m_AbbreviationErr2;
	CComboBox m_AbbreviationErr3;
	CComboBox m_AbbreviationErr4;
	CComboBox m_AbbreviationErr5;
	CComboBox m_AbbreviationErr6;
	/////////////////////////////////////////////
	//개별작성 함수//////////////////////////////
	void GetUserSelectFolder(CString strFilePath);
	void GetUserSelectFolderExcel(CString strFilePath);
	void CreateBitmapInfo(int width, int height, int bpp);
	void DrawImage();
	void ImageViewFunc();
	void OutputListFindNumber();
	void inputKAndS(int number);
	void EnableBox6(bool input);
	void EnableBox1to5(bool input);
	void InputCountryName(bool type);
	vector<CString> SelectControlBox(int SheetNumber);
	CString ImageFolderSelect(CString strFilePath, CString ImageFile);
	CString ImageFileCheck(CString extensionsFolderPath, CString ImageFile);
	void DeleteAllFiles(CString dirName);
	void OnAdjustRightListBoxHScroll();
	void OnAdjustLeftListBoxHScroll();
	/////////////////////////////////////////////


};
