
// autoInputToolDlg.cpp: 구현 파일
//

#include "pch.h"
#include "framework.h"
#include "autoInputTool.h"
#include "autoInputToolDlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 응용 프로그램 정보에 사용되는 CAboutDlg 대화 상자입니다.

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 지원입니다.

// 구현입니다.
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CautoInputToolDlg 대화 상자


IMPLEMENT_DYNAMIC(CautoInputToolDlg, CDialogEx);

CautoInputToolDlg::CautoInputToolDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_AUTOINPUTTOOL_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON1);	//IDR_MAINFRAME
	m_pAutoProxy = nullptr;
}

CautoInputToolDlg::~CautoInputToolDlg()
{
	// 이 대화 상자에 대한 자동화 프록시가 있을 경우 이 대화 상자에 대한
	//  후방 포인터를 null로 설정하여
	//  대화 상자가 삭제되었음을 알 수 있게 합니다.
	if (m_pAutoProxy != nullptr)
		m_pAutoProxy->m_pDialog = nullptr;
}

void CautoInputToolDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_InputAdd, m_ctrlListView);
	DDX_Control(pDX, IDC_InputAdd, m_ctrlListView);
	DDX_Control(pDX, IDC_InputList, m_ListBox);
	DDX_Control(pDX, IDC_OUTPUTLIST, m_ListBox_Out);
	DDX_Control(pDX, IDC_LanguageSelectbox, m_LanguageCtrlBox);
	DDX_Control(pDX, IDC_PLATFORM, m_PlatformListBox);
	DDX_Control(pDX, IDC_Language1, m_Language1);
	DDX_Control(pDX, IDC_Language2, m_Language2);
	DDX_Control(pDX, IDC_Language3, m_Language3);
	DDX_Control(pDX, IDC_Language4, m_Language4);
	DDX_Control(pDX, IDC_Language5, m_Language5);
	DDX_Control(pDX, IDC_Language6, m_Language6);
	DDX_Control(pDX, IDC_TxtOverBox1, m_TxtOverBox1);
	DDX_Control(pDX, IDC_TxtOverBox2, m_TxtOverBox2);
	DDX_Control(pDX, IDC_TxtOverBox3, m_TxtOverBox3);
	DDX_Control(pDX, IDC_TxtOverBox4, m_TxtOverBox4);
	DDX_Control(pDX, IDC_TxtOverBox5, m_TxtOverBox5);
	DDX_Control(pDX, IDC_TxtOverBox6, m_TxtOverBox6);
	DDX_Control(pDX, IDC_WrapErrBox1, m_WrapErrBox1);
	DDX_Control(pDX, IDC_WrapErrBox2, m_WrapErrBox2);
	DDX_Control(pDX, IDC_WrapErrBox3, m_WrapErrBox3);
	DDX_Control(pDX, IDC_WrapErrBox4, m_WrapErrBox4);
	DDX_Control(pDX, IDC_WrapErrBox5, m_WrapErrBox5);
	DDX_Control(pDX, IDC_WrapErrBox6, m_WrapErrBox6);
	DDX_Control(pDX, IDC_NotMatchingBox1, m_NotMatchingBox1);
	DDX_Control(pDX, IDC_NotMatchingBox2, m_NotMatchingBox2);
	DDX_Control(pDX, IDC_NotMatchingBox3, m_NotMatchingBox3);
	DDX_Control(pDX, IDC_NotMatchingBox4, m_NotMatchingBox4);
	DDX_Control(pDX, IDC_NotMatchingBox5, m_NotMatchingBox5);
	DDX_Control(pDX, IDC_NotMatchingBox6, m_NotMatchingBox6);
	DDX_Control(pDX, IDC_AbbreviationErr1, m_AbbreviationErr1);
	DDX_Control(pDX, IDC_AbbreviationErr2, m_AbbreviationErr2);
	DDX_Control(pDX, IDC_AbbreviationErr3, m_AbbreviationErr3);
	DDX_Control(pDX, IDC_AbbreviationErr4, m_AbbreviationErr4);
	DDX_Control(pDX, IDC_AbbreviationErr5, m_AbbreviationErr5);
	DDX_Control(pDX, IDC_AbbreviationErr6, m_AbbreviationErr6);
	DDX_Control(pDX, IDC_EXTENSIONSELECT, m_ExtensionSelect);
	DDX_Control(pDX, IDC_InputAddExcel, m_InputAddExcel);
	DDX_Control(pDX, IDC_FOLDERNUMBER, m_FolderNumber);
	DDX_Control(pDX, IDC_EDIT7, m_EditLog);
	DDX_Control(pDX, IDC_EXCELCOMBO, m_ExcelFileList);
	DDX_Control(pDX, IDC_ResizeLeftValue, m_ReSizeLeftV);
	DDX_Control(pDX, IDC_ResizeRightValue, m_ReSizeRightV);
	DDX_Control(pDX, IDC_ORIMAGECHECK, m_OriImageUse);
	DDX_Control(pDX, IDOK, m_OKButton);
}

BEGIN_MESSAGE_MAP(CautoInputToolDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_LBN_SELCHANGE(IDC_InputAdd, &CautoInputToolDlg::OnLbnSelchangeInputadd)
	ON_BN_CLICKED(IDOK, &CautoInputToolDlg::OnBnClickedOk)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_InputList, &CautoInputToolDlg::OnLvnItemchangedInputlist)
	ON_BN_CLICKED(IDC_LoadFolder, &CautoInputToolDlg::OnBnClickedLoadfolder)
	ON_BN_CLICKED(IDC_AddButton, &CautoInputToolDlg::OnBnClickedAddbutton)
	ON_BN_CLICKED(IDC_RemoveButton, &CautoInputToolDlg::OnBnClickedRemovebutton)
	ON_BN_CLICKED(IDC_ResizeCheckButt, &CautoInputToolDlg::OnBnClickedResizecheckbutt)
	ON_CBN_SELCHANGE(IDC_LanguageSelectbox, &CautoInputToolDlg::OnCbnSelchangeLanguageselectbox)
	ON_WM_DROPFILES()
	ON_LBN_SELCHANGE(IDC_OUTPUTLIST, &CautoInputToolDlg::OnLbnSelchangeOutputlist)
	ON_LBN_DBLCLK(IDC_InputList, &CautoInputToolDlg::OnLbnDblclkInputlist)
	ON_LBN_SELCHANGE(IDC_InputList, &CautoInputToolDlg::SelectedImgPreview)
	ON_CBN_SELCHANGE(IDC_PLATFORM, &CautoInputToolDlg::OnCbnSelchangePlatform)
	ON_BN_CLICKED(IDC_LoadFolderExcel, &CautoInputToolDlg::OnBnClickedLoadfolderexcel)
	ON_CBN_SELCHANGE(IDC_EXTENSIONSELECT, &CautoInputToolDlg::OnCbnSelchangeExtensionselect)
	ON_BN_CLICKED(IDC_ORIMAGECHECK, &CautoInputToolDlg::OnBnClickedOrimagecheck)
	ON_BN_CLICKED(IDC_DELETELOG, &CautoInputToolDlg::OnBnClickedDeletelog)
END_MESSAGE_MAP()

CString strCountryWArr[] = { L"1한국", L"2영국", L"3독일", L"4프랑스", L"5이탈리아", L"6스페인", L"7포르투갈", L"8네덜란드", L"9덴마크", L"10스웨덴", L"11노르웨이",
		L"12폴란드", L"13체코", L"14슬로바키아", L"15러시아", L"16핀란드", L"17헝가리", L"18터키", L"19그리스", L"20슬로베니아", L"21우크라이나", L"22불가리아", L"23루마니아", L"24크로아티아" };
CString strCountryPRMArr[] = { L"1한국", L"2영국", L"3독일", L"4프랑스", L"5이탈리아", L"6스페인", L"7포르투갈", L"8네덜란드", L"9덴마크", L"10스웨덴", L"11노르웨이",
	L"12폴란드", L"13체코", L"14슬로바키아", L"15러시아", L"16헝가리", L"17터키", L" " };

// CautoInputToolDlg 메시지 처리기

BOOL CautoInputToolDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 시스템 메뉴에 "정보..." 메뉴 항목을 추가합니다.

	// IDM_ABOUTBOX는 시스템 명령 범위에 있어야 합니다.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	//ini파일///////////////////////////////////////////////////
	CString strUniqueFilePath;
	strUniqueFilePath.Format(_T("%s\\UserSetData.ini"), GetExePath());
	gv_InIFilePath = strUniqueFilePath;

	CFileFind pFind;
	BOOL bRet = pFind.FindFile(strUniqueFilePath);
	if (!bRet)						//ini파일이 없을경우 새로 만들고 디폴트값을 설정함.
	{
		INIWriteString(_T("AutoInputTool"), _T("PlatForm"), _T("0"), strUniqueFilePath);
		INIWriteString(_T("AutoInputTool"), _T("Language"), _T("0"), strUniqueFilePath);
		INIWriteString(_T("AutoInputTool"), _T("Extension"), _T("0"), strUniqueFilePath);
		INIWriteString(_T("AutoInputTool"), _T("ImageLeftSize"), _T("273"), strUniqueFilePath);
		INIWriteString(_T("AutoInputTool"), _T("ImageRightSize"), _T("190"), strUniqueFilePath);
	}
	else
	{
		//noting
	}

	CString strPlatFormData = INIReadString(_T("AutoInputTool"), _T("PlatForm"), strUniqueFilePath);
	int iPlatFormData = 0;
	iPlatFormData = _ttoi(strPlatFormData);
	////////////////////////////////////////////////////////////
	//File Drop///////////////////////////////////////////////////
	ChangeWindowMessageFilter(0x0049, MSGFLT_ADD);
	ChangeWindowMessageFilter(WM_DROPFILES, MSGFLT_ADD);
	DragAcceptFiles();
	//////////////////////////////////////////////////////////////
	//Platform List///////////////////////////////////////////////
	m_PlatformListBox.AddString(_T("STD5W"));
	m_PlatformListBox.AddString(_T("STD5"));
	m_PlatformListBox.AddString(_T("PRM5"));

	m_PlatformListBox.SetCurSel(iPlatFormData);				//디폴트 값 STD5W
	//////////////////////////////////////////////////////////////
	//Langueage List//////////////////////////////////////////////24개국 ini에서 읽어서 가져오기
	CString strLanguageData = INIReadString(_T("AutoInputTool"), _T("Language"), strUniqueFilePath);
	int iLanguageData = 0;
	iLanguageData = _ttoi(strLanguageData);

	if (iLanguageData == 0)
	{
		InputCountryName(true);
	}
	else
	{
		InputCountryName(false);
	}

	CString strExtensionData = INIReadString(_T("AutoInputTool"), _T("Extension"), strUniqueFilePath);
	int iExtensionData = 0;
	iExtensionData = _ttoi(strExtensionData);
	m_ExtensionSelect.SetCurSel(iExtensionData);					//확장자 디폴트값 PNG
	//////////////////////////////////////////////////////////////
	CString strLeftImageSizeData = INIReadString(_T("AutoInputTool"), _T("ImageLeftSize"), strUniqueFilePath);
	CString strRightImageSizeData = INIReadString(_T("AutoInputTool"), _T("ImageRightSize"), strUniqueFilePath);
	
	m_ReSizeLeftV.SetWindowTextW(strLeftImageSizeData);				//ini파일에서 왼쪽 이미지 사이즈 가져오기
	m_ReSizeRightV.SetWindowTextW(strRightImageSizeData);			//ini파일에서 오른쪽 이미지 사이즈 가져오기

	//검증 결과값 디폴트 값///////////////////////////////////////
	EnableBox1to5(FALSE);
	EnableBox6(FALSE);

	m_TxtOverBox1.SetCurSel(0);
	m_TxtOverBox2.SetCurSel(0);
	m_TxtOverBox3.SetCurSel(0);
	m_TxtOverBox4.SetCurSel(0);
	m_TxtOverBox5.SetCurSel(0);
	m_TxtOverBox6.SetCurSel(0);

	m_WrapErrBox1.SetCurSel(0);
	m_WrapErrBox2.SetCurSel(0);
	m_WrapErrBox3.SetCurSel(0);
	m_WrapErrBox4.SetCurSel(0);
	m_WrapErrBox5.SetCurSel(0);
	m_WrapErrBox6.SetCurSel(0);

	m_NotMatchingBox1.SetCurSel(0);
	m_NotMatchingBox2.SetCurSel(0);
	m_NotMatchingBox3.SetCurSel(0);
	m_NotMatchingBox4.SetCurSel(0);
	m_NotMatchingBox5.SetCurSel(0);
	m_NotMatchingBox6.SetCurSel(0);

	m_AbbreviationErr1.SetCurSel(0);
	m_AbbreviationErr2.SetCurSel(0);
	m_AbbreviationErr3.SetCurSel(0);
	m_AbbreviationErr4.SetCurSel(0);
	m_AbbreviationErr5.SetCurSel(0);
	m_AbbreviationErr6.SetCurSel(0);
	//////////////////////////////////////////////////////////////
	// TODO: 여기에 추가 초기화 작업을 추가합니다.
	
	m_font.CreateFont(12, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH, _T("바탕체"));
	m_OKButton.SetFont(&m_font, TRUE);			//이미지삽입시작 버튼 폰트변경

	return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

void CautoInputToolDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 응용 프로그램의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CautoInputToolDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CautoInputToolDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 컨트롤러에서 해당 개체 중 하나를 계속 사용하고 있을 경우
//  사용자가 UI를 닫을 때 자동화 서버를 종료하면 안 됩니다.  이들
//  메시지 처리기는 프록시가 아직 사용 중인 경우 UI는 숨기지만,
//  UI가 표시되지 않아도 대화 상자는
//  남겨 둡니다.

void CautoInputToolDlg::OnClose()
{
	if (CanExit())
		CDialogEx::OnClose();
}

void CautoInputToolDlg::OnOK()
{
	if (CanExit())
		CDialogEx::OnOK();
}

void CautoInputToolDlg::OnCancel()
{
	if (CanExit())
		CDialogEx::OnCancel();
}

BOOL CautoInputToolDlg::CanExit()
{
	// 프록시 개체가 계속 남아 있으면 자동화 컨트롤러에서는
	//  이 응용 프로그램을 계속 사용합니다.  대화 상자는 남겨 두지만
	//  해당 UI는 숨깁니다.
	if (m_pAutoProxy != nullptr)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}
	return TRUE;
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/13
함수:			CautoInputToolDlg::WindowProc
세부사항:		static, public
설명:			창을 닫을때 확인메시지를 보내는 역할
매개변수:		UINT message, WPARAM wParam, LPARAM lParam
리턴값:			CDialogEx::WindowProc(message, wParam, lParam)
노트:
*****************************************************************************************/
LRESULT CautoInputToolDlg::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	// TODO: 여기에 특수화된 코드를 추가 및/또는 기본 클래스를 호출합니다.
	
	if (message == WM_CLOSE)
	{
		int check = MessageBox(L"프로그램을 종료하시겠습니까?", L"프로그램 종료확인", MB_ICONQUESTION | MB_OKCANCEL);
		if (IDCANCEL == check)
		{
			return 0;
		}
		destroyAllWindows();
		//PostQuitMessage(0);
	}
	else if (message == WM_DESTROY)
	{
		PostQuitMessage(0);
	}
	
	return CDialogEx::WindowProc(message, wParam, lParam);
}

void CautoInputToolDlg::OnLbnSelchangeInputadd()		//주소 드로그앤 드롭(input)
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/11/03
함수:			CautoInputToolDlg::OnBnClickedOk
세부사항:		static, public
설명:			엑셀 삽입 버튼을 클릭했을 경우의 이벤트
매개변수:		없음
리턴값:			없음
노트:			
*****************************************************************************************/
void CautoInputToolDlg::OnBnClickedOk()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	CExcel Excel;

	//엑셀 파일 경로 저장 변수, 이미지 폴더 경로 저장 변수
	CString ExcelPath, ImageFolderPath, cstrExImageFolderPath;
	//엑셀 파일 이름 저장 변수, 엑셀 파일 확장자 제외 이름 저장 변수
	CString ExcelFile, only_ExcelFile;
	//엑셀 입력 확인 변수
	BOOL Insert_Check = FALSE;

	//이미지 폴더 경로
	m_ctrlListView.SetCurSel(0);
	m_ctrlListView.GetText(m_ctrlListView.GetCurSel(), ImageFolderPath);
	cstrExImageFolderPath = ImageFolderPath + L"\\extensions\\";

	//엑셀 폴더 및 파일 경로
	m_InputAddExcel.SetCurSel(0);
	m_InputAddExcel.GetText(m_InputAddExcel.GetCurSel(), ExcelPath);
	m_ExcelFileList.GetLBText(m_ExcelFileList.GetCurSel(), ExcelFile);
	AfxExtractSubString(only_ExcelFile, ExcelFile, 0, '.');
	ExcelPath = ExcelPath + "\\" + ExcelFile;

	//삽입할 이미지가 있는 리스트 박스 개수 저장
	int ListOut_Count = m_ListBox_Out.GetCount();

	//삽입 리스트에 목록이 존재하지 않을때 경고
	if (ListOut_Count == 0)
		MessageBox(L"삽입할 이미지가 없습니다.", L"이미지 선택", MB_ICONERROR);
	else
	{
		//엑셀 파일 실행
		Excel.StartExcel(ExcelPath);

		//실행한 엑셀 파일 시트 개수 파악
		int SheetCount = Excel.SheetsCountExcel();

		//실행한 엑셀 파일 시트 이름 변수 저장
		vector <CString> SheetName;
		SheetName = Excel.SheetNameExcel(SheetCount);

		//플랫폼을 선택하여 이미지가 삽입될 셀 크기 조정(셀 범위는 기존 양식의 셀 범위를 확인하여 설정)
		Excel.CellResizeExcel(SheetCount);

		//리스트의 끝에서부터 이미지 파일의 국가와 시트 이름을 확인하여 같으면 엑셀에 검증결과와 함께 삽입
		for (int i = ListOut_Count - 1; i >= 0; i--)
		{
			//이미지 파일 변수, 임시 변수, 이미지 국가 이름 변수
			CString ImageFile, ImageTemp, ImageNation;
			m_ListBox_Out.GetText(i, ImageFile);
			//이미지 파일에서 ' '을 기준으로 앞의 문자열만 추출하여 임시 변수에 저장
			AfxExtractSubString(ImageTemp, ImageFile, 0, ' ');
			//임시 변수에 저장된 문자열에서 '_' 뒤의 문자열을 이미지 국가 이름 변수에 저장
			ImageNation = ImageTemp.Mid(ImageTemp.Find('_') + 1);

			//이미지 국가 이름과 시트 이름을 비교하여 해당 시트에 입력
			for (int k = 0; k < SheetCount; k++)
			{
				if (ImageNation == SheetName[k])
				{
					CString only_ImageFile, ImageFilePath;
					int SheetNumber, RowNumber;

					//이미지 국가 이름으로 시트 위치 확인(기존 양식의 국가별 시트 번호로 구분)
					SheetNumber = Excel.SheetNumberExcel(ImageNation);

					//이미지 파일 이름에서 맨 앞의 번호를 추출하여 행 번호 지정
					RowNumber = Excel.RowNumberExcel(ImageFile);

					//이미지 파일의 확장자를 제거한 이름 변수에 저장
					AfxExtractSubString(only_ImageFile, ImageFile, 0, '.');

					//이미지 폴더 경로와 이미지 파일 이름을 결합하여 이미지 파일 경로 변수에 저장
					//이미지 파일 폴더에 extensions라는 하위 폴더가 존재할 경우 하위 폴더에서 같은 이름의 이미지로 삽입
					ImageFilePath = ImageFolderSelect(ImageFolderPath, ImageFile);
					
					//검증결과 값 저장
					vector<CString> ResultValue;
					ResultValue = SelectControlBox(SheetNumber);

					for (int j = 0; j < 5; j++)
					{
						CString temp_value = ResultValue[j];
						Excel.ResultValueExcel(SheetNumber, RowNumber, temp_value, j);
					}
					
					//입력된 값을 확인하여 폰트 및 셀 배경색 변경
					Excel.ResultStyleExcel(SheetNumber, RowNumber);

					//이미지 삽입 및 이미지 이름 입력
					Excel.PictureInsertExcel(SheetNumber, RowNumber, only_ImageFile, ImageFilePath);

					//Log에 이미지 입력 현황 출력
					m_EditLog.SetSel(-2, -1);
					m_EditLog.ReplaceSel(only_ExcelFile);
					m_EditLog.ReplaceSel(_T(" 엑셀 파일에 "));
					ImageFilePath.Delete(0, ImageFolderPath.GetLength() + 1);
					m_EditLog.ReplaceSel(ImageFilePath);
					m_EditLog.ReplaceSel(_T("을(를) 입력하였습니다!"));
					m_EditLog.ReplaceSel(_T("\r\n"));

					//삽입을 완료한 이미지는 리스트에서 삭제
					m_ListBox_Out.DeleteString(i);
					Insert_Check = TRUE;
				}
			}
		}
		Excel.CloseExcel();

		//엑셀 파일에 삽입된 이미지가 없을때
		if (Insert_Check == FALSE)
		{
			MessageBox(L"선택한 엑셀 파일에 삽입할 이미지가 없습니다.\r\n엑셀 파일과 이미지 파일을 확인 해주세요", L"엑셀 이미지 확인", MB_ICONERROR);
		}
		//Log에 엑셀 백그라운드 종료 확인
		m_EditLog.ReplaceSel(_T("엑셀 파일이 종료되었습니다."));
		m_EditLog.ReplaceSel(_T("\r\n"));
	}
	GetDlgItem(IDOK)->EnableWindow(FALSE);
	DeleteAllFiles(cstrExImageFolderPath);
}

void CautoInputToolDlg::OnLvnItemchangedInputlist(NMHDR *pNMHDR, LRESULT *pResult)	//input 리스트 표시
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	*pResult = 0;
}


/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/19
함수:			CautoInputToolDlg::OnBnClickedLoadfolder
세부사항:		static, public
설명:			이미지 불러오기 버튼을 클릭했을 경우의 이벤트
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::OnBnClickedLoadfolder()
{
	m_ListBox.ResetContent();			//InputListBox 비워주기
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	int iCheckPathExist = m_ctrlListView.GetCount();
			
	if (iCheckPathExist == 0) {			//입력된 경로가 존재하는지 확인
		MessageBox(L"선택된 경로가 없습니다.", L"입력된 경로 없음", MB_ICONERROR);
		return;
	}
	else {
		CString cstrtempPath = L"", cstrFileExtention = L"";
		CListBox *p_list = (CListBox *)GetDlgItem(IDC_InputAdd);
		p_list->GetText(0, cstrtempPath);
		
		AfxExtractSubString(cstrFileExtention, cstrtempPath, 1, '.');
		
		if (cstrFileExtention == "")	//폴더인지 확인
		{
			GetUserSelectFolder(cstrtempPath);
			if (m_ListBox.GetCount() == 0) {
				MessageBox(L"선택한 폴더내에 이미지파일이 없습니다.", L"이미지파일 없음", MB_ICONERROR);
			}
			//m_ctrlListView.ResetContent();		//폴더입력 주소창 비워주기
			cstrtempPath.Empty();
			OutputListFindNumber();
		}
		else
		{																		//폴더가 아닐경우 경고메시지 출력후 초기화
			MessageBox(L"이미지 폴더가 아닙니다.", L"이미지 폴더가 아님", MB_ICONERROR);
			m_ctrlListView.ResetContent();
		}
		
	}
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/11/14
함수:			CautoInputToolDlg::OnBnClickedAddbutton
세부사항:		static, public
설명:			Add버튼을 클릭했을 경우의 이벤트
매개변수:		없음
리턴값:			없음
노트:			
*****************************************************************************************/
void CautoInputToolDlg::OnBnClickedAddbutton()		//Add 버튼
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	int icount = m_ListBox.GetCount(), iselectedcount = 0;

	for (int i = icount - 1; i >= 0; i--) 
	{
		if (m_ListBox.GetSel(i)) {			//아이템을 선택했을경우
			m_ListBox.GetText(i, strSelectTxt);
			m_ListBox_Out.AddString(strSelectTxt);
			OnAdjustRightListBoxHScroll();
			m_ListBox.DeleteString(i);
			iselectedcount++;
		}	
	}
	if (iselectedcount == 0) {				//선택된 아이템이 하나도 없을경우
		MessageBox(L"선택된 항목이 없습니다.", L"선택된 항목 없음", MB_ICONERROR);
	}
	else if (iselectedcount > 0)
	{
		GetDlgItem(IDC_ResizeCheckButt)->EnableWindow(TRUE);
	}
}


/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/11/14
함수:			CautoInputToolDlg::OnBnClickedRemovebutton
세부사항:		static, public
설명:			Remove 버튼을 클릭했을 경우의 이벤트
매개변수:		없음
리턴값:			없음
노트:			
*****************************************************************************************/
void CautoInputToolDlg::OnBnClickedRemovebutton()	//Remove 버튼
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	int icount = m_ListBox_Out.GetCount(), iselectedcount = 0;

	for (int i = icount - 1; i >= 0; i--)
	{
		if (m_ListBox_Out.GetSel(i)) {			//아이템을 선택했을경우
			m_ListBox_Out.GetText(i, strSelectTxt);
			m_ListBox.AddString(strSelectTxt);
			m_ListBox_Out.DeleteString(i);
			iselectedcount++;
		}
		if (m_ListBox_Out.GetCount() == 0)
		{
			GetDlgItem(IDC_ResizeCheckButt)->EnableWindow(FALSE);
		}
	}
	if (iselectedcount == 0) {					//선택된 아이템이 하나도 없을경우
		MessageBox(L"선택된 항목이 없습니다.", L"선택된 항목 없음", MB_ICONERROR);
	}
}

void CautoInputToolDlg::OnLvnItemchangedAddedfile(NMHDR *pNMHDR, LRESULT *pResult)		//input된 파일 리스트(오른쪽 리스트)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	*pResult = 0;
}


/*****************************************************************************************
최종수정일:		2022/11/14
함수:			autoInputToolDlg::OnBnClickedResizecheckbutt
세부사항:		static, public
설명:			원본 이미지 크기에서 지정한 값의 크기로 이미지 크기 변경 및 변경된 이미지 파일들 저장하는 버튼
매개변수:		없음
리턴값:			없음
노트:			
*****************************************************************************************/
void CautoInputToolDlg::OnBnClickedResizecheckbutt()			//Resize 확정하는 버튼
{
	CString cstrFileName, cstrFileAbsolutePath, cstrreFileName;
	string strFilePath = string(CT2CA(gv_InputFilePath));
	int resizewidth = GetDlgItemInt(IDC_ResizeLeftValue);
	int resizeheight = GetDlgItemInt(IDC_ResizeRightValue);
	
	CString cstrReSizeWidth, cstrReSizeHeight;
	cstrReSizeWidth.Format(_T("%d"), resizewidth); 
	cstrReSizeHeight.Format(_T("%d"), resizeheight);

	m_ListBox_Out.GetText(m_ListBox_Out.GetCurSel(), cstrFileName);
	int ilen = m_ListBox_Out.GetCount();

	CString strMakePath = gv_InputFilePath + "\\extensions\\";
	CString strDirectoy(_T(""));
	CString strCheckPath(_T(""));
	int iPos = 0;


	if (m_OriImageUse.GetCheck())
	{
		MessageBox(L"원본 이미지를 사용합니다.", L"이미지크기변경 없음", MB_OK);
	}
	else
	{
		while (::AfxExtractSubString(strDirectoy, strMakePath, iPos++, _T('\\')))		//해당폴더가 있는지 확인후 없으면 생성
		{
			if (_T("") == strCheckPath)
			{
				strCheckPath = strDirectoy;
			}
			else
			{
				strCheckPath = strCheckPath + _T("\\") + strDirectoy;
			}
			if (INVALID_FILE_ATTRIBUTES == ::GetFileAttributes(strCheckPath))
			{
				BOOL bCreate = ::CreateDirectory(strCheckPath, NULL);
				if (!bCreate)
				{
					CString strMsg;
					strMsg.Format(_T("디렉토리 생성 실패 (%S)"), strCheckPath);
					AfxMessageBox(strMsg);
					break;
				}
			}
		}

		for (int i = 0; i < ilen; i++)
		{
			m_ListBox_Out.GetText(i, cstrFileName);
			cstrFileAbsolutePath = gv_InputFilePath + '\\' + cstrFileName;
			string strFileAbsolutePath = string(CT2CA(cstrFileAbsolutePath));
			//AfxMessageBox(cstrFileAbsolutePath);		//확인용코드

			Mat m_OriImg = imread(strFileAbsolutePath, IMREAD_ANYCOLOR);

			if (mRadio == 0)		//PNG
			{
				resize(m_OriImg, m_reImage, Size(resizewidth, resizeheight));
				AfxExtractSubString(cstrreFileName, cstrFileName, 0, '.');
				string strFileName = string(CT2CA(cstrreFileName));
				imwrite(strFilePath + "\\extensions\\" + strFileName + ".png", m_reImage);
				m_EditLog.ReplaceSel(cstrFileName);
				m_EditLog.ReplaceSel(_T("의 이미지 사이즈를 "));
				m_EditLog.ReplaceSel(cstrReSizeWidth);
				m_EditLog.ReplaceSel(_T(" x "));
				m_EditLog.ReplaceSel(cstrReSizeHeight);
				m_EditLog.ReplaceSel(_T(" 크기로 변경하고 하위폴더에 저장했습니다."));
				m_EditLog.ReplaceSel(_T("\r\n"));
				INIWriteString(_T("AutoInputTool"), _T("Extension"), _T("0"), gv_InIFilePath);
			}
			else if (mRadio == 1)	//JPG
			{
				resize(m_OriImg, m_reImage, Size(resizewidth, resizeheight));
				AfxExtractSubString(cstrreFileName, cstrFileName, 0, '.');
				string strFileName = string(CT2CA(cstrreFileName));
				imwrite(strFilePath + "\\extensions\\" + strFileName + ".jpg", m_reImage);
				m_EditLog.ReplaceSel(cstrFileName);
				m_EditLog.ReplaceSel(_T("의 이미지 사이즈를 "));
				m_EditLog.ReplaceSel(cstrReSizeWidth);
				m_EditLog.ReplaceSel(_T(" x "));
				m_EditLog.ReplaceSel(cstrReSizeHeight);
				m_EditLog.ReplaceSel(_T(" 크기로 변경하고 하위폴더에 저장했습니다."));
				m_EditLog.ReplaceSel(_T("\r\n"));
				INIWriteString(_T("AutoInputTool"), _T("Extension"), _T("1"), gv_InIFilePath);
			}
			else
			{
				//nothing
			}
		}
		INIWriteString(_T("AutoInputTool"), _T("ImageLeftSize"), cstrReSizeWidth, gv_InIFilePath);			//사용자가 최종적으로 사용한 값을 INI파일에 저장
		INIWriteString(_T("AutoInputTool"), _T("ImageRightSize"), cstrReSizeHeight, gv_InIFilePath);		//사용자가 최종적으로 사용한 값을 INI파일에 저장
	}
	if (gv_bImageFolder && gv_bExcelFolder)
	{
		GetDlgItem(IDOK)->EnableWindow(TRUE);		//이미지 삽입시작 버튼 활성화
	}
}


/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/11
함수:			CautoInputToolDlg::OnCbnSelchangeLanguageselectbox
세부사항:		static, public
설명:			언어설정 콤보박스 값을 선택함에 따라 언어를 표시함
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::OnCbnSelchangeLanguageselectbox()		//언어설정 리스트 박스
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	UpdateData();
	CString platformData = L"";
	int ilanguageIndex = 0;

	m_PlatformListBox.GetLBText(m_PlatformListBox.GetCurSel(), platformData);
	ilanguageIndex = m_LanguageCtrlBox.GetCurSel();
	if (ilanguageIndex == 0)
	{
		EnableBox1to5(TRUE);
		EnableBox6(TRUE);
		inputKAndS(0);
	}
	else if (ilanguageIndex == 1)
	{
		EnableBox1to5(TRUE);
		EnableBox6(TRUE);
		inputKAndS(1);
	}
	else if (platformData == _T("STD5W") && ilanguageIndex == 2)
	{
		EnableBox1to5(TRUE);
		EnableBox6(TRUE);
		inputKAndS(2);
	}
	else if (platformData == _T("STD5W") && ilanguageIndex == 3)
	{
		EnableBox1to5(TRUE);
		EnableBox6(TRUE);
		inputKAndS(3);
	}
	else if(platformData == _T("STD5") || platformData == _T("PRM5") && ilanguageIndex == 2)
	{
		EnableBox1to5(TRUE);
		EnableBox6(FALSE);
		inputKAndS(4);
	}
}


/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/11/14
함수:			CautoInputToolDlg::GetUserSelectFolder
세부사항:		static, public
설명:			폴더의 주소값을 얻어와 내용물의 리스트를 리스트박스에 표시함.
매개변수:		CString : strFilePath 폴더의 주소값
리턴값:			없음
노트:			
*****************************************************************************************/
void CautoInputToolDlg::GetUserSelectFolder(CString strFilePath)
{
	gv_InputFilePath = "";		//글로벌 변수 파일경로 초기화
	UpdateData();
	if (strFilePath.IsEmpty()) {
		return;
	}
	// 검색 클래스
	CFileFind file;
	strFilePath += "\\*.*";
	CString strFolderItem = L"", stroOnlyFilePath = L"", cstrSortationName = L"";
	CString strFileName = L"", strFileExt = L"", strFolderNum = L"", strCountryName = L"";
	// CFileFind는 파일, 디렉터리가 존재하면 TRUE 를 반환함
	BOOL bWorking = file.FindFile(strFilePath); //

	CString platformData = L"";
	m_PlatformListBox.GetLBText(m_PlatformListBox.GetCurSel(), platformData);

	bool bFolderName = false, bCountyname = false;

	if (bWorking == 0) {
		return;
	}

	while (bWorking)
	{
		//다음 파일 or 폴더 가 존재하면다면 TRUE 반환
		bWorking = file.FindNextFile();

		// folder 일 경우는 continue
		if (file.IsDirectory() || file.IsDots())
		{
			continue;
		}
		strFolderItem = file.GetFilePath();

		strFileExt = strFolderItem.Mid(strFolderItem.ReverseFind('.'));
		stroOnlyFilePath = strFolderItem.Left(strFolderItem.ReverseFind('\\'));
		gv_InputFilePath = stroOnlyFilePath;	//전체경로에서 파일 이름만 제거하여 글로벌 변수에 대입

		// 파일 일때
		if (!file.IsDots())
		{
			//비교를위해 확장자 전부 대문자로 변경
			strFileExt.MakeUpper();
			if (file.IsDirectory()) {
				continue;
			}
			strFolderItem.Replace(strFileName, _T(""));

		}
		//파일의 이름
		CString _strFileName = file.GetFileName();
		cstrSortationName = _strFileName;

		string strSortationName = string(CT2CA(cstrSortationName));
		strSortationName = strSortationName.substr(0, strSortationName.find(" "));
		cstrSortationName = strSortationName.c_str();

		strFolderNum = cstrSortationName.Left(cstrSortationName.Find('_'));					//~번 이라는 폴더 이름만 잘라서 집어넣음
		strCountryName = cstrSortationName.Mid(cstrSortationName.ReverseFind('_') + 1);		//숫자를 포함한 나라이름을 집어넣음
		//AfxMessageBox(strCountryName); //확인용 코드

		// 현재폴더 상위폴더 썸네일파일은 제외
		if (_strFileName == _T("Thumbs.db")) {
			continue;
		}

		strFolderItem = file.GetFileName();
		//읽어온 파일 이름을 리스트박스에 넣음
		bool bFindFlag = true;

		if (strFileExt == ".PNG" || strFileExt == ".JPG")
		{
			if (platformData == _T("STD5W") && strFolderNum.Right(1) == L"번")
			{
				for (int i = 0; i < MAX_WIDECOUNRY; i++)
				{
					if (strCountryWArr[i] == strCountryName)
					{
						m_ListBox.AddString(strFolderItem);
						OnAdjustLeftListBoxHScroll();
						bFindFlag = false;
					}
				}
				if (bFindFlag)
				{
					MessageBox(_strFileName + L"의 파일명이 형식과 맞지않아 추가되지 않았습니다.", L"파일명이 형식과 맞지않음", MB_ICONERROR);
				}
			}
			else if ((platformData == _T("STD5") || platformData == _T("PRM5")) && strFolderNum.Right(1) == L"번")
			{
				for (int j = 0; j < MAX_PRMCOUNRY; j++)
				{
					if (strCountryPRMArr[j] == strCountryName)
					{
						m_ListBox.AddString(strFolderItem);
						OnAdjustLeftListBoxHScroll();
						bFindFlag = false;
					}
				}
				if (bFindFlag)
				{
					MessageBox(_strFileName + L"의 파일명이 형식과 맞지않아 추가되지 않았습니다.", L"파일명이 형식과 맞지않음", MB_ICONERROR);
				}
			}
			gv_bImageFolder = true;
		}
	}
	file.Close();
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/11/14
함수:			CautoInputToolDlg::GetUserSelectFolderExcel
세부사항:		static, public
설명:			폴더의 주소값을 얻어와 내용물의 리스트를 리스트박스에 표시함.(엑셀파일 전용)
매개변수:		CString : strFilePath 폴더의 주소값
리턴값:			없음
노트:			
*****************************************************************************************/
void CautoInputToolDlg::GetUserSelectFolderExcel(CString strFilePath)
{
	gv_InputExcelFilePath = "";		//글로벌 변수 파일경로 초기화
	if (strFilePath.IsEmpty()) {
		return;
	}
	// 검색 클래스
	CFileFind file;
	strFilePath += "\\*.*";
	CString strFolderItem, strFileExt;
	// CFileFind는 파일, 디렉터리가 존재하면 TRUE 를 반환함
	BOOL bWorking = file.FindFile(strFilePath); //

	CString fileName;
	CString onlyFilePath;

	if (bWorking == 0) {
		return;
	}

	while (bWorking)
	{
		//다음 파일 or 폴더 가 존재하면다면 TRUE 반환
		bWorking = file.FindNextFile();

		// folder 일 경우는 continue
		if (file.IsDirectory() || file.IsDots())
		{
			continue;
		}
		strFolderItem = file.GetFilePath();
		strFileExt = strFolderItem.Mid(strFolderItem.ReverseFind('.'));
		onlyFilePath = strFolderItem.Left(strFolderItem.ReverseFind('\\'));
		gv_InputExcelFilePath = onlyFilePath;	//전체경로에서 파일 이름만 제거하여 글로벌 변수에 대입

		// 파일 일때
		if (!file.IsDots())
		{
			//비교를위해 확장자 전부 대문자로 변경
			strFileExt.MakeUpper();
			if (file.IsDirectory()) {
				continue;
			}
			strFolderItem.Replace(fileName, _T(""));

		}
		//파일의 이름
		CString _fileName = file.GetFileName();


		// 현재폴더 상위폴더 썸네일파일은 제외
		if (_fileName == _T("Thumbs.db")) {
			continue;
		}

		strFolderItem = file.GetFileName();
		//읽어온 파일 이름을 리스트박스에 넣음
		if (strFileExt == ".XLS" || strFileExt == ".XLSX")
		{
			m_ExcelFileList.AddString(strFolderItem);
			gv_bExcelFolder = true;
		}
	}
	file.Close();
}

void CautoInputToolDlg::OnLbnSelchangeOutputlist()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/07/27
함수:			CautoInputToolDlg::OnLbnDblclkInputlist
세부사항:		static, public
설명:			왼쪽 리스트 항목을 더블클릭했을경우의 이벤트
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::OnLbnDblclkInputlist()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	ImageViewFunc();
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/07/27
함수:			CautoInputToolDlg::SelectedImgPreview
세부사항:		static, public
설명:			선택한 항목의 이미지를 Preview창에서 보여줌
매개변수:		없음
리턴값:			없음
노트:			
*****************************************************************************************/
void CautoInputToolDlg::SelectedImgPreview()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	CString cstrItemSelected, cstrFileAbsolutePath;
	CListBox *p_list = (CListBox *)GetDlgItem(IDC_InputList);
	int index = p_list->GetCurSel();

	if (index != LB_ERR) {
		
		p_list->GetText(index, cstrItemSelected);
		//AfxMessageBox(cstrItemSelected);							//확인용 코드
	}
	cstrFileAbsolutePath = gv_InputFilePath + '\\' + cstrItemSelected;
	
	CT2CA pszString(cstrFileAbsolutePath);
	std::string strPath(pszString);
	
	m_matImage = imread(strPath, IMREAD_UNCHANGED);
	
	CreateBitmapInfo(m_matImage.cols, m_matImage.rows, m_matImage.channels() * 8);
	DrawImage();
}


/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/07/27
함수:			CautoInputToolDlg::CreateBitmapInfo
세부사항:		static, public
설명:			pictureControl에서 사용할 BitMap이미지를 생성함.
매개변수:		int width 너비, int height 높이, int bpp 비트
리턴값:			없음
노트:			
*****************************************************************************************/
void CautoInputToolDlg::CreateBitmapInfo(int width, int height, int bpp)
{
	if (m_pBitmapInfo != NULL)
	{
		delete m_pBitmapInfo;
		m_pBitmapInfo = NULL;
	}

	if (bpp == 8)
		m_pBitmapInfo = (BITMAPINFO *) new BYTE[sizeof(BITMAPINFO) + 255 * sizeof(RGBQUAD)];
	else // 24 or 32bit
		m_pBitmapInfo = (BITMAPINFO *) new BYTE[sizeof(BITMAPINFO)];

	m_pBitmapInfo->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	m_pBitmapInfo->bmiHeader.biPlanes = 1;
	m_pBitmapInfo->bmiHeader.biBitCount = bpp;
	m_pBitmapInfo->bmiHeader.biCompression = BI_RGB;
	m_pBitmapInfo->bmiHeader.biSizeImage = 0;
	m_pBitmapInfo->bmiHeader.biXPelsPerMeter = 0;
	m_pBitmapInfo->bmiHeader.biYPelsPerMeter = 0;
	m_pBitmapInfo->bmiHeader.biClrUsed = 0;
	m_pBitmapInfo->bmiHeader.biClrImportant = 0;

	if (bpp == 8)
	{
		for (int i = 0; i < 256; i++)
		{
			m_pBitmapInfo->bmiColors[i].rgbBlue = (BYTE)i;
			m_pBitmapInfo->bmiColors[i].rgbGreen = (BYTE)i;
			m_pBitmapInfo->bmiColors[i].rgbRed = (BYTE)i;
			m_pBitmapInfo->bmiColors[i].rgbReserved = 0;
		}
	}

	m_pBitmapInfo->bmiHeader.biWidth = width;
	m_pBitmapInfo->bmiHeader.biHeight = -height;
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/07/27
함수:			CautoInputToolDlg::DrawImage
세부사항:		static, public
설명:			Dlg에 있는 PictureControl IDC_PC_VIEW에 BitMap 이미지를 그림
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::DrawImage()
{
	CClientDC dc(GetDlgItem(IDC_PC_VIEW));

	CRect rect;
	GetDlgItem(IDC_PC_VIEW)->GetClientRect(&rect);

	SetStretchBltMode(dc.GetSafeHdc(), COLORONCOLOR);
	StretchDIBits(dc.GetSafeHdc(), 0, 0, rect.Width(), rect.Height(), 0, 0, m_matImage.cols, m_matImage.rows, m_matImage.data, m_pBitmapInfo, DIB_RGB_COLORS, SRCCOPY);
}


/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/07/27
함수:			CautoInputToolDlg::PreTranslateMessage
세부사항:		static, public
설명:			Enter키, ESC키 제어
매개변수:		MSG* pMsg
리턴값:			BOOL 키가 눌렸는지의 여부를 리턴
노트:
*****************************************************************************************/
BOOL CautoInputToolDlg::PreTranslateMessage(MSG* pMsg)
{
	// TODO: 여기에 특수화된 코드를 추가 및/또는 기본 클래스를 호출합니다.

	if (pMsg->message == WM_KEYDOWN)
	{
		if (pMsg->wParam == VK_RETURN)		// ENTER키 누를 시
		{
			ImageViewFunc();
			return TRUE;
		}
		else if (pMsg->wParam == VK_ESCAPE) // ESC키 누를 시
			return TRUE;
	}
	return CDialogEx::PreTranslateMessage(pMsg);
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/13
함수:			CautoInputToolDlg::ImageViewFunc
세부사항:		static, public
설명:			이미지를 출력해서 보여주는 함수
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::ImageViewFunc() 
{
	CString cstrFileName, cstrFileAbsolutePath;
	CListBox *p_list = (CListBox *)GetDlgItem(IDC_InputList);
	int index = p_list->GetCurSel();
	string strPrevimg = "";

	if (index != LB_ERR) {
		p_list->GetText(index, cstrFileName);
	}
	cstrFileAbsolutePath = gv_InputFilePath + '\\' + cstrFileName;		//전체경로 + 파일이름
	//AfxMessageBox(FileAbsolutePath);									//경로확인용코드
	
	string strFileAbsolutePath = string(CT2CA(cstrFileAbsolutePath));	//OpenCV에서 사용하기위해 CString -> string 으로 형변환
	string strFileName = string(CT2CA(cstrFileName));

	Mat image = imread(strFileAbsolutePath, IMREAD_ANYCOLOR);			//이미지를 읽어와 Mat Iamge에 넣음

	if (strPrevimg.empty() != 0)
	{
		destroyAllWindows();
	}
		
	namedWindow(strFileName, WINDOW_NORMAL);							//이미지창을 조절가능한 상태로 open함
	imshow(strFileName, image);											//이미지를 보여줌
	resizeWindow(strFileName, 1280, 480);								//이미지를 오픈할때 이미지창의 크기를 1280 480으로 고정함.
	moveWindow(strFileName, 40, 30);
	
	if (waitKey() == 27) {												//ESC키를 입력했을경우
		
		destroyWindow(strFileName);										//창을 닫음
	}
	strPrevimg = strFileName;
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/11/03
함수:			CautoInputToolDlg::OnCbnSelchangePlatform
세부사항:		static, public
설명:			플랫폼 선택에 따라 언어설정을 다르게 해주는 함수
매개변수:		없음
리턴값:			없음
노트:			
*****************************************************************************************/
void CautoInputToolDlg::OnCbnSelchangePlatform()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	UpdateData();

	CString platformData = L"";
	m_PlatformListBox.GetLBText(m_PlatformListBox.GetCurSel(), platformData);

	if (platformData == _T("STD5W"))
	{
		m_LanguageCtrlBox.ResetContent();
		EnableBox1to5(FALSE);
		EnableBox6(FALSE);
		inputKAndS(9);
		InputCountryName(true);
		INIWriteString(_T("AutoInputTool"), _T("PlatForm"), _T("0"), gv_InIFilePath);
		INIWriteString(_T("AutoInputTool"), _T("Language"), _T("0"), gv_InIFilePath);
	}
	else if (platformData == _T("STD5") || platformData == _T("PRM5"))
	{
		m_LanguageCtrlBox.ResetContent();
		inputKAndS(9);
		EnableBox1to5(FALSE);
		EnableBox6(FALSE);
		InputCountryName(false);
		if (platformData == _T("STD5"))
		{
			INIWriteString(_T("AutoInputTool"), _T("PlatForm"), _T("1"), gv_InIFilePath);
		}
		else
		{
			INIWriteString(_T("AutoInputTool"), _T("PlatForm"), _T("2"), gv_InIFilePath);
		}
		INIWriteString(_T("AutoInputTool"), _T("Language"), _T("1"), gv_InIFilePath);
	}
	
}


/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/11
함수:			CautoInputToolDlg::inputKAndS
세부사항:		static, public
설명:			플랫폼 선택에 따라 언어설정을 다르게 해주는 함수
매개변수:		int number
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::inputKAndS(int number)
{
	switch (number)
	{
	case 0:
		m_Language1.SetWindowTextW(_T("1한국"));
		m_Language2.SetWindowTextW(_T("2영국"));
		m_Language3.SetWindowTextW(_T("3독일"));
		m_Language4.SetWindowTextW(_T("4프랑스"));
		m_Language5.SetWindowTextW(_T("5이탈리아"));
		m_Language6.SetWindowTextW(_T("6스페인"));
		break;
	case 1:
		m_Language1.SetWindowTextW(_T("7포르투갈"));
		m_Language2.SetWindowTextW(_T("8네덜란드"));
		m_Language3.SetWindowTextW(_T("9덴마크"));
		m_Language4.SetWindowTextW(_T("10스웨덴"));
		m_Language5.SetWindowTextW(_T("11노르웨이"));
		m_Language6.SetWindowTextW(_T("12폴란드"));
		break;
	case 2:
		m_Language1.SetWindowTextW(_T("13체코"));
		m_Language2.SetWindowTextW(_T("14슬로바키아"));
		m_Language3.SetWindowTextW(_T("15러시아"));
		m_Language4.SetWindowTextW(_T("16핀란드"));
		m_Language5.SetWindowTextW(_T("17헝가리"));
		m_Language6.SetWindowTextW(_T("18터키"));
		break;
	case 3:
		m_Language1.SetWindowTextW(_T("19그리스"));
		m_Language2.SetWindowTextW(_T("20슬로베니아"));
		m_Language3.SetWindowTextW(_T("21우크라이나"));
		m_Language4.SetWindowTextW(_T("22불가리아"));
		m_Language5.SetWindowTextW(_T("23루마니아"));
		m_Language6.SetWindowTextW(_T("24크로아티아"));
		break;
	case 4:
		m_Language1.SetWindowTextW(_T("13체코"));
		m_Language2.SetWindowTextW(_T("14슬로바키아"));
		m_Language3.SetWindowTextW(_T("15러시아"));
		m_Language4.SetWindowTextW(_T("16헝가리"));
		m_Language5.SetWindowTextW(_T("17터키"));
		m_Language6.SetWindowTextW(_T("-"));
		break;
	case 9:
		m_Language1.SetWindowTextW(_T("-"));
		m_Language2.SetWindowTextW(_T("-"));
		m_Language3.SetWindowTextW(_T("-"));
		m_Language4.SetWindowTextW(_T("-"));
		m_Language5.SetWindowTextW(_T("-"));
		m_Language6.SetWindowTextW(_T("-"));
		break;
	default:
		break;
	}
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/11
함수:			CautoInputToolDlg::EnableBox6
세부사항:		static, public
설명:			6번째 Box를 활성화 비활성화 반복하는것을 라인수를 줄이기 위해 만든 함수
매개변수:		bool input
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::EnableBox6(bool input)
{
	GetDlgItem(IDC_TxtOverBox6)->EnableWindow(input);
	GetDlgItem(IDC_WrapErrBox6)->EnableWindow(input);
	GetDlgItem(IDC_NotMatchingBox6)->EnableWindow(input);
	GetDlgItem(IDC_AbbreviationErr6)->EnableWindow(input);
	GetDlgItem(IDC_NOTE6)->EnableWindow(input);
}

/*****************************************************************************************
함수상태:		작성완료
작성자:			최영두 주임연구원
최종수정일:		2022/10/17
함수:			CautoInputToolDlg::EnableBox1to5
세부사항:		static, public
설명:			1~5번째 Box를 활성화 비활성화 반복하는것을 라인수를 줄이기 위해 만든 함수
매개변수:		bool input
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::EnableBox1to5(bool input)
{
	GetDlgItem(IDC_TxtOverBox1)->EnableWindow(input);
	GetDlgItem(IDC_WrapErrBox1)->EnableWindow(input);
	GetDlgItem(IDC_NotMatchingBox1)->EnableWindow(input);
	GetDlgItem(IDC_AbbreviationErr1)->EnableWindow(input);
	GetDlgItem(IDC_NOTE1)->EnableWindow(input);

	GetDlgItem(IDC_TxtOverBox2)->EnableWindow(input);
	GetDlgItem(IDC_WrapErrBox2)->EnableWindow(input);
	GetDlgItem(IDC_NotMatchingBox2)->EnableWindow(input);
	GetDlgItem(IDC_AbbreviationErr2)->EnableWindow(input);
	GetDlgItem(IDC_NOTE2)->EnableWindow(input);

	GetDlgItem(IDC_TxtOverBox3)->EnableWindow(input);
	GetDlgItem(IDC_WrapErrBox3)->EnableWindow(input);
	GetDlgItem(IDC_NotMatchingBox3)->EnableWindow(input);
	GetDlgItem(IDC_AbbreviationErr3)->EnableWindow(input);
	GetDlgItem(IDC_NOTE3)->EnableWindow(input);

	GetDlgItem(IDC_TxtOverBox4)->EnableWindow(input);
	GetDlgItem(IDC_WrapErrBox4)->EnableWindow(input);
	GetDlgItem(IDC_NotMatchingBox4)->EnableWindow(input);
	GetDlgItem(IDC_AbbreviationErr4)->EnableWindow(input);
	GetDlgItem(IDC_NOTE4)->EnableWindow(input);

	GetDlgItem(IDC_TxtOverBox5)->EnableWindow(input);
	GetDlgItem(IDC_WrapErrBox5)->EnableWindow(input);
	GetDlgItem(IDC_NotMatchingBox5)->EnableWindow(input);
	GetDlgItem(IDC_AbbreviationErr5)->EnableWindow(input);
	GetDlgItem(IDC_NOTE5)->EnableWindow(input);
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/25
함수:			CautoInputToolDlg::OnBnClickedLoadfolderexcel
세부사항:		static, public
설명:			선택한 폴더의 엑셀 파일을 콤보박스에 표시하는 함수
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::OnBnClickedLoadfolderexcel()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	m_ExcelFileList.ResetContent();		//ExcelFileList 비워주기
	int iCheckPathExist = m_InputAddExcel.GetCount();

	if (iCheckPathExist == 0) //입력된 경로가 존재하는지 확인
	{
		MessageBox(L"선택된 경로가 없습니다.", L"입력된 경로 없음", MB_ICONERROR);
		return;
	}
	else
	{
		CString cstrtempPath = L"", cstrFileExtention = L"";
		CListBox *p_list = (CListBox *)GetDlgItem(IDC_InputAddExcel);
		p_list->GetText(0, cstrtempPath);

		AfxExtractSubString(cstrFileExtention, cstrtempPath, 1, '.');

		if (cstrFileExtention == "")	//폴더인지 확인
		{
			GetUserSelectFolderExcel(cstrtempPath);
			if (m_ExcelFileList.GetCount() == 0) {
				MessageBox(L"선택한 폴더내에 엑셀파일이 없습니다.", L"엑셀파일 없음", MB_ICONERROR);
			}
			cstrtempPath.Empty();
			m_ExcelFileList.SetCurSel(0);
		}
		else
		{																		//폴더가 아닐경우 경고메시지 출력후 초기화
			MessageBox(L"엑셀 폴더가 아닙니다.", L"엑셀 폴더가 아님", MB_ICONERROR);
			m_ExcelFileList.ResetContent();
		}
	}
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/19
함수:			CautoInputToolDlg::OutputListFindNumber
세부사항:		static, public
설명:			선택된 폴더의 이미지 폴더 번호를 표시해서 사용자에게 보여주는 함수
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::OutputListFindNumber()
{
	CString cstrFindTxt = _T(""), cstrFindNumber = _T("");
	CListBox *p_list = (CListBox *)GetDlgItem(IDC_InputList);
	p_list->GetText(0, cstrFindTxt);

	cstrFindNumber = cstrFindTxt.Left(cstrFindTxt.Find('_')-1);
	//AfxMessageBox(cstrFindNumber);
	m_FolderNumber.SetWindowText(cstrFindNumber);
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/23
함수:			CautoInputToolDlg::OnCbnSelchangeExtensionselect
세부사항:		static, public
설명:			확장자 선택 시 PNG인 경우 mRadio : 0 반환, JPG 선택 시 mRadio : 1 반환
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::OnCbnSelchangeExtensionselect()
{
	CString List_Value;
	int List_index;
	List_index = m_ExtensionSelect.GetCurSel();

	if (List_index == 0)
	{
		mRadio = 0;
	}
	else if (List_index == 1)
	{
		mRadio = 1;
	}
	else
	{
		//nothing
	}
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/21
함수:			CautoInputToolDlg::SelectControlBox
세부사항:		public
설명:			시트번호로 값을 가져올 컨트롤을 선택하여 값 저장
매개변수:		int SheetNumber 시트번호
리턴값:			vector<CString> Value 검증결과/비고 값
노트:
*****************************************************************************************/
vector<CString> CautoInputToolDlg::SelectControlBox(int SheetNumber)
{
	vector<CString> Value;
	Value.resize(5);

	switch (SheetNumber)
	{
	case 1:
		m_TxtOverBox1.GetLBText(m_TxtOverBox1.GetCurSel(), Value[0]);
		m_WrapErrBox1.GetLBText(m_WrapErrBox1.GetCurSel(), Value[1]);
		m_NotMatchingBox1.GetLBText(m_NotMatchingBox1.GetCurSel(), Value[2]);
		m_AbbreviationErr1.GetLBText(m_AbbreviationErr1.GetCurSel(), Value[3]);
		GetDlgItemText(IDC_NOTE1, Value[4]);
		break;
	case 2:
		m_TxtOverBox2.GetLBText(m_TxtOverBox2.GetCurSel(), Value[0]);
		m_WrapErrBox2.GetLBText(m_WrapErrBox2.GetCurSel(), Value[1]);
		m_NotMatchingBox2.GetLBText(m_NotMatchingBox2.GetCurSel(), Value[2]);
		m_AbbreviationErr2.GetLBText(m_AbbreviationErr2.GetCurSel(), Value[3]);
		GetDlgItemText(IDC_NOTE2, Value[4]);
		break;
	case 3:
		m_TxtOverBox3.GetLBText(m_TxtOverBox3.GetCurSel(), Value[0]);
		m_WrapErrBox3.GetLBText(m_WrapErrBox3.GetCurSel(), Value[1]);
		m_NotMatchingBox3.GetLBText(m_NotMatchingBox3.GetCurSel(), Value[2]);
		m_AbbreviationErr3.GetLBText(m_AbbreviationErr3.GetCurSel(), Value[3]);
		GetDlgItemText(IDC_NOTE3, Value[4]);
		break;
	case 4:
		m_TxtOverBox4.GetLBText(m_TxtOverBox4.GetCurSel(), Value[0]);
		m_WrapErrBox4.GetLBText(m_WrapErrBox4.GetCurSel(), Value[1]);
		m_NotMatchingBox4.GetLBText(m_NotMatchingBox4.GetCurSel(), Value[2]);
		m_AbbreviationErr4.GetLBText(m_AbbreviationErr4.GetCurSel(), Value[3]);
		GetDlgItemText(IDC_NOTE4, Value[4]);
		break;
	case 5:
		m_TxtOverBox5.GetLBText(m_TxtOverBox5.GetCurSel(), Value[0]);
		m_WrapErrBox5.GetLBText(m_WrapErrBox5.GetCurSel(), Value[1]);
		m_NotMatchingBox5.GetLBText(m_NotMatchingBox5.GetCurSel(), Value[2]);
		m_AbbreviationErr5.GetLBText(m_AbbreviationErr5.GetCurSel(), Value[3]);
		GetDlgItemText(IDC_NOTE5, Value[4]);
		break;
	case 6:
		m_TxtOverBox6.GetLBText(m_TxtOverBox6.GetCurSel(), Value[0]);
		m_WrapErrBox6.GetLBText(m_WrapErrBox6.GetCurSel(), Value[1]);
		m_NotMatchingBox6.GetLBText(m_NotMatchingBox6.GetCurSel(), Value[2]);
		m_AbbreviationErr6.GetLBText(m_AbbreviationErr6.GetCurSel(), Value[3]);
		GetDlgItemText(IDC_NOTE6, Value[4]);
		break;
	default:
		break;
	}
	return Value;
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/25
함수:			CautoInputToolDlg::InputCountryName
세부사항:		static, public
설명:			다이얼로그에 나라명을 입력해주는 함수
매개변수:		bool type
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::InputCountryName(bool type)
{
	if (type) //STD5W
	{
		for (int i = 0; i < 4; i++) {
			m_LanguageCtrlBox.AddString(strCountryWArr[0 + 6 * i] + ", " + strCountryWArr[1 + 6 * i] + ", " + strCountryWArr[2 + 6 * i] + ", " +
				strCountryWArr[3 + 6 * i] + ", " + strCountryWArr[4 + 6 * i] + ", " + strCountryWArr[5 + 6 * i]);
		}
	}
	else //STD5, PRM5
	{
		for (int i = 0; i < 3; i++) {
			m_LanguageCtrlBox.AddString(strCountryPRMArr[0 + 6 * i] + ", " + strCountryPRMArr[1 + 6 * i] + ", " + strCountryPRMArr[2 + 6 * i] + ", " +
				strCountryPRMArr[3 + 6 * i] + ", " + strCountryPRMArr[4 + 6 * i] + ", " + strCountryPRMArr[5 + 6 * i]);
		}
	}
}

/*****************************************************************************************
함수상태:		작성 완료
최종수정일:		2022/10/28
함수:			CautoInputToolDlg::ImageFolderSelect
세부사항:		static, public
설명:			이미지 폴더의 하위 폴더 extension의 여부를 확인
					존재할 경우 extension폴더의 이미지로 이미지 파일 경로 저장
매개변수:		CString strFilePath 이미지 폴더 경로, CString ImageFile 이미지 파일 이름
리턴값:			CString ImageFilePath 이미지 파일 경로
노트:
*****************************************************************************************/
CString CautoInputToolDlg::ImageFolderSelect(CString strFilePath, CString ImageFile)
{
	CFileFind file;
	CString strFile, ImageFilePath, only_ImageFile;

	AfxExtractSubString(only_ImageFile, ImageFile, 0, '.');

	strFile = strFilePath + _T("\\*.*");

	BOOL bWorking = file.FindFile(strFile);
	BOOL FileExistCheck = FALSE;

	while (bWorking)
	{
		bWorking = file.FindNextFile();

		if (file.IsDots() || file.IsDots())
			continue;

		if (file.IsDirectory())
		{
			CString cstrDir = file.GetFilePath();
			CString FolderName = cstrDir.Mid(cstrDir.ReverseFind('\\'));

			if (FolderName == "\\extensions")
			{
				if (ImageFileCheck(cstrDir, only_ImageFile) == "")
					continue;
				else if (ImageFileCheck(cstrDir, only_ImageFile) != "")
				{
					ImageFilePath = ImageFileCheck(cstrDir, only_ImageFile);
					FileExistCheck = TRUE;
					break;
				}
			}
		}
	}

	if (FileExistCheck == FALSE)
	{
		bWorking = file.FindFile(strFile);

		while (bWorking)
		{
			bWorking = file.FindNextFile();

			if (file.IsDirectory() || file.IsDots())
				continue;

			if (!file.IsDots())
			{
				CString ImageFileName = file.GetFileName();
				if (ImageFileName == ImageFile)
				{
					ImageFilePath = strFilePath + "\\" + ImageFile;
					break;
				}
				else
					continue;
			}
		}
	}
	return ImageFilePath;
	file.Close();
}

/*****************************************************************************************
함수상태:		작성 완료
최종수정일:		2022/10/28
함수:			CautoInputToolDlg::ImageFileCheck
세부사항:		static, public
설명:			extension 폴더에서 변환한 이미지가 있는지 확인
매개변수:		CString extensionsFolderPath  extensions폴더 경로, CString onlyImageFile 이미지 파일 이름
리턴값:			CString Path_Or_Blank 파일이 있으면 경로, 파일이 없으면 값 없음
노트:
*****************************************************************************************/
CString CautoInputToolDlg::ImageFileCheck(CString extensionsFolderPath, CString onlyImageFile)
{
	CFileFind file;
	CString strFile, ImageFilePath, Path_Or_Blank;

	strFile = extensionsFolderPath + _T("\\*.*");

	BOOL bWorking = file.FindFile(strFile);

	while (bWorking)
	{
		bWorking = file.FindNextFile();

		if (file.IsDirectory() || file.IsDots())
			continue;

		if (!file.IsDots())
		{
			CString cstrDir = file.GetFilePath();
			CString exImageFile, only_exImageFile;
			exImageFile = file.GetFileName();
			AfxExtractSubString(only_exImageFile, exImageFile, 0, '.');
			
			if (only_exImageFile == onlyImageFile)
			{
				Path_Or_Blank = cstrDir;
				break;
			}
			else
			{
				Path_Or_Blank = "";
				continue;
			}
		}
	}
	return Path_Or_Blank;
	file.Close();
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/31
함수:			CautoInputToolDlg::OnBnClickedOrimagecheck()
세부사항:		static, public
설명:			원본이미지 사용 체크박스 체크 여부에 따라서 에디트박스 활성화 비활성화
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::OnBnClickedOrimagecheck()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.

	if (m_OriImageUse.GetCheck())
	{
		GetDlgItem(IDC_ResizeLeftValue)->EnableWindow(FALSE);
		GetDlgItem(IDC_ResizeRightValue)->EnableWindow(FALSE);
		GetDlgItem(IDC_EXTENSIONSELECT)->EnableWindow(FALSE);
	}
	else
	{
		GetDlgItem(IDC_ResizeLeftValue)->EnableWindow(TRUE);
		GetDlgItem(IDC_ResizeRightValue)->EnableWindow(TRUE);
		GetDlgItem(IDC_EXTENSIONSELECT)->EnableWindow(TRUE);
	}
}
/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/31
함수:			CautoInputToolDlg::DeleteAllFiles
세부사항:		static, public
설명:			폴더안의 모든 파일을 제거하고 폴더까지 삭제하는 함수
매개변수:		CString dirName
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::DeleteAllFiles(CString dirName)
{
	CFileFind finder;

	BOOL bWorking = finder.FindFile((CString)dirName + "/*.*");
	CString cstrFilePath = dirName;

	while (bWorking)
	{
		bWorking = finder.FindNextFile();
		if (finder.IsDots())
		{
			continue;
		}

		CString filePath = finder.GetFilePath();
		DeleteFile(filePath);
	}
	RemoveDirectory(cstrFilePath);
	finder.Close();
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/11/03
함수:			CautoInputToolDlg::OnBnClickedDeletelog
세부사항:		static, public
설명:			로그삭제 버튼을 눌렀을떄 로그안의 내용을 초기화해주는 함수
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::OnBnClickedDeletelog()
{
	// TODO: 여기에 컨트롤 알림 처리기 코드를 추가합니다.
	m_EditLog.SendMessage(EM_SETREADONLY, (WPARAM)FALSE, (LPARAM)0);
	m_EditLog.SetSel(0, -1, TRUE);
	m_EditLog.Clear();
	m_EditLog.SendMessage(EM_SETREADONLY, (WPARAM)TRUE, (LPARAM)0);
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/11/30
함수:			CautoInputToolDlg::OnAdjustListBoxHScroll
세부사항:		static, public
설명:			오른쪽 리스트박스의 가로 스크롤바를 추가하는 함수
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::OnAdjustRightListBoxHScroll()
{
	CString str; CSize sz; int dx = 0;
	CDC *pDC = m_ListBox_Out.GetDC();

	for (int i = 0; i < m_ListBox_Out.GetCount(); i++)
	{
		m_ListBox_Out.GetText(i, str);
		sz = pDC->GetTextExtent(str);

		if (sz.cx > dx)
			dx = sz.cx;
	}
	m_ListBox_Out.ReleaseDC(pDC);

	if (m_ListBox_Out.GetHorizontalExtent() < dx)
	{
		m_ListBox_Out.SetHorizontalExtent(dx);
		ASSERT(m_at_list.GetHorizontalExtent() == dx);
	}
}
/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/11/30
함수:			CautoInputToolDlg::OnAdjustLeftListBoxHScroll
세부사항:		static, public
설명:			왼쪽 리스트박스의 가로 스크롤바를 추가하는 함수
매개변수:		없음
리턴값:			없음
노트:
*****************************************************************************************/
void CautoInputToolDlg::OnAdjustLeftListBoxHScroll()
{
	CString str; CSize sz; int dx = 0;
	CDC *pDC = m_ListBox.GetDC();

	for (int i = 0; i < m_ListBox.GetCount(); i++)
	{
		m_ListBox.GetText(i, str);
		sz = pDC->GetTextExtent(str);

		if (sz.cx > dx)
			dx = sz.cx;
	}
	m_ListBox.ReleaseDC(pDC);

	if (m_ListBox.GetHorizontalExtent() < dx)
	{
		m_ListBox.SetHorizontalExtent(dx);
		ASSERT(m_at_list.GetHorizontalExtent() == dx);
	}
}