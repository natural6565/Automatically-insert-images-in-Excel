// CExcel.cpp: 구현 파일
//

#include "pch.h"
#include "CExcel.h"
#include <vector>

using namespace std;

// CExcel

IMPLEMENT_DYNAMIC(CExcel, CWnd)

CExcel::CExcel()
{

}

CExcel::~CExcel()
{
}


BEGIN_MESSAGE_MAP(CExcel, CWnd)
END_MESSAGE_MAP()



// CExcel 메시지 처리기

//엑셀 자동화 관련 변수
CApplication a_application;
CWorkbooks a_workbooks;
CWorkbook a_workbook;
CWorksheets a_worksheets;
CWorksheet a_worksheet;
Cnterior a_interior;
CRanges a_ranges;
CRange a_range;
CShapes a_shapes;
CFont0 a_font;

COleVariant covT((short)TRUE);
COleVariant covF((short)FALSE);
COleVariant covO((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

/*****************************************************************************************
함수상태:			작성완료
최종수정일:		2022/10/18
함수:			CExcel::StartExcel
세부사항:		public
설명:			Excel 파일을 실행
매개변수:		CString ExcelFilePath 엑셀 파일 경로
리턴값:			없음
노트:
*****************************************************************************************/
void CExcel::StartExcel(CString ExcelFilePath)
{
	if (!a_application.CreateDispatch(_T("Excel.Application")))
	{
		AfxMessageBox(_T("엑셀을 실행할 수 없습니다."));
		return;
	}
	//실행 과정 시각화(TRUE - 보이게 / FALSE - 보이지 않게)
	//a_application.put_Visible(TRUE);
	//a_application.put_UserControl(TRUE);
	a_application.put_Visible(FALSE);
	a_application.put_UserControl(FALSE);

	a_workbooks = a_application.get_Workbooks();
	a_workbook = a_workbooks.Open(ExcelFilePath, covO, covO, covO, covO, covO, covO, covO, covO, covO, covO, covO, covO, covO, covO);
	a_worksheets = a_workbook.get_Worksheets();
}

/*****************************************************************************************
함수상태:			작성완료
최종수정일:		2022/10/18
함수:			CExcel::CloseExcel
세부사항:		public
설명:			Excel 파일을 종료
매개변수:		없음
리턴값:			없음
노트:			a_workbook.Save() 으로 저장 후 종료
*****************************************************************************************/
void CExcel::CloseExcel()
{
	a_workbook.Save();
	a_font.ReleaseDispatch();
	a_interior.ReleaseDispatch();
	a_shapes.ReleaseDispatch();
	a_ranges.ReleaseDispatch();
	a_range.ReleaseDispatch();
	a_worksheet.ReleaseDispatch();
	a_worksheets.ReleaseDispatch();
	a_workbook.ReleaseDispatch();
	a_workbooks.Close();
	a_workbooks.ReleaseDispatch();
	a_application.Quit();
	a_application.ReleaseDispatch();
	a_application.DetachDispatch();
}


/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/21
함수:			CExcel::CellResizeExcel
세부사항:		public
설명:			엑셀 셀의 크기를 일정한 크기로 변경(전달 받은 양식을 기준으로 설정)
매개변수:		int SheetCount 시트개수, CString Platform 플랫폼
리턴값:			없음
노트:			플랫폼 별로 값이 작성된 셀의 범위가 달라 플랫폼 먼저 선택 필요
*****************************************************************************************/
void CExcel::CellResizeExcel(int SheetCount)
{
	//이미지 삽입 셀 크기 설정에 사용할 변수
	double A1_ColSize = 4.25;
	double A1_RowSize = 15.75;
	double B2_ColSize = 15.75;
	double B2_RowSize = 49.50;
	double C3_ColSize = 33.50;
	double C3_RowSize = 142.50;

	for (int i = 1; i <= SheetCount; i++)
	{
		a_worksheet = a_worksheets.get_Item(COleVariant((short)i));
		a_range = a_worksheet.get_UsedRange();

		LPDISPATCH lpDisp = a_range.SpecialCells(11, covO);
		a_range.AttachDispatch(lpDisp);

		int int_lastrow = a_range.get_Row();
		CString cstr_lastrow;
		cstr_lastrow.Format(_T("%d"), int_lastrow);
		cstr_lastrow = _T("C") + cstr_lastrow;

		a_range.AttachDispatch(a_worksheet.get_Cells(), TRUE);

		a_range = a_worksheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("A1")));
		a_range.put_ColumnWidth(COleVariant(A1_ColSize));
		a_range.put_RowHeight(COleVariant(A1_RowSize));

		a_range = a_worksheet.get_Range(COleVariant(_T("B2")), COleVariant(_T("B2")));
		a_range.put_ColumnWidth(COleVariant(B2_ColSize));
		a_range.put_RowHeight(COleVariant(B2_RowSize));

		a_range = a_worksheet.get_Range(COleVariant(_T("C3")), COleVariant(cstr_lastrow));
		a_range.put_ColumnWidth(COleVariant(C3_ColSize));
		a_range.put_RowHeight(COleVariant(C3_RowSize));
	}
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/21
함수:			CExcel::ResultValueExcel
세부사항:		public
설명:			검증결과 값을 엑셀에 삽입
매개변수:		int SheetNumber 시트번호, int RowNumber 행번호
				CString Result 검증결과, int cell_pos 열번호 지정
리턴값:			없음
노트:
*****************************************************************************************/
void CExcel::ResultValueExcel(int SheetNumber, int RowNumber, CString Result, int cell_pos)
{
	a_worksheet = a_worksheets.get_Item(COleVariant((short)SheetNumber));
	a_range.AttachDispatch(a_worksheet.get_Cells(), TRUE);
	a_range = a_worksheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("L200")));

	if (cell_pos == 4)
	{
		a_range.put_Item(COleVariant((long)(RowNumber)), COleVariant((long)(12)), COleVariant(Result));
	}
	else
	{
		a_range.put_Item(COleVariant((long)(RowNumber)), COleVariant((long)(5 + cell_pos)), COleVariant(Result));
	}
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/21
함수:			CExcel::ResultStyleExcel
세부사항:		public
설명:			검증결과 값에 따라 셀의 폰트와 배경색을 변경
매개변수:		int SheetNumber 시트번호, int RowNumber 행번호
리턴값:			없음
노트:
*****************************************************************************************/
void CExcel::ResultStyleExcel(int SheetNumber, int RowNumber)
{
	a_worksheet = a_worksheets.get_Item(COleVariant((short)SheetNumber));
	a_range.AttachDispatch(a_worksheet.get_Cells(), TRUE);

	int ascii = 65;	//ascii '65' = A
	int col_count = RowNumber;

	for (int i = 0; i < 4; i++)
	{
		CString temp_result;
		a_range = a_worksheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("L200")));
		temp_result = a_range.get_Item(COleVariant((long)col_count), COleVariant((long)(5 + i)));

		int row_ascii = ascii + 4 + i;
		char ch = char(row_ascii);

		CString row, col;
		row = ch;
		col.Format(_T("%d"), col_count);
		row = row + col;
		if (temp_result == "PASS")
		{
			a_range = a_worksheet.get_Range(COleVariant(row), COleVariant(row));
			a_font = a_range.get_Font();
			a_interior = a_range.get_Interior();
			a_font.put_Color(COleVariant((double)RGB(0, 0, 255)));
			a_interior.put_Color(COleVariant((double)RGB(255, 255, 255)));
		}
		else if (temp_result == "FAIL")
		{
			a_range = a_worksheet.get_Range(COleVariant(row), COleVariant(row));
			a_font = a_range.get_Font();
			a_interior = a_range.get_Interior();
			a_font.put_Color(COleVariant((double)RGB(255, 0, 0)));
			a_interior.put_Color(COleVariant((double)RGB(255, 255, 0)));
		}
		else
		{
			a_range = a_worksheet.get_Range(COleVariant(row), COleVariant(row));
			a_font = a_range.get_Font();
			a_interior = a_range.get_Interior();
			a_font.put_Color(COleVariant((double)RGB(0, 0, 0)));
			a_interior.put_Color(COleVariant((double)RGB(255, 255, 255)));
		}
	}
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/19
함수:			CExcel::PictureInsertExcel
세부사항:		public
설명:			이미지 파일을 위치를 지정하여 엑셀에 삽입
매개변수:		int SheetNumber 시트번호, int RowNumber 행번호
				CString ImageFileName 이미지 파일 이름, CString ImageFilePath 이미지 파일 경로
리턴값:			없음
노트:
*****************************************************************************************/
void CExcel::PictureInsertExcel(int SheetNumber, int RowNumber, CString ImageFileName, CString ImageFilePath)
{
	CautoInputToolDlg autoInputToolDlg;

	//이미지 위치 및 크기 설정에 사용할 변수
	float cell_leftpoint = 127.85f;
	float cell_toppoint = 65.75f;
	float cell_height = 142.25f;
	float cell_width = 204.50f;
	float cell_border = 0.25f;

	a_worksheet = a_worksheets.get_Item(COleVariant((short)SheetNumber));
	a_range.AttachDispatch(a_worksheet.get_Cells(), TRUE);
	a_range = a_worksheet.get_Range(COleVariant(_T("A1")), COleVariant(_T("L200")));
	a_shapes = a_worksheet.get_Shapes();

	a_range.put_Item(COleVariant((long)RowNumber), COleVariant((long)4), COleVariant(ImageFileName));

	a_shapes.AddPicture(ImageFilePath, FALSE, TRUE, cell_leftpoint, cell_toppoint + (cell_height*(RowNumber - 3)) + (cell_border*(RowNumber - 3)), cell_width, cell_height);
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/28
함수:			CExcel::SheetNumberExcel
세부사항:		public
설명:			이미지 파일 이름의 국가를 확인하여 시트 번호 확인
매개변수:		CString ImageNation 이미지 파일에서 추출한 국가(ex. 1한국)
리턴값:			int SheetNumber 시트번호
노트:
*****************************************************************************************/
int CExcel::SheetNumberExcel(CString ImageNation)
{
	CString Country;
	int SheetNumber;
	Country = ImageNation;

	if (Country == "1한국" || Country == "7포르투갈" || Country == "13체코" || Country == "19그리스")
		SheetNumber = 1;
	else if (Country == "2영국" || Country == "8네덜란드" || Country == "14슬로바키아" || Country == "20슬로베니아")
		SheetNumber = 2;
	else if (Country == "3독일" || Country == "9덴마크" || Country == "15러시아" || Country == "21우크라이나")
		SheetNumber = 3;
	else if (Country == "4프랑스" || Country == "10스웨덴" || Country == "16핀란드" || Country == "22불가리아" || Country == "16헝가리")
		SheetNumber = 4;
	else if (Country == "5이탈리아" || Country == "11노르웨이" || Country == "17헝가리" || Country == "23루마니아" || Country == "17터키")
		SheetNumber = 5;
	else if (Country == "6스페인" || Country == "12폴란드" || Country == "18터키" || Country == "24크로아티아")
		SheetNumber = 6;

	return SheetNumber;
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/21
함수:			CExcel::RowNumberExcel
세부사항:		public
설명:			이미지 파일 이름의 국가를 확인하여 행번호 확인
매개변수:		CString ImageFileName 이미지 파일 이름
리턴값:			int RowNumber 행번호
노트:
*****************************************************************************************/
int CExcel::RowNumberExcel(CString ImageFileName)
{
	CString Row;
	int pos_Row;

	AfxExtractSubString(Row, ImageFileName, 0, '번');
	pos_Row = _ttoi(Row);

	return pos_Row + 2;
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/21
함수:			CExcel::SheetsCountExcel
세부사항:		public
설명:			엑셀 파일의 시트 개수 확인
매개변수:		없음
리턴값:			int SheetCount 시트개수
노트:
*****************************************************************************************/
int CExcel::SheetsCountExcel()
{
	long SheetCount = a_worksheets.get_Count();
	return (int)SheetCount;
}

/*****************************************************************************************
함수상태:		작성완료
최종수정일:		2022/10/28
함수:			CExcel::SheetsNameExcel
세부사항:		public
설명:			엑셀 파일의 시트 이름 확인
매개변수:		int SheetCount 시트 개수
리턴값:			vector<CString> SheetName 시트이름
노트:
*****************************************************************************************/
vector<CString> CExcel::SheetNameExcel(int SheetCount) 
{
	vector<CString> SheetName;
	SheetName.resize(SheetCount);

	for (int i = 1; i <= SheetCount; i++)
	{
		a_worksheet = a_worksheets.get_Item(COleVariant((short)i));
		SheetName[i-1] = a_worksheet.get_Name();
	}
	return SheetName;
}