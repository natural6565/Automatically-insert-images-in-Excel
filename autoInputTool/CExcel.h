#pragma once

#include <vector>

#include "autoInputTool.h"
#include "autoInputToolDlg.h"

#include "CApplication.h"
#include "CFont0.h"
#include "Cnterior.h"
#include "CRange.h"
#include "CRanges.h"
#include "CShapes.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"

// CExcel
using namespace std;

class CExcel : public CWnd
{
	DECLARE_DYNAMIC(CExcel)

public:
	CExcel();
	virtual ~CExcel();

protected:
	DECLARE_MESSAGE_MAP()

public:
	void StartExcel(CString ExcelFilePath);
	void CloseExcel();
	void CellResizeExcel(int SheetCount);
	void ResultValueExcel(int SheetNumber, int RowNumber, CString Result, int cell_pos);
	void ResultStyleExcel(int SheetNumber, int RowNumber);
	void PictureInsertExcel(int SheetNumber, int RowNumber, CString ImageFileName, CString ImageFilePath);
	int SheetNumberExcel(CString ImageFileName);
	int RowNumberExcel(CString ImageFileName);
	int SheetsCountExcel();
	vector<CString> SheetNameExcel(int SheetCount);
};