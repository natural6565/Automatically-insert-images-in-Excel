
// DlgProxy.cpp: 구현 파일
//

#include "pch.h"
#include "framework.h"
#include "autoInputTool.h"
#include "DlgProxy.h"
#include "autoInputToolDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CautoInputToolDlgAutoProxy

IMPLEMENT_DYNCREATE(CautoInputToolDlgAutoProxy, CCmdTarget)

CautoInputToolDlgAutoProxy::CautoInputToolDlgAutoProxy()
{
	EnableAutomation();

	// 자동화 개체가 활성화되어 있는 동안 계속 응용 프로그램을 실행하기 위해
	//	생성자에서 AfxOleLockApp를 호출합니다.
	AfxOleLockApp();

	// 응용 프로그램의 주 창 포인터를 통해 대화 상자에 대한
	//  액세스를 가져옵니다.  프록시의 내부 포인터를 설정하여
	//  대화 상자를 가리키고 대화 상자의 후방 포인터를 이 프록시로
	//  설정합니다.
	ASSERT_VALID(AfxGetApp()->m_pMainWnd);
	if (AfxGetApp()->m_pMainWnd)
	{
		ASSERT_KINDOF(CautoInputToolDlg, AfxGetApp()->m_pMainWnd);
		if (AfxGetApp()->m_pMainWnd->IsKindOf(RUNTIME_CLASS(CautoInputToolDlg)))
		{
			m_pDialog = reinterpret_cast<CautoInputToolDlg*>(AfxGetApp()->m_pMainWnd);
			m_pDialog->m_pAutoProxy = this;
		}
	}
}

CautoInputToolDlgAutoProxy::~CautoInputToolDlgAutoProxy()
{
	// 모든 개체가 OLE 자동화로 만들어졌을 때 응용 프로그램을 종료하기 위해
	// 	소멸자가 AfxOleUnlockApp를 호출합니다.
	//  이러한 호출로 주 대화 상자가 삭제될 수 있습니다.
	if (m_pDialog != nullptr)
		m_pDialog->m_pAutoProxy = nullptr;
	AfxOleUnlockApp();
}

void CautoInputToolDlgAutoProxy::OnFinalRelease()
{
	// 자동화 개체에 대한 마지막 참조가 해제되면
	// OnFinalRelease가 호출됩니다.  기본 클래스에서 자동으로 개체를 삭제합니다.
	// 기본 클래스를 호출하기 전에 개체에 필요한 추가 정리 작업을
	// 추가하세요.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CautoInputToolDlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CautoInputToolDlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// 참고: IID_IautoInputTool에 대한 지원을 추가하여
//  VBA에서 형식 안전 바인딩을 지원합니다.
//  이 IID는 .IDL 파일에 있는 dispinterface의 GUID와 일치해야 합니다.

// {8aa0cea3-ccc4-4657-a525-3d2f6c01224b}
static const IID IID_IautoInputTool =
{0x8aa0cea3,0xccc4,0x4657,{0xa5,0x25,0x3d,0x2f,0x6c,0x01,0x22,0x4b}};

BEGIN_INTERFACE_MAP(CautoInputToolDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CautoInputToolDlgAutoProxy, IID_IautoInputTool, Dispatch)
END_INTERFACE_MAP()

// The IMPLEMENT_OLECREATE2 macro is defined in pch.h of this project
// {c4fb20f8-1b21-418e-8b79-5b251ec98ec9}
IMPLEMENT_OLECREATE2(CautoInputToolDlgAutoProxy, "autoInputTool.Application", 0xc4fb20f8,0x1b21,0x418e,0x8b,0x79,0x5b,0x25,0x1e,0xc9,0x8e,0xc9)


// CautoInputToolDlgAutoProxy 메시지 처리기
