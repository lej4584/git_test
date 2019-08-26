
// ETMacroDlg.h : 헤더 파일
//

#pragma once
#include "afxcmn.h"
#include "afxwin.h"





// CETMacroDlg 대화 상자
class CETMacroDlg : public CDialogEx
{
// 생성입니다.
public:
	CETMacroDlg(CWnd* pParent = NULL);	// 표준 생성자입니다.
    virtual ~CETMacroDlg();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ETMACRO_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 지원입니다.

    enum E_ThreadState : int
    {
        EThreadState_NotStarted = 0,        ///< 시작된 적 없음.
        EThreadState_Running,               ///< 스레드 동작 중
        EThreadState_Paused,                ///< pause 되었음.
        EThreadState_Stopped,               ///< stop 되었음.
        EThreadState_Finished,              ///< 작업 완료
    };

    struct ButtonPos
    {
        int nx;
        int ny;
    };


public:
    CRect m_rectScreen; // 모니터 해상도
    CRect m_rectClient; // dialog 좌측 상단이 0,0


    HWND m_webHwnd;

    CWinThread* m_pMainThread; // Main Thread
    E_ThreadState m_nCurState; // 현재 Thread 상태
    BOOL m_bContinue;

	// 헬로마켓 웹 버튼 위치
    ButtonPos m_btnWebFrameBar;
    ButtonPos m_btnHelloMarketFirstLogin;
    ButtonPos m_btnHelloMarketLogin;
	ButtonPos m_btnHelloMarketSellItem;
    ButtonPos m_btnHelloMarketGeneralItem;
    ButtonPos m_btnHelloMarketImages;
    ButtonPos m_btnHelloMarketBlank;

	// 번개장터 웹 버튼 위치
    ButtonPos m_btnThunderMarketLogout;
    ButtonPos m_btnThunderMarketFirstLogin;
    ButtonPos m_btnThunderMarketImageUpload;
    ButtonPos m_btnThunderMarketBaseInfo;
    ButtonPos m_btnThunderMarketDetailInfo;
    ButtonPos m_btnThunderMarketRecentlyCity;
    ButtonPos m_btnThunderMarketAccept;


    // 네이버 카페 웹 버튼 위치
    ButtonPos m_btnNaverLogout;
    ButtonPos m_btnNaverLogin;
    ButtonPos m_btnNaverBlank;
    ButtonPos m_btnNaverSellItem;
    ButtonPos m_btnNaverSmartEditorLine;
    ButtonPos m_btnNaverSellTabLine;
    ButtonPos m_btnNaverPhotoUploaderCenter;


    // 네이버 밴드 웹 버튼 위치
    ButtonPos m_btnNaverBandNewPostPos;


    // 카카오스토리 웹 버튼 위치
    ButtonPos m_btnKakaoRobot;
    ButtonPos m_btnKakaoLogin;
    ButtonPos m_btnKakaoLogin2;
    ButtonPos m_btnKakaoSellItem;
    ButtonPos m_btnKakaoSellAccept;



    // 헬로마켓 카테고리 콤보박스
    CComboBox m_combo_hm1;
    CComboBox m_combo_hm2;

    CComboBox m_combo_tm1;
    CComboBox m_combo_tm2;
    CComboBox m_combo_tm3;
    CComboBox m_combo_tm_size;

    CComboBox m_combo_tm_quality;

    CComboBox m_combo_nc;

	CListCtrl m_list_sellitem;

    CStatic m_static_image[6];

	CButton m_check_site0;
	CButton m_check_site1;
	CButton m_check_site2;
	CButton m_check_site3;
	CButton m_check_site4;

    CButton m_check_deliverycost;
    CButton m_check_exchange;

    CButton m_ckeck_hm_skip;

    CButton m_radio_speed0;
    CButton m_radio_speed1;
    CButton m_radio_speed2;

    CEdit m_edit_navercafe_title;
    CEdit m_edit_navercafe_cost;

    float   m_fScale;
    int     m_nCurrentSelectImagesList;
    std::vector<CString> m_vecAddImageName;

public:
    CString GetApplicationDirectory();

    void InitAccount();
    void InitSellItems();
    void InitChromeURL();
    void InitWebButtonPos();

    void InitControls();

    void AddSellItemList();
    void AddItemThunderMarketComboboxs();
    void AddItemHelloMarketComboboxs();
    void AddItemNaverCafeComboboxs();

    void PlayMacro();

    // 헬로마켓 매크로 관련 함수
    void PlayMacroHelloMarket();
    void HelloMarketLogin();

    // 번개장터 매크로 관련 함수
    void PlayMacroThunderMarket();
    void ThunderMarketLogin();

    // 네이버 밴드 매크로 관련 함수
    void PlayMacroNaverBand();
    void NaverBandLogin();

    // 네이버 카페 매크로 관련 함수
    void PlayMacroNaverCafe();
    void NaverCafeLogin();

    // 네이버 카페 매크로 관련 함수
    void PlayMacroKakaoStory();
    void KakaoStoryLogin();

    ////////////////////////////////////////

    void StandBy(int time);

    void MouseLbuttonClickEvent(ButtonPos pos);
    void MouseMoveEvent(ButtonPos pos);
    void KeyboardButtonClickEvent(BYTE byte, int sleep);
    void EventClipboardPaste();
    void EventInputURLEditMode();


    // 엑셀 파싱 함수
    bool ReadExcelAccount(std::wstring wstr);
    bool ReadExcelSellItemsFile();

    bool WriteExcelFile(const std::wstring wstr);

    //클립보드 복사 함수
    int CopyTextToClipboard(CString strCopy);

    void SetItemChangeComboHm1(int nCur);
    void SetItemChangeComboTm1(int nCur);
    void SetItemChangeComboTm2(int nCur);
    void SetItemChangeComboNc(int nCur);

private:
    static UINT procWorkerThread(LPVOID lpParam);



// 구현입니다.
protected:
	HICON m_hIcon;

	// 생성된 메시지 맵 함수
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

   

public:
    afx_msg void OnBnClickedButtonMacroStart();
    afx_msg void OnCbnSelchangeComboHm1();
    afx_msg void OnCbnSelchangeComboTm1();
    afx_msg void OnCbnSelchangeComboTm2();
    afx_msg void OnLvnItemchangedListSellitem(NMHDR *pNMHDR, LRESULT *pResult);
    afx_msg void OnBnClickedButtonExportExcel();
    afx_msg void OnBnClickedButtonExit();
    afx_msg void OnBnClickedCheckSkip();
    afx_msg void OnBnClickedRadioSpeed0();
    afx_msg void OnBnClickedRadioSpeed1();
    afx_msg void OnBnClickedRadioSpeed2();
};
