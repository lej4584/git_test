
// ETMacroDlg.h : ��� ����
//

#pragma once
#include "afxcmn.h"
#include "afxwin.h"





// CETMacroDlg ��ȭ ����
class CETMacroDlg : public CDialogEx
{
// �����Դϴ�.
public:
	CETMacroDlg(CWnd* pParent = NULL);	// ǥ�� �������Դϴ�.
    virtual ~CETMacroDlg();

// ��ȭ ���� �������Դϴ�.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ETMACRO_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV �����Դϴ�.

    enum E_ThreadState : int
    {
        EThreadState_NotStarted = 0,        ///< ���۵� �� ����.
        EThreadState_Running,               ///< ������ ���� ��
        EThreadState_Paused,                ///< pause �Ǿ���.
        EThreadState_Stopped,               ///< stop �Ǿ���.
        EThreadState_Finished,              ///< �۾� �Ϸ�
    };

    struct ButtonPos
    {
        int nx;
        int ny;
    };


public:
    CRect m_rectScreen; // ����� �ػ�
    CRect m_rectClient; // dialog ���� ����� 0,0


    HWND m_webHwnd;

    CWinThread* m_pMainThread; // Main Thread
    E_ThreadState m_nCurState; // ���� Thread ����
    BOOL m_bContinue;

	// ��θ��� �� ��ư ��ġ
    ButtonPos m_btnWebFrameBar;
    ButtonPos m_btnHelloMarketFirstLogin;
    ButtonPos m_btnHelloMarketLogin;
	ButtonPos m_btnHelloMarketSellItem;
    ButtonPos m_btnHelloMarketGeneralItem;
    ButtonPos m_btnHelloMarketImages;
    ButtonPos m_btnHelloMarketBlank;

	// �������� �� ��ư ��ġ
    ButtonPos m_btnThunderMarketLogout;
    ButtonPos m_btnThunderMarketFirstLogin;
    ButtonPos m_btnThunderMarketImageUpload;
    ButtonPos m_btnThunderMarketBaseInfo;
    ButtonPos m_btnThunderMarketDetailInfo;
    ButtonPos m_btnThunderMarketRecentlyCity;
    ButtonPos m_btnThunderMarketAccept;


    // ���̹� ī�� �� ��ư ��ġ
    ButtonPos m_btnNaverLogout;
    ButtonPos m_btnNaverLogin;
    ButtonPos m_btnNaverBlank;
    ButtonPos m_btnNaverSellItem;
    ButtonPos m_btnNaverSmartEditorLine;
    ButtonPos m_btnNaverSellTabLine;
    ButtonPos m_btnNaverPhotoUploaderCenter;


    // ���̹� ��� �� ��ư ��ġ
    ButtonPos m_btnNaverBandNewPostPos;


    // īī�����丮 �� ��ư ��ġ
    ButtonPos m_btnKakaoRobot;
    ButtonPos m_btnKakaoLogin;
    ButtonPos m_btnKakaoLogin2;
    ButtonPos m_btnKakaoSellItem;
    ButtonPos m_btnKakaoSellAccept;



    // ��θ��� ī�װ� �޺��ڽ�
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

    // ��θ��� ��ũ�� ���� �Լ�
    void PlayMacroHelloMarket();
    void HelloMarketLogin();

    // �������� ��ũ�� ���� �Լ�
    void PlayMacroThunderMarket();
    void ThunderMarketLogin();

    // ���̹� ��� ��ũ�� ���� �Լ�
    void PlayMacroNaverBand();
    void NaverBandLogin();

    // ���̹� ī�� ��ũ�� ���� �Լ�
    void PlayMacroNaverCafe();
    void NaverCafeLogin();

    // ���̹� ī�� ��ũ�� ���� �Լ�
    void PlayMacroKakaoStory();
    void KakaoStoryLogin();

    ////////////////////////////////////////

    void StandBy(int time);

    void MouseLbuttonClickEvent(ButtonPos pos);
    void MouseMoveEvent(ButtonPos pos);
    void KeyboardButtonClickEvent(BYTE byte, int sleep);
    void EventClipboardPaste();
    void EventInputURLEditMode();


    // ���� �Ľ� �Լ�
    bool ReadExcelAccount(std::wstring wstr);
    bool ReadExcelSellItemsFile();

    bool WriteExcelFile(const std::wstring wstr);

    //Ŭ������ ���� �Լ�
    int CopyTextToClipboard(CString strCopy);

    void SetItemChangeComboHm1(int nCur);
    void SetItemChangeComboTm1(int nCur);
    void SetItemChangeComboTm2(int nCur);
    void SetItemChangeComboNc(int nCur);

private:
    static UINT procWorkerThread(LPVOID lpParam);



// �����Դϴ�.
protected:
	HICON m_hIcon;

	// ������ �޽��� �� �Լ�
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
