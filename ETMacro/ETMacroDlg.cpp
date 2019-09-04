
// ETMacroDlg.cpp : 구현 파일
//

#include "stdafx.h"
#include "ETMacro.h"
#include "ETMacroDlg.h"
#include "afxdialogex.h"

#include "ExcelFormat.h"
#include "Tokenizer.h"

#include "XLEzAutomation.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CETMacroDlg 대화 상자

using namespace ExcelFormat;
using namespace newSagaUtils;


int GetCStringtoAscii(CString str)
{
	int nAscii = -1;;;;
    if (str == L"a")
        nAscii = 65;
    else if (str == L"b")
        nAscii = 66;
    else if (str == L"c")
        nAscii = 67;
    else if (str == L"d")
        nAscii = 68;
    else if (str == L"e")
        nAscii = 69;
    else if (str == L"f")
        nAscii = 70;
    else if (str == L"g")
        nAscii = 71;
    else if (str == L"h")
        nAscii = 72;
    else if (str == L"i")
        nAscii = 73;
    else if (str == L"j")
        nAscii = 74;
    else if (str == L"k")
        nAscii = 75;
    else if (str == L"l")
        nAscii = 76;
    else if (str == L"m")
        nAscii = 77;
    else if (str == L"n")
        nAscii = 78;
    else if (str == L"o")
        nAscii = 79;
    else if (str == L"p")
        nAscii = 80;
    else if (str == L"q")
        nAscii = 81;
    else if (str == L"r")
        nAscii = 82;
    else if (str == L"s")
        nAscii = 83;
    else if (str == L"t")
        nAscii = 84;
    else if (str == L"u")
        nAscii = 85;
    else if (str == L"v")
        nAscii = 86;
    else if (str == L"w")
        nAscii = 87;
    else if (str == L"x")
        nAscii = 88;
    else if (str == L"y")
        nAscii = 89;
    else if (str == L"z")
        nAscii = 90;

    if (str == L"0")
        nAscii = 48;
    else if (str == L"1")
        nAscii = 49;
    else if (str == L"2")
        nAscii = 50;
    else if (str == L"3")
        nAscii = 51;
    else if (str == L"4")
        nAscii = 52;
    else if (str == L"5")
        nAscii = 53;
    else if (str == L"6")
        nAscii = 54;
    else if (str == L"7")
        nAscii = 55;
    else if (str == L"8")
        nAscii = 56;
    else if (str == L"9")
        nAscii = 57;

    return nAscii;
}




CETMacroDlg::CETMacroDlg(CWnd* pParent /*=NULL*/)
    : CDialogEx(IDD_ETMACRO_DIALOG, pParent)
{
    m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);

    m_pMainThread = NULL;

    m_bContinue = FALSE;

    m_nCurrentSelectImagesList = -1;

}


CETMacroDlg::~CETMacroDlg()
{
    if (m_pMainThread != NULL)
    {
        m_pMainThread = NULL;
    }
}



void CETMacroDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_COMBO_HM_1, m_combo_hm1);
    DDX_Control(pDX, IDC_COMBO_HM_2, m_combo_hm2);

    DDX_Control(pDX, IDC_COMBO_TM_1, m_combo_tm1);
    DDX_Control(pDX, IDC_COMBO_TM_2, m_combo_tm2);
    DDX_Control(pDX, IDC_COMBO_TM_3, m_combo_tm3);
    DDX_Control(pDX, IDC_COMBO_TM_SIZE, m_combo_tm_size);
    DDX_Control(pDX, IDC_COMBO_TM_ITEM_QUALITY, m_combo_tm_quality);

    DDX_Control(pDX, IDC_LIST_SELLITEM, m_list_sellitem);

    DDX_Control(pDX, IDC_CHECK_SITE0, m_check_site0);
    DDX_Control(pDX, IDC_CHECK_SITE1, m_check_site1);
    DDX_Control(pDX, IDC_CHECK_SITE2, m_check_site2);
    DDX_Control(pDX, IDC_CHECK_SITE3, m_check_site3);
    DDX_Control(pDX, IDC_CHECK_SITE4, m_check_site4);
    DDX_Control(pDX, IDC_CHECK_DELIVERY_COST, m_check_deliverycost);
    DDX_Control(pDX, IDC_CHECK_EXCHANGE, m_check_exchange);
    DDX_Control(pDX, IDC_STATIC_IMAGE1, m_static_image[0]);
    DDX_Control(pDX, IDC_STATIC_IMAGE2, m_static_image[1]);
    DDX_Control(pDX, IDC_STATIC_IMAGE3, m_static_image[2]);
    DDX_Control(pDX, IDC_STATIC_IMAGE4, m_static_image[3]);
    DDX_Control(pDX, IDC_STATIC_IMAGE5, m_static_image[4]);
    DDX_Control(pDX, IDC_STATIC_IMAGE6, m_static_image[5]);
    DDX_Control(pDX, IDC_RADIO_SPEED0, m_radio_speed0);
    DDX_Control(pDX, IDC_RADIO_SPEED1, m_radio_speed1);
    DDX_Control(pDX, IDC_RADIO_SPEED2, m_radio_speed2);
    DDX_Control(pDX, IDC_CHECK_SKIP, m_ckeck_hm_skip);
    DDX_Control(pDX, IDC_COMBO_NC, m_combo_nc);
    DDX_Control(pDX, IDC_EDIT_NAVERCAFE_TITLE, m_edit_navercafe_title);
    DDX_Control(pDX, IDC_EDIT_NAVERCAFE_COST, m_edit_navercafe_cost);
}

BEGIN_MESSAGE_MAP(CETMacroDlg, CDialogEx)
    ON_WM_PAINT()
    ON_WM_QUERYDRAGICON()
    ON_BN_CLICKED(IDC_BUTTON_MACRO_START, &CETMacroDlg::OnBnClickedButtonMacroStart)
    ON_CBN_SELCHANGE(IDC_COMBO_HM_1, &CETMacroDlg::OnCbnSelchangeComboHm1)
    ON_CBN_SELCHANGE(IDC_COMBO_TM_1, &CETMacroDlg::OnCbnSelchangeComboTm1)
    ON_CBN_SELCHANGE(IDC_COMBO_TM_2, &CETMacroDlg::OnCbnSelchangeComboTm2)
    ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST_SELLITEM, &CETMacroDlg::OnLvnItemchangedListSellitem)
    ON_BN_CLICKED(IDC_BUTTON_EXPORT_EXCEL, &CETMacroDlg::OnBnClickedButtonExportExcel)
    ON_BN_CLICKED(IDC_BUTTON_EXIT, &CETMacroDlg::OnBnClickedButtonExit)
    ON_BN_CLICKED(IDC_CHECK_SKIP, &CETMacroDlg::OnBnClickedCheckSkip)
    ON_BN_CLICKED(IDC_RADIO_SPEED0, &CETMacroDlg::OnBnClickedRadioSpeed0)
    ON_BN_CLICKED(IDC_RADIO_SPEED1, &CETMacroDlg::OnBnClickedRadioSpeed1)
    ON_BN_CLICKED(IDC_RADIO_SPEED2, &CETMacroDlg::OnBnClickedRadioSpeed2)
END_MESSAGE_MAP()


BOOL CETMacroDlg::OnInitDialog()
{
    CDialogEx::OnInitDialog();

    // 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
    //  프레임워크가 이 작업을 자동으로 수행합니다.
    SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
    SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.


    _wsetlocale(LC_ALL, L"korean");      //지역화 설정을 전역적으로 적용
    wcout.imbue(locale("korean"));        //출력시 부분적 적용
    wcin.imbue(locale("korean"));          //입력시 부분적 적용

    int nWidth = GetSystemMetrics(SM_CXSCREEN);
    int nHeight = GetSystemMetrics(SM_CYSCREEN);

    m_rectScreen.SetRect(0, 0, nWidth, nHeight);
    GetClientRect(m_rectClient);

    InitChromeURL();

    InitAccount();

    InitSellItems();

    InitControls();

    InitWebButtonPos();

    return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 응용 프로그램의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CETMacroDlg::OnPaint()
{
    CPaintDC dc(this);

    POSITION pos = m_list_sellitem.GetFirstSelectedItemPosition();

    while (pos != NULL)
    {
        int nItem = m_list_sellitem.GetNextSelectedItem(pos);

        CString strImages = theApp.m_sellItems[nItem].images;



        CString strImagePath[6];
        for (int img = 0; img < 6; img++)
        {
            int nStart = 0;
            int nCur = strImages.Find(L",", nStart);

            CString strImageFinal;
            if (img != 5)
            {
                CString strCurImage = strImages.Mid(nStart, nCur);

                strImages = strImages.TrimLeft(strCurImage);
                strImages = strImages.TrimLeft(L", ");

                strImageFinal = strCurImage;
            }
            else
            {
                strImageFinal = strImages;
            }


            strImagePath[img] = GetApplicationDirectory() + L"\\Images\\" + strImageFinal;
        }

        CImage Image0;
        CImage Image1;
        CImage Image2;
        CImage Image3;
        CImage Image4;
        CImage Image5;
        HRESULT hResult0 = Image0.Load(strImagePath[0]);
        HRESULT hResult1 = Image1.Load(strImagePath[1]);
        HRESULT hResult2 = Image2.Load(strImagePath[2]);
        HRESULT hResult3 = Image3.Load(strImagePath[3]);
        HRESULT hResult4 = Image4.Load(strImagePath[4]);
        HRESULT hResult5 = Image5.Load(strImagePath[5]);

        if (SUCCEEDED(hResult0))
        {
            CRect rect;
            m_static_image[0].GetWindowRect(rect);

            rect.OffsetRect(-470, -143);
            Image0.StretchBlt(dc.m_hDC, rect.left, rect.top, rect.Width(), rect.Height());
        }

        if (SUCCEEDED(hResult1))
        {
            CRect rect;
            m_static_image[1].GetWindowRect(rect);

            rect.OffsetRect(-470, -143);
            Image1.StretchBlt(dc.m_hDC, rect.left, rect.top, rect.Width(), rect.Height());
        }

        if (SUCCEEDED(hResult2))
        {
            CRect rect;
            m_static_image[2].GetWindowRect(rect);

            rect.OffsetRect(-470, -143);
            Image2.StretchBlt(dc.m_hDC, rect.left, rect.top, rect.Width(), rect.Height());
        }

        if (SUCCEEDED(hResult3))
        {
            CRect rect;
            m_static_image[3].GetWindowRect(rect);

            rect.OffsetRect(-470, -143);
            Image3.StretchBlt(dc.m_hDC, rect.left, rect.top, rect.Width(), rect.Height());
        }

        if (SUCCEEDED(hResult4))
        {
            CRect rect;
            m_static_image[4].GetWindowRect(rect);

            rect.OffsetRect(-470, -143);
            Image4.StretchBlt(dc.m_hDC, rect.left, rect.top, rect.Width(), rect.Height());
        }

        if (SUCCEEDED(hResult5))
        {
            CRect rect;
            m_static_image[5].GetWindowRect(rect);

            rect.OffsetRect(-470, -143);
            Image5.StretchBlt(dc.m_hDC, rect.left, rect.top, rect.Width(), rect.Height());
        }
    }


    CDialogEx::OnPaint();
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CETMacroDlg::OnQueryDragIcon()
{
    return static_cast<HCURSOR>(m_hIcon);
}

bool CETMacroDlg::WriteExcelFile(const std::wstring wstr)
{
    BasicExcel xls;
    xls.New(1); // Sheet 1개일때는 New(1), 2개일때는 New(2), ...
    BasicExcelWorksheet* sheet = xls.GetWorksheet(0); // 1번째 Sheet 얻어올때는 GetWorksheet(0), 2번째 Sheet는 GetWorksheet(1)...
    XLSFormatManager fmt_mgr(xls);

    ExcelFont font_bold;
    font_bold._weight = FW_BOLD;

    CellFormat fmt_bold(fmt_mgr);
    fmt_bold.set_font(font_bold);

    int col, row = 0;
    for (col = 0; col < 3; col++)
    {
        BasicExcelCell* cell = sheet->Cell(row, col);
        cell->Set("TITLE");
        cell->SetFormat(fmt_bold);
    }

    while (++row < 3)
    {
        for (int col = 0; col < 3; col++)
            sheet->Cell(row, col)->Set("Text");
    }
    return xls.SaveAs(wstr.c_str());
}

bool CETMacroDlg::ReadExcelAccount(std::wstring wstr)
{
    std::string str = "";
    str.assign(wstr.begin(), wstr.end());
    BasicExcel xls(str.c_str());
    BasicExcelWorksheet* sheet = xls.GetWorksheet(0);
    XLSFormatManager fmt_mgr(xls);

    // 사이트 이름
    int nRow = 0;
    for (int nCol = 1; nCol < 6; nCol++)
    {
        account acc;
        acc.webName = sheet->Cell(nRow, nCol)->GetWString();

        theApp.m_account.push_back(acc);
    }


    // ID 입력
    nRow = 1;
    int idx = 0;
    for (int nCol = 1; nCol < 6; nCol++)
    {
        account acc;
        std::string str = sheet->Cell(nRow, nCol)->GetString();

        acc.id = str.c_str();

        theApp.m_account[idx].id = acc.id;
        idx++;
    }

    // Password 입력
    nRow = 2;
    idx = 0;
    for (int nCol = 1; nCol < 6; nCol++)
    {
        account acc;
        std::string str = sheet->Cell(nRow, nCol)->GetString();

        acc.password = str.c_str();

        theApp.m_account[idx].password = acc.password;
        idx++;
    }

    // Password 입력
    nRow = 3;
    idx = 0;
    for (int nCol = 1; nCol < 6; nCol++)
    {
        account acc;
        std::string str = sheet->Cell(nRow, nCol)->GetString();

        acc.webURL = str.c_str();

        theApp.m_account[idx].webURL = acc.webURL;
        idx++;
    }


    return true;
}


bool CETMacroDlg::ReadExcelSellItemsFile()
{
    // 엑셀클래스 선언 (FALSE: 처리 과정을 화면에 보이지 않는다)
    CXLEzAutomation dataexcel(FALSE);

    wchar_t chThisPath[256];
    GetCurrentDirectoryW(256, chThisPath);

    CString strThisPath;
    strThisPath.Format(L"%s\\Sell_Items_new.xlsx", chThisPath);
    ///// 위에서 얻은 엑셀파일경로를 바탕으로 파일을 연다.
    BOOL bRet = dataexcel.OpenExcelFile(strThisPath);

    int nTotalCol = 52;

    int nRow = 2;
    while (nRow)
    {
        sellItem item;
        for (int nCol = 1; nCol < nTotalCol; nCol++)
        {
            switch (nCol)
            {
            case 4:// 글제목 (넘버)
            {
                CString strTitleNum = dataexcel.GetCellValue(nCol, nRow);
                item.title = strTitleNum;
                break;
            }
            case 5:// 글제목 (타이틀)
            {
                CString strTitle = dataexcel.GetCellValue(nCol, nRow);
                item.title += L" " + strTitle;

                break;
            }


            case 7:// 가격
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                item.cost = _wtoi(strText);
                break;
            }


            case 13:// 1번째 이미지 (경로 따로 포함, 이후 이미지는 포함안해도됨)
            {
                CString strImageFullPath = dataexcel.GetCellValue(nCol, nRow);

                CString strImageFullPathTemp = strImageFullPath;
                while (1)
                {
                    int nPos = strImageFullPathTemp.Find(L"\\");

                    if (0 > nPos)
                        break;

                    strImageFullPathTemp = strImageFullPathTemp.Mid(nPos + 1, strImageFullPathTemp.GetLength());
                }

                item.images += strImageFullPathTemp + L",";
                break;
            }



            case 14:// 2번째 이미지 (경로 따로 포함, 이후 이미지는 포함안해도됨)
            {
                CString strImageFullPath = dataexcel.GetCellValue(nCol, nRow);

                CString strImageFullPathTemp = strImageFullPath;
                while (1)
                {
                    int nPos = strImageFullPathTemp.Find(L"\\");

                    if (0 > nPos)
                        break;

                    strImageFullPathTemp = strImageFullPathTemp.Mid(nPos + 1, strImageFullPathTemp.GetLength());
                }

                item.images += strImageFullPathTemp + L",";
                break;
            }

            case 15:// 3번째 이미지 (경로 따로 포함, 이후 이미지는 포함안해도됨)
            {
                CString strImageFullPath = dataexcel.GetCellValue(nCol, nRow);

                CString strImageFullPathTemp = strImageFullPath;
                while (1)
                {
                    int nPos = strImageFullPathTemp.Find(L"\\");

                    if (0 > nPos)
                        break;

                    strImageFullPathTemp = strImageFullPathTemp.Mid(nPos + 1, strImageFullPathTemp.GetLength());
                }

                item.images += strImageFullPathTemp + L",";
                break;
            }

            case 16:// 4번째 이미지 (경로 따로 포함, 이후 이미지는 포함안해도됨)
            {
                CString strImageFullPath = dataexcel.GetCellValue(nCol, nRow);

                CString strImageFullPathTemp = strImageFullPath;
                while (1)
                {
                    int nPos = strImageFullPathTemp.Find(L"\\");

                    if (0 > nPos)
                        break;

                    strImageFullPathTemp = strImageFullPathTemp.Mid(nPos + 1, strImageFullPathTemp.GetLength());
                }

                item.images += strImageFullPathTemp + L",";
                break;
            }

            case 17:// 5번째 이미지 (경로 따로 포함, 이후 이미지는 포함안해도됨)
            {
                CString strImageFullPath = dataexcel.GetCellValue(nCol, nRow);

                CString strImageFullPathTemp = strImageFullPath;
                while (1)
                {
                    int nPos = strImageFullPathTemp.Find(L"\\");

                    if (0 > nPos)
                        break;

                    strImageFullPathTemp = strImageFullPathTemp.Mid(nPos + 1, strImageFullPathTemp.GetLength());
                }

                item.images += strImageFullPathTemp + L",";
                break;
            }

            case 18:// 6번째 이미지 (경로 따로 포함, 이후 이미지는 포함안해도됨)
            {
                CString strImageFullPath = dataexcel.GetCellValue(nCol, nRow);

                CString strImageFullPathTemp = strImageFullPath;
                while (1)
                {
                    int nPos = strImageFullPathTemp.Find(L"\\");

                    if (0 > nPos)
                        break;

                    strImageFullPathTemp = strImageFullPathTemp.Mid(nPos + 1, strImageFullPathTemp.GetLength());
                }

                item.images += strImageFullPathTemp;
                break;
            }



            case 23:// 글내용 : 치수(허리)
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                int nInt = _wtoi(str);
                str.Format(L"%d", nInt);
                CString strText = L"어깨(허리) : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 24:// 글내용 : 가슴(허벅지)
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                int nInt = _wtoi(str);
                str.Format(L"%d", nInt);
                CString strText = L"가슴(허벅지) : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 25:// 글내용 : 팔(밑단)
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                int nInt = _wtoi(str);
                str.Format(L"%d", nInt);
                CString strText = L"팔(밑단) : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 26:// 글내용 : 총기장
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                int nInt = _wtoi(str);
                str.Format(L"%d", nInt);
                CString strText = L"총기장 : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 27:// 글내용 : 브랜드
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                CString strText = L"브랜드 : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 28:// 글내용 : 종류
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                CString strText = L"종류 : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 29:// 글내용 : 색상
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                CString strText = L"색상 : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 30:// 글내용 : 소재
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                CString strText = L"소재 : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 31:// 글내용 : 신축성
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                CString strText = L"신축성 : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 32:// 글내용 : 특이사항
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                CString strText = L"특이사항 : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 33:// 글내용 : 택배비
            {
                CString str = dataexcel.GetCellValue(nCol, nRow);
                CString strText = L"택배비 : " + str;
                item.vecDescription.push_back(strText);

                break;
            }

            case 34:// 글내용 : 태그
            {
                // 아이템에도 태그 정보가 들어간다.
                CString strTagAll = dataexcel.GetCellValue(nCol, nRow);

                while (1)
                {
                    int nStart = 0;
                    int nCur = strTagAll.Find(L",", nStart);

                    if (-1 != nCur)
                    {
                        CString strCurTag = strTagAll.Mid(nStart, nCur);

                        strTagAll = strTagAll.TrimLeft(strCurTag);
                        strTagAll = strTagAll.TrimLeft(L",");

                        item.vecTag.push_back(strCurTag);
                    }
                    else
                        break;
                }

                // 마지막남은 태그 아이템 입력
                item.vecTag.push_back(strTagAll);

                CString strText = L"태그 : ";

                for (int tg = 0; tg < int(item.vecTag.size()); tg++)
                {
                    CString str0 = L"#" + item.vecTag[tg] + L" ";

                    strText += str0;
                }

                item.vecDescription.push_back(strText);


                break;
            }

            case 35:// 밴드주소
            {
                item.bandURL = dataexcel.GetCellValue(nCol, nRow);

                break;
            }

            case 36:// 수량
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                item.unit = _wtoi(strText);

                break;
            }


            

            case 37:// 번개장터 1번째 콤보박스
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                item.nThunderMarketFirstComboIndex = _wtoi(strText);
                break;
            }
            case 38:// 번개장터 2번째 콤보박스
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                item.nThunderMarketSecondComboIndex = _wtoi(strText);
                break;
            }
            case 39:// 번개장터 3번째 콤보박스
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                item.nThunderMarketThirdComboIndex = _wtoi(strText);
                break;
            }
            case 40:// 헬로마켓 1번째 콤보박스
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                item.nHelloMarketFirstComboIndex = _wtoi(strText);
                break;
            }
            case 41:// 헬로마켓 2번째 콤보박스
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                item.nHelloMarketSecondComboIndex = _wtoi(strText);
                break;
            }
            case 42:// 중고나라 1번째 콤보박스
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                item.nNaverCafeFirstComboIndex = _wtoi(strText);
                break;
            }

            case 43:// 상세정보
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                item.nSizeComboIndex = _wtoi(strText);
                break;
            }

            case 44:// 상태
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                int nStatus = _wtoi(strText);
                item.itemStatus = (eItemStatus)nStatus;
                break;
            }
            case 45:// 택배비 포함
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                int nRet = _wtoi(strText);
                item.bDeliveryCost = (BOOL)nRet;
                break;
            }
            case 46:// 교환가능
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                int nRet = _wtoi(strText);

                item.bExchange = (BOOL)nRet;
                break;
            }

            case 47:// 헬로마켓 등록 유무
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                int nRet = _wtoi(strText);

                item.bHelloMarketAccept = (BOOL)nRet;
                break;
            }

            case 48:// 번개장터 등록 유무
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                int nRet = _wtoi(strText);

                item.bThunderMarketAccept = (BOOL)nRet;
                break;
            }

            case 49:// 네이버카페 등록 유무
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                int nRet = _wtoi(strText);

                item.bNaverCafeAccept = (BOOL)nRet;
                break;
            }

            case 50:// 네이버밴드 등록 유무
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                int nRet = _wtoi(strText);

                item.bNaverBandAccept = (BOOL)nRet;
                break;
            }

            case 51:// 카카오스토리 등록 유무
            {
                CString strText = dataexcel.GetCellValue(nCol, nRow);
                int nRet = _wtoi(strText);

                item.bKakaoStoryAccept = (BOOL)nRet;
                break;
            }

            default:
                break;
            }
        }

        if (0 == item.cost)
            break;

        theApp.m_sellItems.push_back(item);

        nRow++;
    }

    dataexcel.ReleaseExcel(); ///// 열었던 파일을 다 사용한 후 닫는다.

    return true;
}


void CETMacroDlg::OnBnClickedButtonMacroStart()
{
    // 시작버튼 중복 클릭 방지용 (이미 Thread 동작시에는 Skip)
    if (m_nCurState == EThreadState_NotStarted)
    {
        m_pMainThread = ::AfxBeginThread(procWorkerThread, reinterpret_cast<LPVOID>(this), 0, CREATE_SUSPENDED, 0, NULL);
        ASSERT(m_pMainThread);

        if (m_pMainThread)
        {
            m_nCurState = EThreadState_Running;

            // set continuation flag
            m_bContinue = TRUE;

            // now resume thread
            m_pMainThread->ResumeThread();
        }
    }
}

void CETMacroDlg::InitSellItems()
{
    ReadExcelSellItemsFile();

}

void CETMacroDlg::InitAccount()
{
    std::wstring excelName = L".\\ID_PW.xls";
    ReadExcelAccount(excelName);

}

void CETMacroDlg::AddSellItemList()
{
    m_list_sellitem.InsertColumn(0, _T("순서"), LVCFMT_LEFT, 40);
    m_list_sellitem.InsertColumn(1, _T("글제목"), LVCFMT_LEFT, 120);
    m_list_sellitem.InsertColumn(2, _T("가격"), LVCFMT_LEFT, 80);
    m_list_sellitem.InsertColumn(3, _T("헬로마켓"), LVCFMT_LEFT, 70);
    m_list_sellitem.InsertColumn(4, _T("번개장터"), LVCFMT_LEFT, 70);
    m_list_sellitem.InsertColumn(5, _T("카페"), LVCFMT_LEFT, 70);
    m_list_sellitem.InsertColumn(6, _T("밴드"), LVCFMT_LEFT, 70);
    m_list_sellitem.InsertColumn(7, _T("카스"), LVCFMT_LEFT, 70);

    m_list_sellitem.SetExtendedStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);


    // item 엑셀파일에서 가져온 sell item 정보 리스트
    int nCount = int(theApp.m_sellItems.size());

    for (int n = 0; n < nCount; n++)
    {
        sellItem item = theApp.m_sellItems[n];




        CString strCount;
        strCount.Format(L"%d", n + 1);
        m_list_sellitem.InsertItem(n, strCount);
        m_list_sellitem.SetItem(n, 1, LVIF_TEXT, item.title, 0, NULL, NULL, NULL);

        CString strCost;
        strCost.Format(L"%d 원", item.cost);
        m_list_sellitem.SetItem(n, 2, LVIF_TEXT, strCost, 0, NULL, NULL, NULL);

        // ini 파일로 등록이 된 여부 판단 (구현예정)
        {// 헬로마켓
            CString strStatus;

            // 일단 모든 항목 기본 FALSE 로 처리
            BOOL bRet = TRUE;
            if (bRet)
            {
                strStatus = L"등록완료!";
            }
            else
            {
                strStatus = L"미등록";
            }

            strStatus = L"-";

            m_list_sellitem.SetItem(n, 3, LVIF_TEXT, strStatus, 0, NULL, NULL, NULL);
        }

        {// 번개장터
            CString strStatus;

            // 일단 모든 항목 기본 FALSE 로 처리
            BOOL bRet = FALSE;
            if (bRet)
            {
                strStatus = L"등록완료!";
            }
            else
            {
                strStatus = L"미등록";
            }

            strStatus = L"-";

            m_list_sellitem.SetItem(n, 4, LVIF_TEXT, strStatus, 0, NULL, NULL, NULL);
        }

        {// 네이버카페
            CString strStatus;

            // 일단 모든 항목 기본 FALSE 로 처리
            BOOL bRet = FALSE;
            if (bRet)
            {
                strStatus = L"등록완료!";
            }
            else
            {
                strStatus = L"미등록";
            }

            strStatus = L"-";

            m_list_sellitem.SetItem(n, 5, LVIF_TEXT, strStatus, 0, NULL, NULL, NULL);
        }

        {// 네이버밴드
            CString strStatus;

            // 일단 모든 항목 기본 FALSE 로 처리
            BOOL bRet = FALSE;
            if (bRet)
            {
                strStatus = L"등록완료!";
            }
            else
            {
                strStatus = L"미등록";
            }

            strStatus = L"-";

            m_list_sellitem.SetItem(n, 6, LVIF_TEXT, strStatus, 0, NULL, NULL, NULL);
        }

        {// 카카오스토리
            CString strStatus;

            // 일단 모든 항목 기본 FALSE 로 처리
            BOOL bRet = FALSE;
            if (bRet)
            {
                strStatus = L"등록완료!";
            }
            else
            {
                strStatus = L"미등록";
            }

            strStatus = L"-";

            m_list_sellitem.SetItem(n, 7, LVIF_TEXT, strStatus, 0, NULL, NULL, NULL);
        }
    }
}


void CETMacroDlg::AddItemNaverCafeComboboxs()
{
    m_combo_nc.ResetContent();

    m_combo_nc.AddString(L"--카테고리 선택--");
    m_combo_nc.AddString(L"남성상의");
    m_combo_nc.AddString(L"남성하의");
    m_combo_nc.AddString(L"여성상의");
    m_combo_nc.AddString(L"여성하의");
    m_combo_nc.SetCurSel(0);
}

void CETMacroDlg::AddItemThunderMarketComboboxs()
{
    m_combo_tm1.ResetContent();

    m_combo_tm1.AddString(L"--카테고리 선택--");
    m_combo_tm1.AddString(L"여성의류");
    m_combo_tm1.AddString(L"남성의류");
    m_combo_tm1.AddString(L"패션잡화");
    m_combo_tm1.AddString(L"유아동/출산");
    m_combo_tm1.SetCurSel(0);



    m_combo_tm2.AddString(L"1번 카테고리를 먼저 선택");
    m_combo_tm2.SetCurSel(0);

    m_combo_tm2.EnableWindow(FALSE);

    m_combo_tm3.AddString(L"2번 카테고리를 먼저 선택");
    m_combo_tm3.SetCurSel(0);

    m_combo_tm3.EnableWindow(FALSE);

    m_combo_tm_size.AddString(L"2번 카테고리를 먼저 선택");
    m_combo_tm_size.SetCurSel(0);

    m_combo_tm_size.EnableWindow(FALSE);

    m_combo_tm_quality.AddString(L"중고");
    m_combo_tm_quality.AddString(L"중고+하자(하자가 있는 중고)");
    m_combo_tm_quality.AddString(L"새물품 (미사용");
    m_combo_tm_quality.AddString(L"새것+하자 (새것이고 하자가 있음)");
    m_combo_tm_quality.AddString(L"거의새것 (새것이고 하자가 없음)");
    m_combo_tm_quality.SetCurSel(0);
    
}

void CETMacroDlg::AddItemHelloMarketComboboxs()
{
    m_combo_hm1.ResetContent();

    m_combo_hm1.AddString(L"--카테고리 선택--");
    m_combo_hm1.AddString(L"바이크");
    m_combo_hm1.AddString(L"유아동,완구");
    m_combo_hm1.AddString(L"자동차용품");
    m_combo_hm1.AddString(L"뷰티");
    m_combo_hm1.AddString(L"바이크용품");
    m_combo_hm1.AddString(L"여성의류");
    m_combo_hm1.AddString(L"남성의류");
    m_combo_hm1.AddString(L"신발,가방,잡화");
    m_combo_hm1.AddString(L"휴대폰,태블릿");
    m_combo_hm1.AddString(L"컴퓨터 주변기기");
    m_combo_hm1.AddString(L"카메라");
    m_combo_hm1.AddString(L"디지털,가전");
    m_combo_hm1.AddString(L"게임");
    m_combo_hm1.AddString(L"스포츠,레저");
    m_combo_hm1.AddString(L"가구");
    m_combo_hm1.AddString(L"생활");
    m_combo_hm1.AddString(L"골동품,희귀품");
    m_combo_hm1.AddString(L"여행,숙박");
    m_combo_hm1.AddString(L"티켓");
    m_combo_hm1.AddString(L"재능,서비스");
    m_combo_hm1.AddString(L"도서");
    m_combo_hm1.AddString(L"스타굿즈");
    m_combo_hm1.AddString(L"문구");
    m_combo_hm1.AddString(L"피규어,키덜트");
    m_combo_hm1.AddString(L"CD,DVD");
    m_combo_hm1.AddString(L"음향기기,악기");
    m_combo_hm1.AddString(L"예술,미술");
    m_combo_hm1.AddString(L"반려동물");
    m_combo_hm1.AddString(L"부동산");
    m_combo_hm1.AddString(L"포장식품");
    m_combo_hm1.AddString(L"핸드메이드");
    m_combo_hm1.AddString(L"기타");
    m_combo_hm1.SetCurSel(0);

    m_combo_hm2.AddString(L"1번 카테고리를 먼저 선택");
    m_combo_hm2.SetCurSel(0);

    m_combo_hm2.EnableWindow(FALSE);
}

void CETMacroDlg::InitControls()
{
    m_check_site0.SetCheck(FALSE);
    m_check_site1.SetCheck(FALSE);
    m_check_site2.SetCheck(FALSE);
    m_check_site3.SetCheck(FALSE);
    m_check_site4.SetCheck(FALSE);

    m_edit_navercafe_title.SetWindowText(L"남성상의 브랜드 제품 모음전");
    m_edit_navercafe_cost.SetWindowText(L"10000");

    AddSellItemList();

    AddItemHelloMarketComboboxs();
    AddItemThunderMarketComboboxs();

    AddItemNaverCafeComboboxs();

    m_radio_speed0.SetCheck(TRUE);
}

void CETMacroDlg::InitChromeURL()
{
    //    theApp.m_URLChrome = L"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
    theApp.m_URLChrome = L"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
}


void CETMacroDlg::InitWebButtonPos()
{
    // 헬로마켓 웹 버튼들 중심 위치 초기화

    float fFullHeightGap = 0.063f;

    // 헬로마켓 첫번째 로그인 버튼
    m_btnWebFrameBar.nx = m_rectScreen.Width() * 0.5;
    m_btnWebFrameBar.ny = 10;

    m_btnHelloMarketFirstLogin.nx = m_rectScreen.Width() * 0.688;
    m_btnHelloMarketFirstLogin.ny = m_rectScreen.Height() * (0.122 - fFullHeightGap);

    // 헬로마켓 로그인 다이얼로그의 로그인 버튼
    m_btnHelloMarketLogin.nx = m_rectScreen.CenterPoint().x;
    m_btnHelloMarketLogin.ny = m_rectScreen.Height() * 0.5;

    // 헬로마켓 등록하기 버튼
    m_btnHelloMarketSellItem.nx = m_rectScreen.Width() * 0.71;
    m_btnHelloMarketSellItem.ny = m_rectScreen.Height() * (0.17 - fFullHeightGap);

    // 헬로마켓 일반아이템 버튼
    m_btnHelloMarketGeneralItem.nx = m_rectScreen.Width() * 0.71;
    m_btnHelloMarketGeneralItem.ny = m_rectScreen.Height() * (0.216 - fFullHeightGap);


    // 헬로마켓 대표이미지 버튼										
    m_btnHelloMarketImages.nx = m_rectScreen.Width() * 0.4;
    m_btnHelloMarketImages.ny = m_rectScreen.Height() * (0.43 - fFullHeightGap);

    // 헬로마켓 빈화면 버튼
    m_btnHelloMarketBlank.nx = m_rectScreen.Width() * 0.25;
    m_btnHelloMarketBlank.ny = m_rectScreen.Height() * 0.5;



    // 번개마트 강제 로그아웃 버튼
    m_btnThunderMarketLogout.nx = m_rectScreen.Width() * 0.677;
    m_btnThunderMarketLogout.ny = m_rectScreen.Height() * 0.0722;

    // 번개마트 로그인 다이얼로그 상단 버튼
    m_btnThunderMarketFirstLogin.nx = m_rectScreen.Width() * 0.5;
    m_btnThunderMarketFirstLogin.ny = m_rectScreen.Height() * 0.3;

    // 이미지 업로드 버튼
    m_btnThunderMarketImageUpload.nx = m_rectScreen.Width() * 0.1588;
    m_btnThunderMarketImageUpload.ny = m_rectScreen.Height() * 0.197;

    // 기본정보 탭 위치
    m_btnThunderMarketBaseInfo.nx = m_rectScreen.Width() * 0.503;
    m_btnThunderMarketBaseInfo.ny = m_rectScreen.Height() * 0.1277;

    // 상세정보 탭 위치
    m_btnThunderMarketDetailInfo.nx = m_rectScreen.Width() * 0.5614;
    m_btnThunderMarketDetailInfo.ny = m_rectScreen.Height() * 0.1277;

    // 최근 도시 선택 버튼
    m_btnThunderMarketRecentlyCity.nx = m_rectScreen.Width() * 0.492;
    m_btnThunderMarketRecentlyCity.ny = m_rectScreen.Height() * 0.404;


    // 등록 버튼
    m_btnThunderMarketAccept.nx = m_rectScreen.Width() * 0.7907;
    m_btnThunderMarketAccept.ny = m_rectScreen.Height() * 0.869;



    // 네이버 로그아웃 버튼
    m_btnNaverLogout.nx = m_rectScreen.Width() * 0.752;
    m_btnNaverLogout.ny = m_rectScreen.Height() * 0.2574;


    // 네이버 로그인 버튼
    m_btnNaverLogin.nx = m_rectScreen.Width() * 0.6963;
    m_btnNaverLogin.ny = m_rectScreen.Height() * 0.262;

    // 네이버 빈화면 버튼
    m_btnNaverBlank.nx = m_rectScreen.Width() * 0.1;
    m_btnNaverBlank.ny = m_rectScreen.Height() * 0.5;

    // 네이버 카페 글등록 버튼
    m_btnNaverSellItem.nx = m_rectScreen.Width() * 0.2697;
    m_btnNaverSellItem.ny = m_rectScreen.Height() * 0.5247;

    // 네이버 카페 SmartEditor Line 버튼
    m_btnNaverSmartEditorLine.nx = m_rectScreen.Width() * 0.5864;
    m_btnNaverSmartEditorLine.ny = m_rectScreen.Height() * 0.2407;

    // 네이버 카페 Sell Tab 버튼
    m_btnNaverSellTabLine.nx = m_rectScreen.Width() * 0.527;
    m_btnNaverSellTabLine.ny = m_rectScreen.Height() * 0.35;

    // 네이버 카페 포토 업로더 중심위치 버튼
    m_btnNaverPhotoUploaderCenter.nx = m_rectScreen.Width() * 0.5;
    m_btnNaverPhotoUploaderCenter.ny = m_rectScreen.Height() * 0.5;




    // 네이버 밴드 새로운 소식 클릭 위치
    m_btnNaverBandNewPostPos.nx = m_rectScreen.Width() * 0.5;
    m_btnNaverBandNewPostPos.ny = m_rectScreen.Height() * 0.175;








    // 카카오스토리 로봇 해제 버튼
    m_btnKakaoRobot.nx = m_rectScreen.Width() * 0.5;
    m_btnKakaoRobot.ny = m_rectScreen.Height() * 0.546;


    // 카카오스토리 로그인 버튼
    m_btnKakaoLogin.nx = m_rectScreen.Width() * 0.45;
    m_btnKakaoLogin.ny = m_rectScreen.Height() * 0.536;

    m_btnKakaoLogin2.nx = m_rectScreen.Width() * 0.45;
    m_btnKakaoLogin2.ny = m_rectScreen.Height() * 0.619;



    // 카카오스토리 아이템 등록 버튼
    m_btnKakaoSellItem.nx = m_rectScreen.Width() * 0.345;
    m_btnKakaoSellItem.ny = m_rectScreen.Height() * 0.1898;

    m_btnKakaoSellAccept.nx = m_rectScreen.Width() * 0.536;
    m_btnKakaoSellAccept.ny = m_rectScreen.Height() * 0.7444;

}

void CETMacroDlg::StandBy(int time)
{
    //     _nCoolTime -= time;
    m_fScale = 0.75f;

    m_fScale *= theApp.m_fMacroSpeed;

    int nRet = m_fScale * time;

    Sleep(nRet);
}


void CETMacroDlg::KeyboardButtonClickEvent(BYTE byte, int sleep)
{
    keybd_event(byte, 0, KEYEVENTF_EXTENDEDKEY, 0);
    Sleep(1);
    keybd_event(byte, 0, KEYEVENTF_KEYUP, 0);
    Sleep(1);

    StandBy(sleep);
}


void CETMacroDlg::MouseMoveEvent(ButtonPos pos)
{
    SetCursorPos(pos.nx, pos.ny);
    StandBy(1000);
}

void CETMacroDlg::MouseLbuttonClickEvent(ButtonPos pos)
{
    SetCursorPos(pos.nx, pos.ny);
    StandBy(1000);

    mouse_event(MOUSEEVENTF_LEFTDOWN, pos.nx, pos.ny, 0, 0);
    Sleep(1);
    mouse_event(MOUSEEVENTF_LEFTUP, pos.nx, pos.ny, 0, 0);
    Sleep(1);

    StandBy(2000);
}


void CETMacroDlg::EventInputURLEditMode()
{
    keybd_event(VK_CONTROL, 0, 0, 0);
    keybd_event(76, 0, 0, 0);
    keybd_event(76, 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);

    StandBy(1000);
}

void CETMacroDlg::EventClipboardPaste()
{
    keybd_event(VK_CONTROL, 0, 0, 0);
    keybd_event(86, 0, 0, 0);
    keybd_event(86, 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);

    StandBy(500);
}

// 복사 함수

int CETMacroDlg::CopyTextToClipboard(CString strCopy)
{
    int nLength = strCopy.GetLength() + 1;

    //	char* ap_string = new char[nLength];
    //	char* ap_string = (char*)malloc(sizeof(char) * 100);
    char ap_string[100] = { 0, };
    strcpy(ap_string, CT2A(strCopy));

    // 저장할 문자열의 길이를 구한다. ('\0'까지 포함한 크기)

    //int string_length = strlen(ap_string) + 1;
    int string_length = strlen(ap_string);

    // 클립보드로 문자열을 복사하기 위하여 메모리를 할당한다. 
    // 클립보드에는 핸들을 넣는 형식이라서 HeapAlloc 함수 사용이 블가능하다. 
    //HANDLE h_data = ::GlobalAlloc(GMEM_DDESHARE | GMEM_MOVEABLE, string_length);
    HANDLE h_data = ::GlobalAlloc(GMEM_ZEROINIT | GMEM_FIXED, 100);

    // 할당된 메모리에 문자열을 복사하기 위해서 사용 가능한 주소를 얻는다. 
    char *p_data = (char *)::GlobalLock(h_data);

    if (NULL != p_data)
    {
        // 할당된 메모리 영역에 삽입할 문자열을 복사한다. 
        memcpy(p_data, ap_string, 100);

        // 문자열을 복사하기 위해서 Lock 했던 메모리를 해제한다.
        ::GlobalUnlock(h_data);

        if (::OpenClipboard(m_hWnd))
        {
            ::EmptyClipboard(); // 클립보드를 연다.

            ::SetClipboardData(CF_TEXT, h_data);  // 클립보드에 저장된 기존 문자열을 삭제한다.

                                                  // 클립보드로 문자열을 복사한다

            ::CloseClipboard(); // 클립보드를 닫는다.
        }
    }

    StandBy(500);

    // 이미 복사해둔 note 입력창에 붙여넣기
    EventClipboardPaste();

    //	free(ap_string);

    return 0;
}

//int CETMacroDlg::CopyTextToClipboard(CString strCopy)
//{
//	if (!OpenClipboard())
//	{
//		AfxMessageBox(_T("Cannot open the Clipboard"));
//		return 0;
//	}
//	// Remove the current Clipboard contents
//	if (!EmptyClipboard())
//	{
//		AfxMessageBox(_T("Cannot empty the Clipboard"));
//		return 0;
//	}
//	// Set Data
//	size_t size = (strCopy.GetLength() * 2) + 2;
//	HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE | GMEM_DDESHARE, size);
//	if (hMem)
//	{
//		LPSTR pClipData = (LPSTR)GlobalLock(hMem);
//		pClipData[0] = 0;
//		strncpy(pClipData, ATL::CW2AEX<1024>(strCopy.GetBuffer(0)), size);
//		if (::SetClipboardData(CF_OEMTEXT, hMem) == NULL)
//		{
//			AfxMessageBox(_T("Unable to set Clipboard data"));
//			GlobalUnlock(hMem);
//			GlobalFree(hMem);
//			CloseClipboard();
//			return 0;
//		}
//		GlobalUnlock(hMem);
//		GlobalFree(hMem);
//	}
//	CloseClipboard();
//
//	return 0;
//}


void CETMacroDlg::ThunderMarketLogin()
{
    // 번개 장터 사이트로 이동 (번개장터는 로그아웃이 잘되지않아 강제 로그아웃 버튼을 눌러준다.)
    // 혹시 로그아웃 상태여도 상관없다.
    ShellExecute(m_webHwnd, NULL, theApp.m_URLChrome, theApp.m_account[1].webURL, NULL, SW_SHOWMAXIMIZED);
    StandBy(12000);

    MouseLbuttonClickEvent(m_btnWebFrameBar);

    keybd_event(VK_F11, 0, 0, 0);
    keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
    StandBy(5000);

    MouseLbuttonClickEvent(m_btnThunderMarketLogout);
    StandBy(2000);
    system("taskkill /IM chrome.exe /F");
    StandBy(4000);


    // 번개 장터 사이트로 이동 (웹창이 열려야할 시간이 있어야 하므로 20초 여유)
    ShellExecute(m_webHwnd, NULL, theApp.m_URLChrome, theApp.m_account[1].webURL, NULL, SW_SHOWMAXIMIZED);
    StandBy(15000);

    MouseLbuttonClickEvent(m_btnThunderMarketFirstLogin);

    // 탭 4번 눌러야 로그인입력 화면의 로그인 EditBox 로 이동
    for (int n = 0; n < 4; n++)
    {
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(200);
    }

    StandBy(1000);

    // 아이디 클립보드 복사 
    CString strID = theApp.m_account[1].id;
    CopyTextToClipboard(strID);

    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(500);

    // 패스워드 클립보드 복사 
    CString strPassword = theApp.m_account[1].password;
    CopyTextToClipboard(strPassword);

    KeyboardButtonClickEvent(VK_RETURN, 1000);
    StandBy(7000);
}

void CETMacroDlg::NaverBandLogin()
{
    // 탁 여기부터 진행하면됨 (네이버 밴드)
    ShellExecute(m_webHwnd, NULL, theApp.m_URLChrome, theApp.m_account[3].webURL, NULL, SW_SHOWMAXIMIZED);
    StandBy(17000);

    KeyboardButtonClickEvent(VK_RETURN, 1000);
    StandBy(5000);


    MouseLbuttonClickEvent(m_btnWebFrameBar);

    keybd_event(VK_F11, 0, 0, 0);
    keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
    StandBy(2000);


    // 탭 4번 눌러야 로그인입력 화면의 로그인 EditBox 로 이동
    for (int n = 0; n < 5; n++)
    {
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(200);
    }
    KeyboardButtonClickEvent(VK_RETURN, 1000);
    StandBy(5000);

    // 탭 2번 눌러 로그인 버튼으로 이동
    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(200);
    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(200);

    KeyboardButtonClickEvent(VK_RETURN, 1000);
    StandBy(5000);

    // 아이디 클립보드 복사 
    CString strID = theApp.m_account[3].id;
    CopyTextToClipboard(strID);

    KeyboardButtonClickEvent(VK_RETURN, 1000);
    StandBy(3000);


    // 패스워드 클립보드 복사 
    CString strPassword = theApp.m_account[3].password;
    CopyTextToClipboard(strPassword);

    KeyboardButtonClickEvent(VK_RETURN, 1000);
    StandBy(7000);
}

void CETMacroDlg::PlayMacroNaverBand()
{
    // 네이버밴드 로그인
    NaverBandLogin();
    printf("Access Naver Band\n");

    // 새로운 소식을 남겨보세요 클릭 이벤트
    MouseLbuttonClickEvent(m_btnNaverBandNewPostPos);
    StandBy(3000);


    // Sell Item 목록 추가하기
    int nCount = theApp.m_sellItems.size();
    for (int n = 0; n < nCount; n++)
    {
        sellItem item = theApp.m_sellItems[n];

        BOOL bAccept = item.bNaverBandAccept;

        if (!bAccept)
            continue;

        // 탭두번 눌러서 사진 버튼으로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(500);

        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(500);

        // 사진 버튼 엔터
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(5000);

        // 파일 경로 입력창 활성화 
        KeyboardButtonClickEvent(VK_F4, 1000);

        // 팩스페이스로 삭제
        for (int n = 0; n < 100; n++)
        {
            KeyboardButtonClickEvent(VK_BACK, 30);
        }

        // 이미지경로로 이동
        CString strImagePath = GetApplicationDirectory() + L"\\Images\\";
        CopyTextToClipboard(strImagePath);
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(1000);


        // 파일 입력창으로 이동
        keybd_event(VK_MENU, 0, 0, 0);
        keybd_event(78, 0, 0, 0);
        keybd_event(78, 0, KEYEVENTF_KEYUP, 0);
        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        // 이미지 6장 경로 클립보드에 복사	
        CString strImages = item.images;

        CString strImagesFinal;
        for (int img = 0; img < 6; img++)
        {
            int nStart = 0;
            int nCur = strImages.Find(L",", nStart);

            if (img != 5)
            {
                CString strCurImage = strImages.Mid(nStart, nCur);

                strImages = strImages.TrimLeft(strCurImage);
                strImages = strImages.TrimLeft(L", ");

                CString strTemp = L"\"" + strCurImage + L"\" ";

                strImagesFinal += strTemp;
            }
            else
            {
                CString strTemp = L"\"" + strImages + L"\" ";

                strImagesFinal += strTemp;
            }
        }

        CopyTextToClipboard(strImagesFinal);

        // 이미지 등록 (엔터)
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(10000);

        // 탭 18번 눌러서 첨부하기 버튼으로 포커스
        for (int i = 0; i < 18; i++)
        {
            keybd_event(VK_TAB, 0, 0, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
            StandBy(300);
        }

        // 첨부하기 버튼 엔터
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(1000);

        // 글쓰기
        CString strTitle;
        strTitle.Format(L"제목 : %s", item.title);

        CopyTextToClipboard(strTitle);
        KeyboardButtonClickEvent(VK_RETURN, 300);


        // 상세 내용 줄마다 나누기 
        int nDescriptionLineCount = int(item.vecDescription.size());

        for (int d = 0; d < nDescriptionLineCount; d++)
        {
            CString strLine = item.vecDescription[d];

            CopyTextToClipboard(strLine);
            KeyboardButtonClickEvent(VK_RETURN, 300);
        }

        CString strCost;
        strCost.Format(L"판매가격 : %d 원", item.cost);

        CopyTextToClipboard(strCost);
        KeyboardButtonClickEvent(VK_RETURN, 300);

        // 사진과 글 사이 줄띄우기
        KeyboardButtonClickEvent(VK_RETURN, 300);
        KeyboardButtonClickEvent(VK_RETURN, 300);
    }

    for (int i = 0; i < 14; ++i)
    {
        // 탭 14번 올리기 버튼 포커스
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(500);
    }

    // 등록완료
    printf("등록완료 !!\n");
    KeyboardButtonClickEvent(VK_RETURN, 2000);

    Sleep(4000);
    system("taskkill /IM chrome.exe /F");

    // 가끔 잔존하는 크롬이 있으니 한번더 강제종료
    Sleep(1000);
    system("taskkill /IM chrome.exe /F");
}

void CETMacroDlg::NaverCafeLogin()
{
    // 일단 네이버로 로그인을해서 강제 로그아웃을 해준다. (이미 로그아웃 되있어도 상관없다.)
    CString strNaverWeb = L"https://www.naver.com";
    ShellExecute(m_webHwnd, NULL, theApp.m_URLChrome, strNaverWeb, NULL, SW_SHOWMAXIMIZED);
    StandBy(12000);

    MouseLbuttonClickEvent(m_btnWebFrameBar);

    keybd_event(VK_F11, 0, 0, 0);
    keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
    StandBy(5000);

    // 로그아웃
    MouseLbuttonClickEvent(m_btnNaverLogout);
    StandBy(2000);

    // 강제종료
    system("taskkill /IM chrome.exe /F");
    StandBy(4000);


    // 재로그인
    ShellExecute(m_webHwnd, NULL, theApp.m_URLChrome, strNaverWeb, NULL, SW_SHOWMAXIMIZED);
    StandBy(15000);

    MouseLbuttonClickEvent(m_btnWebFrameBar);

    keybd_event(VK_F11, 0, 0, 0);
    keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
    StandBy(5000);

    MouseLbuttonClickEvent(m_btnNaverLogin);
    StandBy(5000);

    // 아이디 클립보드 복사 
    CString strID = theApp.m_account[2].id;
    CopyTextToClipboard(strID);

    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(500);

    // 패스워드 클립보드 복사 
    CString strPassword = theApp.m_account[2].password;
    CopyTextToClipboard(strPassword);

    KeyboardButtonClickEvent(VK_RETURN, 1000);
    StandBy(7000);

    // 로그인 완료 후 네이버 빈공간 클릭하여 웹 포커스 상태만듬
    MouseLbuttonClickEvent(m_btnNaverBlank);
    StandBy(2000);

    // 네이버 카페 url 을 입력하기 위해 전체화면 해제
    keybd_event(VK_F11, 0, 0, 0);
    keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
    StandBy(3000);


    // 다시 네이버 빈공간 클릭하여 웹 포커스 상태만듬
    MouseLbuttonClickEvent(m_btnNaverBlank);
    StandBy(2000);


    // 네이버 카페 url 입력 모드로 입력함
    CString strWebSellURL = theApp.m_account[2].webURL;

    EventInputURLEditMode();
    CopyTextToClipboard(strWebSellURL);
    KeyboardButtonClickEvent(VK_RETURN, 1000);

    // 다시 네이버 빈공간 클릭하여 웹 포커스 상태만듬
    MouseLbuttonClickEvent(m_btnNaverBlank);
    StandBy(2000);

    // 전체 화면으로 전환 하여 로그인 완료
    keybd_event(VK_F11, 0, 0, 0);
    keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
    StandBy(3000);
}

void CETMacroDlg::KakaoStoryLogin()
{
    ShellExecute(m_webHwnd, NULL, theApp.m_URLChrome, theApp.m_account[4].webURL, NULL, SW_SHOWMAXIMIZED);
    StandBy(17000);

    MouseLbuttonClickEvent(m_btnWebFrameBar);

    keybd_event(VK_F11, 0, 0, 0);
    keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
    StandBy(5000);

    // 탭한번 눌러서 로그인 위치로 이동
    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(1000);

    // 아이디 클립보드 복사 
    CString strID = theApp.m_account[4].id;
    CopyTextToClipboard(strID);

    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(500);

    // 패스워드 클립보드 복사 
    CString strPassword = theApp.m_account[4].password;
    CopyTextToClipboard(strPassword);

//     KeyboardButtonClickEvent(VK_RETURN, 1000);
//     StandBy(7000);

    HDC hdc = ::GetWindowDC(this->m_webHwnd);
    CDC* pDC = GetDC();
    COLORREF color;

    color = GetPixel(hdc, m_btnKakaoLogin.nx, m_btnKakaoLogin.ny);

    int nColor[3] = { GetRValue(color), GetGValue(color), GetBValue(color) };

    // 컬러값이 노란색이 나오면
    if (nColor[0] > 200 && nColor[1] > 190 && nColor[2] < 90)
    {
        MouseLbuttonClickEvent(m_btnKakaoLogin);
        StandBy(10000);
    }
    else
    {
        // 포토 업로더 포커스를 위해서 중앙 클릭
        MouseLbuttonClickEvent(m_btnKakaoRobot);
        StandBy(1000);


        MouseLbuttonClickEvent(m_btnKakaoLogin2);
        StandBy(10000);
    }


//     // 포토 업로더 포커스를 위해서 중앙 클릭
//     MouseLbuttonClickEvent(m_btnKakaoRobot);
//     StandBy(1000);
// 
// 
//     MouseLbuttonClickEvent(m_btnKakaoLogin2);
//     StandBy(10000);
}

void CETMacroDlg::PlayMacroKakaoStory()
{
    // 카카오스토리 로그인
    KakaoStoryLogin();

    // Sell Item 목록 추가하기
    int nCount = theApp.m_sellItems.size();
    for (int n = 0; n < nCount; n++)
    {
        sellItem item = theApp.m_sellItems[n];

        BOOL bAccept = item.bKakaoStoryAccept;

        if (!bAccept)
            continue;

        printf("카카오스토리 판매 등록 시작 (%d / %d) !!\n\n\n", (n + 1), nCount);


        // 포토 업로더 포커스를 위해서 중앙 클릭
        MouseLbuttonClickEvent(m_btnKakaoSellItem);
        StandBy(4000);


        CopyTextToClipboard(item.title);
        StandBy(1000);

        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(1000);

        // 상세 내용 줄마다 나누기 
        int nDescriptionLineCount = int(item.vecDescription.size());

        for (int d = 0; d < nDescriptionLineCount; d++)
        {
            CString strLine = item.vecDescription[d];

            CopyTextToClipboard(strLine);
            KeyboardButtonClickEvent(VK_RETURN, 300);
        }

        CString strCost;
        strCost.Format(L"판매가격 : %d 원", item.cost);

        CopyTextToClipboard(strCost);
        KeyboardButtonClickEvent(VK_RETURN, 300);





        // 이미지 등록 버튼 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(500);

        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(5000);

        // 파일 경로 입력창 활성화 
        KeyboardButtonClickEvent(VK_F4, 1000);

        // 팩스페이스로 삭제
        for (int n = 0; n < 100; n++)
        {
            KeyboardButtonClickEvent(VK_BACK, 30);
        }


        // 이미지경로로 이동
        CString strImagePath = GetApplicationDirectory() + L"\\Images\\";
        CopyTextToClipboard(strImagePath);
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(1000);


        // 파일 입력창으로 이동
        keybd_event(VK_MENU, 0, 0, 0);
        keybd_event(78, 0, 0, 0);
        keybd_event(78, 0, KEYEVENTF_KEYUP, 0);
        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        // 이미지 6장 경로 클립보드에 복사	
        CString strImages = item.images;

        CString strImagesFinal;
        for (int img = 0; img < 6; img++)
        {
            int nStart = 0;
            int nCur = strImages.Find(L",", nStart);

            if (img != 5)
            {
                CString strCurImage = strImages.Mid(nStart, nCur);

                strImages = strImages.TrimLeft(strCurImage);
                strImages = strImages.TrimLeft(L", ");

                CString strTemp = L"\"" + strCurImage + L"\" ";

                strImagesFinal += strTemp;
            }
            else
            {
                CString strTemp = L"\"" + strImages + L"\" ";

                strImagesFinal += strTemp;
            }
        }

        CopyTextToClipboard(strImagesFinal);

        printf("이미지 삽입 끝 !!\n");
        // 이미지 등록 (엔터)
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(10000);




        for (int t = 0; t < 31; t++)
        {
            // 탭 7번 올리기 버튼 포커스
            keybd_event(VK_TAB, 0, 0, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
            StandBy(500);
        }

        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(10000);
    }

    // 강제종료
    system("taskkill /IM chrome.exe /F");
    StandBy(1000);
    system("taskkill /IM chrome.exe /F");
    StandBy(1000);
}

void CETMacroDlg::PlayMacroNaverCafe()
{
    // 네이버밴드 로그인
    NaverCafeLogin();

    printf("네이버카페 판매 등록 시작  !!\n\n\n");

    // 로그인 완료 후 네이버 빈공간 클릭하여 웹 포커스 상태만듬
    StandBy(2000);
    MouseLbuttonClickEvent(m_btnNaverSellItem);
    StandBy(5000);



    // 네이버 빈공간 클릭하여 웹 포커스 상태만듬
    MouseLbuttonClickEvent(m_btnNaverBlank);
    StandBy(2000);

    // ctrl + home 버튼으로 최상단으로 이동
    keybd_event(VK_CONTROL, 0, 0, 0);
    keybd_event(VK_HOME, 0, 0, 0);
    keybd_event(VK_HOME, 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
    StandBy(2000);


    // SmartEditor 라인에 포커스를 두기 위해서 쿨릭
    MouseLbuttonClickEvent(m_btnNaverSmartEditorLine);
    StandBy(1000);


    // 탭두번 눌러서 게시판 선택으로 이동
    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(1000);

    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(1000);

    // 엔터를 눌러 콤보박스 리스트를 열어서 이동
    // 이렇게해야 중간이동시 팝업이 안뜬다.
    KeyboardButtonClickEvent(VK_RETURN, 1000);
    StandBy(1000);

    int nNcComboCur = theApp.m_sellItems[0].nNaverCafeFirstComboIndex;
    for (int c = 0; c < nNcComboCur; c++)
    {
        keybd_event(VK_DOWN, 0, 0, 0);
        keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0);
        StandBy(100);

    }

    KeyboardButtonClickEvent(VK_RETURN, 1000);
    StandBy(1000);


    KeyboardButtonClickEvent(VK_RETURN, 1000);
    StandBy(3000);



    // 네이버 빈공간 클릭하여 웹 포커스 상태만듬
    MouseLbuttonClickEvent(m_btnNaverBlank);
    StandBy(2000);

    // ctrl + home 버튼으로 최상단으로 이동
    keybd_event(VK_CONTROL, 0, 0, 0);
    keybd_event(VK_HOME, 0, 0, 0);
    keybd_event(VK_HOME, 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
    StandBy(2000);


    // SmartEditor 라인에 포커스를 두기 위해서 쿨릭
    MouseLbuttonClickEvent(m_btnNaverSmartEditorLine);
    StandBy(1000);



    for (int t = 0; t < 4; t++)
    {
        // 탭네번 눌러서 제목입력으로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(500);
    }


    printf("통합 타이틀 입력 !!\n");
//     CopyTextToClipboard(item.title);
    CString strMainTitle;
    m_edit_navercafe_title.GetWindowText(strMainTitle);

    CopyTextToClipboard(strMainTitle);


    // sell tab 라인에 포커스를 두기 위해서 쿨릭
    MouseLbuttonClickEvent(m_btnNaverSellTabLine);
    StandBy(1000);


    for (int t = 0; t < 6; t++)
    {
        // 탭 6번 눌러서 가격입력으로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(500);
    }



    printf("판매가격 입력 !!\n");
    CString strMainCost;
    m_edit_navercafe_cost.GetWindowText(strMainCost);

    // 태그 내용 복사
    CopyTextToClipboard(strMainCost);
    StandBy(1000);


    // 탭 2번 눌러서 네이버페이 체츠
    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(500);
    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(500);

    printf("네이버페이 체크 !!\n");
    KeyboardButtonClickEvent(VK_SPACE, 1000);
    StandBy(1000);

    // 탭 2번 눌러서 정보동의 체츠
    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(500);
    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(500);

    printf("정보동의 체크 !!\n");
    KeyboardButtonClickEvent(VK_SPACE, 1000);
    StandBy(1000);




    // Sell Item 목록 추가하기
    int nCount = theApp.m_sellItems.size();
    for (int n = 0; n < nCount; n++)
    {
        sellItem item = theApp.m_sellItems[n];

        BOOL bAccept = item.bNaverCafeAccept;

        if (!bAccept)
            continue;

        // 네이버 빈공간 클릭하여 웹 포커스 상태만듬
        MouseLbuttonClickEvent(m_btnNaverBlank);
        StandBy(2000);

        // ctrl + home 버튼으로 최상단으로 이동
        keybd_event(VK_CONTROL, 0, 0, 0);
        keybd_event(VK_HOME, 0, 0, 0);
        keybd_event(VK_HOME, 0, KEYEVENTF_KEYUP, 0);
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
        StandBy(2000);



        // sell tab 라인에 포커스를 두기 위해서 쿨릭
        MouseLbuttonClickEvent(m_btnNaverSellTabLine);
        StandBy(1000);

        // sell tab 부터 탭 14번 눌러 사진등록 버튼으로 이동
        for (int t = 0; t < 14; t++)
        {
            // 탭 14번 눌러서 사진 버튼으로 이동
            keybd_event(VK_TAB, 0, 0, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
            StandBy(500);
        }

        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(5000);


        // 포토 업로더 포커스를 위해서 중앙 클릭
        MouseLbuttonClickEvent(m_btnNaverPhotoUploaderCenter);
        StandBy(1000);

        // 탭 1번 눌러서 포토 업로더 버튼으로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(500);

        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(5000);

        // 파일 경로 입력창 활성화 
        KeyboardButtonClickEvent(VK_F4, 1000);

        // 팩스페이스로 삭제
        for (int n = 0; n < 100; n++)
        {
            KeyboardButtonClickEvent(VK_BACK, 30);
        }

        // 이미지경로로 이동
        CString strImagePath = GetApplicationDirectory() + L"\\Images\\";
        CopyTextToClipboard(strImagePath);
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(1000);


        // 파일 입력창으로 이동
        keybd_event(VK_MENU, 0, 0, 0);
        keybd_event(78, 0, 0, 0);
        keybd_event(78, 0, KEYEVENTF_KEYUP, 0);
        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        // 이미지 6장 경로 클립보드에 복사	
        CString strImages = item.images;

        CString strImagesFinal;
        for (int img = 0; img < 6; img++)
        {
            int nStart = 0;
            int nCur = strImages.Find(L",", nStart);

            if (img != 5)
            {
                CString strCurImage = strImages.Mid(nStart, nCur);

                strImages = strImages.TrimLeft(strCurImage);
                strImages = strImages.TrimLeft(L", ");

                CString strTemp = L"\"" + strCurImage + L"\" ";

                strImagesFinal += strTemp;
            }
            else
            {
                CString strTemp = L"\"" + strImages + L"\" ";

                strImagesFinal += strTemp;
            }
        }

        CopyTextToClipboard(strImagesFinal);

        // 이미지 등록 (엔터)
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(12000);


        // 포토 업로더 포커스를 위해서 중앙 클릭
        MouseLbuttonClickEvent(m_btnNaverPhotoUploaderCenter);
        StandBy(1000);



        for (int t = 0; t < 9; t++)
        {
            // 탭 7번 올리기 버튼 포커스
            keybd_event(VK_TAB, 0, 0, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
            StandBy(500);
        }

        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(12000);

        printf("이미지 삽입 끝 !!\n");

        // 사진과 글 사이 줄띄우기
        KeyboardButtonClickEvent(VK_RETURN, 300);

        CString strTitle;
        strTitle.Format(L"제목 : %s", item.title);

        CopyTextToClipboard(strTitle);
        KeyboardButtonClickEvent(VK_RETURN, 300);


        // 상세 내용 줄마다 나누기 
        int nDescriptionLineCount = int(item.vecDescription.size());

        for (int d = 0; d < nDescriptionLineCount; d++)
        {
            CString strLine = item.vecDescription[d];

            CopyTextToClipboard(strLine);
            KeyboardButtonClickEvent(VK_RETURN, 300);
        }

        CString strCost;
        strCost.Format(L"판매가격 : %d 원", item.cost);

        CopyTextToClipboard(strCost);
        KeyboardButtonClickEvent(VK_RETURN, 300);

        // 사진과 글 사이 줄띄우기
        KeyboardButtonClickEvent(VK_RETURN, 300);
        KeyboardButtonClickEvent(VK_RETURN, 300);
    }

    // 하이퍼링크 입력창으로 내용입력창 탈출 (ctrl + k) + (shift + tab)
    keybd_event(VK_CONTROL, 0, 0, 0);
    keybd_event(75, 0, 0, 0);
    keybd_event(75, 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
    StandBy(2000);

    keybd_event(VK_SHIFT, 0, 0, 0);
    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
    StandBy(2000);

    for (int t = 0; t < 4; t++)
    {
        // 탭 4번 올리기 버튼 포커스
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(500);
    }

    // 등록완료
    printf("등록완료 !!\n");
    KeyboardButtonClickEvent(VK_RETURN, 2000);

    // 강제종료
    system("taskkill /IM chrome.exe /F");
    StandBy(1000);
    system("taskkill /IM chrome.exe /F");
    StandBy(1000);
}

void CETMacroDlg::PlayMacroThunderMarket()
{
    // 번개장터 로그인
    ThunderMarketLogin();

    CString strWebURL = theApp.m_account[1].webURL;
    CString strSellURL = L"/sale/product/register";

    CString strWebSellURL = strWebURL + strSellURL;

    // Sell Item 목록 추가하기
    int nCount = theApp.m_sellItems.size();
    for (int n = 0; n < nCount; n++)
    {
        sellItem item = theApp.m_sellItems[n];

        BOOL bAccept = item.bThunderMarketAccept;

        if (!bAccept)
            continue;

        printf("번개장터 판매 등록 시작 (%d / %d) !!\n\n\n", (n + 1), nCount);

        EventInputURLEditMode();
        CopyTextToClipboard(strWebSellURL);
        KeyboardButtonClickEvent(VK_RETURN, 1000);

        StandBy(2000);

        MouseLbuttonClickEvent(m_btnWebFrameBar);
        StandBy(1000);

        keybd_event(VK_F11, 0, 0, 0);
        keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
        StandBy(5000);



        MouseLbuttonClickEvent(m_btnThunderMarketImageUpload);
        StandBy(2000);


        // 파일 경로 입력창 활성화 
        KeyboardButtonClickEvent(VK_F4, 1000);

        // 팩스페이스로 삭제
        for (int n = 0; n < 100; n++)
        {
            KeyboardButtonClickEvent(VK_BACK, 30);
        }


        // 이미지경로로 이동
        CString strImagePath = GetApplicationDirectory() + L"\\Images\\";
        CopyTextToClipboard(strImagePath);
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(1000);


        // 파일 입력창으로 이동
        keybd_event(VK_MENU, 0, 0, 0);
        keybd_event(78, 0, 0, 0);
        keybd_event(78, 0, KEYEVENTF_KEYUP, 0);
        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        // 이미지 6장 경로 클립보드에 복사	
        CString strImages = item.images;

        CString strImagesFinal;
        for (int img = 0; img < 6; img++)
        {
            int nStart = 0;
            int nCur = strImages.Find(L",", nStart);

            if (img != 5)
            {
                CString strCurImage = strImages.Mid(nStart, nCur);

                strImages = strImages.TrimLeft(strCurImage);
                strImages = strImages.TrimLeft(L", ");

                CString strTemp = L"\"" + strCurImage + L"\" ";

                strImagesFinal += strTemp;
            }
            else
            {
                CString strTemp = L"\"" + strImages + L"\" ";

                strImagesFinal += strTemp;
            }
        }

        CopyTextToClipboard(strImagesFinal);

        printf("이미지 삽입 끝 !!\n");
        // 이미지 등록 (엔터)
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(10000);

        // 기본정보 탭으로 이동
        MouseLbuttonClickEvent(m_btnThunderMarketBaseInfo);
        StandBy(2000);

        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);


        printf("1번째 카테고리 콤보박스 입력 !!\n");
        // 제 1 카테고리로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        // 첫번째 콤보박스 Down 키 카운트
        int nFirstComboCount = item.nThunderMarketFirstComboIndex;
        for (int c = 0; c < nFirstComboCount; c++)
        {
            keybd_event(VK_DOWN, 0, 0, 0);
            keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0);
            StandBy(100);
        }
        StandBy(500);

        printf("2번째 카테고리 콤보박스 입력 !!\n");
        // 제 2 카테고리로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        // 두번째 콤보박스 Down 키 카운트
        int nSecondComboCount = item.nThunderMarketSecondComboIndex - 1;
        for (int c = 0; c < nSecondComboCount; c++)
        {
            keybd_event(VK_DOWN, 0, 0, 0);
            keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0);
            StandBy(100);
        }

        printf("3번째 카테고리 콤보박스 입력 !!\n");
        // 제 3 카테고리로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        // 세번째 콤보박스 Down 키 카운트
        int nThirdComboCount = item.nThunderMarketThirdComboIndex;
        for (int c = 0; c < nThirdComboCount; c++)
        {
            keybd_event(VK_DOWN, 0, 0, 0);
            keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0);
            StandBy(100);
        }
        StandBy(1000);


        printf("최근 지역 선택 !!\n");
        // 최근 지역으로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(200);

        // 최근 지역 창활성화
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(5000);

       
        printf("최근 지역 선택(부산역) !!\n");
        // 가장 상단 위치 클릭
        MouseLbuttonClickEvent(m_btnThunderMarketRecentlyCity);
        StandBy(4000);


        printf("아이템 상태 선택 !!\n");
        // 상태 콤보박스로 이동
        for (int s = 0; s < 3; s++)
        {
            keybd_event(VK_TAB, 0, 0, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
            StandBy(400);
        }

        // 세번째 콤보박스 Down 키 카운트
        int nItemStatus = item.itemStatus;
        for (int c = 0; c < nItemStatus; c++)
        {
            keybd_event(VK_DOWN, 0, 0, 0);
            keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0);
            StandBy(100);
        }
        StandBy(1000);



        printf("타이틀 입력 !!\n");
        // 타이틀 입력으로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        CopyTextToClipboard(item.title);
        StandBy(1000);


        printf("가격 입력 !!\n");
        // 가격 입력으로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        CString strCost;
        strCost.Format(L"%d", item.cost);

        // 태그 내용 복사
        CopyTextToClipboard(strCost);
        StandBy(1000);


        printf("택배비 체크 입력 !!\n");
        // 택배비 체크 입력으로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        BOOL bDeliveryCheck = item.bDeliveryCost;
        if (bDeliveryCheck)
        {
            KeyboardButtonClickEvent(VK_SPACE, 1000);
            StandBy(1000);
        }

        printf("교환 체크 입력 !!\n");
        // 택배비 체크 입력으로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);
        
        BOOL bExchangeCheck = item.bExchange;
        if (bExchangeCheck)
        {
            KeyboardButtonClickEvent(VK_SPACE, 1000);
            StandBy(1000);
        }


        printf("상세내용 입력 !!\n");
        // 타이틀 입력으로 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        // 상세 내용 줄마다 나누기 (클립보드 복사시 죽는 현상 있음)
        int nDescriptionLineCount = int(item.vecDescription.size());

        for (int d = 0; d < nDescriptionLineCount; d++)
        {
            CString strLine = item.vecDescription[d];

            CopyTextToClipboard(strLine);
            KeyboardButtonClickEvent(VK_RETURN, 300);
        }

        printf("태그 입력 !!\n");
        // 태그 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(100);

        int nTagCount = int(item.vecTag.size());

        for (int t = 0; t < nTagCount; t++)
        {
            // 태그는 최대 5개까지 밖에 못올린다.
            if (t == 5)
                break;

            CString strTag = item.vecTag[t];
            // 태그 내용 복사
            CopyTextToClipboard(strTag);

            // 다음태그 입력으로 넘어가려면 엔터를 눌러야한다.
            KeyboardButtonClickEvent(VK_RETURN, 300);
        }

        // ctrl + home 버튼으로 가장 최상단으로 휠 이동
        ButtonPos ptCenter;
        ptCenter.nx = m_rectScreen.Width() * 0.5;
        ptCenter.ny = m_rectScreen.Height() * 0.5;

        MouseLbuttonClickEvent(ptCenter);
        StandBy(1000);

        keybd_event(VK_CONTROL, 0, 0, 0);
        keybd_event(VK_HOME, 0, 0, 0);
        keybd_event(VK_HOME, 0, KEYEVENTF_KEYUP, 0);
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
        StandBy(2000);

        // 상세정보 탭으로 이동
        MouseLbuttonClickEvent(m_btnThunderMarketDetailInfo);
        StandBy(2000);


        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(100);


        // 사이즈 콤보박스 Down 키 카운트
        int nItemSize = item.nSizeComboIndex + 1;
        for (int c = 0; c < nItemSize; c++)
        {
            keybd_event(VK_DOWN, 0, 0, 0);
            keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0);
            StandBy(100);
        }
        StandBy(1000);

        // 등록 완료
        MouseLbuttonClickEvent(m_btnThunderMarketAccept);
        StandBy(5000);

        keybd_event(VK_F11, 0, 0, 0);
        keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
        StandBy(5000);
    }
}


void CETMacroDlg::HelloMarketLogin()
{
    // 헬로 마켓 사이트로 이동 (웹창이 열려야할 시간이 있어야 하므로 20초 여유)
    ShellExecute(m_webHwnd, NULL, theApp.m_URLChrome, theApp.m_account[0].webURL, NULL, SW_SHOWMAXIMIZED);

    StandBy(17000);

    MouseLbuttonClickEvent(m_btnWebFrameBar);

    keybd_event(VK_F11, 0, 0, 0);
    keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
    StandBy(5000);


    // 위의 주석은 텝으로 이동하여 버튼을 클릭하는부분
    // 허용/거부 이런 창이 가끔뜨면 오류를 범하므로 로그인으로 직접
    // 마우스 클릭
    MouseLbuttonClickEvent(m_btnHelloMarketFirstLogin);

    // 탭 10번 눌러야 로그인입력 화면의 로그인 EditBox 로 이동
    for (int n = 0; n < 10; n++)
    {
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(100);
    }

    StandBy(1000);

    // 아이디 클립보드 복사 
    CString strID = theApp.m_account[0].id;
    CopyTextToClipboard(strID);


    keybd_event(VK_TAB, 0, 0, 0);
    keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
    StandBy(500);


    // 패스워드 클립보드 복사 
    CString strPassword = theApp.m_account[0].password;
    CopyTextToClipboard(strPassword);

    MouseLbuttonClickEvent(m_btnHelloMarketLogin);
}


void CETMacroDlg::PlayMacroHelloMarket()
{
    // 헬로 마켓 로그인
    printf("헬로 마켓 로그인 !!\n");
    HelloMarketLogin();

    StandBy(5000);

    BOOL bFisrtStart = TRUE;
    // Sell Item 목록 추가하기
    int nCount = theApp.m_sellItems.size();
    for (int n = 0; n < nCount; n++)
    {
        printf("헬로마켓 판매 등록 시작 (%d / %d) !!\n\n\n", (n + 1), nCount);

        sellItem item = theApp.m_sellItems[n];

        BOOL bAccept = item.bHelloMarketAccept;

        if (!bAccept)
            continue;

        printf("일반판매 등록 진입 !!\n");
        // 최초 아이템 등록시에는 판매등록 및 일반아이템 버튼을 누르면
        // 상세등록 화면으로 전환되나 1회 등록 이후 재버튼 클릭시 반응이없다.
        // 대신 F5로 리프레시하면 상세등록 모드로 바뀐다.
        if (bFisrtStart)
        {
            MouseLbuttonClickEvent(m_btnHelloMarketSellItem);
            StandBy(2000);

            // 일반아이템으로 포커스 이동
            MouseLbuttonClickEvent(m_btnHelloMarketGeneralItem);
            StandBy(3000);
        }
        else
        {
            MouseLbuttonClickEvent(m_btnHelloMarketBlank);
            StandBy(2000);

            keybd_event(VK_F11, 0, 0, 0);
            keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
            StandBy(3000);

            CString strWebURL = theApp.m_account[0].webURL;
            CString strSellURL = L"item/form.hm";

            CString strWebSellURL = strWebURL + strSellURL;

            EventInputURLEditMode();
            CopyTextToClipboard(strWebSellURL);
            KeyboardButtonClickEvent(VK_RETURN, 1000);
            StandBy(5000);

            MouseLbuttonClickEvent(m_btnWebFrameBar);
            StandBy(2000);

            keybd_event(VK_F11, 0, 0, 0);
            keybd_event(VK_F11, 0, KEYEVENTF_KEYUP, 0);
            StandBy(3000);


//             KeyboardButtonClickEvent(VK_F5, 1000);
//             KeyboardButtonClickEvent(VK_F5, 1000);
//             StandBy(3000);
        }


        printf("이미지 삽입 !!\n");

        // 이미지 삽입 버튼 클릭
        MouseLbuttonClickEvent(m_btnHelloMarketImages);
        StandBy(2000);

        // 파일 경로 입력창 활성화 
        KeyboardButtonClickEvent(VK_F4, 1000);

        // 팩스페이스로 삭제
        for (int n = 0; n < 100; n++)
        {
            KeyboardButtonClickEvent(VK_BACK, 30);
        }

        // 파일명 입력창 활성화
        //keybd_event(VK_MENU, 0, 0, 0);
        //keybd_event(78, 0, 0, 0);
        //keybd_event(78, 0, KEYEVENTF_KEYUP, 0);
        //keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);


        // 이미지경로로 이동
        CString strImagePath = GetApplicationDirectory() + L"\\Images\\";
        CopyTextToClipboard(strImagePath);
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(1000);


        // 파일 입력창으로 이동
        keybd_event(VK_MENU, 0, 0, 0);
        keybd_event(78, 0, 0, 0);
        keybd_event(78, 0, KEYEVENTF_KEYUP, 0);
        keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        // 이미지 6장 경로 클립보드에 복사	
        CString strImages = item.images;

        CString strImagesFinal;
        for (int img = 0; img < 6; img++)
        {
            int nStart = 0;
            int nCur = strImages.Find(L",", nStart);

            if (img != 5)
            {
                CString strCurImage = strImages.Mid(nStart, nCur);

                strImages = strImages.TrimLeft(strCurImage);
                strImages = strImages.TrimLeft(L", ");

                CString strTemp = L"\"" + strCurImage + L"\" ";

                strImagesFinal += strTemp;
            }
            else
            {
                CString strTemp = L"\"" + strImages + L"\" ";

                strImagesFinal += strTemp;
            }
        }

        CopyTextToClipboard(strImagesFinal);

        printf("이미지 삽입 끝 !!\n");
        // 이미지 등록 (엔터)
        KeyboardButtonClickEvent(VK_RETURN, 1000);
        StandBy(10000);


        printf("타이틀 입력 !!\n");
        // 제목 입력
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        CopyTextToClipboard(item.title);

        StandBy(1000);
        // 제목에 따른 연결 카테고리 Skip

        int nCheck = m_ckeck_hm_skip.GetCheck();

        if (!nCheck)
        {
            for (int s = 0; s < 4; s++)
            {
                keybd_event(VK_TAB, 0, 0, 0);
                keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
                StandBy(400);
            }
        }

        StandBy(2000);


        printf("1번째 카테고리 콤보박스 입력 !!\n");

        // 카테고리 1 
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        // 첫번째 콤보박스 Down 키 카운트
        int nFirstComboCount = item.nHelloMarketFirstComboIndex;

        for (int c = 0; c < nFirstComboCount; c++)
        {
            // 카테고리 1 
            keybd_event(VK_DOWN, 0, 0, 0);
            keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0);
            StandBy(200);
        }
        StandBy(1000);

        // 카테고리 2
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(1000);

        printf("2번째 카테고리 콤보박스 입력 !!\n");

        // 두번째 콤보박스 Down 키 카운트
        int nSecondComboCount = item.nHelloMarketSecondComboIndex;

        for (int c = 0; c < nSecondComboCount; c++)
        {
            // 카테고리 2
            keybd_event(VK_DOWN, 0, 0, 0);
            keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, 0);
            StandBy(200);
        }

        int nCur = item.nHelloMarketFirstComboIndex;
        if (7 == nCur)
        {
            // 중간 링크 Skip
            keybd_event(VK_TAB, 0, 0, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
            StandBy(200);
        }
        else
        {
            // 중간 링크 Skip
            keybd_event(VK_TAB, 0, 0, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
            StandBy(1000);


            // 상세 내용 내용 입력창
            keybd_event(VK_TAB, 0, 0, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
            StandBy(1000);
        }

        printf("상세설명 입력 !!\n");

        // 상세 내용 줄마다 나누기 (클립보드 복사시 죽는 현상 있음)
        int nDescriptionLineCount = int(item.vecDescription.size());

        for (int d = 0; d < nDescriptionLineCount; d++)
        {
            CString strLine = item.vecDescription[d];

            CopyTextToClipboard(strLine);
            KeyboardButtonClickEvent(VK_RETURN, 500);
        }

        printf("태그 입력 !!\n");

        // 태그 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(100);

        int nTagCount = int(item.vecTag.size());

        for (int t = 0; t < nTagCount; t++)
        {
            // 태그는 최대 5개까지 밖에 못올린다.
            if (t == 5)
                break;

            CString strTag = item.vecTag[t];
            // 태그 내용 복사
            CopyTextToClipboard(strTag);

            // 다음태그 입력으로 넘어가려면 엔터를 눌러야한다.
            KeyboardButtonClickEvent(VK_RETURN, 300);
        }

        printf("가격표 입력 !!\n");

        // 가격표 이동
        keybd_event(VK_TAB, 0, 0, 0);
        keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
        StandBy(100);

        CString strCost;
        strCost.Format(L"%d", item.cost);

        // 태그 내용 복사
        CopyTextToClipboard(strCost);

        printf("가격표 입력 종료 ㅣ %s!!\n", strCost);


        StandBy(1000);

        // 등록완료까지 tab으로 이동
        for (int e = 0; e < 3; e++)
        {
            // 가격표 이동
            keybd_event(VK_TAB, 0, 0, 0);
            keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0);
            StandBy(500);
        }

        printf("등록완료 !!\n");

        // 등록완료 엔터
        KeyboardButtonClickEvent(VK_RETURN, 1000);


        printf("==========================================\n");

        StandBy(4000);

        bFisrtStart = FALSE;
    }

    system("taskkill /IM chrome.exe /F");

    // 가끔 잔존하는 크롬이 있으니 한번더 강제종료
    Sleep(1000);
    system("taskkill /IM chrome.exe /F");

    Sleep(1000);
}

void CETMacroDlg::PlayMacro()
{
    // 기존에 열려있는 크롬창 전부 강제종료
    system("taskkill /IM chrome.exe /F");

    // 가끔 잔존하는 크롬이 있으니 한번더 강제종료
    Sleep(1000);
    system("taskkill /IM chrome.exe /F");

    Sleep(1000);

    BOOL bHelloMarketLogin = FALSE;
    BOOL bThunderMarketLogin = FALSE;
    BOOL bNaverCafeLogin = FALSE;
    BOOL bNaverBandLogin = FALSE;
    BOOL bKakaoStoryLogin = FALSE;
    int nCount = theApp.m_sellItems.size();
    for (int n = 0; n < nCount; n++)
    {
        sellItem item = theApp.m_sellItems[n];

        BOOL b0 = item.bHelloMarketAccept;
        BOOL b1 = item.bThunderMarketAccept;
        BOOL b2 = item.bNaverCafeAccept;
        BOOL b3 = item.bNaverBandAccept;
        BOOL b4 = item.bKakaoStoryAccept;

        if (b0)
            bHelloMarketLogin = TRUE;

        if (b1)
            bThunderMarketLogin = TRUE;

        if (b2)
            bNaverCafeLogin = TRUE;
 
        if (b3)
            bNaverBandLogin = TRUE;

        if (b4)
            bKakaoStoryLogin = TRUE;
    }


    // 헬로 마켓 매크로 시작
    if(bHelloMarketLogin)
        PlayMacroHelloMarket();

    // 번개 장터 매크로 시작
    if(bThunderMarketLogin)
        PlayMacroThunderMarket();

    // 네이버 카페 매크로 시작
    if(bNaverCafeLogin)
        PlayMacroNaverCafe();

    // 네이버 밴드 매크로 시작
    if (bNaverBandLogin)
        PlayMacroNaverBand();

    // 카카오스토리 매크로 시작
    if (bKakaoStoryLogin)
        PlayMacroKakaoStory();

}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// worker thread
UINT CETMacroDlg::procWorkerThread(LPVOID lpParam)
{
    CETMacroDlg* pObj = reinterpret_cast<CETMacroDlg*>(lpParam);

    while (1)
    {
        if (!pObj->m_bContinue)
            continue;







        // 매크로 시작 된다.
        pObj->PlayMacro();

        Sleep(1);

        pObj->m_bContinue = FALSE;
    }

    return 0;
}




void CETMacroDlg::OnCbnSelchangeComboHm1()
{
    int nCur = m_combo_hm1.GetCurSel();

    if (0 == nCur)
    {
        m_combo_hm2.ResetContent();
        m_combo_hm2.AddString(L"1번 카테고리를 먼저 선택");
        m_combo_hm2.SetCurSel(0);

        m_combo_hm2.EnableWindow(FALSE);

        return;
    }

    m_combo_hm2.ResetContent();
    m_combo_hm2.EnableWindow(TRUE);

    if (1 == nCur)
    {//모터사이클,용품
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"125cc 이하");
        m_combo_hm2.AddString(L"125cc 초과");
        m_combo_hm2.AddString(L"500cc 초과");

        m_combo_hm2.SetCurSel(0);
    }
    else if (2 == nCur)
    {//유아동,완구
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"신생아, 유아의류 (유아동, 완구)");
        m_combo_hm2.AddString(L"아동의류 (유아동, 완구)");
        m_combo_hm2.AddString(L"유아동잡화 (유아동, 완구)");
        m_combo_hm2.AddString(L"유아동생활용품 (유아동, 완구)");
        m_combo_hm2.AddString(L"완구, 인형 (유아동, 완구)");
        m_combo_hm2.AddString(L"임부복, 출산용품 (유아동, 완구)");

        m_combo_hm2.SetCurSel(0);
    }
    else if (3 == nCur)
    {//자동차용품
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"블랙박스,네비게이션 ( 자동차용품 )");
        m_combo_hm2.AddString(L"타이어,휠,체인 ( 자동차용품 )");
        m_combo_hm2.AddString(L"카오디오,AV ( 자동차용품 )");
        m_combo_hm2.AddString(L"실내부품,용품 ( 자동차용품 )");
        m_combo_hm2.AddString(L"외장부품,용품 ( 자동차용품 )");
        m_combo_hm2.AddString(L"세차,청소용품 ( 자동차용품 )");
        m_combo_hm2.AddString(L"기타 자동차용품 ( 자동차용품 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (4 == nCur)
    {//뷰티
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"여성화장품 ( 뷰티 )");
        m_combo_hm2.AddString(L"메이크업 ( 뷰티 )");
        m_combo_hm2.AddString(L"남성화장품 ( 뷰티 )");
        m_combo_hm2.AddString(L"향수,헤어,바디 ( 뷰티 )");
        m_combo_hm2.AddString(L"기타 뷰티 ( 뷰티 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (5 == nCur)
    {//바이크용품
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"라이더헬맷");
        m_combo_hm2.AddString(L"라이더의류");
        m_combo_hm2.AddString(L"라이더신발, 잡화");
        m_combo_hm2.AddString(L"바이크용품, 부품");
        m_combo_hm2.AddString(L"기타바이크용품");

        m_combo_hm2.SetCurSel(0);
    }
    else if (6 == nCur)
    {//여성의류
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"여성코트,아우터 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성티셔츠 ( 여성의류 )");
        m_combo_hm2.AddString(L"남방,블라우스 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성니트 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성조끼 ( 여성의류 )");
        m_combo_hm2.AddString(L"원피스,정장 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성바지 ( 여성의류 )");
        m_combo_hm2.AddString(L"스커트,치마 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성트레이닝복 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성속옷 ( 여성의류 )");
        m_combo_hm2.AddString(L"기타 여성의류 ( 여성의류 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (7 == nCur)
    {//남성의류
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"남성코트,아우터 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성티셔츠 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성남방 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성니트 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성바지 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성정장 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성트레이닝복 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성속옷 ( 남성의류 )");
        m_combo_hm2.AddString(L"기타 남성의류 ( 남성의류 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (8 == nCur)
    {//신발,가방,잡화
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"여성신발 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"남성신발 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"운동화,기능화 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"가방 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"시계,보석 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"지갑,벨트 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"모자,안경 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"기타 잡화 ( 신발,가방,잡화 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (9 == nCur)
    {//휴대폰,태블릿
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"삼성 ( 휴대폰,태블릿 )");
        m_combo_hm2.AddString(L"애플 ( 휴대폰,태블릿 )");
        m_combo_hm2.AddString(L"LG ( 휴대폰,태블릿 )");
        m_combo_hm2.AddString(L"기타 휴대폰 ( 휴대폰,태블릿 )");
        m_combo_hm2.AddString(L"태블릿 ( 휴대폰,태블릿 )");
        m_combo_hm2.AddString(L"액세서리,주변기기 ( 휴대폰,태블릿 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (10 == nCur)
    {//컴퓨터,주변기기
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"노트북 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"데스크탑 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"모니터 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"컴퓨터부품 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"프린터,오피스기기 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"저장장치 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"소프트웨어 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"학습기기 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"기타 기기,용품 ( 컴퓨터,주변기기 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (11 == nCur)
    {//카메라
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"일반디카 ( 카메라 )");
        m_combo_hm2.AddString(L"DSLR ( 카메라 )");
        m_combo_hm2.AddString(L"필름카메라 ( 카메라 )");
        m_combo_hm2.AddString(L"카메라렌즈 ( 카메라 )");
        m_combo_hm2.AddString(L"카메라액세서리 ( 카메라 )");
        m_combo_hm2.AddString(L"캠코더 ( 카메라 )");
        m_combo_hm2.AddString(L"기타 광학용품 ( 카메라 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (12 == nCur)
    {//디지털,가전
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"TV ( 디지털,가전 )");
        m_combo_hm2.AddString(L"청소기 ( 디지털,가전 )");
        m_combo_hm2.AddString(L"냉장고 ( 디지털,가전 )");
        m_combo_hm2.AddString(L"세탁기 ( 디지털,가전 )");
        m_combo_hm2.AddString(L"주방조리가전 ( 디지털,가전 )");
        m_combo_hm2.AddString(L"건강,계절가전 ( 디지털,가전 )");
        m_combo_hm2.AddString(L"MP3,iPod ( 디지털,가전 )");
        m_combo_hm2.AddString(L"기타 디지털,가전 ( 디지털,가전 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (13 == nCur)
    {//게임
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"PC게임 ( 게임 )");
        m_combo_hm2.AddString(L"닌텐도 ( 게임 )");
        m_combo_hm2.AddString(L"플레이스테이션 ( 게임 )");
        m_combo_hm2.AddString(L"PSP ( 게임 )");
        m_combo_hm2.AddString(L"Wii ( 게임 )");
        m_combo_hm2.AddString(L"Xbox ( 게임 )");
        m_combo_hm2.AddString(L"보드,퍼즐 ( 게임 )");
        m_combo_hm2.AddString(L"기타 게임관련 ( 게임 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (14 == nCur)
    {//스포츠,레저
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"자전거 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"등산 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"캠핑 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"골프 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"낚시 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"스키,보드 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"수상스포츠 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"축구 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"야구 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"농구 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"인라인,X게임 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"헬스,요가 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"기타 스포츠 ( 스포츠,레저 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (15 == nCur)
    {//가구
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"침실가구 ( 가구 )");
        m_combo_hm2.AddString(L"거실가구 ( 가구 )");
        m_combo_hm2.AddString(L"수납가구 ( 가구 )");
        m_combo_hm2.AddString(L"주방가구 ( 가구 )");
        m_combo_hm2.AddString(L"책상,책장 ( 가구 )");
        m_combo_hm2.AddString(L"의자 ( 가구 )");
        m_combo_hm2.AddString(L"기타 가구 ( 가구 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (16 == nCur)
    {//생활
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"침구,커튼,카펫트 ( 생활 )");
        m_combo_hm2.AddString(L"원예,수예 ( 생활 )");
        m_combo_hm2.AddString(L"세탁,청소용품 ( 생활 )");
        m_combo_hm2.AddString(L"욕실용품 ( 생활 )");
        m_combo_hm2.AddString(L"주방용품 ( 생활 )");
        m_combo_hm2.AddString(L"인테리어소품 ( 생활 )");
        m_combo_hm2.AddString(L"생활,수납용품 ( 생활 )");
        m_combo_hm2.AddString(L"공구,연장 ( 생활 )");
        m_combo_hm2.AddString(L"기타 생활 ( 생활 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (17 == nCur)
    {//골동품,희귀품
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"골동품 ( 골동품,희귀품 )");
        m_combo_hm2.AddString(L"희귀품 ( 골동품,희귀품 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (18 == nCur)
    {//여행,숙박
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"내가찍은여행사진 ( 여행,숙박 )");
        m_combo_hm2.AddString(L"여행정보,가이드 ( 여행,숙박 )");
        m_combo_hm2.AddString(L"대명,한화숙박권 ( 여행,숙박 )");
        m_combo_hm2.AddString(L"리조트,호텔 ( 여행,숙박 )");
        m_combo_hm2.AddString(L"펜션,캠핑,기타숙박 ( 여행,숙박 )");
        m_combo_hm2.AddString(L"해외숙박 ( 여행,숙박 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (19 == nCur)
    {//티켓
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"영화,공연,전시 ( 티켓 )");
        m_combo_hm2.AddString(L"스포츠,레저 ( 티켓 )");
        m_combo_hm2.AddString(L"테마파크,워터파크 ( 티켓 )");
        m_combo_hm2.AddString(L"e티켓,상품권 ( 티켓 )");
        m_combo_hm2.AddString(L"기타 티켓 ( 티켓 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (20 == nCur)
    {//재능,서비스
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"과외 ( 재능,서비스 )");
        m_combo_hm2.AddString(L"아르바이트 ( 재능,서비스 )");
        m_combo_hm2.AddString(L"전문스킬,프리랜서 ( 재능,서비스 )");
        m_combo_hm2.AddString(L"가사,생활도움 ( 재능,서비스 )");
        m_combo_hm2.AddString(L"기타 재능공유 ( 재능,서비스 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (21 == nCur)
    {//도서
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"일반도서 ( 도서 )");
        m_combo_hm2.AddString(L"교재,전문 ( 도서 )");
        m_combo_hm2.AddString(L"유아동,전집 ( 도서 )");
        m_combo_hm2.AddString(L"만화책 ( 도서 )");
        m_combo_hm2.AddString(L"여행,취미 ( 도서 )");
        m_combo_hm2.AddString(L"잡지 ( 도서 )");
        m_combo_hm2.AddString(L"외국도서 ( 도서 )");
        m_combo_hm2.AddString(L"기타 도서 ( 도서 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (22 == nCur)
    {//스타굿즈
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"포토카드,인스,포스터 ( 스타굿즈 )");
        m_combo_hm2.AddString(L"음반 ( 스타굿즈 )");
        m_combo_hm2.AddString(L"응원도구 ( 스타굿즈 )");
        m_combo_hm2.AddString(L"의류 ( 스타굿즈 )");
        m_combo_hm2.AddString(L"잡화,악세서리 ( 스타굿즈 )");
        m_combo_hm2.AddString(L"기타 스타굿즈 ( 스타굿즈 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (23 == nCur)
    {//문구
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"문구용품 ( 문구 )");
        m_combo_hm2.AddString(L"사무용품 ( 문구 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (24 == nCur)
    {//피규어,키덜트
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"피규어 ( 피규어,키덜트 )");
        m_combo_hm2.AddString(L"프라모델,레고 ( 피규어,키덜트 )");
        m_combo_hm2.AddString(L"RC,드론 ( 피규어,키덜트 )");
        m_combo_hm2.AddString(L"기타 키덜트 ( 피규어,키덜트 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (25 == nCur)
    {//CD,DVD
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"CD,LP ( CD,DVD )");
        m_combo_hm2.AddString(L"DVD ( CD,DVD )");
        m_combo_hm2.AddString(L"유아동CD,DVD ( CD,DVD )");
        m_combo_hm2.AddString(L"교육콘텐츠 ( CD,DVD )");
        m_combo_hm2.AddString(L"기타 CD,DVD ( CD,DVD )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (26 == nCur)
    {//음향기기,악기
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"헤드폰,이어폰 ( 음향기기,악기 )");
        m_combo_hm2.AddString(L"스피커,오디오 ( 음향기기,악기 )");
        m_combo_hm2.AddString(L"기타 음향기기 ( 음향기기,악기 )");
        m_combo_hm2.AddString(L"피아노,건반악기 ( 음향기기,악기 )");
        m_combo_hm2.AddString(L"바이올린,현악기 ( 음향기기,악기 )");
        m_combo_hm2.AddString(L"그외 악기 ( 음향기기,악기 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (27 == nCur)
    {//예술,미술
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"예술,미술작품 ( 예술,미술 )");
        m_combo_hm2.AddString(L"미술용품 ( 예술,미술 )");
        m_combo_hm2.AddString(L"기타 예술,미술 ( 예술,미술 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (28 == nCur)
    {//반려동물
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"반려동물용품 ( 반려동물 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (29 == nCur)
    {//부동산
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"매매 ( 부동산 )");
        m_combo_hm2.AddString(L"전,월세 ( 부동산 )");
        m_combo_hm2.AddString(L"쉐어,룸메이트 ( 부동산 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (30 == nCur)
    {//포장식품
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"건강기능식품 ( 포장식품 )");
        m_combo_hm2.AddString(L"기타 포장식품 ( 포장식품 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (31 == nCur)
    {//핸드메이드
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"수제패션,소품 ( 핸드메이드 )");
        m_combo_hm2.AddString(L"기타 수제품 ( 핸드메이드 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (32 == nCur)
    {//기타
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"기타 ( 기타 )");

        m_combo_hm2.SetCurSel(0);
    }
}

void CETMacroDlg::OnCbnSelchangeComboTm1()
{
    int nCurTm1 = m_combo_tm1.GetCurSel();

    if (0 == nCurTm1)
    {
        m_combo_tm2.ResetContent();
        m_combo_tm2.AddString(L"1번 카테고리를 먼저 선택");
        m_combo_tm2.SetCurSel(0);
        m_combo_tm2.EnableWindow(FALSE);

        m_combo_tm3.ResetContent();
        m_combo_tm3.AddString(L"2번 카테고리를 먼저 선택");
        m_combo_tm3.SetCurSel(0);
        m_combo_tm3.EnableWindow(FALSE);

        m_combo_tm_size.ResetContent();
        m_combo_tm_size.AddString(L"2번 카테고리를 먼저 선택");
        m_combo_tm_size.SetCurSel(0);
        m_combo_tm_size.EnableWindow(FALSE);

        return;
    }

    m_combo_tm2.ResetContent();
    m_combo_tm2.EnableWindow(TRUE);

    if (1 == nCurTm1)
    {//여성의류
        m_combo_tm2.AddString(L"카테고리");
        m_combo_tm2.AddString(L"긴팔 티셔츠(여성의류)");
        m_combo_tm2.AddString(L"반팔 티셔츠(여성의류)");
        m_combo_tm2.AddString(L"맨투맨/후드티(여성의류)");
        m_combo_tm2.AddString(L"원피스(여성의류)");
        m_combo_tm2.AddString(L"블라우스(여성의류)");
        m_combo_tm2.AddString(L"셔츠/남방(여성의류)");
        m_combo_tm2.AddString(L"니트/스웨터(여성의류)");
        m_combo_tm2.AddString(L"가디건(여성의류)");
        m_combo_tm2.AddString(L"조끼/베스트(여성의류)");
        m_combo_tm2.AddString(L"야상/점퍼/패딩(여성의류)");
        m_combo_tm2.AddString(L"자켓(여성의류)");
        m_combo_tm2.AddString(L"코트(여성의류)");
        m_combo_tm2.AddString(L"스커트/치마(여성의류)");
        m_combo_tm2.AddString(L"청바지/스키니진(여성의류)");
        m_combo_tm2.AddString(L"면/캐주얼 바지(여성의류)");
        m_combo_tm2.AddString(L"반바지/핫팬츠(여성의류)");
        m_combo_tm2.AddString(L"레깅스(여성의류)");
        m_combo_tm2.AddString(L"비지니스 정장(여성의류)");
        m_combo_tm2.AddString(L"트레이닝(여성의류)");
        m_combo_tm2.AddString(L"언더웨어/속옷(여성의류)");
        m_combo_tm2.AddString(L"빅사이즈(여성의류)");
        m_combo_tm2.AddString(L"테마/이벤트옷(여성의류)");

        m_combo_tm2.SetCurSel(0);
    }
    else if (2 == nCurTm1)
    {//남성의류
        m_combo_tm2.AddString(L"카테고리");
        m_combo_tm2.AddString(L"긴팔 티셔츠(남성의류)");
        m_combo_tm2.AddString(L"반팔 티셔츠(남성의류)");
        m_combo_tm2.AddString(L"맨투맨/후드티(남성의류)");
        m_combo_tm2.AddString(L"셔츠/남방(남성의류)");
        m_combo_tm2.AddString(L"니트/스웨터(남성의류)");
        m_combo_tm2.AddString(L"가디건(남성의류)");
        m_combo_tm2.AddString(L"조끼/베스트(남성의류)");
        m_combo_tm2.AddString(L"점퍼/야상/패딩(남성의류)");
        m_combo_tm2.AddString(L"자켓(남성의류)");
        m_combo_tm2.AddString(L"코트(남성의류)");
        m_combo_tm2.AddString(L"청바지(긴)(남성의류)");
        m_combo_tm2.AddString(L"면/캐주얼 바지(남성의류)");
        m_combo_tm2.AddString(L"반바지/7~9부(남성의류)");
        m_combo_tm2.AddString(L"비지니스 정장(남성의류)");
        m_combo_tm2.AddString(L"트레이닝(남성의류)");
        m_combo_tm2.AddString(L"언더웨어/속옷(남성의류)");
        m_combo_tm2.AddString(L"빅사이즈(남성의류)");
        m_combo_tm2.AddString(L"테마/이벤트옷(남성의류)");

        m_combo_tm2.SetCurSel(0);
    }
    else if (3 == nCurTm1)
    {//패션잡화
        m_combo_tm2.AddString(L"카테고리");
        m_combo_tm2.AddString(L"여성가방(패션잡화)");
        m_combo_tm2.AddString(L"남성가방(패션잡화)");
        m_combo_tm2.AddString(L"여행용가방(패션잡화)");
        m_combo_tm2.AddString(L"운동화(패션잡화)");
        m_combo_tm2.AddString(L"여성화(패션잡화)");
        m_combo_tm2.AddString(L"남성화(패션잡화)");
        m_combo_tm2.AddString(L"지갑(패션잡화)");
        m_combo_tm2.AddString(L"모자(패션잡화)");
        m_combo_tm2.AddString(L"안경/선글라스(패션잡화)");
        m_combo_tm2.AddString(L"주얼리(패션잡화)");
        m_combo_tm2.AddString(L"시계(패션잡화)");
        m_combo_tm2.AddString(L"벨트/장갑(패션잡화)");


        m_combo_tm2.SetCurSel(0);
    }
    else if (4 == nCurTm1)
    {//유아동/출산
        m_combo_tm2.AddString(L"카테고리");
        m_combo_tm2.AddString(L"베이비의류(유아동/출산)");
        m_combo_tm2.AddString(L"여아의류(유아동/출산)");
        m_combo_tm2.AddString(L"남아의류(유아동/출산)");
        m_combo_tm2.AddString(L"여주니어의류(유아동/출산)");
        m_combo_tm2.AddString(L"남주니어의류(유아동/출산)");
        m_combo_tm2.AddString(L"유아동신발(유아동/출산)");
        m_combo_tm2.AddString(L"유아동용품(유아동/출산)");
        m_combo_tm2.AddString(L"출산(유아동/출산)");
        m_combo_tm2.AddString(L"교육/완구(유아동/출산)");
        m_combo_tm2.AddString(L"기저귀(유아동/출산)");


        m_combo_tm2.SetCurSel(0);
    }

}


void CETMacroDlg::OnCbnSelchangeComboTm2()
{
    int nCurTm1 = m_combo_tm1.GetCurSel();
    int nCurTm2 = m_combo_tm2.GetCurSel();

    if (0 == nCurTm2)
    {
        m_combo_tm3.ResetContent();
        m_combo_tm3.AddString(L"2번 카테고리를 먼저 선택");
        m_combo_tm3.SetCurSel(0);
        m_combo_tm3.EnableWindow(FALSE);

        m_combo_tm_size.ResetContent();
        m_combo_tm_size.AddString(L"2번 카테고리를 먼저 선택");
        m_combo_tm_size.SetCurSel(0);
        m_combo_tm_size.EnableWindow(FALSE);

        return;
    }

    m_combo_tm3.ResetContent();
    m_combo_tm3.EnableWindow(TRUE);

    m_combo_tm_size.ResetContent();
    m_combo_tm_size.EnableWindow(TRUE);

    // 1번 카테고리와 2번 카테고리의 인덱스에따라서 3번, Size 콤보박스가 다르게 생성됨(주의해야함)
    if (1 == nCurTm1 && 1 == nCurTm2)
    {//여성의류 && 긴팔 티셔츠(여성의류)
        m_combo_tm3.AddString(L"무지/기본 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"라운드 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"브이넥 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"스프라이프 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"폴라 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"기타(긴팔 티셔츠(여성의류)");
    }
    else if (1 == nCurTm1 && 2 == nCurTm2)
    {//여성의류 && 반팔 티셔츠(여성의류)
        m_combo_tm3.AddString(L"무지/기본 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"라운드 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"브이넥 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"카라 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"스프라이프 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"민소매/나시 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"기타(반팔 티셔츠(여성의류)");
    }

    else if (1 == nCurTm1 && 3 == nCurTm2)
    {//여성의류 && 맨투맨/후드티(여성의류)
        m_combo_tm3.AddString(L"맨투맨 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"후드 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"후드 집업(여성의류)");
        m_combo_tm3.AddString(L"카라 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"기타(여성의류)");
    }

    else if (1 == nCurTm1 && 4 == nCurTm2)
    {//여성의류 && 원피스(여성의류)
        m_combo_tm3.AddString(L"캐주얼 원피스(여성의류)");
        m_combo_tm3.AddString(L"미니 원피스(여성의류)");
        m_combo_tm3.AddString(L"롱 원피스(여성의류)");
        m_combo_tm3.AddString(L"나시/탑 원피스(여성의류)");
        m_combo_tm3.AddString(L"쉬폰/레이스 원피스(여성의류)");
        m_combo_tm3.AddString(L"럭셔리 원피스(여성의류)");
        m_combo_tm3.AddString(L"후드/니트 원피스(여성의류)");
        m_combo_tm3.AddString(L"청 원피스(여성의류)");
        m_combo_tm3.AddString(L"프린트 원피스(여성의류)");
        m_combo_tm3.AddString(L"투피스 원피스(여성의류)");
        m_combo_tm3.AddString(L"기타(원피스)(여성의류)");
    }

    else if (1 == nCurTm1 && 5 == nCurTm2)
    {//여성의류 && 블라우스(여성의류)
        m_combo_tm3.AddString(L"쉬폰/시스루 블라우스(여성의류)");
        m_combo_tm3.AddString(L"레이스 블라우스(여성의류)");
        m_combo_tm3.AddString(L"프릴/셔링 블라우스(여성의류)");
        m_combo_tm3.AddString(L"프린트 블라우스(여성의류)");
        m_combo_tm3.AddString(L"오프숄더 블라우스(여성의류)");
        m_combo_tm3.AddString(L"민소매/홀터넥 블라우스(여성의류)");
        m_combo_tm3.AddString(L"기타(블라우스)(여성의류)");
    }

    else if (1 == nCurTm1 && 6 == nCurTm2)
    {//여성의류 && 셔츠/남방(여성의류)
        m_combo_tm3.AddString(L"무지/기본 셔츠(여성의류)");
        m_combo_tm3.AddString(L"루즈핏/박시 셔츠(여성의류)");
        m_combo_tm3.AddString(L"체크 셔츠(여성의류)");
        m_combo_tm3.AddString(L"청/데님 셔츠(여성의류)");
        m_combo_tm3.AddString(L"스트라이프 셔츠(여성의류)");
        m_combo_tm3.AddString(L"기타(셔츠/남방)(여성의류)");
    }

    else if (1 == nCurTm1 && 7 == nCurTm2)
    {//여성의류 && 니트/스웨터(여성의류)
        m_combo_tm3.AddString(L"라운드넥 니트(여성의류)");
        m_combo_tm3.AddString(L"브이넥 니트(여성의류)");
        m_combo_tm3.AddString(L"오프숄더 니트(여성의류)");
        m_combo_tm3.AddString(L"폴라/터틀(여성의류)");
        m_combo_tm3.AddString(L"롱 니트(여성의류)");
        m_combo_tm3.AddString(L"루즈핏 니트(여성의류)");
        m_combo_tm3.AddString(L"기타(니트/스웨터)(여성의류)");
    }

    else if (1 == nCurTm1 && 8 == nCurTm2)
    {//여성의류 && 가디건(여성의류)
        m_combo_tm3.AddString(L"라운드넥 가디건(여성의류)");
        m_combo_tm3.AddString(L"브이넥 가디건(여성의류)");
        m_combo_tm3.AddString(L"루즈핏/박시 가디건(여성의류)");
        m_combo_tm3.AddString(L"롱 가디건(여성의류)");
        m_combo_tm3.AddString(L"후드 가디건(여성의류)");
        m_combo_tm3.AddString(L"기타(가디건)(여성의류)");
    }

    else if (1 == nCurTm1 && 9 == nCurTm2)
    {//여성의류 && 조끼/베스트(여성의류)
        m_combo_tm3.AddString(L"니트 조끼(여성의류)");
        m_combo_tm3.AddString(L"청/데님 조끼(여성의류)");
        m_combo_tm3.AddString(L"퍼 조끼(여성의류)");
        m_combo_tm3.AddString(L"패딩 조끼(여성의류)");
        m_combo_tm3.AddString(L"기타(조끼/베스트)(여성의류)");
    }

    else if (1 == nCurTm1 && 10 == nCurTm2)
    {//여성의류 && 야상/점퍼/패딩(여성의류)
        m_combo_tm3.AddString(L"야상/사파리(여성의류)");
        m_combo_tm3.AddString(L"야구점퍼(여성의류)");
        m_combo_tm3.AddString(L"패딩(여성의류)");
        m_combo_tm3.AddString(L"바람막이(여성의류)");
        m_combo_tm3.AddString(L"기타(야상/점퍼/패딩)(여성의류)");
    }

    else if (1 == nCurTm1 && 11 == nCurTm2)
    {//여성의류 && 자켓(여성의류)
        m_combo_tm3.AddString(L"기본/테일러드 자켓(여성의류)");
        m_combo_tm3.AddString(L"청/데님자켓(여성의류)");
        m_combo_tm3.AddString(L"트위드/체크자켓(여성의류)");
        m_combo_tm3.AddString(L"가죽/라이더(여성의류)");
        m_combo_tm3.AddString(L"기타(자켓)(자켓(여성의류))");
    }

    else if (1 == nCurTm1 && 12 == nCurTm2)
    {//여성의류 && 코트(여성의류)
        m_combo_tm3.AddString(L"트렌치 코트(여성의류)");
        m_combo_tm3.AddString(L"반/하프 코트(여성의류)");
        m_combo_tm3.AddString(L"롱 코트(여성의류)");
        m_combo_tm3.AddString(L"케이프/망토(여성의류)");
        m_combo_tm3.AddString(L"무스탕(여성의류)");
        m_combo_tm3.AddString(L"모피(여성의류)");
        m_combo_tm3.AddString(L"기타(코트)(코트(여성의류))");
    }

    else if (1 == nCurTm1 && 13 == nCurTm2)
    {//여성의류 && 스커트/치마(여성의류)
        m_combo_tm3.AddString(L"쉬폰/레이스(여성의류)");
        m_combo_tm3.AddString(L"플레어 스커트(여성의류)");
        m_combo_tm3.AddString(L"롱 스커트(여성의류)");
        m_combo_tm3.AddString(L"미니 스커트(여성의류)");
        m_combo_tm3.AddString(L"모직/니트 스커트(여성의류)");
        m_combo_tm3.AddString(L"플리츠(주름)(여성의류)");
        m_combo_tm3.AddString(L"청스커트(여성의류)");
        m_combo_tm3.AddString(L"기타(스커트/치마)(여성의류)");
    }

    else if (1 == nCurTm1 && 14 == nCurTm2)
    {//여성의류 && 청바지/스키니(긴)(여성의류)
        m_combo_tm3.AddString(L"스키니진(여성의류)");
        m_combo_tm3.AddString(L"일자 청바지(여성의류)");
        m_combo_tm3.AddString(L"부츠컷 청바지(여성의류)");
        m_combo_tm3.AddString(L"배기/카고(여성의류)");
        m_combo_tm3.AddString(L"하이웨스트 진(여성의류)");
        m_combo_tm3.AddString(L"기타(청바지/스키니(긴))(여성의류)");
    }

    else if (1 == nCurTm1 && 15 == nCurTm2)
    {//여성의류 && 면/캐주얼 바지(긴)(여성의류)
        m_combo_tm3.AddString(L"일자바지/슬렉스(여성의류)");
        m_combo_tm3.AddString(L"점프 수트/멜빵(여성의류)");
        m_combo_tm3.AddString(L"통/와이드 팬츠(여성의류)");
        m_combo_tm3.AddString(L"배기 팬츠(여성의류)");
        m_combo_tm3.AddString(L"하이웨스트 팬츠(여성의류)");
        m_combo_tm3.AddString(L"가죽/모직 바지(여성의류)");
        m_combo_tm3.AddString(L"기타(면/캐주얼 바지(긴))(여성의류)");
    }

    else if (1 == nCurTm1 && 16 == nCurTm2)
    {//여성의류 && 반바지/핫팬츠(여성의류)
        m_combo_tm3.AddString(L"면 반바지(여성의류)");
        m_combo_tm3.AddString(L"청 반바지(여성의류)");
        m_combo_tm3.AddString(L"핫 팬츠(여성의류)");
        m_combo_tm3.AddString(L"치마 바지(여성의류)");
        m_combo_tm3.AddString(L"가죽/모직 반바지(여성의류)");
        m_combo_tm3.AddString(L"기타(반바지/핫팬츠)(여성의류)");
    }

    else if (1 == nCurTm1 && 17 == nCurTm2)
    {//여성의류 && 레깅스(여성의류)
        m_combo_tm3.AddString(L"무지 레깅스(여성의류)");
        m_combo_tm3.AddString(L"치마 레깅스(여성의류)");
        m_combo_tm3.AddString(L"프린트 레깅스(여성의류)");
        m_combo_tm3.AddString(L"기모/밍크 레깅스(여성의류)");
        m_combo_tm3.AddString(L"가죽 레깅스(여성의류)");
        m_combo_tm3.AddString(L"기타(레깅스)(여성의류)");
    }

    else if (1 == nCurTm1 && 18 == nCurTm2)
    {//여성의류 && 비즈니스 정장(여성의류)
        m_combo_tm3.AddString(L"정장 세트(비즈니스 정장)(여성의류)");
        m_combo_tm3.AddString(L"정장 원피스(여성의류)");
        m_combo_tm3.AddString(L"정장 자켓(여성의류)");
        m_combo_tm3.AddString(L"정장 블라우스(여성의류)");
        m_combo_tm3.AddString(L"정장 바지/슬랙스(여성의류)");
        m_combo_tm3.AddString(L"정장 치마(여성의류)");
        m_combo_tm3.AddString(L"기타(비즈니스 정장)(여성의류)");
    }

    else if (1 == nCurTm1 && 19 == nCurTm2)
    {//여성의류 && 트레이닝(여성의류)
        m_combo_tm3.AddString(L"트레이닝 상의(여성의류)");
        m_combo_tm3.AddString(L"트레이닝 하의(여성의류)");
        m_combo_tm3.AddString(L"트레이닝 세트(여성의류)");
        m_combo_tm3.AddString(L"기타(트레이닝)(여성의류)");
    }

    else if (1 == nCurTm1 && 20 == nCurTm2)
    {//여성의류 && 언더웨어/속옷(여성의류)
        m_combo_tm3.AddString(L"브라(언더웨어/속옷)(여성의류)");
        m_combo_tm3.AddString(L"팬티(언더웨어/속옷)(여성의류)");
        m_combo_tm3.AddString(L"브라팬티 세트(여성의류)");
        m_combo_tm3.AddString(L"보정 속옷(여성의류)");
        m_combo_tm3.AddString(L"잠옷/이지웨어(여성의류)");
        m_combo_tm3.AddString(L"런닝/내의(여성의류)");
        m_combo_tm3.AddString(L"슬립/캐미솔(여성의류)");
        m_combo_tm3.AddString(L"속치마/속바지(여성의류)");
        m_combo_tm3.AddString(L"기능성/히트텍(여성의류)");
        m_combo_tm3.AddString(L"기타(언더웨어/속옷)(여성의류)");

        m_combo_tm_size.AddString(L"70A");
        m_combo_tm_size.AddString(L"70B");
        m_combo_tm_size.AddString(L"70C");
        m_combo_tm_size.AddString(L"70D");
        m_combo_tm_size.AddString(L"70E");
        m_combo_tm_size.AddString(L"70F");
        m_combo_tm_size.AddString(L"75A");
        m_combo_tm_size.AddString(L"75B");
        m_combo_tm_size.AddString(L"75C");
        m_combo_tm_size.AddString(L"75D");
        m_combo_tm_size.AddString(L"75E");
        m_combo_tm_size.AddString(L"75F");
        m_combo_tm_size.AddString(L"80A");
        m_combo_tm_size.AddString(L"80B");
        m_combo_tm_size.AddString(L"80C");
        m_combo_tm_size.AddString(L"80D");
        m_combo_tm_size.AddString(L"80E");
        m_combo_tm_size.AddString(L"80F");
        m_combo_tm_size.AddString(L"85A");
        m_combo_tm_size.AddString(L"85B");
        m_combo_tm_size.AddString(L"85C");
        m_combo_tm_size.AddString(L"85D");
        m_combo_tm_size.AddString(L"85E");
        m_combo_tm_size.AddString(L"85F");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (1 == nCurTm1 && 21 == nCurTm2)
    {//여성의류 && 빅사이즈(여성의류)
        m_combo_tm3.AddString(L"긴팔 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"반팔 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"맨투맨/후드티(여성의류)");
        m_combo_tm3.AddString(L"니트/가디건(여성의류)");
        m_combo_tm3.AddString(L"자켓/점퍼(여성의류)");
        m_combo_tm3.AddString(L"패딩/코트(여성의류)");
        m_combo_tm3.AddString(L"면/캐주얼(여성의류)");
        m_combo_tm3.AddString(L"청바지(여성의류)");
        m_combo_tm3.AddString(L"반바지(여성의류)");
        m_combo_tm3.AddString(L"정장/팬츠(여성의류)");
        m_combo_tm3.AddString(L"트레이닝복(여성의류)");
        m_combo_tm3.AddString(L"언더웨어/속옷(여성의류)");
        m_combo_tm3.AddString(L"기타(빅사이즈)(여성의류)");
    }

    else if (1 == nCurTm1 && 22 == nCurTm2)
    {//여성의류 && 테마/이벤트 의류(여성의류)
        m_combo_tm3.AddString(L"우비/레인코트(여성의류)");
        m_combo_tm3.AddString(L"교복(여성의류)");
        m_combo_tm3.AddString(L"유니폼/작업복(여성의류)");
        m_combo_tm3.AddString(L"생활/전통(여성의류)");
        m_combo_tm3.AddString(L"웨딩 드레스(여성의류)");
        m_combo_tm3.AddString(L"코디(여성의류)");
        m_combo_tm3.AddString(L"옷장정리/급처(여성의류)");
        m_combo_tm3.AddString(L"기타(테마/이벤트 의류)(여성의류)");

        m_combo_tm_size.AddString(L"없음");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    if (1 == nCurTm1 && (12 >= nCurTm2 || 18 == nCurTm2 || 19 == nCurTm2 || 21 == nCurTm2))
    {
        m_combo_tm_size.AddString(L"44");
        m_combo_tm_size.AddString(L"55");
        m_combo_tm_size.AddString(L"66");
        m_combo_tm_size.AddString(L"77");
        m_combo_tm_size.AddString(L"88");
        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }
    else if (1 == nCurTm1 && (13 <= nCurTm2 && 17 >= nCurTm2))
    {
        m_combo_tm_size.AddString(L"44");
        m_combo_tm_size.AddString(L"55");
        m_combo_tm_size.AddString(L"66");
        m_combo_tm_size.AddString(L"77");
        m_combo_tm_size.AddString(L"88");
        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"23");
        m_combo_tm_size.AddString(L"24");
        m_combo_tm_size.AddString(L"25");
        m_combo_tm_size.AddString(L"26");
        m_combo_tm_size.AddString(L"27");
        m_combo_tm_size.AddString(L"28");
        m_combo_tm_size.AddString(L"29");
        m_combo_tm_size.AddString(L"30");
        m_combo_tm_size.AddString(L"31");
        m_combo_tm_size.AddString(L"32");
        m_combo_tm_size.AddString(L"33");
        m_combo_tm_size.AddString(L"34");
        m_combo_tm_size.AddString(L"35");
        m_combo_tm_size.AddString(L"36");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }


    // 남성의류
    else if (2 == nCurTm1 && 1 == nCurTm2)
    {//남성의류 && 긴팔 티셔츠(남성의류)
        m_combo_tm3.AddString(L"라운드넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"브이넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"카라넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"기타(긴팔 티셔츠(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 2 == nCurTm2)
    {//남성의류 && 반팔 티셔츠(남성의류)
        m_combo_tm3.AddString(L"라운드넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"카라넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"브이넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"민소매/ 나시 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"기타(반팔 티셔츠(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 3 == nCurTm2)
    {//남성의류 && 맨투맨/후드티(남성의류)
        m_combo_tm3.AddString(L"맨투맨 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"후드 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"후드 집업(남성의류)");
        m_combo_tm3.AddString(L"기타(맨투맨/후드티)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 4 == nCurTm2)
    {//남성의류 && 셔츠/남방(남성의류)
        m_combo_tm3.AddString(L"솔리드(단색)(남성의류)");
        m_combo_tm3.AddString(L"스트라이프 셔츠(남성의류)");
        m_combo_tm3.AddString(L"린넨/마 셔츠(남성의류)");
        m_combo_tm3.AddString(L"청/데님 셔츠(남성의류)");
        m_combo_tm3.AddString(L"헨리넥 셔츠(남성의류)");
        m_combo_tm3.AddString(L"체크 셔츠(남성의류)");
        m_combo_tm3.AddString(L"기타(셔츠/남방)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 5 == nCurTm2)
    {//남성의류 && 니트/스웨터(남성의류)
        m_combo_tm3.AddString(L"라운드넥 니트(남성의류)");
        m_combo_tm3.AddString(L"브이넥 니트(남성의류)");
        m_combo_tm3.AddString(L"집업 니트(남성의류)");
        m_combo_tm3.AddString(L"카라넥 니트(남성의류)");
        m_combo_tm3.AddString(L"폴라 니트(남성의류)");
        m_combo_tm3.AddString(L"기타(니트/스웨터)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 6 == nCurTm2)
    {//남성의류 && 가디건(남성의류)
        m_combo_tm3.AddString(L"브이넥 가디건(남성의류)");
        m_combo_tm3.AddString(L"라운드 가디건(남성의류)");
        m_combo_tm3.AddString(L"집업 가디건(남성의류)");
        m_combo_tm3.AddString(L"후드 가디건(남성의류)");
        m_combo_tm3.AddString(L"기타(가디건)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 7 == nCurTm2)
    {//남성의류 && 조끼/베스트(남성의류)
        m_combo_tm3.AddString(L"니트 조끼(남성의류)");
        m_combo_tm3.AddString(L"청/데님 조끼(남성의류)");
        m_combo_tm3.AddString(L"브이넥 조끼(남성의류)");
        m_combo_tm3.AddString(L"패딩 조끼(남성의류)");
        m_combo_tm3.AddString(L"기타(조끼/베스트)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 8 == nCurTm2)
    {//남성의류 && 점퍼/야상/패딩(남성의류)
        m_combo_tm3.AddString(L"바람막이(남성의류)");
        m_combo_tm3.AddString(L"패딩 점퍼(남성의류)");
        m_combo_tm3.AddString(L"다운 점퍼(남성의류)");
        m_combo_tm3.AddString(L"야구 점퍼(남성의류)");
        m_combo_tm3.AddString(L"블루종/항공점퍼(남성의류)");
        m_combo_tm3.AddString(L"야상/사파리(남성의류)");
        m_combo_tm3.AddString(L"기타(점퍼/야상/패딩)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 9 == nCurTm2)
    {//남성의류 && 자켓(남성의류)
        m_combo_tm3.AddString(L"캐주얼 자켓(남성의류)");
        m_combo_tm3.AddString(L"청/데님 자켓(남성의류)");
        m_combo_tm3.AddString(L"가죽 자켓(남성의류)");
        m_combo_tm3.AddString(L"차이나/노카라 자켓(남성의류)");
        m_combo_tm3.AddString(L"린넨/마 자켓(남성의류)");
        m_combo_tm3.AddString(L"후드/져지 자켓(남성의류)");
        m_combo_tm3.AddString(L"기타(자켓)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 10 == nCurTm2)
    {//남성의류 && 코트(남성의류)
        m_combo_tm3.AddString(L"모직 코트(남성의류)");
        m_combo_tm3.AddString(L"트렌치 코트(남성의류)");
        m_combo_tm3.AddString(L"하프 코트(남성의류)");
        m_combo_tm3.AddString(L"캐시미어 코트(남성의류)");
        m_combo_tm3.AddString(L"기타(코트)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 11 == nCurTm2)
    {//남성의류 && 청바지(긴)(남성의류)
        m_combo_tm3.AddString(L"일자 청바지(남성의류)");
        m_combo_tm3.AddString(L"스키니진(남성의류)");
        m_combo_tm3.AddString(L"빈티지/구제 청바지(남성의류)");
        m_combo_tm3.AddString(L"배기 청바지(남성의류)");
        m_combo_tm3.AddString(L"블랙/그레이진(남성의류)");
        m_combo_tm3.AddString(L"부츠컷 청바지(남성의류)");
        m_combo_tm3.AddString(L"기타(청바지(긴)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"23");
        m_combo_tm_size.AddString(L"24");
        m_combo_tm_size.AddString(L"25");
        m_combo_tm_size.AddString(L"26");
        m_combo_tm_size.AddString(L"27");
        m_combo_tm_size.AddString(L"28");
        m_combo_tm_size.AddString(L"29");
        m_combo_tm_size.AddString(L"30");
        m_combo_tm_size.AddString(L"31");
        m_combo_tm_size.AddString(L"32");
        m_combo_tm_size.AddString(L"33");
        m_combo_tm_size.AddString(L"34");
        m_combo_tm_size.AddString(L"35");
        m_combo_tm_size.AddString(L"36");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 12 == nCurTm2)
    {//남성의류 && 면/캐주얼 바지(긴)(남성의류)
        m_combo_tm3.AddString(L"면바지(남성의류)");
        m_combo_tm3.AddString(L"슬랙스(남성의류)");
        m_combo_tm3.AddString(L"배기/조거 바지(남성의류)");
        m_combo_tm3.AddString(L"기모 바지(남성의류)");
        m_combo_tm3.AddString(L"카고 바지(남성의류)");
        m_combo_tm3.AddString(L"기타(면/캐주얼 바지(긴))(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"23");
        m_combo_tm_size.AddString(L"24");
        m_combo_tm_size.AddString(L"25");
        m_combo_tm_size.AddString(L"26");
        m_combo_tm_size.AddString(L"27");
        m_combo_tm_size.AddString(L"28");
        m_combo_tm_size.AddString(L"29");
        m_combo_tm_size.AddString(L"30");
        m_combo_tm_size.AddString(L"31");
        m_combo_tm_size.AddString(L"32");
        m_combo_tm_size.AddString(L"33");
        m_combo_tm_size.AddString(L"34");
        m_combo_tm_size.AddString(L"35");
        m_combo_tm_size.AddString(L"36");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 13 == nCurTm2)
    {//남성의류 && 반바지/7~9부(남성의류)
        m_combo_tm3.AddString(L"면 반바지(남성의류)");
        m_combo_tm3.AddString(L"청/데님 반바지(남성의류)");
        m_combo_tm3.AddString(L"밴딩 반바지(남성의류)");
        m_combo_tm3.AddString(L"스포츠 반바지(남성의류)");
        m_combo_tm3.AddString(L"기타(반바지/7~9부)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"23");
        m_combo_tm_size.AddString(L"24");
        m_combo_tm_size.AddString(L"25");
        m_combo_tm_size.AddString(L"26");
        m_combo_tm_size.AddString(L"27");
        m_combo_tm_size.AddString(L"28");
        m_combo_tm_size.AddString(L"29");
        m_combo_tm_size.AddString(L"30");
        m_combo_tm_size.AddString(L"31");
        m_combo_tm_size.AddString(L"32");
        m_combo_tm_size.AddString(L"33");
        m_combo_tm_size.AddString(L"34");
        m_combo_tm_size.AddString(L"35");
        m_combo_tm_size.AddString(L"36");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 14 == nCurTm2)
    {//남성의류 && 비즈니스 정장(남성의류)
        m_combo_tm3.AddString(L"정장 자켓(남성의류)");
        m_combo_tm3.AddString(L"정장 바지(남성의류)");
        m_combo_tm3.AddString(L"정장 베스트(남성의류)");
        m_combo_tm3.AddString(L"기타(비즈니스 정장)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 15 == nCurTm2)
    {//남성의류 && 트레이닝(남성의류)
        m_combo_tm3.AddString(L"트레이닝 상의(남성의류)");
        m_combo_tm3.AddString(L"트레이닝 하의(남성의류)");
        m_combo_tm3.AddString(L"트레이닝 세트(남성의류)");
        m_combo_tm3.AddString(L"기타(트레이닝)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 16 == nCurTm2)
    {//남성의류 && 언더웨어/속옷(남성의류)
        m_combo_tm3.AddString(L"런닝(남성의류)");
        m_combo_tm3.AddString(L"드로즈/삼각(남성의류)");
        m_combo_tm3.AddString(L"트렁크(남성의류)");
        m_combo_tm3.AddString(L"잠옷/이지웨어(남성의류)");
        m_combo_tm3.AddString(L"런닝/팬티(남성의류)");
        m_combo_tm3.AddString(L"기능성/히트텍(남성의류)");
        m_combo_tm3.AddString(L"기타(언더웨어/속옷)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 17 == nCurTm2)
    {//남성의류 && 빅사이즈(남성의류)
        m_combo_tm3.AddString(L"긴팔 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"반팔 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"맨투맨 후드(남성의류)");
        m_combo_tm3.AddString(L"니트/가디건(남성의류)");
        m_combo_tm3.AddString(L"자켓/점퍼(남성의류)");
        m_combo_tm3.AddString(L"패딩/코트(남성의류)");
        m_combo_tm3.AddString(L"면/캐주얼(남성의류)");
        m_combo_tm3.AddString(L"청바지(남성의류)");
        m_combo_tm3.AddString(L"정장/팬츠(남성의류)");
        m_combo_tm3.AddString(L"트레이닝복(남성의류)");
        m_combo_tm3.AddString(L"기타(빅사이즈)(남성의류)");

        m_combo_tm_size.AddString(L"없음");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 18 == nCurTm2)
    {//남성의류 && 테마/이벤트 의류(남성의류)
        m_combo_tm3.AddString(L"우비/레인코트(남성의류)");
        m_combo_tm3.AddString(L"교복(남성의류)");
        m_combo_tm3.AddString(L"유니폼/작업복(남성의류)");
        m_combo_tm3.AddString(L"생활/전통(남성의류)");
        m_combo_tm3.AddString(L"군복(남성의류)");
        m_combo_tm3.AddString(L"기타(테마/이벤트 의류)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }


    // 패션잡화
    else if (3 == nCurTm1 && 1 == nCurTm2)
    {//패션잡화 && 여성가방(패션잡화)
        m_combo_tm3.AddString(L"숄더백(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"크로스백(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"클러치백(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"토트백(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"백팩(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"파우치(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"미니백(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"기타(여성가방(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 2 == nCurTm2)
    {//패션잡화 && 남성가방(패션잡화)
        m_combo_tm3.AddString(L"백팩(남성가방(패션잡화))");
        m_combo_tm3.AddString(L"크로스백(남성가방(패션잡화))");
        m_combo_tm3.AddString(L"숄더백(남성가방(패션잡화))");
        m_combo_tm3.AddString(L"비즈니스가방(남성가방(패션잡화))");
        m_combo_tm3.AddString(L"클러치백(남성가방(패션잡화))");
        m_combo_tm3.AddString(L"기타(남성가방(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 3 == nCurTm2)
    {//패션잡화 && 여행용가방/소품(패션잡화)
        m_combo_tm3.AddString(L"하드 캐리어(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"소프트 캐리어(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"이민/유학용(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"여행용 백팩(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"여행용 크로스(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"여행용 파우치(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"기타(여행용가방/소품(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 4 == nCurTm2)
    {//패션잡화 && 운동화/캐주얼화(패션잡화)
        m_combo_tm3.AddString(L"런닝화/워킹화(운동화/캐주얼화(패션잡화))");
        m_combo_tm3.AddString(L"농구화(운동화/캐주얼화(패션잡화))");
        m_combo_tm3.AddString(L"캐주얼화(운동화/캐주얼화(패션잡화))");
        m_combo_tm3.AddString(L"기타(운동화/캐주얼화(패션잡화)");

        m_combo_tm_size.AddString(L"200.0");
        m_combo_tm_size.AddString(L"205.0");
        m_combo_tm_size.AddString(L"210.0");
        m_combo_tm_size.AddString(L"215.0");
        m_combo_tm_size.AddString(L"220.0");
        m_combo_tm_size.AddString(L"225.0");
        m_combo_tm_size.AddString(L"230.0");
        m_combo_tm_size.AddString(L"235.0");
        m_combo_tm_size.AddString(L"240.0");
        m_combo_tm_size.AddString(L"245.0");
        m_combo_tm_size.AddString(L"250.0");
        m_combo_tm_size.AddString(L"255.0");
        m_combo_tm_size.AddString(L"260.0");
        m_combo_tm_size.AddString(L"265.0");
        m_combo_tm_size.AddString(L"270.0");
        m_combo_tm_size.AddString(L"275.0");
        m_combo_tm_size.AddString(L"280.0");
        m_combo_tm_size.AddString(L"285.0");
        m_combo_tm_size.AddString(L"290.0");
        m_combo_tm_size.AddString(L"295.0");
        m_combo_tm_size.AddString(L"300.0");
        m_combo_tm_size.AddString(L"305.0");
        m_combo_tm_size.AddString(L"310.0");
        m_combo_tm_size.AddString(L"315.0");
        m_combo_tm_size.AddString(L"320.0");
        m_combo_tm_size.AddString(L"325.0");
        m_combo_tm_size.AddString(L"330.0");
        m_combo_tm_size.AddString(L"335.0");
        m_combo_tm_size.AddString(L"340.0");
        m_combo_tm_size.AddString(L"345.0");
        m_combo_tm_size.AddString(L"350.0");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 5 == nCurTm2)
    {//패션잡화 && 여성화(패션잡화)
        m_combo_tm3.AddString(L"단화/플랫슈즈(여성화(패션잡화))");
        m_combo_tm3.AddString(L"하이힐(여성화(패션잡화))");
        m_combo_tm3.AddString(L"미들굽 펌프스(여성화(패션잡화))");
        m_combo_tm3.AddString(L"샌들/슬리퍼(여성화(패션잡화))");
        m_combo_tm3.AddString(L"플랫폼 슈즈(여성화(패션잡화))");
        m_combo_tm3.AddString(L"워커(여성화(패션잡화))");
        m_combo_tm3.AddString(L"앵클부츠(여성화(패션잡화))");
        m_combo_tm3.AddString(L"롱/미들 부츠(여성화(패션잡화))");
        m_combo_tm3.AddString(L"기타(여성화(패션잡화)");

        m_combo_tm_size.AddString(L"200.0");
        m_combo_tm_size.AddString(L"205.0");
        m_combo_tm_size.AddString(L"210.0");
        m_combo_tm_size.AddString(L"215.0");
        m_combo_tm_size.AddString(L"220.0");
        m_combo_tm_size.AddString(L"225.0");
        m_combo_tm_size.AddString(L"230.0");
        m_combo_tm_size.AddString(L"235.0");
        m_combo_tm_size.AddString(L"240.0");
        m_combo_tm_size.AddString(L"245.0");
        m_combo_tm_size.AddString(L"250.0");
        m_combo_tm_size.AddString(L"255.0");
        m_combo_tm_size.AddString(L"260.0");
        m_combo_tm_size.AddString(L"265.0");
        m_combo_tm_size.AddString(L"270.0");
        m_combo_tm_size.AddString(L"275.0");
        m_combo_tm_size.AddString(L"280.0");
        m_combo_tm_size.AddString(L"285.0");
        m_combo_tm_size.AddString(L"290.0");
        m_combo_tm_size.AddString(L"295.0");
        m_combo_tm_size.AddString(L"300.0");
        m_combo_tm_size.AddString(L"305.0");
        m_combo_tm_size.AddString(L"310.0");
        m_combo_tm_size.AddString(L"315.0");
        m_combo_tm_size.AddString(L"320.0");
        m_combo_tm_size.AddString(L"325.0");
        m_combo_tm_size.AddString(L"330.0");
        m_combo_tm_size.AddString(L"335.0");
        m_combo_tm_size.AddString(L"340.0");
        m_combo_tm_size.AddString(L"345.0");
        m_combo_tm_size.AddString(L"350.0");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 6 == nCurTm2)
    {//패션잡화 && 남성화(패션잡화)
        m_combo_tm3.AddString(L"정장구두(남성화(패션잡화))");
        m_combo_tm3.AddString(L"수제화(남성화(패션잡화))");
        m_combo_tm3.AddString(L"샌들/슬리퍼(남성화(패션잡화))");
        m_combo_tm3.AddString(L"로퍼/슬립온(남성화(패션잡화))");
        m_combo_tm3.AddString(L"워커(남성화(패션잡화))");
        m_combo_tm3.AddString(L"부츠(남성화(패션잡화))");
        m_combo_tm3.AddString(L"기타(남성화(패션잡화)");

        m_combo_tm_size.AddString(L"200.0");
        m_combo_tm_size.AddString(L"205.0");
        m_combo_tm_size.AddString(L"210.0");
        m_combo_tm_size.AddString(L"215.0");
        m_combo_tm_size.AddString(L"220.0");
        m_combo_tm_size.AddString(L"225.0");
        m_combo_tm_size.AddString(L"230.0");
        m_combo_tm_size.AddString(L"235.0");
        m_combo_tm_size.AddString(L"240.0");
        m_combo_tm_size.AddString(L"245.0");
        m_combo_tm_size.AddString(L"250.0");
        m_combo_tm_size.AddString(L"255.0");
        m_combo_tm_size.AddString(L"260.0");
        m_combo_tm_size.AddString(L"265.0");
        m_combo_tm_size.AddString(L"270.0");
        m_combo_tm_size.AddString(L"275.0");
        m_combo_tm_size.AddString(L"280.0");
        m_combo_tm_size.AddString(L"285.0");
        m_combo_tm_size.AddString(L"290.0");
        m_combo_tm_size.AddString(L"295.0");
        m_combo_tm_size.AddString(L"300.0");
        m_combo_tm_size.AddString(L"305.0");
        m_combo_tm_size.AddString(L"310.0");
        m_combo_tm_size.AddString(L"315.0");
        m_combo_tm_size.AddString(L"320.0");
        m_combo_tm_size.AddString(L"325.0");
        m_combo_tm_size.AddString(L"330.0");
        m_combo_tm_size.AddString(L"335.0");
        m_combo_tm_size.AddString(L"340.0");
        m_combo_tm_size.AddString(L"345.0");
        m_combo_tm_size.AddString(L"350.0");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 7 == nCurTm2)
    {//패션잡화 && 지갑(패션잡화)
        m_combo_tm3.AddString(L"여성 장지갑(지갑(패션잡화))");
        m_combo_tm3.AddString(L"여성 중/반지갑(지갑(패션잡화))");
        m_combo_tm3.AddString(L"남자 장지갑(지갑(패션잡화))");
        m_combo_tm3.AddString(L"남자 중/반지갑(지갑(패션잡화))");
        m_combo_tm3.AddString(L"카드/명합 지갑(지갑(패션잡화))");
        m_combo_tm3.AddString(L"기타(지갑)(지갑(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 8 == nCurTm2)
    {//패션잡화 && 모자(패션잡화)
        m_combo_tm3.AddString(L"야구모자/군모(모자(패션잡화))");
        m_combo_tm3.AddString(L"스냅백(모자(패션잡화))");
        m_combo_tm3.AddString(L"패션모자(모자(패션잡화))");
        m_combo_tm3.AddString(L"왕골/바캉스모자(모자(패션잡화))");
        m_combo_tm3.AddString(L"비니(모자(패션잡화))");
        m_combo_tm3.AddString(L"털/방울(모자(패션잡화))");
        m_combo_tm3.AddString(L"기타(모자)(모자(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 9 == nCurTm2)
    {//패션잡화 && 안경/선글라스(패션잡화)
        m_combo_tm3.AddString(L"안경(뿔테)(안경/선글라스(패션잡화))");
        m_combo_tm3.AddString(L"안경(금속테)(안경/선글라스(패션잡화))");
        m_combo_tm3.AddString(L"선글라스(안경/선글라스(패션잡화))");
        m_combo_tm3.AddString(L"기타(안경/선글라스(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 10 == nCurTm2)
    {//패션잡화 && 주얼리/액세서리(패션잡화)
        m_combo_tm3.AddString(L"반지(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"귀걸이(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"목걸이(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"팔찌/발찌(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"귀금속/보석(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"커플용 주얼리(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"남성용 주얼리(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"헤어 액세서리(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"기타(안경/주얼리/액세서리(패션잡화)");

        m_combo_tm_size.AddString(L"1");
        m_combo_tm_size.AddString(L"2");
        m_combo_tm_size.AddString(L"3");
        m_combo_tm_size.AddString(L"4");
        m_combo_tm_size.AddString(L"5");
        m_combo_tm_size.AddString(L"6");
        m_combo_tm_size.AddString(L"7");
        m_combo_tm_size.AddString(L"8");
        m_combo_tm_size.AddString(L"9");
        m_combo_tm_size.AddString(L"10");
        m_combo_tm_size.AddString(L"11");
        m_combo_tm_size.AddString(L"12");
        m_combo_tm_size.AddString(L"13");
        m_combo_tm_size.AddString(L"14");
        m_combo_tm_size.AddString(L"15");
        m_combo_tm_size.AddString(L"17");
        m_combo_tm_size.AddString(L"18");
        m_combo_tm_size.AddString(L"19");
        m_combo_tm_size.AddString(L"20");
        m_combo_tm_size.AddString(L"21");
        m_combo_tm_size.AddString(L"22");
        m_combo_tm_size.AddString(L"23");
        m_combo_tm_size.AddString(L"24");
        m_combo_tm_size.AddString(L"25");
        m_combo_tm_size.AddString(L"26");
        m_combo_tm_size.AddString(L"27");
        m_combo_tm_size.AddString(L"28");
        m_combo_tm_size.AddString(L"29");
        m_combo_tm_size.AddString(L"30");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 11 == nCurTm2)
    {//패션잡화 && 시계(패션잡화)
        m_combo_tm3.AddString(L"메탈시계(시계(패션잡화))");
        m_combo_tm3.AddString(L"가죽시계(시계(패션잡화)");
        m_combo_tm3.AddString(L"스포츠/방수시계(시계(패션잡화)");
        m_combo_tm3.AddString(L"뱅글/팔찌형시계(시계(패션잡화)");
        m_combo_tm3.AddString(L"젤리/우레탄시계(시계(패션잡화)");
        m_combo_tm3.AddString(L"기타(시계(패션잡화)");

        m_combo_tm_size.AddString(L"남성");
        m_combo_tm_size.AddString(L"여성");
        m_combo_tm_size.AddString(L"공용");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 12 == nCurTm2)
    {//패션잡화 && 벨트/장갑/스타킹/기타(패션잡화)
        m_combo_tm3.AddString(L"여성 벨트(패션잡화)");
        m_combo_tm3.AddString(L"남성 벨트(패션잡화)");
        m_combo_tm3.AddString(L"스카프/머플러(패션잡화)");
        m_combo_tm3.AddString(L"넥타이(패션잡화)");
        m_combo_tm3.AddString(L"장갑(패션잡화)");
        m_combo_tm3.AddString(L"양말/스타킹(패션잡화)");
        m_combo_tm3.AddString(L"우산/양산(패션잡화)");
        m_combo_tm3.AddString(L"기타(잡화)(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }


    // 유아동/출산
    else if (4 == nCurTm1 && 1 == nCurTm2)
    {//유아동/출산 && 베이비의류(0-2세)(유아동/출산)
        m_combo_tm3.AddString(L"유아내의/속옷(유아동/출산)");
        m_combo_tm3.AddString(L"유아상의(유아동/출산)");
        m_combo_tm3.AddString(L"유아하의(유아동/출산)");
        m_combo_tm3.AddString(L"우주복/슈트(유아동/출산)");
        m_combo_tm3.AddString(L"자켓/점퍼(유아동/출산)");
        m_combo_tm3.AddString(L"원피스(유아동/출산)");
        m_combo_tm3.AddString(L"정장/드레스(유아동/출산)");
        m_combo_tm3.AddString(L"베이비수영복(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"50");
        m_combo_tm_size.AddString(L"60");
        m_combo_tm_size.AddString(L"70");
        m_combo_tm_size.AddString(L"80");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"3M");
        m_combo_tm_size.AddString(L"6M");
        m_combo_tm_size.AddString(L"9M");
        m_combo_tm_size.AddString(L"12M");
        m_combo_tm_size.AddString(L"18M");
        m_combo_tm_size.AddString(L"24M");
        m_combo_tm_size.AddString(L"2T");
        m_combo_tm_size.AddString(L"3T");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 2 == nCurTm2)
    {//유아동/출산 && 여아의류(3-6세)(유아동/출산)
        m_combo_tm3.AddString(L"티셔츠(유아동/출산)");
        m_combo_tm3.AddString(L"팬츠(유아동/출산)");
        m_combo_tm3.AddString(L"원피스(유아동/출산)");
        m_combo_tm3.AddString(L"블라우스(유아동/출산)");
        m_combo_tm3.AddString(L"니트/스웨터(유아동/출산)");
        m_combo_tm3.AddString(L"스커트/치마(유아동/출산)");
        m_combo_tm3.AddString(L"가디건/조끼(유아동/출산)");
        m_combo_tm3.AddString(L"자켓/점퍼(유아동/출산)");
        m_combo_tm3.AddString(L"코트/정장(유아동/출산)");
        m_combo_tm3.AddString(L"상하복세트(유아동/출산)");
        m_combo_tm3.AddString(L"속옷/잠옷(유아동/출산)");
        m_combo_tm3.AddString(L"수영복(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"120");
        m_combo_tm_size.AddString(L"130");
        m_combo_tm_size.AddString(L"140");
        m_combo_tm_size.AddString(L"3호");
        m_combo_tm_size.AddString(L"5호");
        m_combo_tm_size.AddString(L"7호");
        m_combo_tm_size.AddString(L"9호");
        m_combo_tm_size.AddString(L"11호");
        m_combo_tm_size.AddString(L"13호");
        m_combo_tm_size.AddString(L"15호");
        m_combo_tm_size.AddString(L"3T");
        m_combo_tm_size.AddString(L"4T");
        m_combo_tm_size.AddString(L"5T");
        m_combo_tm_size.AddString(L"6T");
        m_combo_tm_size.AddString(L"7T");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 3 == nCurTm2)
    {//유아동/출산 && 남아의류(3-6세)(유아동/출산)
        m_combo_tm3.AddString(L"티셔츠(유아동/출산)");
        m_combo_tm3.AddString(L"팬츠(유아동/출산)");
        m_combo_tm3.AddString(L"셔츠/남방(유아동/출산)");
        m_combo_tm3.AddString(L"니트/스웨터(유아동/출산)");
        m_combo_tm3.AddString(L"가디건/조끼(유아동/출산)");
        m_combo_tm3.AddString(L"자켓/점퍼(유아동/출산)");
        m_combo_tm3.AddString(L"코트/정장(유아동/출산)");
        m_combo_tm3.AddString(L"상하복세트(유아동/출산)");
        m_combo_tm3.AddString(L"속옷/잠옷(유아동/출산)");
        m_combo_tm3.AddString(L"수영복(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"120");
        m_combo_tm_size.AddString(L"130");
        m_combo_tm_size.AddString(L"140");
        m_combo_tm_size.AddString(L"3호");
        m_combo_tm_size.AddString(L"5호");
        m_combo_tm_size.AddString(L"7호");
        m_combo_tm_size.AddString(L"9호");
        m_combo_tm_size.AddString(L"11호");
        m_combo_tm_size.AddString(L"13호");
        m_combo_tm_size.AddString(L"15호");
        m_combo_tm_size.AddString(L"3T");
        m_combo_tm_size.AddString(L"4T");
        m_combo_tm_size.AddString(L"5T");
        m_combo_tm_size.AddString(L"6T");
        m_combo_tm_size.AddString(L"7T");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 4 == nCurTm2)
    {//유아동/출산 && 여주니어의류(7세~)(유아동/출산)
        m_combo_tm3.AddString(L"티셔츠(유아동/출산)");
        m_combo_tm3.AddString(L"팬츠(유아동/출산)");
        m_combo_tm3.AddString(L"원피스(유아동/출산)");
        m_combo_tm3.AddString(L"블라우스(유아동/출산)");
        m_combo_tm3.AddString(L"니트/스웨터(유아동/출산)");
        m_combo_tm3.AddString(L"스커트/치마(유아동/출산)");
        m_combo_tm3.AddString(L"가디건/조끼(유아동/출산)");
        m_combo_tm3.AddString(L"자켓/점퍼(유아동/출산)");
        m_combo_tm3.AddString(L"코트/정장(유아동/출산)");
        m_combo_tm3.AddString(L"상하복세트(유아동/출산)");
        m_combo_tm3.AddString(L"속옷/잠옷(유아동/출산)");
        m_combo_tm3.AddString(L"수영복(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"140");
        m_combo_tm_size.AddString(L"150");
        m_combo_tm_size.AddString(L"160");
        m_combo_tm_size.AddString(L"9호");
        m_combo_tm_size.AddString(L"11호");
        m_combo_tm_size.AddString(L"13호");
        m_combo_tm_size.AddString(L"15호");
        m_combo_tm_size.AddString(L"17호");
        m_combo_tm_size.AddString(L"3T");
        m_combo_tm_size.AddString(L"4T");
        m_combo_tm_size.AddString(L"5T");
        m_combo_tm_size.AddString(L"6T");
        m_combo_tm_size.AddString(L"7T");
        m_combo_tm_size.AddString(L"8T");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 5 == nCurTm2)
    {//유아동/출산 && 남주니어의류(7세~)(유아동/출산)
        m_combo_tm3.AddString(L"티셔츠(유아동/출산)");
        m_combo_tm3.AddString(L"팬츠(유아동/출산)");
        m_combo_tm3.AddString(L"셔츠/남방(유아동/출산)");
        m_combo_tm3.AddString(L"니트/스웨터(유아동/출산)");
        m_combo_tm3.AddString(L"가디건/조끼(유아동/출산)");
        m_combo_tm3.AddString(L"자켓/점퍼(유아동/출산)");
        m_combo_tm3.AddString(L"코트/정장(유아동/출산)");
        m_combo_tm3.AddString(L"상하복세트(유아동/출산)");
        m_combo_tm3.AddString(L"속옷/잠옷(유아동/출산)");
        m_combo_tm3.AddString(L"수영복(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"140");
        m_combo_tm_size.AddString(L"150");
        m_combo_tm_size.AddString(L"160");
        m_combo_tm_size.AddString(L"9호");
        m_combo_tm_size.AddString(L"11호");
        m_combo_tm_size.AddString(L"13호");
        m_combo_tm_size.AddString(L"15호");
        m_combo_tm_size.AddString(L"17호");
        m_combo_tm_size.AddString(L"3T");
        m_combo_tm_size.AddString(L"4T");
        m_combo_tm_size.AddString(L"5T");
        m_combo_tm_size.AddString(L"6T");
        m_combo_tm_size.AddString(L"7T");
        m_combo_tm_size.AddString(L"8T");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 6 == nCurTm2)
    {//유아동/출산 && 유아동신발/잡화(유아동/출산)
        m_combo_tm3.AddString(L"신발(유아동/출산)");
        m_combo_tm3.AddString(L"가방(유아동/출산)");
        m_combo_tm3.AddString(L"모자(유아동/출산)");
        m_combo_tm3.AddString(L"양말(유아동/출산)");
        m_combo_tm3.AddString(L"시계(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"80.0");
        m_combo_tm_size.AddString(L"85.0");
        m_combo_tm_size.AddString(L"90.0");
        m_combo_tm_size.AddString(L"95.0");
        m_combo_tm_size.AddString(L"100.0");
        m_combo_tm_size.AddString(L"105.0");
        m_combo_tm_size.AddString(L"110.0");
        m_combo_tm_size.AddString(L"115.0");
        m_combo_tm_size.AddString(L"120.0");
        m_combo_tm_size.AddString(L"125.0");
        m_combo_tm_size.AddString(L"130.0");
        m_combo_tm_size.AddString(L"135.0");
        m_combo_tm_size.AddString(L"140.0");
        m_combo_tm_size.AddString(L"145.0");
        m_combo_tm_size.AddString(L"150.0");
        m_combo_tm_size.AddString(L"155.0");
        m_combo_tm_size.AddString(L"160.0");
        m_combo_tm_size.AddString(L"165.0");
        m_combo_tm_size.AddString(L"170.0");
        m_combo_tm_size.AddString(L"175.0");
        m_combo_tm_size.AddString(L"180.0");
        m_combo_tm_size.AddString(L"185.0");
        m_combo_tm_size.AddString(L"190.0");
        m_combo_tm_size.AddString(L"195.0");
        m_combo_tm_size.AddString(L"200.0");
        m_combo_tm_size.AddString(L"205.0");
        m_combo_tm_size.AddString(L"210.0");
        m_combo_tm_size.AddString(L"215.0");
        m_combo_tm_size.AddString(L"220.0");
        m_combo_tm_size.AddString(L"225.0");
        m_combo_tm_size.AddString(L"230.0");
        m_combo_tm_size.AddString(L"235.0");
        m_combo_tm_size.AddString(L"240.0");
        m_combo_tm_size.AddString(L"245.0");
        m_combo_tm_size.AddString(L"250.0");
        m_combo_tm_size.AddString(L"255.0");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 7 == nCurTm2)
    {//유아동/출산 && 유아동용품(유아동/출산)
        m_combo_tm3.AddString(L"유모차(유아동/출산)");
        m_combo_tm3.AddString(L"카시트(유아동/출산)");
        m_combo_tm3.AddString(L"아기띠(유아동/출산)");
        m_combo_tm3.AddString(L"보행기/쏘서(유아동/출산)");
        m_combo_tm3.AddString(L"가구/침대(유아동/출산)");
        m_combo_tm3.AddString(L"스킨케어(유아동/출산)");
        m_combo_tm3.AddString(L"목욕/구강용품(유아동/출산)");
        m_combo_tm3.AddString(L"세탁/위생용품(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"0~6개월");
        m_combo_tm_size.AddString(L"7~12개월");
        m_combo_tm_size.AddString(L"13~24개월");
        m_combo_tm_size.AddString(L"25~36개월");
        m_combo_tm_size.AddString(L"37~48개월");
        m_combo_tm_size.AddString(L"5~7세");
        m_combo_tm_size.AddString(L"8~10세");
        m_combo_tm_size.AddString(L"11~13세");
        m_combo_tm_size.AddString(L"전체이용가");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 8 == nCurTm2)
    {//유아동/출산 && 출산/임부용품(유아동/출산)
        m_combo_tm3.AddString(L"겉싸개/속싸개(유아동/출산)");
        m_combo_tm3.AddString(L"배냇저고리(유아동/출산)");
        m_combo_tm3.AddString(L"딸랑이/모빌(유아동/출산)");
        m_combo_tm3.AddString(L"손발싸개(유아동/출산)");
        m_combo_tm3.AddString(L"이불/침구(유아동/출산)");
        m_combo_tm3.AddString(L"임부의류/속옷(유아동/출산)");
        m_combo_tm3.AddString(L"임부스킨케어(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 9 == nCurTm2)
    {//유아동/출산 && 교육/완구/인형(유아동/출산)
        m_combo_tm3.AddString(L"신생아완구(유아동/출산)");
        m_combo_tm3.AddString(L"교육완구(유아동/출산)");
        m_combo_tm3.AddString(L"도서/CD(유아동/출산)");
        m_combo_tm3.AddString(L"인형(유아용)(유아동/출산)");
        m_combo_tm3.AddString(L"자전거(유아동/출산)");
        m_combo_tm3.AddString(L"볼텐트/놀이터(유아동/출산)");
        m_combo_tm3.AddString(L"퍼즐/블록(유아동/출산)");
        m_combo_tm3.AddString(L"물놀이용품(유아동/출산)");
        m_combo_tm3.AddString(L"스포츠완구(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"0~6개월");
        m_combo_tm_size.AddString(L"7~12개월");
        m_combo_tm_size.AddString(L"13~24개월");
        m_combo_tm_size.AddString(L"25~36개월");
        m_combo_tm_size.AddString(L"37~48개월");
        m_combo_tm_size.AddString(L"5~7세");
        m_combo_tm_size.AddString(L"8~10세");
        m_combo_tm_size.AddString(L"11~13세");
        m_combo_tm_size.AddString(L"전체이용가");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 10 == nCurTm2)
    {//유아동/출산 && 기저귀/수유/이유식(유아동/출산)
        m_combo_tm3.AddString(L"기저귀(유아동/출산)");
        m_combo_tm3.AddString(L"물티슈(유아동/출산)");
        m_combo_tm3.AddString(L"분유수유용품(유아동/출산)");
        m_combo_tm3.AddString(L"모유수유용품(유아동/출산)");
        m_combo_tm3.AddString(L"이유식용품(유아동/출산)");
        m_combo_tm3.AddString(L"젖병세정용품(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }
}



void CETMacroDlg::OnLvnItemchangedListSellitem(NMHDR *pNMHDR, LRESULT *pResult)
{
    LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);

    for (int i = 0; i < m_list_sellitem.GetItemCount(); i++)
    {
        if (m_list_sellitem.GetItemState(i, LVIS_SELECTED) & LVIS_SELECTED)
        {
            // 헬로마켓 등록 유무
            BOOL bHelloMarketAccept = theApp.m_sellItems[i].bHelloMarketAccept;
            m_check_site0.SetCheck(bHelloMarketAccept);

            // 번개장터 등록 유무
            BOOL bThunderMarketAccept = theApp.m_sellItems[i].bThunderMarketAccept;
            m_check_site1.SetCheck(bThunderMarketAccept);

            // 네이버카페 등록 유무
            BOOL bNaverCafeAccept = theApp.m_sellItems[i].bNaverCafeAccept;
            m_check_site2.SetCheck(bNaverCafeAccept);

            // 네이버밴드 등록 유무
            BOOL bNaverBandAccept = theApp.m_sellItems[i].bNaverBandAccept;
            m_check_site3.SetCheck(bNaverBandAccept);

            // 카카오스토리 등록 유무
            BOOL bKakaoStoryAccept = theApp.m_sellItems[i].bKakaoStoryAccept;
            m_check_site4.SetCheck(bKakaoStoryAccept);

            Invalidate();

            SetItemChangeComboHm1(i);

            SetItemChangeComboTm1(i);
            SetItemChangeComboTm2(i);

            SetItemChangeComboNc(i);
        }
    }




    *pResult = 0;
}

CString CETMacroDlg::GetApplicationDirectory()
{
    wchar_t chThisPath[256];
    GetCurrentDirectoryW(256, chThisPath);

    CString path = chThisPath;

    return path;
}

void CETMacroDlg::OnBnClickedButtonExportExcel()
{
    // 엑셀클래스 선언 (FALSE: 처리 과정을 화면에 보이지 않는다)
    CXLEzAutomation dataexcel(FALSE);

    wchar_t chThisPath[256];
    GetCurrentDirectoryW(256, chThisPath);

    CString strThisPath;
    strThisPath.Format(L"%s\\Category_All.xls", chThisPath);
    ///// 위에서 얻은 엑셀파일경로를 바탕으로 파일을 연다.
    BOOL bRet = dataexcel.OpenExcelFile(strThisPath);

    int nCur = 0;
    CString strText;
    // 번개장터 카테고리1 저장
    nCur = m_combo_tm1.GetCurSel();
    strText.Format(L"%d", nCur);
    dataexcel.SetCellValue(1, 2, strText);

    // 번개장터 카테고리2 저장
    nCur = m_combo_tm2.GetCurSel();
    strText.Format(L"%d", nCur);
    dataexcel.SetCellValue(2, 2, strText);

    // 번개장터 카테고리3 저장
    nCur = m_combo_tm3.GetCurSel();
    strText.Format(L"%d", nCur);
    dataexcel.SetCellValue(3, 2, strText);


    // 헬로마켓 카테고리1 저장
    nCur = m_combo_hm1.GetCurSel();
    strText.Format(L"%d", nCur);
    dataexcel.SetCellValue(4, 2, strText);

    // 헬로마켓 카테고리2 저장
    nCur = m_combo_hm2.GetCurSel();
    strText.Format(L"%d", nCur);
    dataexcel.SetCellValue(5, 2, strText);

    ///////////////////////////////////////////////////////////
    // 중고나라 카테고리 저장
    nCur = m_combo_nc.GetCurSel();
    if (nCur == 1)
    {// 남성상의
        nCur = 89;
        strText.Format(L"%d", nCur);
        dataexcel.SetCellValue(6, 2, strText);
    }
    else if (nCur == 2)
    {// 남성하의
        nCur = 90;
        strText.Format(L"%d", nCur);
        dataexcel.SetCellValue(6, 2, strText);
    }
    else if (nCur == 3)
    {// 여성상의
        nCur = 82;
        strText.Format(L"%d", nCur);
        dataexcel.SetCellValue(6, 2, strText);
    }
    else
    {// 여성하의
        nCur = 83;
        strText.Format(L"%d", nCur);
        dataexcel.SetCellValue(6, 2, strText);
    }


    // 번개장터 상세정보(사이즈) 저장
    nCur = m_combo_tm_size.GetCurSel();
    strText.Format(L"%d", nCur);
    dataexcel.SetCellValue(7, 2, strText);

    // 번개장터 상태 저장
    nCur = m_combo_tm_quality.GetCurSel();
    strText.Format(L"%d", nCur);
    dataexcel.SetCellValue(8, 2, strText);

    // 택배비포함
    nCur = m_check_deliverycost.GetCheck();
    strText.Format(L"%d", nCur);
    dataexcel.SetCellValue(9, 2, strText);

    // 택배비포함
    nCur = m_check_exchange.GetCheck();
    strText.Format(L"%d", nCur);
    dataexcel.SetCellValue(10, 2, strText);

    ////////////////////////////////////////////////////////

    dataexcel.ReleaseExcel(); ///// 열었던 파일을 다 사용한 후 닫는다.
}


void CETMacroDlg::SetItemChangeComboHm1(int nCur)
{
    m_combo_hm2.ResetContent();
    m_combo_hm2.EnableWindow(TRUE);

    int idxHm1 = theApp.m_sellItems[nCur].nHelloMarketFirstComboIndex;
    int idxHm2 = theApp.m_sellItems[nCur].nHelloMarketSecondComboIndex;
    m_combo_hm1.SetCurSel(idxHm1);

    if (1 == idxHm1)
    {//모터사이클,용품
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"125cc 이하");
        m_combo_hm2.AddString(L"125cc 초과");
        m_combo_hm2.AddString(L"500cc 초과");

        m_combo_hm2.SetCurSel(0);
    }
    else if (2 == idxHm1)
    {//유아동,완구
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"신생아, 유아의류 (유아동, 완구)");
        m_combo_hm2.AddString(L"아동의류 (유아동, 완구)");
        m_combo_hm2.AddString(L"유아동잡화 (유아동, 완구)");
        m_combo_hm2.AddString(L"유아동생활용품 (유아동, 완구)");
        m_combo_hm2.AddString(L"완구, 인형 (유아동, 완구)");
        m_combo_hm2.AddString(L"임부복, 출산용품 (유아동, 완구)");

        m_combo_hm2.SetCurSel(0);
    }
    else if (3 == idxHm1)
    {//자동차용품
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"블랙박스,네비게이션 ( 자동차용품 )");
        m_combo_hm2.AddString(L"타이어,휠,체인 ( 자동차용품 )");
        m_combo_hm2.AddString(L"카오디오,AV ( 자동차용품 )");
        m_combo_hm2.AddString(L"실내부품,용품 ( 자동차용품 )");
        m_combo_hm2.AddString(L"외장부품,용품 ( 자동차용품 )");
        m_combo_hm2.AddString(L"세차,청소용품 ( 자동차용품 )");
        m_combo_hm2.AddString(L"기타 자동차용품 ( 자동차용품 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (4 == idxHm1)
    {//뷰티
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"여성화장품 ( 뷰티 )");
        m_combo_hm2.AddString(L"메이크업 ( 뷰티 )");
        m_combo_hm2.AddString(L"남성화장품 ( 뷰티 )");
        m_combo_hm2.AddString(L"향수,헤어,바디 ( 뷰티 )");
        m_combo_hm2.AddString(L"기타 뷰티 ( 뷰티 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (5 == idxHm1)
    {//바이크용품
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"라이더헬맷");
        m_combo_hm2.AddString(L"라이더의류");
        m_combo_hm2.AddString(L"라이더신발, 잡화");
        m_combo_hm2.AddString(L"바이크용품, 부품");
        m_combo_hm2.AddString(L"기타바이크용품");

        m_combo_hm2.SetCurSel(0);
    }
    else if (6 == idxHm1)
    {//여성의류
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"여성코트,아우터 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성티셔츠 ( 여성의류 )");
        m_combo_hm2.AddString(L"남방,블라우스 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성니트 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성조끼 ( 여성의류 )");
        m_combo_hm2.AddString(L"원피스,정장 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성바지 ( 여성의류 )");
        m_combo_hm2.AddString(L"스커트,치마 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성트레이닝복 ( 여성의류 )");
        m_combo_hm2.AddString(L"여성속옷 ( 여성의류 )");
        m_combo_hm2.AddString(L"기타 여성의류 ( 여성의류 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (7 == idxHm1)
    {//남성의류
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"남성코트,아우터 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성티셔츠 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성남방 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성니트 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성바지 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성정장 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성트레이닝복 ( 남성의류 )");
        m_combo_hm2.AddString(L"남성속옷 ( 남성의류 )");
        m_combo_hm2.AddString(L"기타 남성의류 ( 남성의류 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (8 == idxHm1)
    {//신발,가방,잡화
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"여성신발 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"남성신발 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"운동화,기능화 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"가방 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"시계,보석 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"지갑,벨트 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"모자,안경 ( 신발,가방,잡화 )");
        m_combo_hm2.AddString(L"기타 잡화 ( 신발,가방,잡화 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (9 == idxHm1)
    {//휴대폰,태블릿
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"삼성 ( 휴대폰,태블릿 )");
        m_combo_hm2.AddString(L"애플 ( 휴대폰,태블릿 )");
        m_combo_hm2.AddString(L"LG ( 휴대폰,태블릿 )");
        m_combo_hm2.AddString(L"기타 휴대폰 ( 휴대폰,태블릿 )");
        m_combo_hm2.AddString(L"태블릿 ( 휴대폰,태블릿 )");
        m_combo_hm2.AddString(L"액세서리,주변기기 ( 휴대폰,태블릿 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (10 == idxHm1)
    {//컴퓨터,주변기기
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"노트북 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"데스크탑 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"모니터 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"컴퓨터부품 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"프린터,오피스기기 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"저장장치 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"소프트웨어 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"학습기기 ( 컴퓨터,주변기기 )");
        m_combo_hm2.AddString(L"기타 기기,용품 ( 컴퓨터,주변기기 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (11 == idxHm1)
    {//카메라
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"일반디카 ( 카메라 )");
        m_combo_hm2.AddString(L"DSLR ( 카메라 )");
        m_combo_hm2.AddString(L"필름카메라 ( 카메라 )");
        m_combo_hm2.AddString(L"카메라렌즈 ( 카메라 )");
        m_combo_hm2.AddString(L"카메라액세서리 ( 카메라 )");
        m_combo_hm2.AddString(L"캠코더 ( 카메라 )");
        m_combo_hm2.AddString(L"기타 광학용품 ( 카메라 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (12 == idxHm1)
    {//디지털,가전
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"TV ( 디지털,가전 )");
        m_combo_hm2.AddString(L"청소기 ( 디지털,가전 )");
        m_combo_hm2.AddString(L"냉장고 ( 디지털,가전 )");
        m_combo_hm2.AddString(L"세탁기 ( 디지털,가전 )");
        m_combo_hm2.AddString(L"주방조리가전 ( 디지털,가전 )");
        m_combo_hm2.AddString(L"건강,계절가전 ( 디지털,가전 )");
        m_combo_hm2.AddString(L"MP3,iPod ( 디지털,가전 )");
        m_combo_hm2.AddString(L"기타 디지털,가전 ( 디지털,가전 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (13 == idxHm1)
    {//게임
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"PC게임 ( 게임 )");
        m_combo_hm2.AddString(L"닌텐도 ( 게임 )");
        m_combo_hm2.AddString(L"플레이스테이션 ( 게임 )");
        m_combo_hm2.AddString(L"PSP ( 게임 )");
        m_combo_hm2.AddString(L"Wii ( 게임 )");
        m_combo_hm2.AddString(L"Xbox ( 게임 )");
        m_combo_hm2.AddString(L"보드,퍼즐 ( 게임 )");
        m_combo_hm2.AddString(L"기타 게임관련 ( 게임 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (14 == idxHm1)
    {//스포츠,레저
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"자전거 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"등산 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"캠핑 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"골프 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"낚시 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"스키,보드 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"수상스포츠 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"축구 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"야구 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"농구 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"인라인,X게임 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"헬스,요가 ( 스포츠,레저 )");
        m_combo_hm2.AddString(L"기타 스포츠 ( 스포츠,레저 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (15 == idxHm1)
    {//가구
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"침실가구 ( 가구 )");
        m_combo_hm2.AddString(L"거실가구 ( 가구 )");
        m_combo_hm2.AddString(L"수납가구 ( 가구 )");
        m_combo_hm2.AddString(L"주방가구 ( 가구 )");
        m_combo_hm2.AddString(L"책상,책장 ( 가구 )");
        m_combo_hm2.AddString(L"의자 ( 가구 )");
        m_combo_hm2.AddString(L"기타 가구 ( 가구 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (16 == idxHm1)
    {//생활
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"침구,커튼,카펫트 ( 생활 )");
        m_combo_hm2.AddString(L"원예,수예 ( 생활 )");
        m_combo_hm2.AddString(L"세탁,청소용품 ( 생활 )");
        m_combo_hm2.AddString(L"욕실용품 ( 생활 )");
        m_combo_hm2.AddString(L"주방용품 ( 생활 )");
        m_combo_hm2.AddString(L"인테리어소품 ( 생활 )");
        m_combo_hm2.AddString(L"생활,수납용품 ( 생활 )");
        m_combo_hm2.AddString(L"공구,연장 ( 생활 )");
        m_combo_hm2.AddString(L"기타 생활 ( 생활 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (17 == idxHm1)
    {//골동품,희귀품
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"골동품 ( 골동품,희귀품 )");
        m_combo_hm2.AddString(L"희귀품 ( 골동품,희귀품 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (18 == idxHm1)
    {//여행,숙박
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"내가찍은여행사진 ( 여행,숙박 )");
        m_combo_hm2.AddString(L"여행정보,가이드 ( 여행,숙박 )");
        m_combo_hm2.AddString(L"대명,한화숙박권 ( 여행,숙박 )");
        m_combo_hm2.AddString(L"리조트,호텔 ( 여행,숙박 )");
        m_combo_hm2.AddString(L"펜션,캠핑,기타숙박 ( 여행,숙박 )");
        m_combo_hm2.AddString(L"해외숙박 ( 여행,숙박 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (19 == idxHm1)
    {//티켓
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"영화,공연,전시 ( 티켓 )");
        m_combo_hm2.AddString(L"스포츠,레저 ( 티켓 )");
        m_combo_hm2.AddString(L"테마파크,워터파크 ( 티켓 )");
        m_combo_hm2.AddString(L"e티켓,상품권 ( 티켓 )");
        m_combo_hm2.AddString(L"기타 티켓 ( 티켓 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (20 == idxHm1)
    {//재능,서비스
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"과외 ( 재능,서비스 )");
        m_combo_hm2.AddString(L"아르바이트 ( 재능,서비스 )");
        m_combo_hm2.AddString(L"전문스킬,프리랜서 ( 재능,서비스 )");
        m_combo_hm2.AddString(L"가사,생활도움 ( 재능,서비스 )");
        m_combo_hm2.AddString(L"기타 재능공유 ( 재능,서비스 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (21 == idxHm1)
    {//도서
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"일반도서 ( 도서 )");
        m_combo_hm2.AddString(L"교재,전문 ( 도서 )");
        m_combo_hm2.AddString(L"유아동,전집 ( 도서 )");
        m_combo_hm2.AddString(L"만화책 ( 도서 )");
        m_combo_hm2.AddString(L"여행,취미 ( 도서 )");
        m_combo_hm2.AddString(L"잡지 ( 도서 )");
        m_combo_hm2.AddString(L"외국도서 ( 도서 )");
        m_combo_hm2.AddString(L"기타 도서 ( 도서 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (22 == idxHm1)
    {//스타굿즈
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"포토카드,인스,포스터 ( 스타굿즈 )");
        m_combo_hm2.AddString(L"음반 ( 스타굿즈 )");
        m_combo_hm2.AddString(L"응원도구 ( 스타굿즈 )");
        m_combo_hm2.AddString(L"의류 ( 스타굿즈 )");
        m_combo_hm2.AddString(L"잡화,악세서리 ( 스타굿즈 )");
        m_combo_hm2.AddString(L"기타 스타굿즈 ( 스타굿즈 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (23 == idxHm1)
    {//문구
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"문구용품 ( 문구 )");
        m_combo_hm2.AddString(L"사무용품 ( 문구 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (24 == idxHm1)
    {//피규어,키덜트
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"피규어 ( 피규어,키덜트 )");
        m_combo_hm2.AddString(L"프라모델,레고 ( 피규어,키덜트 )");
        m_combo_hm2.AddString(L"RC,드론 ( 피규어,키덜트 )");
        m_combo_hm2.AddString(L"기타 키덜트 ( 피규어,키덜트 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (25 == idxHm1)
    {//CD,DVD
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"CD,LP ( CD,DVD )");
        m_combo_hm2.AddString(L"DVD ( CD,DVD )");
        m_combo_hm2.AddString(L"유아동CD,DVD ( CD,DVD )");
        m_combo_hm2.AddString(L"교육콘텐츠 ( CD,DVD )");
        m_combo_hm2.AddString(L"기타 CD,DVD ( CD,DVD )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (26 == idxHm1)
    {//음향기기,악기
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"헤드폰,이어폰 ( 음향기기,악기 )");
        m_combo_hm2.AddString(L"스피커,오디오 ( 음향기기,악기 )");
        m_combo_hm2.AddString(L"기타 음향기기 ( 음향기기,악기 )");
        m_combo_hm2.AddString(L"피아노,건반악기 ( 음향기기,악기 )");
        m_combo_hm2.AddString(L"바이올린,현악기 ( 음향기기,악기 )");
        m_combo_hm2.AddString(L"그외 악기 ( 음향기기,악기 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (27 == idxHm1)
    {//예술,미술
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"예술,미술작품 ( 예술,미술 )");
        m_combo_hm2.AddString(L"미술용품 ( 예술,미술 )");
        m_combo_hm2.AddString(L"기타 예술,미술 ( 예술,미술 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (28 == idxHm1)
    {//반려동물
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"반려동물용품 ( 반려동물 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (29 == idxHm1)
    {//부동산
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"매매 ( 부동산 )");
        m_combo_hm2.AddString(L"전,월세 ( 부동산 )");
        m_combo_hm2.AddString(L"쉐어,룸메이트 ( 부동산 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (30 == idxHm1)
    {//포장식품
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"건강기능식품 ( 포장식품 )");
        m_combo_hm2.AddString(L"기타 포장식품 ( 포장식품 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (31 == idxHm1)
    {//핸드메이드
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"수제패션,소품 ( 핸드메이드 )");
        m_combo_hm2.AddString(L"기타 수제품 ( 핸드메이드 )");

        m_combo_hm2.SetCurSel(0);
    }
    else if (32 == idxHm1)
    {//기타
        m_combo_hm2.AddString(L"카테고리");
        m_combo_hm2.AddString(L"기타 ( 기타 )");
    }

    m_combo_hm2.SetCurSel(idxHm2);
}


void CETMacroDlg::SetItemChangeComboTm1(int nCur)
{
    int idxTm1 = theApp.m_sellItems[nCur].nThunderMarketFirstComboIndex;
    int idxTm2 = theApp.m_sellItems[nCur].nThunderMarketSecondComboIndex;
    m_combo_tm1.SetCurSel(idxTm1);

    int nCurTm1 = m_combo_tm1.GetCurSel();

    m_combo_tm2.ResetContent();
    m_combo_tm2.EnableWindow(TRUE);

    if (1 == nCurTm1)
    {//여성의류
        m_combo_tm2.AddString(L"카테고리");
        m_combo_tm2.AddString(L"긴팔 티셔츠(여성의류)");
        m_combo_tm2.AddString(L"반팔 티셔츠(여성의류)");
        m_combo_tm2.AddString(L"맨투맨/후드티(여성의류)");
        m_combo_tm2.AddString(L"원피스(여성의류)");
        m_combo_tm2.AddString(L"블라우스(여성의류)");
        m_combo_tm2.AddString(L"셔츠/남방(여성의류)");
        m_combo_tm2.AddString(L"니트/스웨터(여성의류)");
        m_combo_tm2.AddString(L"가디건(여성의류)");
        m_combo_tm2.AddString(L"조끼/베스트(여성의류)");
        m_combo_tm2.AddString(L"야상/점퍼/패딩(여성의류)");
        m_combo_tm2.AddString(L"자켓(여성의류)");
        m_combo_tm2.AddString(L"코트(여성의류)");
        m_combo_tm2.AddString(L"스커트/치마(여성의류)");
        m_combo_tm2.AddString(L"청바지/스키니진(여성의류)");
        m_combo_tm2.AddString(L"면/캐주얼 바지(여성의류)");
        m_combo_tm2.AddString(L"반바지/핫팬츠(여성의류)");
        m_combo_tm2.AddString(L"레깅스(여성의류)");
        m_combo_tm2.AddString(L"비지니스 정장(여성의류)");
        m_combo_tm2.AddString(L"트레이닝(여성의류)");
        m_combo_tm2.AddString(L"언더웨어/속옷(여성의류)");
        m_combo_tm2.AddString(L"빅사이즈(여성의류)");
        m_combo_tm2.AddString(L"테마/이벤트옷(여성의류)");

        m_combo_tm2.SetCurSel(0);
    }
    else if (2 == nCurTm1)
    {//남성의류
        m_combo_tm2.AddString(L"카테고리");
        m_combo_tm2.AddString(L"긴팔 티셔츠(남성의류)");
        m_combo_tm2.AddString(L"반팔 티셔츠(남성의류)");
        m_combo_tm2.AddString(L"맨투맨/후드티(남성의류)");
        m_combo_tm2.AddString(L"셔츠/남방(남성의류)");
        m_combo_tm2.AddString(L"니트/스웨터(남성의류)");
        m_combo_tm2.AddString(L"가디건(남성의류)");
        m_combo_tm2.AddString(L"조끼/베스트(남성의류)");
        m_combo_tm2.AddString(L"점퍼/야상/패딩(남성의류)");
        m_combo_tm2.AddString(L"자켓(남성의류)");
        m_combo_tm2.AddString(L"코트(남성의류)");
        m_combo_tm2.AddString(L"청바지(긴)(남성의류)");
        m_combo_tm2.AddString(L"면/캐주얼 바지(남성의류)");
        m_combo_tm2.AddString(L"반바지/7~9부(남성의류)");
        m_combo_tm2.AddString(L"비지니스 정장(남성의류)");
        m_combo_tm2.AddString(L"트레이닝(남성의류)");
        m_combo_tm2.AddString(L"언더웨어/속옷(남성의류)");
        m_combo_tm2.AddString(L"빅사이즈(남성의류)");
        m_combo_tm2.AddString(L"테마/이벤트옷(남성의류)");

        m_combo_tm2.SetCurSel(0);
    }
    else if (3 == nCurTm1)
    {//패션잡화
        m_combo_tm2.AddString(L"카테고리");
        m_combo_tm2.AddString(L"여성가방(패션잡화)");
        m_combo_tm2.AddString(L"남성가방(패션잡화)");
        m_combo_tm2.AddString(L"여행용가방(패션잡화)");
        m_combo_tm2.AddString(L"운동화(패션잡화)");
        m_combo_tm2.AddString(L"여성화(패션잡화)");
        m_combo_tm2.AddString(L"남성화(패션잡화)");
        m_combo_tm2.AddString(L"지갑(패션잡화)");
        m_combo_tm2.AddString(L"모자(패션잡화)");
        m_combo_tm2.AddString(L"안경/선글라스(패션잡화)");
        m_combo_tm2.AddString(L"주얼리(패션잡화)");
        m_combo_tm2.AddString(L"시계(패션잡화)");
        m_combo_tm2.AddString(L"벨트/장갑(패션잡화)");


        m_combo_tm2.SetCurSel(0);
    }
    else if (4 == nCurTm1)
    {//유아동/출산
        m_combo_tm2.AddString(L"카테고리");
        m_combo_tm2.AddString(L"베이비의류(유아동/출산)");
        m_combo_tm2.AddString(L"여아의류(유아동/출산)");
        m_combo_tm2.AddString(L"남아의류(유아동/출산)");
        m_combo_tm2.AddString(L"여주니어의류(유아동/출산)");
        m_combo_tm2.AddString(L"남주니어의류(유아동/출산)");
        m_combo_tm2.AddString(L"유아동신발(유아동/출산)");
        m_combo_tm2.AddString(L"유아동용품(유아동/출산)");
        m_combo_tm2.AddString(L"출산(유아동/출산)");
        m_combo_tm2.AddString(L"교육/완구(유아동/출산)");
        m_combo_tm2.AddString(L"기저귀(유아동/출산)");


        m_combo_tm2.SetCurSel(0);
    }

    m_combo_tm2.SetCurSel(idxTm2);
}


void CETMacroDlg::SetItemChangeComboNc(int nCur)
{
    int idxNc = theApp.m_sellItems[nCur].nNaverCafeFirstComboIndex;

    if (89 == idxNc)
        m_combo_nc.SetCurSel(1);
    else if(90 == idxNc)
        m_combo_nc.SetCurSel(2);
    else if (82 == idxNc)
        m_combo_nc.SetCurSel(3);
    else
        m_combo_nc.SetCurSel(4);
}

void CETMacroDlg::SetItemChangeComboTm2(int nCur)
{
    int nCurTm1 = m_combo_tm1.GetCurSel();
    int nCurTm2 = m_combo_tm2.GetCurSel();

    int idxTm3 = theApp.m_sellItems[nCur].nThunderMarketThirdComboIndex;
    int idxTmSize = theApp.m_sellItems[nCur].nSizeComboIndex;


    m_combo_tm3.ResetContent();
    m_combo_tm3.EnableWindow(TRUE);

    m_combo_tm_size.ResetContent();
    m_combo_tm_size.EnableWindow(TRUE);

    // 1번 카테고리와 2번 카테고리의 인덱스에따라서 3번, Size 콤보박스가 다르게 생성됨(주의해야함)
    if (1 == nCurTm1 && 1 == nCurTm2)
    {//여성의류 && 긴팔 티셔츠(여성의류)
        m_combo_tm3.AddString(L"무지/기본 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"라운드 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"브이넥 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"스프라이프 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"폴라 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"기타(긴팔 티셔츠(여성의류)");
    }
    else if (1 == nCurTm1 && 2 == nCurTm2)
    {//여성의류 && 반팔 티셔츠(여성의류)
        m_combo_tm3.AddString(L"무지/기본 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"라운드 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"브이넥 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"카라 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"스프라이프 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"민소매/나시 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"기타(반팔 티셔츠(여성의류)");
    }

    else if (1 == nCurTm1 && 3 == nCurTm2)
    {//여성의류 && 맨투맨/후드티(여성의류)
        m_combo_tm3.AddString(L"맨투맨 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"후드 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"후드 집업(여성의류)");
        m_combo_tm3.AddString(L"카라 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"기타(여성의류)");
    }

    else if (1 == nCurTm1 && 4 == nCurTm2)
    {//여성의류 && 원피스(여성의류)
        m_combo_tm3.AddString(L"캐주얼 원피스(여성의류)");
        m_combo_tm3.AddString(L"미니 원피스(여성의류)");
        m_combo_tm3.AddString(L"롱 원피스(여성의류)");
        m_combo_tm3.AddString(L"나시/탑 원피스(여성의류)");
        m_combo_tm3.AddString(L"쉬폰/레이스 원피스(여성의류)");
        m_combo_tm3.AddString(L"럭셔리 원피스(여성의류)");
        m_combo_tm3.AddString(L"후드/니트 원피스(여성의류)");
        m_combo_tm3.AddString(L"청 원피스(여성의류)");
        m_combo_tm3.AddString(L"프린트 원피스(여성의류)");
        m_combo_tm3.AddString(L"투피스 원피스(여성의류)");
        m_combo_tm3.AddString(L"기타(원피스)(여성의류)");
    }

    else if (1 == nCurTm1 && 5 == nCurTm2)
    {//여성의류 && 블라우스(여성의류)
        m_combo_tm3.AddString(L"쉬폰/시스루 블라우스(여성의류)");
        m_combo_tm3.AddString(L"레이스 블라우스(여성의류)");
        m_combo_tm3.AddString(L"프릴/셔링 블라우스(여성의류)");
        m_combo_tm3.AddString(L"프린트 블라우스(여성의류)");
        m_combo_tm3.AddString(L"오프숄더 블라우스(여성의류)");
        m_combo_tm3.AddString(L"민소매/홀터넥 블라우스(여성의류)");
        m_combo_tm3.AddString(L"기타(블라우스)(여성의류)");
    }

    else if (1 == nCurTm1 && 6 == nCurTm2)
    {//여성의류 && 셔츠/남방(여성의류)
        m_combo_tm3.AddString(L"무지/기본 셔츠(여성의류)");
        m_combo_tm3.AddString(L"루즈핏/박시 셔츠(여성의류)");
        m_combo_tm3.AddString(L"체크 셔츠(여성의류)");
        m_combo_tm3.AddString(L"청/데님 셔츠(여성의류)");
        m_combo_tm3.AddString(L"스트라이프 셔츠(여성의류)");
        m_combo_tm3.AddString(L"기타(셔츠/남방)(여성의류)");
    }

    else if (1 == nCurTm1 && 7 == nCurTm2)
    {//여성의류 && 니트/스웨터(여성의류)
        m_combo_tm3.AddString(L"라운드넥 니트(여성의류)");
        m_combo_tm3.AddString(L"브이넥 니트(여성의류)");
        m_combo_tm3.AddString(L"오프숄더 니트(여성의류)");
        m_combo_tm3.AddString(L"폴라/터틀(여성의류)");
        m_combo_tm3.AddString(L"롱 니트(여성의류)");
        m_combo_tm3.AddString(L"루즈핏 니트(여성의류)");
        m_combo_tm3.AddString(L"기타(니트/스웨터)(여성의류)");
    }

    else if (1 == nCurTm1 && 8 == nCurTm2)
    {//여성의류 && 가디건(여성의류)
        m_combo_tm3.AddString(L"라운드넥 가디건(여성의류)");
        m_combo_tm3.AddString(L"브이넥 가디건(여성의류)");
        m_combo_tm3.AddString(L"루즈핏/박시 가디건(여성의류)");
        m_combo_tm3.AddString(L"롱 가디건(여성의류)");
        m_combo_tm3.AddString(L"후드 가디건(여성의류)");
        m_combo_tm3.AddString(L"기타(가디건)(여성의류)");
    }

    else if (1 == nCurTm1 && 9 == nCurTm2)
    {//여성의류 && 조끼/베스트(여성의류)
        m_combo_tm3.AddString(L"니트 조끼(여성의류)");
        m_combo_tm3.AddString(L"청/데님 조끼(여성의류)");
        m_combo_tm3.AddString(L"퍼 조끼(여성의류)");
        m_combo_tm3.AddString(L"패딩 조끼(여성의류)");
        m_combo_tm3.AddString(L"기타(조끼/베스트)(여성의류)");
    }

    else if (1 == nCurTm1 && 10 == nCurTm2)
    {//여성의류 && 야상/점퍼/패딩(여성의류)
        m_combo_tm3.AddString(L"야상/사파리(여성의류)");
        m_combo_tm3.AddString(L"야구점퍼(여성의류)");
        m_combo_tm3.AddString(L"패딩(여성의류)");
        m_combo_tm3.AddString(L"바람막이(여성의류)");
        m_combo_tm3.AddString(L"기타(야상/점퍼/패딩)(여성의류)");
    }

    else if (1 == nCurTm1 && 11 == nCurTm2)
    {//여성의류 && 자켓(여성의류)
        m_combo_tm3.AddString(L"기본/테일러드 자켓(여성의류)");
        m_combo_tm3.AddString(L"청/데님자켓(여성의류)");
        m_combo_tm3.AddString(L"트위드/체크자켓(여성의류)");
        m_combo_tm3.AddString(L"가죽/라이더(여성의류)");
        m_combo_tm3.AddString(L"기타(자켓)(자켓(여성의류))");
    }

    else if (1 == nCurTm1 && 12 == nCurTm2)
    {//여성의류 && 코트(여성의류)
        m_combo_tm3.AddString(L"트렌치 코트(여성의류)");
        m_combo_tm3.AddString(L"반/하프 코트(여성의류)");
        m_combo_tm3.AddString(L"롱 코트(여성의류)");
        m_combo_tm3.AddString(L"케이프/망토(여성의류)");
        m_combo_tm3.AddString(L"무스탕(여성의류)");
        m_combo_tm3.AddString(L"모피(여성의류)");
        m_combo_tm3.AddString(L"기타(코트)(코트(여성의류))");
    }

    else if (1 == nCurTm1 && 13 == nCurTm2)
    {//여성의류 && 스커트/치마(여성의류)
        m_combo_tm3.AddString(L"쉬폰/레이스(여성의류)");
        m_combo_tm3.AddString(L"플레어 스커트(여성의류)");
        m_combo_tm3.AddString(L"롱 스커트(여성의류)");
        m_combo_tm3.AddString(L"미니 스커트(여성의류)");
        m_combo_tm3.AddString(L"모직/니트 스커트(여성의류)");
        m_combo_tm3.AddString(L"플리츠(주름)(여성의류)");
        m_combo_tm3.AddString(L"청스커트(여성의류)");
        m_combo_tm3.AddString(L"기타(스커트/치마)(여성의류)");
    }

    else if (1 == nCurTm1 && 14 == nCurTm2)
    {//여성의류 && 청바지/스키니(긴)(여성의류)
        m_combo_tm3.AddString(L"스키니진(여성의류)");
        m_combo_tm3.AddString(L"일자 청바지(여성의류)");
        m_combo_tm3.AddString(L"부츠컷 청바지(여성의류)");
        m_combo_tm3.AddString(L"배기/카고(여성의류)");
        m_combo_tm3.AddString(L"하이웨스트 진(여성의류)");
        m_combo_tm3.AddString(L"기타(청바지/스키니(긴))(여성의류)");
    }

    else if (1 == nCurTm1 && 15 == nCurTm2)
    {//여성의류 && 면/캐주얼 바지(긴)(여성의류)
        m_combo_tm3.AddString(L"일자바지/슬렉스(여성의류)");
        m_combo_tm3.AddString(L"점프 수트/멜빵(여성의류)");
        m_combo_tm3.AddString(L"통/와이드 팬츠(여성의류)");
        m_combo_tm3.AddString(L"배기 팬츠(여성의류)");
        m_combo_tm3.AddString(L"하이웨스트 팬츠(여성의류)");
        m_combo_tm3.AddString(L"가죽/모직 바지(여성의류)");
        m_combo_tm3.AddString(L"기타(면/캐주얼 바지(긴))(여성의류)");
    }

    else if (1 == nCurTm1 && 16 == nCurTm2)
    {//여성의류 && 반바지/핫팬츠(여성의류)
        m_combo_tm3.AddString(L"면 반바지(여성의류)");
        m_combo_tm3.AddString(L"청 반바지(여성의류)");
        m_combo_tm3.AddString(L"핫 팬츠(여성의류)");
        m_combo_tm3.AddString(L"치마 바지(여성의류)");
        m_combo_tm3.AddString(L"가죽/모직 반바지(여성의류)");
        m_combo_tm3.AddString(L"기타(반바지/핫팬츠)(여성의류)");
    }

    else if (1 == nCurTm1 && 17 == nCurTm2)
    {//여성의류 && 레깅스(여성의류)
        m_combo_tm3.AddString(L"무지 레깅스(여성의류)");
        m_combo_tm3.AddString(L"치마 레깅스(여성의류)");
        m_combo_tm3.AddString(L"프린트 레깅스(여성의류)");
        m_combo_tm3.AddString(L"기모/밍크 레깅스(여성의류)");
        m_combo_tm3.AddString(L"가죽 레깅스(여성의류)");
        m_combo_tm3.AddString(L"기타(레깅스)(여성의류)");
    }

    else if (1 == nCurTm1 && 18 == nCurTm2)
    {//여성의류 && 비즈니스 정장(여성의류)
        m_combo_tm3.AddString(L"정장 세트(비즈니스 정장)(여성의류)");
        m_combo_tm3.AddString(L"정장 원피스(여성의류)");
        m_combo_tm3.AddString(L"정장 자켓(여성의류)");
        m_combo_tm3.AddString(L"정장 블라우스(여성의류)");
        m_combo_tm3.AddString(L"정장 바지/슬랙스(여성의류)");
        m_combo_tm3.AddString(L"정장 치마(여성의류)");
        m_combo_tm3.AddString(L"기타(비즈니스 정장)(여성의류)");
    }

    else if (1 == nCurTm1 && 19 == nCurTm2)
    {//여성의류 && 트레이닝(여성의류)
        m_combo_tm3.AddString(L"트레이닝 상의(여성의류)");
        m_combo_tm3.AddString(L"트레이닝 하의(여성의류)");
        m_combo_tm3.AddString(L"트레이닝 세트(여성의류)");
        m_combo_tm3.AddString(L"기타(트레이닝)(여성의류)");
    }

    else if (1 == nCurTm1 && 20 == nCurTm2)
    {//여성의류 && 언더웨어/속옷(여성의류)
        m_combo_tm3.AddString(L"브라(언더웨어/속옷)(여성의류)");
        m_combo_tm3.AddString(L"팬티(언더웨어/속옷)(여성의류)");
        m_combo_tm3.AddString(L"브라팬티 세트(여성의류)");
        m_combo_tm3.AddString(L"보정 속옷(여성의류)");
        m_combo_tm3.AddString(L"잠옷/이지웨어(여성의류)");
        m_combo_tm3.AddString(L"런닝/내의(여성의류)");
        m_combo_tm3.AddString(L"슬립/캐미솔(여성의류)");
        m_combo_tm3.AddString(L"속치마/속바지(여성의류)");
        m_combo_tm3.AddString(L"기능성/히트텍(여성의류)");
        m_combo_tm3.AddString(L"기타(언더웨어/속옷)(여성의류)");

        m_combo_tm_size.AddString(L"70A");
        m_combo_tm_size.AddString(L"70B");
        m_combo_tm_size.AddString(L"70C");
        m_combo_tm_size.AddString(L"70D");
        m_combo_tm_size.AddString(L"70E");
        m_combo_tm_size.AddString(L"70F");
        m_combo_tm_size.AddString(L"75A");
        m_combo_tm_size.AddString(L"75B");
        m_combo_tm_size.AddString(L"75C");
        m_combo_tm_size.AddString(L"75D");
        m_combo_tm_size.AddString(L"75E");
        m_combo_tm_size.AddString(L"75F");
        m_combo_tm_size.AddString(L"80A");
        m_combo_tm_size.AddString(L"80B");
        m_combo_tm_size.AddString(L"80C");
        m_combo_tm_size.AddString(L"80D");
        m_combo_tm_size.AddString(L"80E");
        m_combo_tm_size.AddString(L"80F");
        m_combo_tm_size.AddString(L"85A");
        m_combo_tm_size.AddString(L"85B");
        m_combo_tm_size.AddString(L"85C");
        m_combo_tm_size.AddString(L"85D");
        m_combo_tm_size.AddString(L"85E");
        m_combo_tm_size.AddString(L"85F");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (1 == nCurTm1 && 21 == nCurTm2)
    {//여성의류 && 빅사이즈(여성의류)
        m_combo_tm3.AddString(L"긴팔 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"반팔 티셔츠(여성의류)");
        m_combo_tm3.AddString(L"맨투맨/후드티(여성의류)");
        m_combo_tm3.AddString(L"니트/가디건(여성의류)");
        m_combo_tm3.AddString(L"자켓/점퍼(여성의류)");
        m_combo_tm3.AddString(L"패딩/코트(여성의류)");
        m_combo_tm3.AddString(L"면/캐주얼(여성의류)");
        m_combo_tm3.AddString(L"청바지(여성의류)");
        m_combo_tm3.AddString(L"반바지(여성의류)");
        m_combo_tm3.AddString(L"정장/팬츠(여성의류)");
        m_combo_tm3.AddString(L"트레이닝복(여성의류)");
        m_combo_tm3.AddString(L"언더웨어/속옷(여성의류)");
        m_combo_tm3.AddString(L"기타(빅사이즈)(여성의류)");
    }

    else if (1 == nCurTm1 && 22 == nCurTm2)
    {//여성의류 && 테마/이벤트 의류(여성의류)
        m_combo_tm3.AddString(L"우비/레인코트(여성의류)");
        m_combo_tm3.AddString(L"교복(여성의류)");
        m_combo_tm3.AddString(L"유니폼/작업복(여성의류)");
        m_combo_tm3.AddString(L"생활/전통(여성의류)");
        m_combo_tm3.AddString(L"웨딩 드레스(여성의류)");
        m_combo_tm3.AddString(L"코디(여성의류)");
        m_combo_tm3.AddString(L"옷장정리/급처(여성의류)");
        m_combo_tm3.AddString(L"기타(테마/이벤트 의류)(여성의류)");

        m_combo_tm_size.AddString(L"없음");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    if (1 == nCurTm1 && (12 >= nCurTm2 || 18 == nCurTm2 || 19 == nCurTm2 || 21 == nCurTm2))
    {
        m_combo_tm_size.AddString(L"44");
        m_combo_tm_size.AddString(L"55");
        m_combo_tm_size.AddString(L"66");
        m_combo_tm_size.AddString(L"77");
        m_combo_tm_size.AddString(L"88");
        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }
    else if (1 == nCurTm1 && (13 <= nCurTm2 && 17 >= nCurTm2))
    {
        m_combo_tm_size.AddString(L"44");
        m_combo_tm_size.AddString(L"55");
        m_combo_tm_size.AddString(L"66");
        m_combo_tm_size.AddString(L"77");
        m_combo_tm_size.AddString(L"88");
        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"23");
        m_combo_tm_size.AddString(L"24");
        m_combo_tm_size.AddString(L"25");
        m_combo_tm_size.AddString(L"26");
        m_combo_tm_size.AddString(L"27");
        m_combo_tm_size.AddString(L"28");
        m_combo_tm_size.AddString(L"29");
        m_combo_tm_size.AddString(L"30");
        m_combo_tm_size.AddString(L"31");
        m_combo_tm_size.AddString(L"32");
        m_combo_tm_size.AddString(L"33");
        m_combo_tm_size.AddString(L"34");
        m_combo_tm_size.AddString(L"35");
        m_combo_tm_size.AddString(L"36");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }


    // 남성의류
    else if (2 == nCurTm1 && 1 == nCurTm2)
    {//남성의류 && 긴팔 티셔츠(남성의류)
        m_combo_tm3.AddString(L"라운드넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"브이넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"카라넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"기타(긴팔 티셔츠(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 2 == nCurTm2)
    {//남성의류 && 반팔 티셔츠(남성의류)
        m_combo_tm3.AddString(L"라운드넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"카라넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"브이넥 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"민소매/ 나시 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"기타(반팔 티셔츠(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 3 == nCurTm2)
    {//남성의류 && 맨투맨/후드티(남성의류)
        m_combo_tm3.AddString(L"맨투맨 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"후드 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"후드 집업(남성의류)");
        m_combo_tm3.AddString(L"기타(맨투맨/후드티)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 4 == nCurTm2)
    {//남성의류 && 셔츠/남방(남성의류)
        m_combo_tm3.AddString(L"솔리드(단색)(남성의류)");
        m_combo_tm3.AddString(L"스트라이프 셔츠(남성의류)");
        m_combo_tm3.AddString(L"린넨/마 셔츠(남성의류)");
        m_combo_tm3.AddString(L"청/데님 셔츠(남성의류)");
        m_combo_tm3.AddString(L"헨리넥 셔츠(남성의류)");
        m_combo_tm3.AddString(L"체크 셔츠(남성의류)");
        m_combo_tm3.AddString(L"기타(셔츠/남방)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 5 == nCurTm2)
    {//남성의류 && 니트/스웨터(남성의류)
        m_combo_tm3.AddString(L"라운드넥 니트(남성의류)");
        m_combo_tm3.AddString(L"브이넥 니트(남성의류)");
        m_combo_tm3.AddString(L"집업 니트(남성의류)");
        m_combo_tm3.AddString(L"카라넥 니트(남성의류)");
        m_combo_tm3.AddString(L"폴라 니트(남성의류)");
        m_combo_tm3.AddString(L"기타(니트/스웨터)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 6 == nCurTm2)
    {//남성의류 && 가디건(남성의류)
        m_combo_tm3.AddString(L"브이넥 가디건(남성의류)");
        m_combo_tm3.AddString(L"라운드 가디건(남성의류)");
        m_combo_tm3.AddString(L"집업 가디건(남성의류)");
        m_combo_tm3.AddString(L"후드 가디건(남성의류)");
        m_combo_tm3.AddString(L"기타(가디건)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 7 == nCurTm2)
    {//남성의류 && 조끼/베스트(남성의류)
        m_combo_tm3.AddString(L"니트 조끼(남성의류)");
        m_combo_tm3.AddString(L"청/데님 조끼(남성의류)");
        m_combo_tm3.AddString(L"브이넥 조끼(남성의류)");
        m_combo_tm3.AddString(L"패딩 조끼(남성의류)");
        m_combo_tm3.AddString(L"기타(조끼/베스트)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 8 == nCurTm2)
    {//남성의류 && 점퍼/야상/패딩(남성의류)
        m_combo_tm3.AddString(L"바람막이(남성의류)");
        m_combo_tm3.AddString(L"패딩 점퍼(남성의류)");
        m_combo_tm3.AddString(L"다운 점퍼(남성의류)");
        m_combo_tm3.AddString(L"야구 점퍼(남성의류)");
        m_combo_tm3.AddString(L"블루종/항공점퍼(남성의류)");
        m_combo_tm3.AddString(L"야상/사파리(남성의류)");
        m_combo_tm3.AddString(L"기타(점퍼/야상/패딩)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 9 == nCurTm2)
    {//남성의류 && 자켓(남성의류)
        m_combo_tm3.AddString(L"캐주얼 자켓(남성의류)");
        m_combo_tm3.AddString(L"청/데님 자켓(남성의류)");
        m_combo_tm3.AddString(L"가죽 자켓(남성의류)");
        m_combo_tm3.AddString(L"차이나/노카라 자켓(남성의류)");
        m_combo_tm3.AddString(L"린넨/마 자켓(남성의류)");
        m_combo_tm3.AddString(L"후드/져지 자켓(남성의류)");
        m_combo_tm3.AddString(L"기타(자켓)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 10 == nCurTm2)
    {//남성의류 && 코트(남성의류)
        m_combo_tm3.AddString(L"모직 코트(남성의류)");
        m_combo_tm3.AddString(L"트렌치 코트(남성의류)");
        m_combo_tm3.AddString(L"하프 코트(남성의류)");
        m_combo_tm3.AddString(L"캐시미어 코트(남성의류)");
        m_combo_tm3.AddString(L"기타(코트)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 11 == nCurTm2)
    {//남성의류 && 청바지(긴)(남성의류)
        m_combo_tm3.AddString(L"일자 청바지(남성의류)");
        m_combo_tm3.AddString(L"스키니진(남성의류)");
        m_combo_tm3.AddString(L"빈티지/구제 청바지(남성의류)");
        m_combo_tm3.AddString(L"배기 청바지(남성의류)");
        m_combo_tm3.AddString(L"블랙/그레이진(남성의류)");
        m_combo_tm3.AddString(L"부츠컷 청바지(남성의류)");
        m_combo_tm3.AddString(L"기타(청바지(긴)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"23");
        m_combo_tm_size.AddString(L"24");
        m_combo_tm_size.AddString(L"25");
        m_combo_tm_size.AddString(L"26");
        m_combo_tm_size.AddString(L"27");
        m_combo_tm_size.AddString(L"28");
        m_combo_tm_size.AddString(L"29");
        m_combo_tm_size.AddString(L"30");
        m_combo_tm_size.AddString(L"31");
        m_combo_tm_size.AddString(L"32");
        m_combo_tm_size.AddString(L"33");
        m_combo_tm_size.AddString(L"34");
        m_combo_tm_size.AddString(L"35");
        m_combo_tm_size.AddString(L"36");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 12 == nCurTm2)
    {//남성의류 && 면/캐주얼 바지(긴)(남성의류)
        m_combo_tm3.AddString(L"면바지(남성의류)");
        m_combo_tm3.AddString(L"슬랙스(남성의류)");
        m_combo_tm3.AddString(L"배기/조거 바지(남성의류)");
        m_combo_tm3.AddString(L"기모 바지(남성의류)");
        m_combo_tm3.AddString(L"카고 바지(남성의류)");
        m_combo_tm3.AddString(L"기타(면/캐주얼 바지(긴))(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"23");
        m_combo_tm_size.AddString(L"24");
        m_combo_tm_size.AddString(L"25");
        m_combo_tm_size.AddString(L"26");
        m_combo_tm_size.AddString(L"27");
        m_combo_tm_size.AddString(L"28");
        m_combo_tm_size.AddString(L"29");
        m_combo_tm_size.AddString(L"30");
        m_combo_tm_size.AddString(L"31");
        m_combo_tm_size.AddString(L"32");
        m_combo_tm_size.AddString(L"33");
        m_combo_tm_size.AddString(L"34");
        m_combo_tm_size.AddString(L"35");
        m_combo_tm_size.AddString(L"36");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 13 == nCurTm2)
    {//남성의류 && 반바지/7~9부(남성의류)
        m_combo_tm3.AddString(L"면 반바지(남성의류)");
        m_combo_tm3.AddString(L"청/데님 반바지(남성의류)");
        m_combo_tm3.AddString(L"밴딩 반바지(남성의류)");
        m_combo_tm3.AddString(L"스포츠 반바지(남성의류)");
        m_combo_tm3.AddString(L"기타(반바지/7~9부)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"23");
        m_combo_tm_size.AddString(L"24");
        m_combo_tm_size.AddString(L"25");
        m_combo_tm_size.AddString(L"26");
        m_combo_tm_size.AddString(L"27");
        m_combo_tm_size.AddString(L"28");
        m_combo_tm_size.AddString(L"29");
        m_combo_tm_size.AddString(L"30");
        m_combo_tm_size.AddString(L"31");
        m_combo_tm_size.AddString(L"32");
        m_combo_tm_size.AddString(L"33");
        m_combo_tm_size.AddString(L"34");
        m_combo_tm_size.AddString(L"35");
        m_combo_tm_size.AddString(L"36");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 14 == nCurTm2)
    {//남성의류 && 비즈니스 정장(남성의류)
        m_combo_tm3.AddString(L"정장 자켓(남성의류)");
        m_combo_tm3.AddString(L"정장 바지(남성의류)");
        m_combo_tm3.AddString(L"정장 베스트(남성의류)");
        m_combo_tm3.AddString(L"기타(비즈니스 정장)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 15 == nCurTm2)
    {//남성의류 && 트레이닝(남성의류)
        m_combo_tm3.AddString(L"트레이닝 상의(남성의류)");
        m_combo_tm3.AddString(L"트레이닝 하의(남성의류)");
        m_combo_tm3.AddString(L"트레이닝 세트(남성의류)");
        m_combo_tm3.AddString(L"기타(트레이닝)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 16 == nCurTm2)
    {//남성의류 && 언더웨어/속옷(남성의류)
        m_combo_tm3.AddString(L"런닝(남성의류)");
        m_combo_tm3.AddString(L"드로즈/삼각(남성의류)");
        m_combo_tm3.AddString(L"트렁크(남성의류)");
        m_combo_tm3.AddString(L"잠옷/이지웨어(남성의류)");
        m_combo_tm3.AddString(L"런닝/팬티(남성의류)");
        m_combo_tm3.AddString(L"기능성/히트텍(남성의류)");
        m_combo_tm3.AddString(L"기타(언더웨어/속옷)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 17 == nCurTm2)
    {//남성의류 && 빅사이즈(남성의류)
        m_combo_tm3.AddString(L"긴팔 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"반팔 티셔츠(남성의류)");
        m_combo_tm3.AddString(L"맨투맨 후드(남성의류)");
        m_combo_tm3.AddString(L"니트/가디건(남성의류)");
        m_combo_tm3.AddString(L"자켓/점퍼(남성의류)");
        m_combo_tm3.AddString(L"패딩/코트(남성의류)");
        m_combo_tm3.AddString(L"면/캐주얼(남성의류)");
        m_combo_tm3.AddString(L"청바지(남성의류)");
        m_combo_tm3.AddString(L"정장/팬츠(남성의류)");
        m_combo_tm3.AddString(L"트레이닝복(남성의류)");
        m_combo_tm3.AddString(L"기타(빅사이즈)(남성의류)");

        m_combo_tm_size.AddString(L"없음");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (2 == nCurTm1 && 18 == nCurTm2)
    {//남성의류 && 테마/이벤트 의류(남성의류)
        m_combo_tm3.AddString(L"우비/레인코트(남성의류)");
        m_combo_tm3.AddString(L"교복(남성의류)");
        m_combo_tm3.AddString(L"유니폼/작업복(남성의류)");
        m_combo_tm3.AddString(L"생활/전통(남성의류)");
        m_combo_tm3.AddString(L"군복(남성의류)");
        m_combo_tm3.AddString(L"기타(테마/이벤트 의류)(남성의류)");

        m_combo_tm_size.AddString(L"XS");
        m_combo_tm_size.AddString(L"S");
        m_combo_tm_size.AddString(L"M");
        m_combo_tm_size.AddString(L"L");
        m_combo_tm_size.AddString(L"XL");
        m_combo_tm_size.AddString(L"XXL");
        m_combo_tm_size.AddString(L"85");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"95");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"105");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }


    // 패션잡화
    else if (3 == nCurTm1 && 1 == nCurTm2)
    {//패션잡화 && 여성가방(패션잡화)
        m_combo_tm3.AddString(L"숄더백(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"크로스백(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"클러치백(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"토트백(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"백팩(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"파우치(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"미니백(여성가방(패션잡화))");
        m_combo_tm3.AddString(L"기타(여성가방(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 2 == nCurTm2)
    {//패션잡화 && 남성가방(패션잡화)
        m_combo_tm3.AddString(L"백팩(남성가방(패션잡화))");
        m_combo_tm3.AddString(L"크로스백(남성가방(패션잡화))");
        m_combo_tm3.AddString(L"숄더백(남성가방(패션잡화))");
        m_combo_tm3.AddString(L"비즈니스가방(남성가방(패션잡화))");
        m_combo_tm3.AddString(L"클러치백(남성가방(패션잡화))");
        m_combo_tm3.AddString(L"기타(남성가방(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 3 == nCurTm2)
    {//패션잡화 && 여행용가방/소품(패션잡화)
        m_combo_tm3.AddString(L"하드 캐리어(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"소프트 캐리어(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"이민/유학용(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"여행용 백팩(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"여행용 크로스(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"여행용 파우치(여행용가방/소품(패션잡화))");
        m_combo_tm3.AddString(L"기타(여행용가방/소품(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 4 == nCurTm2)
    {//패션잡화 && 운동화/캐주얼화(패션잡화)
        m_combo_tm3.AddString(L"런닝화/워킹화(운동화/캐주얼화(패션잡화))");
        m_combo_tm3.AddString(L"농구화(운동화/캐주얼화(패션잡화))");
        m_combo_tm3.AddString(L"캐주얼화(운동화/캐주얼화(패션잡화))");
        m_combo_tm3.AddString(L"기타(운동화/캐주얼화(패션잡화)");

        m_combo_tm_size.AddString(L"200.0");
        m_combo_tm_size.AddString(L"205.0");
        m_combo_tm_size.AddString(L"210.0");
        m_combo_tm_size.AddString(L"215.0");
        m_combo_tm_size.AddString(L"220.0");
        m_combo_tm_size.AddString(L"225.0");
        m_combo_tm_size.AddString(L"230.0");
        m_combo_tm_size.AddString(L"235.0");
        m_combo_tm_size.AddString(L"240.0");
        m_combo_tm_size.AddString(L"245.0");
        m_combo_tm_size.AddString(L"250.0");
        m_combo_tm_size.AddString(L"255.0");
        m_combo_tm_size.AddString(L"260.0");
        m_combo_tm_size.AddString(L"265.0");
        m_combo_tm_size.AddString(L"270.0");
        m_combo_tm_size.AddString(L"275.0");
        m_combo_tm_size.AddString(L"280.0");
        m_combo_tm_size.AddString(L"285.0");
        m_combo_tm_size.AddString(L"290.0");
        m_combo_tm_size.AddString(L"295.0");
        m_combo_tm_size.AddString(L"300.0");
        m_combo_tm_size.AddString(L"305.0");
        m_combo_tm_size.AddString(L"310.0");
        m_combo_tm_size.AddString(L"315.0");
        m_combo_tm_size.AddString(L"320.0");
        m_combo_tm_size.AddString(L"325.0");
        m_combo_tm_size.AddString(L"330.0");
        m_combo_tm_size.AddString(L"335.0");
        m_combo_tm_size.AddString(L"340.0");
        m_combo_tm_size.AddString(L"345.0");
        m_combo_tm_size.AddString(L"350.0");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 5 == nCurTm2)
    {//패션잡화 && 여성화(패션잡화)
        m_combo_tm3.AddString(L"단화/플랫슈즈(여성화(패션잡화))");
        m_combo_tm3.AddString(L"하이힐(여성화(패션잡화))");
        m_combo_tm3.AddString(L"미들굽 펌프스(여성화(패션잡화))");
        m_combo_tm3.AddString(L"샌들/슬리퍼(여성화(패션잡화))");
        m_combo_tm3.AddString(L"플랫폼 슈즈(여성화(패션잡화))");
        m_combo_tm3.AddString(L"워커(여성화(패션잡화))");
        m_combo_tm3.AddString(L"앵클부츠(여성화(패션잡화))");
        m_combo_tm3.AddString(L"롱/미들 부츠(여성화(패션잡화))");
        m_combo_tm3.AddString(L"기타(여성화(패션잡화)");

        m_combo_tm_size.AddString(L"200.0");
        m_combo_tm_size.AddString(L"205.0");
        m_combo_tm_size.AddString(L"210.0");
        m_combo_tm_size.AddString(L"215.0");
        m_combo_tm_size.AddString(L"220.0");
        m_combo_tm_size.AddString(L"225.0");
        m_combo_tm_size.AddString(L"230.0");
        m_combo_tm_size.AddString(L"235.0");
        m_combo_tm_size.AddString(L"240.0");
        m_combo_tm_size.AddString(L"245.0");
        m_combo_tm_size.AddString(L"250.0");
        m_combo_tm_size.AddString(L"255.0");
        m_combo_tm_size.AddString(L"260.0");
        m_combo_tm_size.AddString(L"265.0");
        m_combo_tm_size.AddString(L"270.0");
        m_combo_tm_size.AddString(L"275.0");
        m_combo_tm_size.AddString(L"280.0");
        m_combo_tm_size.AddString(L"285.0");
        m_combo_tm_size.AddString(L"290.0");
        m_combo_tm_size.AddString(L"295.0");
        m_combo_tm_size.AddString(L"300.0");
        m_combo_tm_size.AddString(L"305.0");
        m_combo_tm_size.AddString(L"310.0");
        m_combo_tm_size.AddString(L"315.0");
        m_combo_tm_size.AddString(L"320.0");
        m_combo_tm_size.AddString(L"325.0");
        m_combo_tm_size.AddString(L"330.0");
        m_combo_tm_size.AddString(L"335.0");
        m_combo_tm_size.AddString(L"340.0");
        m_combo_tm_size.AddString(L"345.0");
        m_combo_tm_size.AddString(L"350.0");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 6 == nCurTm2)
    {//패션잡화 && 남성화(패션잡화)
        m_combo_tm3.AddString(L"정장구두(남성화(패션잡화))");
        m_combo_tm3.AddString(L"수제화(남성화(패션잡화))");
        m_combo_tm3.AddString(L"샌들/슬리퍼(남성화(패션잡화))");
        m_combo_tm3.AddString(L"로퍼/슬립온(남성화(패션잡화))");
        m_combo_tm3.AddString(L"워커(남성화(패션잡화))");
        m_combo_tm3.AddString(L"부츠(남성화(패션잡화))");
        m_combo_tm3.AddString(L"기타(남성화(패션잡화)");

        m_combo_tm_size.AddString(L"200.0");
        m_combo_tm_size.AddString(L"205.0");
        m_combo_tm_size.AddString(L"210.0");
        m_combo_tm_size.AddString(L"215.0");
        m_combo_tm_size.AddString(L"220.0");
        m_combo_tm_size.AddString(L"225.0");
        m_combo_tm_size.AddString(L"230.0");
        m_combo_tm_size.AddString(L"235.0");
        m_combo_tm_size.AddString(L"240.0");
        m_combo_tm_size.AddString(L"245.0");
        m_combo_tm_size.AddString(L"250.0");
        m_combo_tm_size.AddString(L"255.0");
        m_combo_tm_size.AddString(L"260.0");
        m_combo_tm_size.AddString(L"265.0");
        m_combo_tm_size.AddString(L"270.0");
        m_combo_tm_size.AddString(L"275.0");
        m_combo_tm_size.AddString(L"280.0");
        m_combo_tm_size.AddString(L"285.0");
        m_combo_tm_size.AddString(L"290.0");
        m_combo_tm_size.AddString(L"295.0");
        m_combo_tm_size.AddString(L"300.0");
        m_combo_tm_size.AddString(L"305.0");
        m_combo_tm_size.AddString(L"310.0");
        m_combo_tm_size.AddString(L"315.0");
        m_combo_tm_size.AddString(L"320.0");
        m_combo_tm_size.AddString(L"325.0");
        m_combo_tm_size.AddString(L"330.0");
        m_combo_tm_size.AddString(L"335.0");
        m_combo_tm_size.AddString(L"340.0");
        m_combo_tm_size.AddString(L"345.0");
        m_combo_tm_size.AddString(L"350.0");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 7 == nCurTm2)
    {//패션잡화 && 지갑(패션잡화)
        m_combo_tm3.AddString(L"여성 장지갑(지갑(패션잡화))");
        m_combo_tm3.AddString(L"여성 중/반지갑(지갑(패션잡화))");
        m_combo_tm3.AddString(L"남자 장지갑(지갑(패션잡화))");
        m_combo_tm3.AddString(L"남자 중/반지갑(지갑(패션잡화))");
        m_combo_tm3.AddString(L"카드/명합 지갑(지갑(패션잡화))");
        m_combo_tm3.AddString(L"기타(지갑)(지갑(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 8 == nCurTm2)
    {//패션잡화 && 모자(패션잡화)
        m_combo_tm3.AddString(L"야구모자/군모(모자(패션잡화))");
        m_combo_tm3.AddString(L"스냅백(모자(패션잡화))");
        m_combo_tm3.AddString(L"패션모자(모자(패션잡화))");
        m_combo_tm3.AddString(L"왕골/바캉스모자(모자(패션잡화))");
        m_combo_tm3.AddString(L"비니(모자(패션잡화))");
        m_combo_tm3.AddString(L"털/방울(모자(패션잡화))");
        m_combo_tm3.AddString(L"기타(모자)(모자(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 9 == nCurTm2)
    {//패션잡화 && 안경/선글라스(패션잡화)
        m_combo_tm3.AddString(L"안경(뿔테)(안경/선글라스(패션잡화))");
        m_combo_tm3.AddString(L"안경(금속테)(안경/선글라스(패션잡화))");
        m_combo_tm3.AddString(L"선글라스(안경/선글라스(패션잡화))");
        m_combo_tm3.AddString(L"기타(안경/선글라스(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 10 == nCurTm2)
    {//패션잡화 && 주얼리/액세서리(패션잡화)
        m_combo_tm3.AddString(L"반지(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"귀걸이(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"목걸이(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"팔찌/발찌(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"귀금속/보석(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"커플용 주얼리(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"남성용 주얼리(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"헤어 액세서리(주얼리/액세서리(패션잡화))");
        m_combo_tm3.AddString(L"기타(안경/주얼리/액세서리(패션잡화)");

        m_combo_tm_size.AddString(L"1");
        m_combo_tm_size.AddString(L"2");
        m_combo_tm_size.AddString(L"3");
        m_combo_tm_size.AddString(L"4");
        m_combo_tm_size.AddString(L"5");
        m_combo_tm_size.AddString(L"6");
        m_combo_tm_size.AddString(L"7");
        m_combo_tm_size.AddString(L"8");
        m_combo_tm_size.AddString(L"9");
        m_combo_tm_size.AddString(L"10");
        m_combo_tm_size.AddString(L"11");
        m_combo_tm_size.AddString(L"12");
        m_combo_tm_size.AddString(L"13");
        m_combo_tm_size.AddString(L"14");
        m_combo_tm_size.AddString(L"15");
        m_combo_tm_size.AddString(L"17");
        m_combo_tm_size.AddString(L"18");
        m_combo_tm_size.AddString(L"19");
        m_combo_tm_size.AddString(L"20");
        m_combo_tm_size.AddString(L"21");
        m_combo_tm_size.AddString(L"22");
        m_combo_tm_size.AddString(L"23");
        m_combo_tm_size.AddString(L"24");
        m_combo_tm_size.AddString(L"25");
        m_combo_tm_size.AddString(L"26");
        m_combo_tm_size.AddString(L"27");
        m_combo_tm_size.AddString(L"28");
        m_combo_tm_size.AddString(L"29");
        m_combo_tm_size.AddString(L"30");
        m_combo_tm_size.AddString(L"FREE");
        m_combo_tm_size.AddString(L"알수없음");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 11 == nCurTm2)
    {//패션잡화 && 시계(패션잡화)
        m_combo_tm3.AddString(L"메탈시계(시계(패션잡화))");
        m_combo_tm3.AddString(L"가죽시계(시계(패션잡화)");
        m_combo_tm3.AddString(L"스포츠/방수시계(시계(패션잡화)");
        m_combo_tm3.AddString(L"뱅글/팔찌형시계(시계(패션잡화)");
        m_combo_tm3.AddString(L"젤리/우레탄시계(시계(패션잡화)");
        m_combo_tm3.AddString(L"기타(시계(패션잡화)");

        m_combo_tm_size.AddString(L"남성");
        m_combo_tm_size.AddString(L"여성");
        m_combo_tm_size.AddString(L"공용");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (3 == nCurTm1 && 12 == nCurTm2)
    {//패션잡화 && 벨트/장갑/스타킹/기타(패션잡화)
        m_combo_tm3.AddString(L"여성 벨트(패션잡화)");
        m_combo_tm3.AddString(L"남성 벨트(패션잡화)");
        m_combo_tm3.AddString(L"스카프/머플러(패션잡화)");
        m_combo_tm3.AddString(L"넥타이(패션잡화)");
        m_combo_tm3.AddString(L"장갑(패션잡화)");
        m_combo_tm3.AddString(L"양말/스타킹(패션잡화)");
        m_combo_tm3.AddString(L"우산/양산(패션잡화)");
        m_combo_tm3.AddString(L"기타(잡화)(패션잡화)");

        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }


    // 유아동/출산
    else if (4 == nCurTm1 && 1 == nCurTm2)
    {//유아동/출산 && 베이비의류(0-2세)(유아동/출산)
        m_combo_tm3.AddString(L"유아내의/속옷(유아동/출산)");
        m_combo_tm3.AddString(L"유아상의(유아동/출산)");
        m_combo_tm3.AddString(L"유아하의(유아동/출산)");
        m_combo_tm3.AddString(L"우주복/슈트(유아동/출산)");
        m_combo_tm3.AddString(L"자켓/점퍼(유아동/출산)");
        m_combo_tm3.AddString(L"원피스(유아동/출산)");
        m_combo_tm3.AddString(L"정장/드레스(유아동/출산)");
        m_combo_tm3.AddString(L"베이비수영복(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"50");
        m_combo_tm_size.AddString(L"60");
        m_combo_tm_size.AddString(L"70");
        m_combo_tm_size.AddString(L"80");
        m_combo_tm_size.AddString(L"90");
        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"3M");
        m_combo_tm_size.AddString(L"6M");
        m_combo_tm_size.AddString(L"9M");
        m_combo_tm_size.AddString(L"12M");
        m_combo_tm_size.AddString(L"18M");
        m_combo_tm_size.AddString(L"24M");
        m_combo_tm_size.AddString(L"2T");
        m_combo_tm_size.AddString(L"3T");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 2 == nCurTm2)
    {//유아동/출산 && 여아의류(3-6세)(유아동/출산)
        m_combo_tm3.AddString(L"티셔츠(유아동/출산)");
        m_combo_tm3.AddString(L"팬츠(유아동/출산)");
        m_combo_tm3.AddString(L"원피스(유아동/출산)");
        m_combo_tm3.AddString(L"블라우스(유아동/출산)");
        m_combo_tm3.AddString(L"니트/스웨터(유아동/출산)");
        m_combo_tm3.AddString(L"스커트/치마(유아동/출산)");
        m_combo_tm3.AddString(L"가디건/조끼(유아동/출산)");
        m_combo_tm3.AddString(L"자켓/점퍼(유아동/출산)");
        m_combo_tm3.AddString(L"코트/정장(유아동/출산)");
        m_combo_tm3.AddString(L"상하복세트(유아동/출산)");
        m_combo_tm3.AddString(L"속옷/잠옷(유아동/출산)");
        m_combo_tm3.AddString(L"수영복(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"120");
        m_combo_tm_size.AddString(L"130");
        m_combo_tm_size.AddString(L"140");
        m_combo_tm_size.AddString(L"3호");
        m_combo_tm_size.AddString(L"5호");
        m_combo_tm_size.AddString(L"7호");
        m_combo_tm_size.AddString(L"9호");
        m_combo_tm_size.AddString(L"11호");
        m_combo_tm_size.AddString(L"13호");
        m_combo_tm_size.AddString(L"15호");
        m_combo_tm_size.AddString(L"3T");
        m_combo_tm_size.AddString(L"4T");
        m_combo_tm_size.AddString(L"5T");
        m_combo_tm_size.AddString(L"6T");
        m_combo_tm_size.AddString(L"7T");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 3 == nCurTm2)
    {//유아동/출산 && 남아의류(3-6세)(유아동/출산)
        m_combo_tm3.AddString(L"티셔츠(유아동/출산)");
        m_combo_tm3.AddString(L"팬츠(유아동/출산)");
        m_combo_tm3.AddString(L"셔츠/남방(유아동/출산)");
        m_combo_tm3.AddString(L"니트/스웨터(유아동/출산)");
        m_combo_tm3.AddString(L"가디건/조끼(유아동/출산)");
        m_combo_tm3.AddString(L"자켓/점퍼(유아동/출산)");
        m_combo_tm3.AddString(L"코트/정장(유아동/출산)");
        m_combo_tm3.AddString(L"상하복세트(유아동/출산)");
        m_combo_tm3.AddString(L"속옷/잠옷(유아동/출산)");
        m_combo_tm3.AddString(L"수영복(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"100");
        m_combo_tm_size.AddString(L"110");
        m_combo_tm_size.AddString(L"120");
        m_combo_tm_size.AddString(L"130");
        m_combo_tm_size.AddString(L"140");
        m_combo_tm_size.AddString(L"3호");
        m_combo_tm_size.AddString(L"5호");
        m_combo_tm_size.AddString(L"7호");
        m_combo_tm_size.AddString(L"9호");
        m_combo_tm_size.AddString(L"11호");
        m_combo_tm_size.AddString(L"13호");
        m_combo_tm_size.AddString(L"15호");
        m_combo_tm_size.AddString(L"3T");
        m_combo_tm_size.AddString(L"4T");
        m_combo_tm_size.AddString(L"5T");
        m_combo_tm_size.AddString(L"6T");
        m_combo_tm_size.AddString(L"7T");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 4 == nCurTm2)
    {//유아동/출산 && 여주니어의류(7세~)(유아동/출산)
        m_combo_tm3.AddString(L"티셔츠(유아동/출산)");
        m_combo_tm3.AddString(L"팬츠(유아동/출산)");
        m_combo_tm3.AddString(L"원피스(유아동/출산)");
        m_combo_tm3.AddString(L"블라우스(유아동/출산)");
        m_combo_tm3.AddString(L"니트/스웨터(유아동/출산)");
        m_combo_tm3.AddString(L"스커트/치마(유아동/출산)");
        m_combo_tm3.AddString(L"가디건/조끼(유아동/출산)");
        m_combo_tm3.AddString(L"자켓/점퍼(유아동/출산)");
        m_combo_tm3.AddString(L"코트/정장(유아동/출산)");
        m_combo_tm3.AddString(L"상하복세트(유아동/출산)");
        m_combo_tm3.AddString(L"속옷/잠옷(유아동/출산)");
        m_combo_tm3.AddString(L"수영복(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"140");
        m_combo_tm_size.AddString(L"150");
        m_combo_tm_size.AddString(L"160");
        m_combo_tm_size.AddString(L"9호");
        m_combo_tm_size.AddString(L"11호");
        m_combo_tm_size.AddString(L"13호");
        m_combo_tm_size.AddString(L"15호");
        m_combo_tm_size.AddString(L"17호");
        m_combo_tm_size.AddString(L"3T");
        m_combo_tm_size.AddString(L"4T");
        m_combo_tm_size.AddString(L"5T");
        m_combo_tm_size.AddString(L"6T");
        m_combo_tm_size.AddString(L"7T");
        m_combo_tm_size.AddString(L"8T");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 5 == nCurTm2)
    {//유아동/출산 && 남주니어의류(7세~)(유아동/출산)
        m_combo_tm3.AddString(L"티셔츠(유아동/출산)");
        m_combo_tm3.AddString(L"팬츠(유아동/출산)");
        m_combo_tm3.AddString(L"셔츠/남방(유아동/출산)");
        m_combo_tm3.AddString(L"니트/스웨터(유아동/출산)");
        m_combo_tm3.AddString(L"가디건/조끼(유아동/출산)");
        m_combo_tm3.AddString(L"자켓/점퍼(유아동/출산)");
        m_combo_tm3.AddString(L"코트/정장(유아동/출산)");
        m_combo_tm3.AddString(L"상하복세트(유아동/출산)");
        m_combo_tm3.AddString(L"속옷/잠옷(유아동/출산)");
        m_combo_tm3.AddString(L"수영복(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"140");
        m_combo_tm_size.AddString(L"150");
        m_combo_tm_size.AddString(L"160");
        m_combo_tm_size.AddString(L"9호");
        m_combo_tm_size.AddString(L"11호");
        m_combo_tm_size.AddString(L"13호");
        m_combo_tm_size.AddString(L"15호");
        m_combo_tm_size.AddString(L"17호");
        m_combo_tm_size.AddString(L"3T");
        m_combo_tm_size.AddString(L"4T");
        m_combo_tm_size.AddString(L"5T");
        m_combo_tm_size.AddString(L"6T");
        m_combo_tm_size.AddString(L"7T");
        m_combo_tm_size.AddString(L"8T");
        m_combo_tm_size.AddString(L"알수없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 6 == nCurTm2)
    {//유아동/출산 && 유아동신발/잡화(유아동/출산)
        m_combo_tm3.AddString(L"신발(유아동/출산)");
        m_combo_tm3.AddString(L"가방(유아동/출산)");
        m_combo_tm3.AddString(L"모자(유아동/출산)");
        m_combo_tm3.AddString(L"양말(유아동/출산)");
        m_combo_tm3.AddString(L"시계(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"80.0");
        m_combo_tm_size.AddString(L"85.0");
        m_combo_tm_size.AddString(L"90.0");
        m_combo_tm_size.AddString(L"95.0");
        m_combo_tm_size.AddString(L"100.0");
        m_combo_tm_size.AddString(L"105.0");
        m_combo_tm_size.AddString(L"110.0");
        m_combo_tm_size.AddString(L"115.0");
        m_combo_tm_size.AddString(L"120.0");
        m_combo_tm_size.AddString(L"125.0");
        m_combo_tm_size.AddString(L"130.0");
        m_combo_tm_size.AddString(L"135.0");
        m_combo_tm_size.AddString(L"140.0");
        m_combo_tm_size.AddString(L"145.0");
        m_combo_tm_size.AddString(L"150.0");
        m_combo_tm_size.AddString(L"155.0");
        m_combo_tm_size.AddString(L"160.0");
        m_combo_tm_size.AddString(L"165.0");
        m_combo_tm_size.AddString(L"170.0");
        m_combo_tm_size.AddString(L"175.0");
        m_combo_tm_size.AddString(L"180.0");
        m_combo_tm_size.AddString(L"185.0");
        m_combo_tm_size.AddString(L"190.0");
        m_combo_tm_size.AddString(L"195.0");
        m_combo_tm_size.AddString(L"200.0");
        m_combo_tm_size.AddString(L"205.0");
        m_combo_tm_size.AddString(L"210.0");
        m_combo_tm_size.AddString(L"215.0");
        m_combo_tm_size.AddString(L"220.0");
        m_combo_tm_size.AddString(L"225.0");
        m_combo_tm_size.AddString(L"230.0");
        m_combo_tm_size.AddString(L"235.0");
        m_combo_tm_size.AddString(L"240.0");
        m_combo_tm_size.AddString(L"245.0");
        m_combo_tm_size.AddString(L"250.0");
        m_combo_tm_size.AddString(L"255.0");


        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 7 == nCurTm2)
    {//유아동/출산 && 유아동용품(유아동/출산)
        m_combo_tm3.AddString(L"유모차(유아동/출산)");
        m_combo_tm3.AddString(L"카시트(유아동/출산)");
        m_combo_tm3.AddString(L"아기띠(유아동/출산)");
        m_combo_tm3.AddString(L"보행기/쏘서(유아동/출산)");
        m_combo_tm3.AddString(L"가구/침대(유아동/출산)");
        m_combo_tm3.AddString(L"스킨케어(유아동/출산)");
        m_combo_tm3.AddString(L"목욕/구강용품(유아동/출산)");
        m_combo_tm3.AddString(L"세탁/위생용품(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"0~6개월");
        m_combo_tm_size.AddString(L"7~12개월");
        m_combo_tm_size.AddString(L"13~24개월");
        m_combo_tm_size.AddString(L"25~36개월");
        m_combo_tm_size.AddString(L"37~48개월");
        m_combo_tm_size.AddString(L"5~7세");
        m_combo_tm_size.AddString(L"8~10세");
        m_combo_tm_size.AddString(L"11~13세");
        m_combo_tm_size.AddString(L"전체이용가");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 8 == nCurTm2)
    {//유아동/출산 && 출산/임부용품(유아동/출산)
        m_combo_tm3.AddString(L"겉싸개/속싸개(유아동/출산)");
        m_combo_tm3.AddString(L"배냇저고리(유아동/출산)");
        m_combo_tm3.AddString(L"딸랑이/모빌(유아동/출산)");
        m_combo_tm3.AddString(L"손발싸개(유아동/출산)");
        m_combo_tm3.AddString(L"이불/침구(유아동/출산)");
        m_combo_tm3.AddString(L"임부의류/속옷(유아동/출산)");
        m_combo_tm3.AddString(L"임부스킨케어(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 9 == nCurTm2)
    {//유아동/출산 && 교육/완구/인형(유아동/출산)
        m_combo_tm3.AddString(L"신생아완구(유아동/출산)");
        m_combo_tm3.AddString(L"교육완구(유아동/출산)");
        m_combo_tm3.AddString(L"도서/CD(유아동/출산)");
        m_combo_tm3.AddString(L"인형(유아용)(유아동/출산)");
        m_combo_tm3.AddString(L"자전거(유아동/출산)");
        m_combo_tm3.AddString(L"볼텐트/놀이터(유아동/출산)");
        m_combo_tm3.AddString(L"퍼즐/블록(유아동/출산)");
        m_combo_tm3.AddString(L"물놀이용품(유아동/출산)");
        m_combo_tm3.AddString(L"스포츠완구(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"0~6개월");
        m_combo_tm_size.AddString(L"7~12개월");
        m_combo_tm_size.AddString(L"13~24개월");
        m_combo_tm_size.AddString(L"25~36개월");
        m_combo_tm_size.AddString(L"37~48개월");
        m_combo_tm_size.AddString(L"5~7세");
        m_combo_tm_size.AddString(L"8~10세");
        m_combo_tm_size.AddString(L"11~13세");
        m_combo_tm_size.AddString(L"전체이용가");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    else if (4 == nCurTm1 && 10 == nCurTm2)
    {//유아동/출산 && 기저귀/수유/이유식(유아동/출산)
        m_combo_tm3.AddString(L"기저귀(유아동/출산)");
        m_combo_tm3.AddString(L"물티슈(유아동/출산)");
        m_combo_tm3.AddString(L"분유수유용품(유아동/출산)");
        m_combo_tm3.AddString(L"모유수유용품(유아동/출산)");
        m_combo_tm3.AddString(L"이유식용품(유아동/출산)");
        m_combo_tm3.AddString(L"젖병세정용품(유아동/출산)");
        m_combo_tm3.AddString(L"기타(유아동/출산)");


        m_combo_tm_size.AddString(L"없음");

        m_combo_tm3.SetCurSel(0);
        m_combo_tm_size.SetCurSel(0);
    }

    m_combo_tm3.SetCurSel(idxTm3);
    m_combo_tm_size.SetCurSel(idxTmSize);
}



void CETMacroDlg::OnBnClickedButtonExit()
{
    if (m_pMainThread != NULL)
    {
        m_pMainThread = NULL;
    }

    exit(0);
}


void CETMacroDlg::OnBnClickedCheckSkip()
{
    
}


void CETMacroDlg::OnBnClickedRadioSpeed0()
{
    theApp.m_fMacroSpeed = 1.f;
}


void CETMacroDlg::OnBnClickedRadioSpeed1()
{
    theApp.m_fMacroSpeed = 0.9f;
}


void CETMacroDlg::OnBnClickedRadioSpeed2()
{
    theApp.m_fMacroSpeed = 0.8f;
}
