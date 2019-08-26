
// ETMacro.h : PROJECT_NAME 응용 프로그램에 대한 주 헤더 파일입니다.
//

#pragma once

#ifndef __AFXWIN_H__
#error "PCH에 대해 이 파일을 포함하기 전에 'stdafx.h'를 포함합니다."
#endif

#include "resource.h"		// 주 기호입니다.


// CETMacroApp:
// 이 클래스의 구현에 대해서는 ETMacro.cpp을 참조하십시오.
//

enum eItemStatus
{
    EItemState_0 = 0,   // 중고
    EItemState_1,       // 중고 + 하자
    EItemState_2,       // 새물품
    EItemState_3,       // 새것 + 하자
    EItemState_4,       // 거의 새것

};

struct account
{
    CString webName;
    CString webURL;

    CString id;
    CString password;
};

struct sellItem
{
    CString title;      // 글제목
    std::vector<CString> vecDescription;// 글내용 (줄마다 벡터에 추가)

    CString bandURL;    // 밴드주소
    int unit;           // 수량
    int cost;           // 가격

    std::vector<CString> vecTag;        // 테그

                                        // 번개장터 1번째 콤보박스
    int nThunderMarketFirstComboIndex;
    // 번개장터 2번째 콤보박스
    int nThunderMarketSecondComboIndex;
    // 번개장터 3번째 콤보박스
    int nThunderMarketThirdComboIndex;

    // 헬로마켓 1번째 콤보박스
    int nHelloMarketFirstComboIndex;
    // 헬로마켓 2번째 콤보박스
    int nHelloMarketSecondComboIndex;

    // 중고나라 1번째 콤보박스
    int nNaverCafeFirstComboIndex;

    // 이미지는 6개로 무조껀 고정
    CString images;
    CString imagePath;

    // 상세정보
    int nSizeComboIndex;

    // 상태
    eItemStatus itemStatus;

    // 택배비 포함
    BOOL bDeliveryCost;

    // 교환가능
    BOOL bExchange;

    // 중고나라 접속
    BOOL bConnectUsedWorld;

    // 밴드 접속
    BOOL bConnectNaverBand;

    // 헬로마켓 등록 유무
    BOOL bHelloMarketAccept;

    // 번개장터 등록 유무
    BOOL bThunderMarketAccept;

    // 네이버카페 등록 유무
    BOOL bNaverCafeAccept;

    // 네이버밴드 등록 유무
    BOOL bNaverBandAccept;

    // 카카오스토리 등록 유무
    BOOL bKakaoStoryAccept;

    sellItem()
    {
        unit = 1;
        cost = 0;

        nThunderMarketFirstComboIndex = 0;
        nThunderMarketSecondComboIndex = 0;
        nThunderMarketThirdComboIndex = 0;

        nHelloMarketFirstComboIndex = 0;
        nHelloMarketSecondComboIndex = 0;

        nNaverCafeFirstComboIndex = 0;
        nSizeComboIndex = 0;

        itemStatus = EItemState_0;

        bDeliveryCost = FALSE;
        bExchange = FALSE;
        bConnectUsedWorld = FALSE;
        bConnectNaverBand = FALSE;


        bHelloMarketAccept = TRUE;
        bThunderMarketAccept = TRUE;
        bNaverCafeAccept = TRUE;
        bNaverBandAccept = TRUE;
        bKakaoStoryAccept = TRUE;
    }

};

class CETMacroApp : public CWinApp
{
public:
    CETMacroApp();

    // 재정의입니다.
public:
    virtual BOOL InitInstance();

    // 구현입니다.

    DECLARE_MESSAGE_MAP()

public:
    std::vector<account> m_account;
    std::vector<sellItem> m_sellItems;


    CString m_URLChrome;

    float m_fMacroSpeed;
};

extern CETMacroApp theApp;