
// ETMacro.h : PROJECT_NAME ���� ���α׷��� ���� �� ��� �����Դϴ�.
//

#pragma once

#ifndef __AFXWIN_H__
#error "PCH�� ���� �� ������ �����ϱ� ���� 'stdafx.h'�� �����մϴ�."
#endif

#include "resource.h"		// �� ��ȣ�Դϴ�.


// CETMacroApp:
// �� Ŭ������ ������ ���ؼ��� ETMacro.cpp�� �����Ͻʽÿ�.
//

enum eItemStatus
{
    EItemState_0 = 0,   // �߰�
    EItemState_1,       // �߰� + ����
    EItemState_2,       // ����ǰ
    EItemState_3,       // ���� + ����
    EItemState_4,       // ���� ����

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
    CString title;      // ������
    std::vector<CString> vecDescription;// �۳��� (�ٸ��� ���Ϳ� �߰�)

    CString bandURL;    // ����ּ�
    int unit;           // ����
    int cost;           // ����

    std::vector<CString> vecTag;        // �ױ�

                                        // �������� 1��° �޺��ڽ�
    int nThunderMarketFirstComboIndex;
    // �������� 2��° �޺��ڽ�
    int nThunderMarketSecondComboIndex;
    // �������� 3��° �޺��ڽ�
    int nThunderMarketThirdComboIndex;

    // ��θ��� 1��° �޺��ڽ�
    int nHelloMarketFirstComboIndex;
    // ��θ��� 2��° �޺��ڽ�
    int nHelloMarketSecondComboIndex;

    // �߰��� 1��° �޺��ڽ�
    int nNaverCafeFirstComboIndex;

    // �̹����� 6���� ������ ����
    CString images;
    CString imagePath;

    // ������
    int nSizeComboIndex;

    // ����
    eItemStatus itemStatus;

    // �ù�� ����
    BOOL bDeliveryCost;

    // ��ȯ����
    BOOL bExchange;

    // �߰��� ����
    BOOL bConnectUsedWorld;

    // ��� ����
    BOOL bConnectNaverBand;

    // ��θ��� ��� ����
    BOOL bHelloMarketAccept;

    // �������� ��� ����
    BOOL bThunderMarketAccept;

    // ���̹�ī�� ��� ����
    BOOL bNaverCafeAccept;

    // ���̹���� ��� ����
    BOOL bNaverBandAccept;

    // īī�����丮 ��� ����
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

    // �������Դϴ�.
public:
    virtual BOOL InitInstance();

    // �����Դϴ�.

    DECLARE_MESSAGE_MAP()

public:
    std::vector<account> m_account;
    std::vector<sellItem> m_sellItems;


    CString m_URLChrome;

    float m_fMacroSpeed;
};

extern CETMacroApp theApp;