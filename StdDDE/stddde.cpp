// stddde.cpp
#include "stdafx.h"

#include "stddde.h"

#include <algorithm>
#include <vector>

#define _CRTDBG_MAP_ALLOC
#include <stdlib.h>
#include <crtdbg.h>

#ifdef _DEBUG
#define DDE_DEBUG_NEW new ( _NORMAL_BLOCK , __FILE__ , __LINE__ )
#define new DDE_DEBUG_NEW
#else
#define DDE_DEBUG_NEW new
#endif

//
// Format lists
//

static const WORD SysFormatList[] = {
    CF_TEXT,
    CF_UNICODETEXT,
    NULL
};

//
// Structure used to hold a clipboard id and its text name
//

typedef struct _CFTAGNAME {
    WORD    wFmt;
    LPCTSTR pszName;
} CFTAGNAME, FAR *PCFTAGNAME;

//
// Standard format name lookup table
//

CFTAGNAME CFNames[] = {
    CF_TEXT,        SZCF_TEXT,
    CF_BITMAP,      SZCF_BITMAP,
    CF_METAFILEPICT,SZCF_METAFILEPICT,
    CF_SYLK,        SZCF_SYLK,
    CF_DIF,         SZCF_DIF,
    CF_TIFF,        SZCF_TIFF,
    CF_OEMTEXT,     SZCF_OEMTEXT,
    CF_DIB,         SZCF_DIB,
    CF_PALETTE,     SZCF_PALETTE,
    CF_PENDATA,     SZCF_PENDATA,
    CF_RIFF,        SZCF_RIFF,
    CF_WAVE,        SZCF_WAVE,
    CF_UNICODETEXT, SZCF_UNICODETEXT,
    NULL,           NULL
    };

////////////////////////////////////////////////////////////////////////////////////
//
// ********** The Barfy bit *********************
//
// We only support one server per app right now
// sooooo: here's the global we use to find it in the
// hateful DDE callback routine
// Let's see if I can get away with this

static CDDEServer* pTheServer = NULL;

////////////////////////////////////////////////////////////////////////////////////
//
// CDDECountedObject

CDDECountedObject::CDDECountedObject()
{
    m_iRefCount = 0;
}

CDDECountedObject::~CDDECountedObject()
{
    _ASSERT(m_iRefCount == 0);
}

int CDDECountedObject::AddRef()
{
    _ASSERT(m_iRefCount < 1000); // sanity check
    return ++m_iRefCount;
}

int CDDECountedObject::Release()
{
    int i = --m_iRefCount;
    _ASSERT(m_iRefCount >= 0);
    if (m_iRefCount == 0) {
        delete this;
    }
    return i;
}

////////////////////////////////////////////////////////////////////////////////////
//
// CHSZ

CHSZ::CHSZ()
{
    m_hsz = NULL;
    m_dwDDEInstance = 0;
}

CHSZ::CHSZ(CDDEServer* pServer, LPCTSTR szName)
{
    m_dwDDEInstance = pServer->m_dwDDEInstance;
    m_hsz = ::DdeCreateStringHandle(m_dwDDEInstance,
                                    szName,
                                    DDE_CODEPAGE);
    _ASSERT(m_hsz);
}

void CHSZ::Create(CDDEServer* pServer, LPCTSTR szName)
{
    if (m_hsz) {
        ::DdeFreeStringHandle(pServer->m_dwDDEInstance, m_hsz);
    }
    m_dwDDEInstance = pServer->m_dwDDEInstance;
    m_hsz = ::DdeCreateStringHandle(m_dwDDEInstance,
                                    szName,
                                    DDE_CODEPAGE);
    _ASSERT(m_hsz);
}

CHSZ::~CHSZ()
{
    if (m_hsz) {
        ::DdeFreeStringHandle(m_dwDDEInstance, m_hsz);
    }
}

////////////////////////////////////////////////////////////////////////////////////
//
// CDDEItem

CDDEItem::CDDEItem()
{
    m_pTopic = NULL;
}

CDDEItem::~CDDEItem()
{
}

void CDDEItem::Create(LPCTSTR pszName)
{
    m_strName = pszName;
}

BOOL CDDEItem::Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData)
{
    return FALSE;
}

BOOL CDDEItem::Poke(UINT wFmt, void* pData, DWORD dwSize)
{
    return FALSE;
}

BOOL CDDEItem::IsSupportedFormat(WORD wFormat) const
{
    const WORD* pFmt = GetFormatList();
    if (!pFmt) return FALSE;
    while(*pFmt) {
        if (*pFmt == wFormat) return TRUE;
        pFmt++;
    }
    return FALSE;
}

BOOL CDDEItem::CanAdvise(UINT wFmt) const
{
    return IsSupportedFormat(wFmt);
}

void CDDEItem::PostAdvise()
{
    if (m_pTopic == NULL) return;
    m_pTopic->PostAdvise(this);
}

HDDEDATA CDDEItem::CreateDdeData(UINT wFmt, HSZ hszItem, void* pData, DWORD dwSize) const
{
    HDDEDATA hData = ::DdeCreateDataHandle(m_pTopic->m_pServer->m_dwDDEInstance,
                                           (LPBYTE)pData, dwSize, 0,
                                           hszItem, wFmt, 0);
    return hData;
}

HDDEDATA CDDEItem::CreateDdeData(UINT wFmt, HSZ hszItem, const CDDEString& s) const
{
    if (wFmt == CF_TEXT) {
        return CreateDdeDataA(hszItem, s);
    }
    else if (wFmt == CF_UNICODETEXT) {
        return CreateDdeDataW(hszItem, s);
    }
    return FALSE;
}

HDDEDATA CDDEItem::CreateDdeDataA(HSZ hszItem, const CDDEString& s) const
{
#ifdef _UNICODE
    int bytes = WideCharToMultiByte(CP_ACP, 0, s.c_str(), (int)s.size(),
        NULL, 0, NULL, NULL);
    std::vector<char> buf(bytes + 1);
    buf[bytes] = 0;
    WideCharToMultiByte(CP_ACP, 0, s.c_str(), (int)s.size(),
        &buf[0], bytes, NULL, NULL);
    return CreateDdeData(CF_TEXT, hszItem, &buf[0], bytes + 1);
#else
    return CreateDdeData(CF_TEXT, hszItem, (void*)s.c_str(), (DWORD)(s.size() + 1));
#endif
}

HDDEDATA CDDEItem::CreateDdeDataW(HSZ hszItem, const CDDEString& s) const
{
#ifdef _UNICODE
    return CreateDdeData(CF_UNICODETEXT, hszItem, (void*)s.c_str(),
                         (DWORD)((s.size() + 1) * sizeof(TCHAR)));
#else
    int cch = MultiByteToWideChar(CP_ACP, 0, s.c_str(), (int)s.size(), NULL, 0);
    int bytes = (cch + 1) * sizeof(WCHAR);
    std::vector<WCHAR> buf(cch + 1);
    buf[cch] = 0;
    MultiByteToWideChar(CP_ACP, 0, s.c_str(), (int)s.size(), &buf[0], cch);
    return CreateDdeData(CF_UNICODETEXT, hszItem, &buf[0], bytes);
#endif
}

////////////////////////////////////////////////////////////////////////////////////
//
// CDDEStringItem

BOOL CDDEStringItem::Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData)
{
    if (!IsSupportedFormat(wFmt)) return FALSE;
    *phReturnData = CreateDdeData(wFmt, hszItem, m_strData);
    return TRUE;
}

BOOL CDDEStringItem::Poke(UINT wFmt, void* pData, DWORD dwSize)
{
    if (!IsSupportedFormat(wFmt)) return FALSE;
    _ASSERT(pData);
    if (wFmt == DDE_CF_TEXT) {
        m_strData = (const TCHAR*)pData;
    }
    else {
#ifdef _UNICODE
        // convert ANSI to UNICODE
        int cch = MultiByteToWideChar(CP_ACP, 0, (LPCCH)pData, dwSize, NULL, 0);
        m_strData.resize(cch);
        MultiByteToWideChar(CP_ACP, 0, (LPCCH)pData, dwSize, &m_strData[0], cch);
#else
        // convert UNICODE to ANSI
        int bytes = WideCharToMultiByte(CP_ACP, 0, (LPCWCH)pData, dwSize,
                                        NULL, 0, NULL, NULL);
        m_strData.resize(bytes);
        WideCharToMultiByte(CP_ACP, 0, (LPCWCH)pData, dwSize,
                            &m_strData[0], bytes, NULL, NULL);
#endif
    }
    OnPoke();
    return TRUE;
}

const WORD* CDDEStringItem::GetFormatList() const
{
    return SysFormatList;
}

////////////////////////////////////////////////////////////////////////////////////
//
// CDDEItemList

////////////////////////////////////////////////////////////////////////////////////
//
// CDDETopic

CDDETopic::CDDETopic()
{
}

CDDETopic::~CDDETopic()
{
}

void CDDETopic::Create(LPCTSTR pszName)
{
    m_strName = pszName;
}

BOOL CDDETopic::AddItem(CDDEItem* pNewItem)
{
    _ASSERT(pNewItem);

    //
    // See if we already have this item
    //

    CDDEItemList::iterator it = std::find(m_ItemList.begin(),
                                          m_ItemList.end(),
                                          pNewItem);
    if (it != m_ItemList.end()) return TRUE; // already have it

    //
    // Add the new item
    //

    m_ItemList.push_back(pNewItem);
    pNewItem->m_pTopic = this;

    return TRUE;
}

BOOL CDDETopic::Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData)
{
    //
    // See if we have this item
    //
    CDDEString strItem = m_pServer->StringFromHsz(hszItem);
    CDDEItem* pItem = FindItem(strItem.c_str());
    if (!pItem) return FALSE;

    return pItem->Request(wFmt, hszItem, phReturnData);
}

BOOL CDDETopic::Poke(UINT wFmt, LPCTSTR pszItem,
                     void* pData, DWORD dwSize)
{
    //
    // See if we have this item
    //

    CDDEItem* pItem = FindItem(pszItem);
    if (!pItem) return FALSE;

    return pItem->Poke(wFmt, pData, dwSize);
}

BOOL CDDETopic::Exec(void* pData, DWORD dwSize)
{
    return FALSE;
}

CDDEItem* CDDETopic::FindItem(LPCTSTR pszItem)
{
    CDDEItemList::iterator it = m_ItemList.begin();
    for (; it != m_ItemList.end(); ++it) {
        CDDEItem* pItem = *it;
        if (_tcsicmp(pItem->m_strName.c_str(), pszItem) == 0) return pItem;
    }
    return NULL;
}

const CDDEItem* CDDETopic::FindItem(LPCTSTR pszItem) const
{
    CDDEItemList::const_iterator it = m_ItemList.begin();
    for (; it != m_ItemList.end(); ++it) {
        const CDDEItem* pItem = *it;
        if (_tcsicmp(pItem->m_strName.c_str(), pszItem) == 0) return pItem;
    }
    return NULL;
}


BOOL CDDETopic::CanAdvise(UINT wFmt, LPCTSTR pszItem)
{
    //
    // See if we have this item
    //

    const CDDEItem* pItem = FindItem(pszItem);
    if (!pItem) return FALSE;

    return pItem->CanAdvise(wFmt);
}

void CDDETopic::PostAdvise(CDDEItem* pItem)
{
    _ASSERT(m_pServer);
    _ASSERT(pItem);
    m_pServer->PostAdvise(this, pItem);
}

////////////////////////////////////////////////////////////////////////////////////
//
// CDDETopicList

////////////////////////////////////////////////////////////////////////////////////
//
// CDDEConv

CDDEConv::CDDEConv()
{
    m_pServer = NULL;
    m_hConv = NULL;
    m_hszTopic = NULL;
}

CDDEConv::CDDEConv(CDDEServer* pServer)
{
    m_pServer = pServer;
    m_hConv = NULL;
    m_hszTopic = NULL;
}

CDDEConv::CDDEConv(CDDEServer* pServer, HCONV hConv, HSZ hszTopic)
{
    m_pServer = pServer;
    m_hConv = hConv;
    m_hszTopic = hszTopic;
}

CDDEConv::~CDDEConv()
{
    Terminate();
}

BOOL CDDEConv::Terminate()
{
    if (m_hConv) {

        //
        // Terminate this conversation
        //

        if (!::DdeDisconnect(m_hConv)) {
            _ASSERT_EXPR(FALSE, _T("DdeDisconnect Failed."));
        }

        m_hConv = NULL;

        return TRUE;
    }

    return FALSE; // wasn't active
}

BOOL CDDEConv::ConnectTo(LPCTSTR pszService, LPCTSTR pszTopic)
{
    _ASSERT(pszService);
    _ASSERT(pszTopic);
    _ASSERT(m_pServer);
    _ASSERT(!m_hConv);

    CHSZ hszService (m_pServer, pszService);
    CHSZ hszTopic(m_pServer, pszTopic);

    //
    // Try to connect
    //

    DWORD dwErr = 0;
    m_hConv = ::DdeConnect(m_pServer->m_dwDDEInstance,
                           hszService,
                           hszTopic,
                           NULL);

    if (!m_hConv) {

        dwErr = GetLastError();

    }

    if (!m_hConv) {
        m_pServer->Status(_T("Failed to connect to %s|%s. Error %u"),
                          (LPCTSTR) pszService,
                          (LPCTSTR) pszTopic,
                          dwErr);
        return FALSE;
    }

    //
    // Add this conversation to the server list
    //

    m_pServer->AddConversation(this);
    return TRUE;
}

BOOL CDDEConv::AdviseData(UINT wFmt, LPCTSTR pszTopic, LPCTSTR pszItem,
                          void* pData, DWORD dwSize)
{
    return FALSE;
}

BOOL CDDEConv::Request(UINT wFmt, LPCTSTR pszItem, void** ppData, DWORD* pdwSize)
{
    _ASSERT(m_pServer);
    _ASSERT(pszItem);
    _ASSERT(ppData);
    _ASSERT(pdwSize);

    CHSZ hszItem (m_pServer, pszItem);
    HDDEDATA hData = ::DdeClientTransaction(NULL,
                                            0,
                                            m_hConv,
                                            hszItem,
                                            wFmt,
                                            XTYP_REQUEST,
                                            m_pServer->GetTimeout(),
                                            NULL);

    if (!hData) {

        // Failed
        *pdwSize = 0;
        *ppData = NULL;
        return FALSE;
    }

    //
    // Copy the result data
    //

    BYTE* pData = ::DdeAccessData(hData, pdwSize);
    _ASSERT(*pdwSize);
    *ppData = malloc(*pdwSize);
    _ASSERT(*ppData);
    memcpy(*ppData, pData, *pdwSize);
    ::DdeUnaccessData(hData);

    return TRUE;
}

BOOL CDDEConv::Advise(UINT wFmt, LPCTSTR pszItem)
{
    _ASSERT(m_pServer);
    _ASSERT(pszItem);

    CHSZ hszItem (m_pServer, pszItem);
    HDDEDATA hData = ::DdeClientTransaction(NULL,
                                            0,
                                            m_hConv,
                                            hszItem,
                                            wFmt,
                                            XTYP_ADVSTART,
                                            m_pServer->GetTimeout(),
                                            NULL);

    if (!hData) {
        // Failed
        return FALSE;
    }
    return TRUE;
}

BOOL CDDEConv::Exec(UINT wFmt, LPCTSTR pszCmd)
{
    //
    // Send the command
    //

    HDDEDATA hData = ::DdeClientTransaction((BYTE*)pszCmd,
                                            (DWORD)((_tcslen(pszCmd) + 1) * sizeof(TCHAR)),
                                            m_hConv,
                                            0,
                                            wFmt,
                                            XTYP_EXECUTE,
                                            m_pServer->GetTimeout(),
                                            NULL);

    if (!hData) {
        // Failed
        return FALSE;
    }
    return TRUE;
}

BOOL CDDEConv::Poke(UINT wFmt, LPCTSTR pszItem, void* pData, DWORD dwSize)
{
    //
    // Send the command
    //

    CHSZ hszItem (m_pServer, pszItem);
    HDDEDATA hData = ::DdeClientTransaction((BYTE*)pData,
                                            dwSize,
                                            m_hConv,
                                            hszItem,
                                            wFmt,
                                            XTYP_POKE,
                                            m_pServer->GetTimeout(),
                                            NULL);

    if (!hData) {
        // Failed
        return FALSE;
    }
    return TRUE;
}

////////////////////////////////////////////////////////////////////////////////////
//
// CDDEConvList

////////////////////////////////////////////////////////////////////////////////////
//
// Topics and items to support the 'system' topic

//
// Generic system topic items
//

const WORD* CDDESystemItem::GetFormatList() const
{
    return SysFormatList;
}

//
// Specific system topic items
//

BOOL CDDESystemItem_TopicList::Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData)
{
    if (!IsSupportedFormat(wFmt)) {
        return FALSE;
    }

    //
    // Return the list of topics for this service
    //

    CDDEString strTopics;
    strTopics.clear();
    _ASSERT(m_pTopic);
    CDDEServer* pServer = m_pTopic->m_pServer;
    _ASSERT(pServer);
    CDDETopicList::iterator it = pServer->m_TopicList.begin();
    int items = 0;
    for (; it != pServer->m_TopicList.end(); ++it) {

        CDDETopic* pTopic = *it;

        //
        // put in a tab delimiter unless this is the first item
        //

        if (items != 0) strTopics += SZ_TAB;

        //
        // Copy the string name of the item
        //

        strTopics += pTopic->m_strName;

        items++;
    }

    //
    // Set up the return info
    //

    *phReturnData = CreateDdeData(wFmt, hszItem, strTopics);
    return TRUE;
}

BOOL CDDESystemItem_ItemList::Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData)
{
    if (!IsSupportedFormat(wFmt)) {
        return FALSE;
    }

    //
    // Return the list of items in this topic
    //

    CDDEString strItems;
    strItems.clear();
    _ASSERT(m_pTopic);
    CDDEItemList::iterator it = m_pTopic->m_ItemList.begin();
    int items = 0;
    for (; it != m_pTopic->m_ItemList.end(); ++it) {

        CDDEItem* pItem = *it;

        //
        // put in a tab delimiter unless this is the first item
        //

        if (items != 0) strItems += SZ_TAB;

        //
        // Copy the string name of the item
        //

        strItems += pItem->m_strName;

        items++;
    }

    //
    // Set up the return info
    //

    *phReturnData = CreateDdeData(wFmt, hszItem, strItems);
    return TRUE;
}

BOOL CDDESystemItem_FormatList::Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData)
{
    if (!IsSupportedFormat(wFmt)) {
        return FALSE;
    }

    //
    // Return the list of formats in this topic
    //

    CDDEString strFormats;
    strFormats.clear();
    _ASSERT(m_pTopic);
    CDDEItemList::iterator it = m_pTopic->m_ItemList.begin();
    int iFormats = 0;
    WORD wFmtList[100];
    for (; it != m_pTopic->m_ItemList.end(); ++it) {

        CDDEItem* pItem = *it;

        //
        // get the format list for this item
        //

        const WORD* pItemFmts = pItem->GetFormatList();
        if (pItemFmts) {

            //
            // Add each format to the list if we don't have it already
            //

            while (*pItemFmts) {

                //
                // See if we have it
                //

                int i;
                for (i = 0; i < iFormats; i++) {
                    if (wFmtList[i] == *pItemFmts) break; // have it already
                }

                if (i == iFormats) {

                    //
                    // This is a new one
                    //

                    wFmtList[iFormats] = *pItemFmts;

                    //
                    // Add the string name to the list
                    //

                    //
                    // put in a tab delimiter unless this is the first item
                    //

                    if (iFormats != 0) strFormats += SZ_TAB;

                    //
                    // Copy the string name of the item
                    //

                    strFormats += DDEGetFormatName(*pItemFmts);

                    iFormats++;
                }

                pItemFmts++;
            }
        }
    }

    //
    // Set up the return info
    //
    *phReturnData = CreateDdeData(wFmt, hszItem, strFormats);
    return TRUE;
}

BOOL CDDEServerSystemTopic::Request(UINT wFmt, HSZ hszItem, HDDEDATA* phReturnData)
{
    CDDEString strItem = m_pServer->StringFromHsz(hszItem);
    m_pServer->Status(_T("System topic request: %s"), strItem.c_str());
    return CDDETopic::Request(wFmt, hszItem, phReturnData);
}

////////////////////////////////////////////////////////////////////////////////////
//
// CDDEServer

CDDEServer::CDDEServer()
{
    m_bInitialized = FALSE;
    m_dwDDEInstance = 0;
    m_dwTimeout = DDE_TIMEOUT;
}

CDDEServer::~CDDEServer()
{
    Shutdown();
}

void CDDEServer::Shutdown()
{
    if (m_bInitialized) {

        //
        // Terminate all conversations
        //

        CDDEConvList::iterator it = m_ConvList.begin();
        for (; it != m_ConvList.end(); ++it) {
            CDDEConv*& pConv = *it;
            _ASSERT(pConv);
            pConv->Terminate();
            pConv->Release();
            pConv = NULL;
        }
        m_ConvList.clear();

        //
        // Unregister the service name
        //

        ::DdeNameService(m_dwDDEInstance,
                         m_hszServiceName,
                         NULL,
                         DNS_UNREGISTER);

        //
        // Release DDEML
        //

        ::DdeUninitialize(m_dwDDEInstance);

        m_bInitialized = FALSE;
    }
}

BOOL CDDEServer::Create(LPCTSTR pszServiceName,
                        DWORD dwFilterFlags/* = 0 */,
                        DWORD* pdwDDEInst/* = NULL */,
                        DWORD dwTimeout/* = DDE_TIMEOUT */)
{
    //
    // make sure we are alone in the world
    //

    if (pTheServer != NULL) {
        _ASSERT_EXPR(FALSE, _T("Already got a server!"));
        return FALSE;
    } else {
        pTheServer = this;
    }

    //
    // Make sure the application hasn't requested any filter options
    // which will prevent us from working correctly.
    //

    dwFilterFlags &= !( CBF_FAIL_CONNECTIONS
                      | CBF_SKIP_CONNECT_CONFIRMS
                      | CBF_SKIP_DISCONNECTS);

    //
    // Initialize DDEML.  Note that DDEML doesn't make any callbacks
    // during initialization so we don't need to worry about the
    // custom callback yet.
    //

    UINT uiResult;
    uiResult = ::DdeInitialize(&m_dwDDEInstance,
                               (PFNCALLBACK)&StdDDECallback,
                               dwFilterFlags,
                               0);

    if (uiResult != DMLERR_NO_ERROR) return FALSE;

    //
    // Return the DDE instance id if it was requested
    //

    if (pdwDDEInst) {
        *pdwDDEInst = m_dwDDEInstance;
    }

    m_dwTimeout = dwTimeout;

    //
    // Copy the service name and create a DDE name handle for it
    //

    m_strServiceName = pszServiceName;
    m_hszServiceName.Create(this, m_strServiceName.c_str());

    //
    // Add all the system topic to the service tree
    //

    //
    // Create a system topic
    //

    m_SystemTopic.Create(SZDDESYS_TOPIC);
    AddTopic(&m_SystemTopic);

    //
    // Create some system topic items
    //

    m_SystemItemTopics.Create(SZDDESYS_ITEM_TOPICS);
    m_SystemTopic.AddItem(&m_SystemItemTopics);

    m_SystemItemSysItems.Create(SZDDESYS_ITEM_SYSITEMS);
    m_SystemTopic.AddItem(&m_SystemItemSysItems);

    m_SystemItemItems.Create(SZDDE_ITEM_ITEMLIST);
    m_SystemTopic.AddItem(&m_SystemItemItems);

    m_SystemItemFormats.Create(SZDDESYS_ITEM_FORMATS);
    m_SystemTopic.AddItem(&m_SystemItemFormats);

    //
    // Register the name of our service
    //

    ::DdeNameService(m_dwDDEInstance,
                     m_hszServiceName,
                     NULL,
                     DNS_REGISTER);

    m_bInitialized = TRUE;

    //
    // See if any derived class wants to add anything
    //

    return OnCreate();
}

void PrintStatus(CDDEServer* pServ, WORD wType)
{
    switch (wType) {
    case XTYP_REGISTER:
        pServ->Status(_T("Callback %4.4XH: XTYP_REGISTER"), wType);
        break;
    case XTYP_CONNECT_CONFIRM:
        pServ->Status(_T("Callback %4.4XH: XTYP_CONNECT_CONFIRM"), wType);
        break;
    case XTYP_DISCONNECT:
        pServ->Status(_T("Callback %4.4XH: XTYP_DISCONNECT"), wType);
        break;
    case XTYP_WILDCONNECT:
        pServ->Status(_T("Callback %4.4XH: XTYP_WILDCONNECT"), wType);
        break;
    case XTYP_ADVSTART:
        pServ->Status(_T("Callback %4.4XH: XTYP_ADVSTART"), wType);
        break;
    case XTYP_CONNECT:
        pServ->Status(_T("Callback %4.4XH: XTYP_CONNECT"), wType);
        break;
    case XTYP_EXECUTE:
        pServ->Status(_T("Callback %4.4XH: XTYP_EXECUTE"), wType);
        break;
    case XTYP_REQUEST:
        pServ->Status(_T("Callback %4.4XH: XTYP_REQUEST"), wType);
        break;
    case XTYP_ADVREQ:
        pServ->Status(_T("Callback %4.4XH: XTYP_ADVREQ"), wType);
        break;
    case XTYP_ADVDATA:
        pServ->Status(_T("Callback %4.4XH: XTYP_ADVDATA"), wType);
        break;
    case XTYP_POKE:
        pServ->Status(_T("Callback %4.4XH: XTYP_POKE"), wType);
        break;
    default:
        pServ->Status(_T("Callback %4.4XH"), wType);
        break;
    }
}


//
// Callback function
// Note: this is a static
//

HDDEDATA CALLBACK CDDEServer::StdDDECallback(WORD wType,
                                            WORD wFmt,
                                            HCONV hConv,
                                            HSZ hsz1,
                                            HSZ hsz2,
                                            HDDEDATA hData,
                                            DWORD dwData1,
                                            DWORD dwData2)
{
    HDDEDATA hDdeData = NULL;
    UINT ui = 0;
    DWORD dwErr = 0;

    //
    // get a pointer to the server
    //

    CDDEServer* pServ = pTheServer; // BARF BARF BARF
    _ASSERT(pServ);
    PrintStatus(pServ, wType);

    switch (wType) {
    case XTYP_CONNECT_CONFIRM:

        //
        // Add a new conversation to the list
        //

        pServ->Status(_T("Connect to %s"), pServ->StringFromHsz(hsz1).c_str());
        pServ->AddConversation(hConv, hsz1);
        break;

    case XTYP_DISCONNECT:

        //
        // get some info on why it disconnected
        //

        CONVINFO ci;
        memset(&ci, 0, sizeof(ci));
        ci.cb = sizeof(ci);
        ui = ::DdeQueryConvInfo(hConv, wType, &ci);
        dwErr = pServ->GetLastError();

        //
        // Remove a conversation from the list
        //

        pServ->Status(_T("Disconnect"));
        pServ->OnConvDisconnected(hConv);
        break;

    case XTYP_WILDCONNECT:

        //
        // We only support wild connects to either a NULL service
        // name or to the name of our own service.
        //

        if ((hsz2 == NULL)
        || !::DdeCmpStringHandles(hsz2, pServ->m_hszServiceName)) {

            pServ->Status(_T("Wild connect to %s"), pServ->StringFromHsz(hsz1).c_str());
            return pServ->DoWildConnect(hsz1);

        }
        break;

        //
        // For all other messages we see if we want them here
        // and if not, they get passed on to the user callback
        // if one is defined.
        //

    case XTYP_ADVSTART:
    case XTYP_CONNECT:
    case XTYP_EXECUTE:
    case XTYP_REQUEST:
    case XTYP_ADVREQ:
    case XTYP_ADVDATA:
    case XTYP_POKE:

        //
        // Try and process them here first.
        //

        if (pServ->DoCallback(wType,
                       wFmt,
                       hConv,
                       hsz1,
                       hsz2,
                       hData,
                       &hDdeData)) {

            return hDdeData;
        }

        //
        // Fall Through to allow the custom callback a chance
        //

    default:

        return pServ->CustomCallback(wType,
                                      wFmt,
                                      hConv,
                                      hsz1,
                                      hsz2,
                                      hData,
                                      dwData1,
                                      dwData2);
    }

    return (HDDEDATA) NULL;
}

CDDEConv* CDDEServer::AddConversation(HCONV hConv, HSZ hszTopic)
{
    //
    // create a new conversation object
    //

    CDDEConv* pConv = new CDDEConv(this, hConv, hszTopic);
    _ASSERT(pConv);
    pConv->AddRef();

    //
    // Add it into the list
    //

    m_ConvList.push_back(pConv);

    return pConv;
}

CDDEConv* CDDEServer::AddConversation(CDDEConv* pNewConv)
{
    _ASSERT(pNewConv);
    pNewConv->AddRef();

    //
    // Add it into the list
    //

    m_ConvList.push_back(pNewConv);

    return pNewConv;
}

BOOL CDDEServer::OnConvDisconnected(HCONV hConv)
{
    if (!hConv) {
        return FALSE;
    }

    //
    // Try to find the conversation in the list
    //

    CDDEConv* pConv = NULL;
    CDDEConvList::iterator it = m_ConvList.begin();
    for (; it != m_ConvList.end(); ++it) {
        pConv = *it;
        if (pConv->m_hConv == hConv) {
            it = m_ConvList.erase(it);
            pConv->Terminate();
            pConv->Release();
            return TRUE;
        }
    }

    //
    // Not in the list
    //

    return FALSE;
}

BOOL CDDEServer::RemoveConversation(CDDEConv* pConv)
{
    if (!pConv) {
        return FALSE;
    }

    CDDEConvList::iterator it = m_ConvList.begin();
    for (; it != m_ConvList.end(); ++it) {
        if (pConv == *it) {
            it = m_ConvList.erase(it);
            pConv->Release();
            return TRUE;
        }
    }

    return FALSE;
}

HDDEDATA CDDEServer::DoWildConnect(HSZ hszTopic)
{
    //
    // See how many topics we will be returning
    //

    size_t iTopics = 0;
    CDDEString strTopic = _T("<null>");
    if (hszTopic == NULL) {

        //
        // Count all the topics we have
        //

        iTopics = m_TopicList.size();

    } else {

        //
        // See if we have this topic in our list
        //

        strTopic = StringFromHsz(hszTopic);
        CDDETopic* pTopic = FindTopic(strTopic.c_str());
        if(pTopic) {
            iTopics++;
        }
    }

    //
    // If we have no match or no topics at all, just return
    // NULL now to refuse the connect
    //

    if (!iTopics) {
        Status(_T("Wild connect to %s refused"), strTopic.c_str());
        return (HDDEDATA) NULL;
    }

    //
    // Allocate a chunk of DDE data big enough for all the HSZPAIRS
    // we'll be sending back plus space for a NULL entry on the end
    //

    HDDEDATA hData = ::DdeCreateDataHandle(m_dwDDEInstance,
                                           NULL,
                                           (DWORD)((iTopics + 1) * sizeof(HSZPAIR)),
                                           0,
                                           NULL,
                                           0,
                                           0);

    //
    // Check we actually got it.
    //

    if (!hData) return (HDDEDATA) NULL;

    HSZPAIR* pHszPair = (PHSZPAIR) DdeAccessData(hData, NULL);

    //
    // Copy the topic data
    //

    if (hszTopic == NULL) {

        //
        // Copy all the topics we have (includes the system topic)
        //

        CDDETopicList::iterator it = m_TopicList.begin();
        for (; it != m_TopicList.end(); ++it) {

            CDDETopic* pTopic = *it;
            pHszPair->hszSvc = ::DdeCreateStringHandle(m_dwDDEInstance,
                                                       m_strServiceName.c_str(),
                                                       DDE_CODEPAGE);
            pHszPair->hszTopic = ::DdeCreateStringHandle(m_dwDDEInstance,
                                                         pTopic->m_strName.c_str(),
                                                         DDE_CODEPAGE);

            pHszPair++;
        }

    } else {

        //
        // Just copy the one topic asked for
        //

        pHszPair->hszSvc = m_hszServiceName;
        pHszPair->hszTopic = hszTopic;

        pHszPair++;

    }

    //
    // Put the terminator on the end
    //

    pHszPair->hszSvc = NULL;
    pHszPair->hszTopic = NULL;

    //
    // Finished with the data block
    //

    ::DdeUnaccessData(hData);

    //
    // Return the block handle
    //

    return hData;
}

CDDETopic* CDDEServer::FindTopic(LPCTSTR pszTopic)
{
    CDDETopicList::iterator it = m_TopicList.begin();
    for (; it != m_TopicList.end(); ++it) {
        CDDETopic* pTopic = *it;
        if (_tcsicmp(pTopic->m_strName.c_str(), pszTopic) == 0) return pTopic;
    }
    return NULL;
}

const CDDETopic* CDDEServer::FindTopic(LPCTSTR pszTopic) const
{
    CDDETopicList::const_iterator it = m_TopicList.begin();
    for (; it != m_TopicList.end(); ++it) {
        const CDDETopic* pTopic = *it;
        if (_tcsicmp(pTopic->m_strName.c_str(), pszTopic) == 0) return pTopic;
    }
    return NULL;
}

BOOL CDDEServer::DoCallback(WORD wType,
                            WORD wFmt,
                            HCONV hConv,
                            HSZ hszTopic,
                            HSZ hszItem,
                            HDDEDATA hData,
                            HDDEDATA *phReturnData)
{
    //
    // See if we know the topic
    //

    CDDEString strTopic = StringFromHsz(hszTopic);

    //
    // See if this is an execute request
    //

    if (wType == XTYP_EXECUTE) {

        //
        // Call the exec function to process it
        //

        Status(_T("Exec"));
        DWORD dwLength = 0;
        void* pData = ::DdeAccessData(hData, &dwLength);
        BOOL b = Exec(strTopic.c_str(), pData, dwLength);
        ::DdeUnaccessData(hData);

        if (b) {

            *phReturnData = (HDDEDATA) DDE_FACK;
            return FALSE;

        }

        //
        // Either no handler or it didn't get handled by the function
        //

        Status(_T("Exec failed"));
        *phReturnData = (HDDEDATA) DDE_FNOTPROCESSED;
        return FALSE;
    }

    //
    // See if this is a connect request. Accept it if it is.
    //

    if (wType == XTYP_CONNECT) {

        if (!FindTopic(strTopic.c_str())) return FALSE; // unknown topic
        *phReturnData = (HDDEDATA) TRUE;
        return TRUE;
    }

    //
    // For any other transaction we need to be sure this is an
    // item we support and in some cases, that the format requested
    // is supported for that item.
    //

    CDDEString strItem = StringFromHsz(hszItem);

    //
    // Now just do whatever is required for each specific transaction
    //

    BOOL b = FALSE;
    DWORD dwLength = 0;
    void* pData = NULL;

    switch (wType) {
    case XTYP_ADVSTART:

        //
        // Confirm that the supported topic/item pair is OK and
        // that the format is supported

        if (!CanAdvise(wFmt, strTopic.c_str(), strItem.c_str())) {

            Status(_T("Can't advise on %s|%s"), strTopic.c_str(), strItem.c_str());
            return FALSE;
        }

        //
        // Start an advise request.  Topic/item and format are ok.
        //

        *phReturnData = (HDDEDATA) TRUE;
        break;

    case XTYP_POKE:

        //
        // Some data for one of our items.
        //

        pData = ::DdeAccessData(hData, &dwLength);
        b = Poke(wFmt, strTopic.c_str(), strItem.c_str(), pData, dwLength);
        ::DdeUnaccessData(hData);

        if (!b) {

            //
            // Nobody took the data.
            // Maybe its not a supported item or format
            //

            Status(_T("Poke %s|%s failed"), strTopic.c_str(), strItem.c_str());
            return FALSE;

        }

        //
        // Data at the server has changed.  See if we
        // did this ourself (from a poke) or if it's from
        // someone else.  If it came from elsewhere then post
        // an advise notice of the change.
        //

        CONVINFO ci;
        ci.cb = sizeof(CONVINFO);
        if (::DdeQueryConvInfo(hConv, (DWORD)QID_SYNC, &ci)) {

            if (! (ci.wStatus & ST_ISSELF)) {

                //
                // It didn't come from us
                //

                ::DdePostAdvise(m_dwDDEInstance,
                              hszTopic,
                              hszItem);
            }
        }

        *phReturnData = (HDDEDATA) DDE_FACK; // say we took it
        break;

    case XTYP_ADVDATA:

        //
        // A server topic/item has changed value
        //

        pData = ::DdeAccessData(hData, &dwLength);
        b = AdviseData(wFmt, hConv, strTopic.c_str(), strItem.c_str(), pData, dwLength);
        ::DdeUnaccessData(hData);

        if (!b) {

            //
            // Nobody took the data.
            // Maybe its not of interrest
            //

            Status(_T("AdviseData %s|%s failed"), strTopic.c_str(), strItem.c_str());
            *phReturnData = (HDDEDATA) DDE_FNOTPROCESSED;

        } else {

            *phReturnData = (HDDEDATA) DDE_FACK; // say we took it
        }
        break;

    case XTYP_ADVREQ:
    case XTYP_REQUEST:

        //
        // Attempt to start an advise or get the data on a topic/item
        // See if we have a request function for this item or
        // a generic one for the topic
        //

        Status(_T("Request %s|%s"), strTopic.c_str(), strItem.c_str());
        dwLength = 0;
        if (!Request(wFmt, strTopic.c_str(), hszItem, phReturnData)) {

            //
            // Nobody accepted the request
            // Maybe unsupported topic/item or bad format
            //

            Status(_T("Request %s|%s failed"), strTopic.c_str(), strItem.c_str());
            *phReturnData = NULL;
            return FALSE;

        }

        break;

    default:
        break;
    }

    //
    // Say we processed the transaction in some way
    //

    return TRUE;

}

BOOL CDDEServer::AddTopic(CDDETopic* pNewTopic)
{
    _ASSERT(pNewTopic);

    //
    // See if we already have this topic
    //

    CDDETopicList::iterator it = std::find(m_TopicList.begin(),
                                           m_TopicList.end(),
                                           pNewTopic);
    if (it != m_TopicList.end()) return TRUE; // already have it

    //
    // Add the new topic
    //

    m_TopicList.push_back(pNewTopic);
    pNewTopic->m_pServer = this;

    pNewTopic->AddItem(&m_SystemItemItems);
    pNewTopic->AddItem(&m_SystemItemFormats);

    return TRUE;
}

CDDEString CDDEServer::StringFromHsz(HSZ hsz)
{
    CDDEString str = _T("<null>");

    //
    // Get the length of the string
    //

    DWORD dwLen = ::DdeQueryString(m_dwDDEInstance,
                                   hsz,
                                   NULL,
                                   0,
                                   DDE_CODEPAGE);

    if (dwLen == 0) return str;

    //
    // get the text
    //

    str.resize(dwLen+1);

    //
    // Get the string text
    //

    DWORD dw = ::DdeQueryString(m_dwDDEInstance,
                                hsz,
                                (LPTSTR)str.c_str(),
                                dwLen+1,
                                DDE_CODEPAGE);

    if (dw == 0) str = _T("<error>");

    return str;
}

BOOL CDDEServer::Request(UINT wFmt, LPCTSTR pszTopic, HSZ hszItem,
                         HDDEDATA* phReturnData)
{
    //
    // See if we have a topic that matches
    //

    CDDETopic* pTopic = FindTopic(pszTopic);
    if (!pTopic) return FALSE;

    return pTopic->Request(wFmt, hszItem, phReturnData);
}

BOOL CDDEServer::Poke(UINT wFmt, LPCTSTR pszTopic, LPCTSTR pszItem,
                      void* pData, DWORD dwSize)
{
    //
    // See if we have a topic that matches
    //

    CDDETopic* pTopic = FindTopic(pszTopic);
    if (!pTopic) return FALSE;

    return pTopic->Poke(wFmt, pszItem, pData, dwSize);
}

BOOL CDDEServer::Exec(LPCTSTR pszTopic, void* pData, DWORD dwSize)
{
    //
    // See if we have a topic that matches
    //

    CDDETopic* pTopic = FindTopic(pszTopic);
    if (!pTopic) return FALSE;

    return pTopic->Exec(pData, dwSize);
}

BOOL CDDEServer::CanAdvise(UINT wFmt, LPCTSTR pszTopic, LPCTSTR pszItem)
{
    //
    // See if we have a topic that matches
    //

    CDDETopic* pTopic = FindTopic(pszTopic);
    if (!pTopic) return FALSE;

    return pTopic->CanAdvise(wFmt, pszItem);
}

void CDDEServer::PostAdvise(CDDETopic* pTopic, CDDEItem* pItem)
{
    _ASSERT(pTopic);
    _ASSERT(pItem);

    BOOL result = ::DdePostAdvise(m_dwDDEInstance,
                                  ::DdeCreateStringHandle(m_dwDDEInstance,
                                                          pTopic->m_strName.c_str(),
                                                          DDE_CODEPAGE),
                                  ::DdeCreateStringHandle(m_dwDDEInstance,
                                                          pItem->m_strName.c_str(),
                                                          DDE_CODEPAGE));

    if (!result) {
        Status(_T("DdePostAdvise error: %d"), GetLastError());
    }
}

CDDEString DDEGetFormatName(WORD wFmt)
{
    CDDEString strName;
    PCFTAGNAME pCTN;

    //
    // Try for a standard one first
    //

    pCTN = CFNames;
    while (pCTN->wFmt) {
        if (pCTN->wFmt == wFmt) {
            strName = pCTN->pszName;
            return strName;
        }
        pCTN++;
    }

    //
    // See if it's a registered one
    //

    TCHAR buf[256];
    if (::GetClipboardFormatName(wFmt, buf, _countof(buf))) {
        strName = buf;
    }

    return strName;
}

CDDEConv* CDDEServer::FindConversation(HCONV hConv)
{
    CDDEConvList::iterator it = m_ConvList.begin();
    for (; it != m_ConvList.end(); ++it) {

        CDDEConv* pConv = *it;
        _ASSERT(pConv);
        if (pConv->m_hConv == hConv) return pConv;
    }
    return NULL;
}

BOOL CDDEServer::AdviseData(UINT wFmt, HCONV hConv,
                            LPCTSTR pszTopic, LPCTSTR pszItem,
                            void* pData, DWORD dwSize)
{
    //
    // See if we know this conversation
    //

    CDDEConv* pConv = FindConversation(hConv);
    if (!pConv) return FALSE;

    return pConv->AdviseData(wFmt, pszTopic, pszItem, pData, dwSize);
}