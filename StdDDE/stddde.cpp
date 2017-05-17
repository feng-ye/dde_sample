// stddde.cpp

#include "stdafx.h"
#include "stddde.h"

//
// Constants
//

#define DDE_TIMEOUT     5000 // 5 seconds

#ifdef _UNICODE
#define DDE_CODEPAGE    CP_WINUNICODE
#else
#define DDE_CODEPAGE    CP_WINANSI
#endif

//
// Format lists
//

static WORD SysFormatList[] = {
    CF_TEXT,
    NULL};

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
    ATLASSERT(m_iRefCount == 0);
}

int CDDECountedObject::AddRef()
{
    ATLASSERT(m_iRefCount < 1000); // sanity check
    return ++m_iRefCount;
}

int CDDECountedObject::Release()
{
    int i = --m_iRefCount;
    ATLASSERT(m_iRefCount >= 0);
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
    ATLASSERT(m_hsz);
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
    ATLASSERT(m_hsz);
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

BOOL CDDEItem::Request(UINT wFmt, void** ppData, DWORD* pdwSize)
{
    return FALSE;
}

BOOL CDDEItem::Poke(UINT wFmt, void* pData, DWORD dwSize)
{
    return FALSE;
}

BOOL CDDEItem::IsSupportedFormat(WORD wFormat)
{
    WORD* pFmt = GetFormatList();
    if (!pFmt) return FALSE;
    while(*pFmt) {
        if (*pFmt == wFormat) return TRUE;
        pFmt++;
    }
    return FALSE;
}

BOOL CDDEItem::CanAdvise(UINT wFmt)
{
    return IsSupportedFormat(wFmt);
}

void CDDEItem::PostAdvise()
{
    if (m_pTopic == NULL) return;
    m_pTopic->PostAdvise(this);
}

////////////////////////////////////////////////////////////////////////////////////
//
// CDDEStringItem

WORD* CDDEStringItem::GetFormatList()
{
    return SysFormatList; // CF_TEXT
}

BOOL CDDEStringItem::Request(UINT wFmt, void** ppData, DWORD* pdwSize)
{
    ATLASSERT(wFmt == CF_TEXT);
    ATLASSERT(ppData);
    *ppData = (void*)(LPCTSTR)m_strData;
    *pdwSize = m_strData.GetLength() + 1; // allow for the null
    return TRUE;
}

BOOL CDDEStringItem::Poke(UINT wFmt, void* pData, DWORD dwSize)
{
    ATLASSERT(wFmt == CF_TEXT);
    ATLASSERT(pData);
    m_strData = (LPTSTR)pData;
    OnPoke();
    return TRUE;
}

void CDDEStringItem::SetData(LPCTSTR pszData)
{
    ATLASSERT(pszData);
    m_strData = pszData;
    PostAdvise();
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
    ATLASSERT(pNewItem);

    //
    // See if we already have this item
    //

    POSITION pos = m_ItemList.Find(pNewItem);
    if (pos) return TRUE; // already have it

    //
    // Add the new item
    //

    m_ItemList.AddTail(pNewItem);
    pNewItem->m_pTopic = this;

    return TRUE;
}

BOOL CDDETopic::Request(UINT wFmt, LPCTSTR pszItem,
                            void** ppData, DWORD* pdwSize)
{
    //
    // See if we have this item
    //

    CDDEItem* pItem = FindItem(pszItem);
    if (!pItem) return FALSE;

    return pItem->Request(wFmt, ppData, pdwSize);
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
    POSITION pos = m_ItemList.GetHeadPosition();
    while (pos) {
        CDDEItem* pItem = m_ItemList.GetNext(pos);
        if (pItem->m_strName.CompareNoCase(pszItem) == 0) return pItem;
    }
    return NULL;
}

BOOL CDDETopic::CanAdvise(UINT wFmt, LPCTSTR pszItem)
{
    //
    // See if we have this item
    //

    CDDEItem* pItem = FindItem(pszItem);
    if (!pItem) return FALSE;

    return pItem->CanAdvise(wFmt);
}

void CDDETopic::PostAdvise(CDDEItem* pItem)
{
    ATLASSERT(m_pServer);
    ATLASSERT(pItem);
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

        ::DdeDisconnect(m_hConv);

        //
        // Tell the server
        //

        ATLASSERT(m_pServer);
        m_pServer->RemoveConversation(m_hConv);

        m_hConv = NULL;

        return TRUE;
    }

    return FALSE; // wasn't active
}

BOOL CDDEConv::ConnectTo(LPCTSTR pszService, LPCTSTR pszTopic)
{
    ATLASSERT(pszService);
    ATLASSERT(pszTopic);
    ATLASSERT(m_pServer);
    ATLASSERT(!m_hConv);

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

BOOL CDDEConv::Request(LPCTSTR pszItem, void** ppData, DWORD* pdwSize)
{
    ATLASSERT(m_pServer);
    ATLASSERT(pszItem);
    ATLASSERT(ppData);
    ATLASSERT(pdwSize);

    CHSZ hszItem (m_pServer, pszItem);
    HDDEDATA hData = ::DdeClientTransaction(NULL,
                                            0,
                                            m_hConv,
                                            hszItem,
                                            CF_TEXT,
                                            XTYP_REQUEST,
                                            DDE_TIMEOUT,
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
    ATLASSERT(*pdwSize);
    *ppData = new char[*pdwSize];
    ATLASSERT(*ppData);
    memcpy(*ppData, pData, *pdwSize);
    ::DdeUnaccessData(hData);

    return TRUE;
}

BOOL CDDEConv::Advise(LPCTSTR pszItem)
{
    ATLASSERT(m_pServer);
    ATLASSERT(pszItem);

    CHSZ hszItem (m_pServer, pszItem);
    HDDEDATA hData = ::DdeClientTransaction(NULL,
                                            0,
                                            m_hConv,
                                            hszItem,
                                            CF_TEXT,
                                            XTYP_ADVSTART,
                                            DDE_TIMEOUT,
                                            NULL);

    if (!hData) {
        // Failed
        return FALSE;
    }
    return TRUE;
}

BOOL CDDEConv::Exec(LPCTSTR pszCmd)
{
    //
    // Send the command
    //

    HDDEDATA hData = ::DdeClientTransaction((BYTE*)pszCmd,
                                            (DWORD)((_tcslen(pszCmd) + 1) * sizeof(TCHAR)),
                                            m_hConv,
                                            0,
                                            CF_TEXT,
                                            XTYP_EXECUTE,
                                            DDE_TIMEOUT,
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
                                            DDE_TIMEOUT,
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

WORD* CDDESystemItem::GetFormatList()
{
    return SysFormatList;
}

//
// Specific system topic items
//

BOOL CDDESystemItem_TopicList::Request(UINT wFmt, void** ppData, DWORD* pdwSize)
{
    //
    // Return the list of topics for this service
    //

    static CString strTopics;
    strTopics = "";
    ATLASSERT(m_pTopic);
    CDDEServer* pServer = m_pTopic->m_pServer;
    ATLASSERT(pServer);
    POSITION pos = pServer->m_TopicList.GetHeadPosition();
    int items = 0;
    while (pos) {

        CDDETopic* pTopic = pServer->m_TopicList.GetNext(pos);

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

    *ppData = (void*)(LPCTSTR)strTopics;
    *pdwSize = strTopics.GetLength() + 1; // include room for the NULL
    return TRUE;
}

BOOL CDDESystemItem_ItemList::Request(UINT wFmt, void** ppData, DWORD* pdwSize)
{
    //
    // Return the list of items in this topic
    //

    static CString strItems;
    strItems = "";
    ATLASSERT(m_pTopic);
    POSITION pos = m_pTopic->m_ItemList.GetHeadPosition();
    int items = 0;
    while (pos) {

        CDDEItem* pItem = m_pTopic->m_ItemList.GetNext(pos);

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

    *ppData = (void*)(LPCTSTR)strItems;
    *pdwSize = strItems.GetLength() + 1; // include romm for the NULL
    return TRUE;
}

BOOL CDDESystemItem_FormatList::Request(UINT wFmt, void** ppData, DWORD* pdwSize)
{
    //
    // Return the list of formats in this topic
    //

    static CString strFormats;
    strFormats = "";
    ATLASSERT(m_pTopic);
    POSITION pos = m_pTopic->m_ItemList.GetHeadPosition();
    int iFormats = 0;
    WORD wFmtList[100];
    while (pos) {

        CDDEItem* pItem = m_pTopic->m_ItemList.GetNext(pos);

        //
        // get the format list for this item
        //

        WORD* pItemFmts = pItem->GetFormatList();
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

                    strFormats += ::GetFormatName(*pItemFmts);

                    iFormats++;
                }

                pItemFmts++;
            }
        }
    }

    //
    // Set up the return info
    //

    *ppData = (void*)(LPCTSTR)strFormats;
    *pdwSize = strFormats.GetLength() + 1; // include romm for the NULL
    return TRUE;
}

BOOL CDDEServerSystemTopic::Request(UINT wFmt, LPCTSTR pszItem,
                                    void** ppData, DWORD* pdwSize)
{
    m_pServer->Status(_T("System topic request: %s"), pszItem);
    return CDDETopic::Request(wFmt, pszItem, ppData, pdwSize);
}

////////////////////////////////////////////////////////////////////////////////////
//
// CDDEServer

CDDEServer::CDDEServer(LPCTSTR pszServiceName)
{
    m_bInitialized = FALSE;
    m_strServiceName = pszServiceName;
    m_dwDDEInstance = 0;
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

        POSITION pos = m_ConvList.GetHeadPosition();
        while (pos) {
            CDDEConv* pConv = m_ConvList.GetNext(pos);
            ATLASSERT(pConv);
            pConv->Terminate();
        }

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
                        DWORD* pdwDDEInst/* = NULL */)
{
    //
    // make sure we are alone in the world
    //

    if (pTheServer != NULL) {
        ATLTRACE("Already got a server!\n");
        ATLASSERT(0);
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

    //
    // Copy the service name and create a DDE name handle for it
    //

    m_strServiceName = pszServiceName;
    m_hszServiceName.Create(this, m_strServiceName);

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
    ATLASSERT(pServ);
    pServ->Status(_T("Callback %4.4XH"), wType);

    switch (wType) {
    case XTYP_CONNECT_CONFIRM:

        //
        // Add a new conversation to the list
        //

        pServ->Status(_T("Connect to %s"), (LPCTSTR)pServ->StringFromHsz(hsz1));
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
        pServ->RemoveConversation(hConv);
        break;

    case XTYP_WILDCONNECT:

        //
        // We only support wild connects to either a NULL service
        // name or to the name of our own service.
        //

        if ((hsz2 == NULL)
        || !::DdeCmpStringHandles(hsz2, pServ->m_hszServiceName)) {

            pServ->Status(_T("Wild connect to %s"), (LPCTSTR)pServ->StringFromHsz(hsz1));
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
    ATLASSERT(pConv);
    pConv->AddRef();

    //
    // Add it into the list
    //

    m_ConvList.AddTail(pConv);

    return pConv;
}

CDDEConv* CDDEServer::AddConversation(CDDEConv* pNewConv)
{
    ATLASSERT(pNewConv);
    pNewConv->AddRef();

    //
    // Add it into the list
    //

    m_ConvList.AddTail(pNewConv);

    return pNewConv;
}

BOOL CDDEServer::RemoveConversation(HCONV hConv)
{
    //
    // Try to find the conversation in the list
    //

    CDDEConv* pConv = NULL;
    POSITION pos = m_ConvList.GetHeadPosition();
    while (pos) {
        pConv = m_ConvList.GetNext(pos);
        if (pConv->m_hConv == hConv) {
            m_ConvList.RemoveAt(m_ConvList.Find(pConv));
            pConv->Release();
            return TRUE;
        }
    }

    //
    // Not in the list
    //

    return FALSE;
}

HDDEDATA CDDEServer::DoWildConnect(HSZ hszTopic)
{
    //
    // See how many topics we will be returning
    //

    size_t iTopics = 0;
    CString strTopic = _T("<null>");
    if (hszTopic == NULL) {

        //
        // Count all the topics we have
        //

        iTopics = m_TopicList.GetCount();

    } else {

        //
        // See if we have this topic in our list
        //

        strTopic = StringFromHsz(hszTopic);
        CDDETopic* pTopic = FindTopic(strTopic);
        if(pTopic) {
            iTopics++;
        }
    }

    //
    // If we have no match or no topics at all, just return
    // NULL now to refuse the connect
    //

    if (!iTopics) {
        Status(_T("Wild connect to %s refused"), (LPCTSTR)strTopic);
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

        POSITION pos = m_TopicList.GetHeadPosition();
        while (pos) {

            CDDETopic* pTopic = m_TopicList.GetNext(pos);
            pHszPair->hszSvc = ::DdeCreateStringHandle(m_dwDDEInstance,
                                                       (LPTSTR)(LPCTSTR)m_strServiceName,
                                                       DDE_CODEPAGE);
            pHszPair->hszTopic = ::DdeCreateStringHandle(m_dwDDEInstance,
                                                         (LPTSTR)(LPCTSTR)pTopic->m_strName,
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
    POSITION pos = m_TopicList.GetHeadPosition();
    while (pos) {
        CDDETopic* pTopic = m_TopicList.GetNext(pos);
        if (pTopic->m_strName.CompareNoCase(pszTopic) == 0) return pTopic;
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

    CString strTopic = StringFromHsz(hszTopic);

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
        BOOL b = Exec(strTopic, pData, dwLength);
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

        if (!FindTopic(strTopic)) return FALSE; // unknown topic
        *phReturnData = (HDDEDATA) TRUE;
        return TRUE;
    }

    //
    // For any other transaction we need to be sure this is an
    // item we support and in some cases, that the format requested
    // is supported for that item.
    //

    CString strItem = StringFromHsz(hszItem);

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

        if (!CanAdvise(wFmt, strTopic, strItem)) {

            Status(_T("Can't advise on %s|%s"), (LPCTSTR)strTopic, (LPCTSTR)strItem);
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
        b = Poke(wFmt, strTopic, strItem, pData, dwLength);
        ::DdeUnaccessData(hData);

        if (!b) {

            //
            // Nobody took the data.
            // Maybe its not a supported item or format
            //

            Status(_T("Poke %s|%s failed"), (LPCTSTR)strTopic, (LPCTSTR)strItem);
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
        b = AdviseData(wFmt, hConv, strTopic, strItem, pData, dwLength);
        ::DdeUnaccessData(hData);

        if (!b) {

            //
            // Nobody took the data.
            // Maybe its not of interrest
            //

            Status(_T("AdviseData %s|%s failed"), (LPCTSTR)strTopic, (LPCTSTR)strItem);
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

        Status(_T("Request %s|%s"), (LPCTSTR)strTopic, (LPCTSTR)strItem);
        dwLength = 0;
        if (!Request(wFmt, strTopic, strItem, &pData, &dwLength)) {

            //
            // Nobody accepted the request
            // Maybe unsupported topic/item or bad format
            //

            Status(_T("Request %s|%s failed"), (LPCTSTR)strTopic, (LPCTSTR)strItem);
            *phReturnData = NULL;
            return FALSE;

        }

        //
        // There is some data so build a DDE data object to return
        //

        *phReturnData = ::DdeCreateDataHandle(m_dwDDEInstance,
                                              (unsigned char*)pData,
                                              dwLength,
                                              0,
                                              hszItem,
                                              wFmt,
                                              0);

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
    ATLASSERT(pNewTopic);

    //
    // See if we already have this topic
    //

    POSITION pos = m_TopicList.Find(pNewTopic);
    if (pos) return TRUE; // already have it

    //
    // Add the new topic
    //

    m_TopicList.AddTail(pNewTopic);
    pNewTopic->m_pServer = this;

    pNewTopic->AddItem(&m_SystemItemItems);
    pNewTopic->AddItem(&m_SystemItemFormats);

    return TRUE;
}

CString CDDEServer::StringFromHsz(HSZ hsz)
{
    CString str = _T("<null>");

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

    LPTSTR pBuf = str.GetBufferSetLength(dwLen+1);
    ATLASSERT(pBuf);

    //
    // Get the string text
    //

    DWORD dw = ::DdeQueryString(m_dwDDEInstance,
                                 hsz,
                                 pBuf,
                                 dwLen+1,
                                 DDE_CODEPAGE);

    //
    // Tidy up
    //

    str.ReleaseBuffer();

    if (dw == 0) str = _T("<error>");

    return str;
}

BOOL CDDEServer::Request(UINT wFmt, LPCTSTR pszTopic, LPCTSTR pszItem,
                         void** ppData, DWORD* pdwSize)
{
    //
    // See if we have a topic that matches
    //

    CDDETopic* pTopic = FindTopic(pszTopic);
    if (!pTopic) return FALSE;

    return pTopic->Request(wFmt, pszItem, ppData, pdwSize);
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
    ATLASSERT(pTopic);
    ATLASSERT(pItem);

    ::DdePostAdvise(m_dwDDEInstance,
                    ::DdeCreateStringHandle(m_dwDDEInstance,
                                            (LPTSTR)(LPCTSTR)pTopic->m_strName,
                                            DDE_CODEPAGE),
                    ::DdeCreateStringHandle(m_dwDDEInstance,
                                            (LPTSTR)(LPCTSTR)pItem->m_strName,
                                            DDE_CODEPAGE));

}

CString GetFormatName(WORD wFmt)
{
    CString strName = "";
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
    POSITION pos = m_ConvList.GetHeadPosition();
    while (pos) {

        CDDEConv* pConv = m_ConvList.GetNext(pos);
        ATLASSERT(pConv);
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