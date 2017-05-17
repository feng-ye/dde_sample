// myclient.h

class CMyClient : public CDDEServer
{
public:
    CMyClient();
    virtual void Status(LPCTSTR pszFormat, ...);
};
