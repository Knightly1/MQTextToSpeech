#pragma once

#include <intsafe.h>
#include <atlcomcli.h>
#include <sperror.h>

inline HRESULT SpGetCategoryFromId(
    const WCHAR * pszCategoryId,
    ISpObjectTokenCategory ** ppCategory,
    BOOL fCreateIfNotExist = FALSE)
{
    HRESULT hr;
    
    CComPtr<ISpObjectTokenCategory> cpTokenCategory;
    hr = cpTokenCategory.CoCreateInstance(CLSID_SpObjectTokenCategory);
    
    if (SUCCEEDED(hr))
    {
        hr = cpTokenCategory->SetId(pszCategoryId, fCreateIfNotExist);
    }
    
    if (SUCCEEDED(hr))
    {
        *ppCategory = cpTokenCategory.Detach();
    }
    
    return hr;
}


inline HRESULT SpEnumTokens(
    const WCHAR * pszCategoryId, 
    const WCHAR * pszReqAttribs, 
    const WCHAR * pszOptAttribs, 
    IEnumSpObjectTokens ** ppEnum)
{
    HRESULT hr = S_OK;
    
    CComPtr<ISpObjectTokenCategory> cpCategory;
    hr = SpGetCategoryFromId(pszCategoryId, &cpCategory);
    
    if (SUCCEEDED(hr))
    {
        hr = cpCategory->EnumTokens(
                    pszReqAttribs,
                    pszOptAttribs,
                    ppEnum);
    }
    
    return hr;
}

inline HRESULT SpFindBestToken(
    const WCHAR * pszCategoryId,
    const WCHAR * pszReqAttribs,
    const WCHAR * pszOptAttribs,
    ISpObjectToken **ppObjectToken)
{
    HRESULT hr = S_OK;

    const WCHAR *pszVendorPreferred = L"VendorPreferred";
    const ULONG ulLenVendorPreferred = (ULONG) wcslen(pszVendorPreferred);

    // append VendorPreferred to the end of pszOptAttribs to force this preference
    ULONG ulLen;
    if (pszOptAttribs)
    {
        hr = ULongAdd((ULONG)wcslen(pszOptAttribs), ulLenVendorPreferred, &ulLen);
        if (SUCCEEDED(hr))
        {
            hr = ULongAdd(ulLen, 1 + 1, &ulLen); // including 1 char here for null terminator
        }
    }
    else
    {
        hr = ULongAdd(ulLenVendorPreferred, 1, &ulLen); // including 1 char here for null terminator
    }
    if (FAILED(hr))
    {
        hr = E_INVALIDARG;
    }
    else
    {
        WCHAR *pszOptAttribsVendorPref = new WCHAR[ulLen];
        if (pszOptAttribsVendorPref)
        {
            if (pszOptAttribs)
            {
                StringCchCopyW (pszOptAttribsVendorPref, ulLen, pszOptAttribs);
                StringCchCatW (pszOptAttribsVendorPref, ulLen, L";");
                StringCchCatW (pszOptAttribsVendorPref, ulLen, pszVendorPreferred);
            }
            else
            {
                StringCchCopyW (pszOptAttribsVendorPref, ulLen, pszVendorPreferred);
            }
        }
        else
        {
            hr = E_OUTOFMEMORY;
        }

        CComPtr<IEnumSpObjectTokens> cpEnum;
        if (SUCCEEDED(hr))
        {
            hr = SpEnumTokens(pszCategoryId, pszReqAttribs, pszOptAttribsVendorPref, &cpEnum);
        }

        delete[] pszOptAttribsVendorPref;

        if (SUCCEEDED(hr))
        {
            hr = cpEnum->Next(1, ppObjectToken, NULL);
            if (hr == S_FALSE)
            {
                *ppObjectToken = NULL;
                hr = SPERR_NOT_FOUND;
            }
        }
    }

    return hr;
}