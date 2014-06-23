//
// Set MAPI Address List Search Order
//
// Shawn Poulson <spoulson@explodingcoder.com>, 2008.10.24
//
// Modified to work with Outlook 2010 by CL <cleitet@gmail.com> 2012-10-15
//

#include "stdafx.h"
#include <mapix.h>
#include <mapiutil.h>
#include <mapiguid.h>
#include <mapidefs.h>
#include <string>
#include <iostream>
#include <list>
#include <algorithm>
//Start modified by CL

#include <INITGUID.H>

//End modified by CL

using namespace std;

//Start modified by CL

// http://blogs.msdn.com/b/stephen_griffin/archive/2010/09/13/you-chose-wisely.aspx
// Capone profile section
// {00020D0A-0000-0000-C000-000000000046}
DEFINE_OLEGUID(IID_CAPONE_PROF, 0x00020d0a, 0, 0);

// See http://blogs.msdn.com/b/stephen_griffin/archive/2011/04/13/setsearchpath-not-really.aspx
#define PR_AB_CHOOSE_DIRECTORY_AUTOMATICALLY PROP_TAG( PT_BOOLEAN, 0x3D1C)

#define PR_AB_SEARCH_PATH_CUSTOMIZATION PROP_TAG( PT_LONG, 0x3D1B)

typedef enum _SearchPathReorderType
{
                SEARCHPATHREORDERTYPE_RAW = 0,
                SEARCHPATHREORDERTYPE_ACCT_PREFERGAL,
                SEARCHPATHREORDERTYPE_ACCT_PREFERCONTACTS,
} SearchPathReorderType;

//End modified by CL

STDMETHODIMP MAPILogon(LPMAPISESSION *lppSession);
void MAPILogoff(IMAPISession &Session);
STDMETHODIMP SetAddressListSearchOrder(IMAPISession &Session, const list<string> &SearchList);
SRowSet *AllocSRowSet(const list<SRow> &SRowList, const LPVOID pParent);
STDMETHODIMP CopySBinary(SBinary &sbDest, const SBinary &sbSrc, const LPVOID pParent);
string GetFilename(const char *Pathname);
//Start modified by CL

int Setcustomization(IMAPISession &lpSession);

//End modified by CL


int main(int argc, char *argv[]) {
   HRESULT hr;

   if (argc == 1) {
      cout << "Set MAPI address list search order" << endl;
      cout << "Shawn Poulson <spoulson@explodingcoder.com>, 2008.10.24" << endl;
      //Start modified by CL

      cout << "Modified to work with Outlook2010 by CL <cleitet@gmail.com> 2012-10-16" <<endl;

      //End modified by CL

      cout << endl;
      cout << "Usage: " << GetFilename(argv[0]) << " \"Address List A\" [ \"Address List B\" ...]" << endl;
      cout << endl;
      cout << "Example lists:" << endl;
      cout << " All Contacts           (All Outlook contacts folders)" << endl;
      cout << " Contacts               (Default Outlook contacts)" << endl;
      cout << " Global Address List" << endl;
      cout << " All Address Lists      (All lists defined in Exchange)" << endl;
      cout << " All Users              (All Exchange users)" << endl;
      return 1;
   }

   // Initialize MAPI
   hr = MAPIInitialize(NULL);
   if (FAILED(hr)) {
      cerr << "Error initializing MAPI" << endl;
      goto Exit;
   }

   // Logon to MAPI with default profile
   LPMAPISESSION lpSession;
   hr = MAPILogon(&lpSession);
   if (FAILED(hr)) goto Exit;

   if (lpSession != NULL) {

      // Compile command line arguments to SearchList
      list<string> SearchList;
      for (int i = 1; i < argc; i++) {
         SearchList.push_back(argv[i]);
      }

      // Save SearchList
      SetAddressListSearchOrder(*lpSession, SearchList);

	  //Start modified by CL

	  Setcustomization(*lpSession);

	  //End modified by CL

      // Clean up
      MAPILogoff(*lpSession);
      hr = lpSession->Release();
      if (FAILED(hr)) {
         cerr << "Warning: lpSession->Release() failed" << endl;
      }
   }
   else {
      cerr << "Error logging on to MAPI" << endl;
      goto Exit;
   }

Exit:
   MAPIUninitialize();
   return 0;
}

// Logon to MAPI session with default profile
STDMETHODIMP MAPILogon(LPMAPISESSION *lppSession) {
   HRESULT hr = MAPILogonEx(NULL, NULL, NULL, MAPI_USE_DEFAULT, lppSession);
   if (FAILED(hr)) {
      cerr << "Error logging on to MAPI." << endl;
   }
   return hr;
}

// Logoff MAPI session
void MAPILogoff(IMAPISession &Session) {
   HRESULT hr = Session.Logoff(NULL, NULL, 0);
   if (FAILED(hr)) {
      cerr << "Warning: MAPI log off failed" << endl;
   }
}

  //Start modified by CL

  // Set address list search order to Custom
  int Setcustomization(IMAPISession &lpSession) {
  HRESULT hr;
  LPPROFSECT lpProfileSection = NULL;
  LPSPropValue lpPropValue = NULL;
  LONG FAR * ulPropCnt = NULL;
  LPSPropValue FAR * pProps = NULL;

  hr = lpSession.OpenProfileSection((LPMAPIUID)&IID_CAPONE_PROF, NULL, MAPI_MODIFY  , &lpProfileSection);
  if (FAILED (hr)) {
	  cerr << "Error: Could not open the CAPONE profile section" <<endl;
	  return 1;
  }

  //hr = HrGetOneProp(lpProfileSection, PR_AB_CHOOSE_DIRECTORY_AUTOMATICALLY, &lpPropValue);// good for select automatically
  hr = HrGetOneProp(lpProfileSection, PR_AB_SEARCH_PATH_CUSTOMIZATION, &lpPropValue);
  if (FAILED (hr)) {
	  lpProfileSection->Release();
	  cerr << "Error: Could not open the property of the address book to set address list search order to Custom" <<endl;
	  return 2;
  }

  //cout << "Server DN: %d\n", lpPropValue->Value.b;
  lpPropValue->Value.l = SEARCHPATHREORDERTYPE_RAW;
  hr = HrSetOneProp(lpProfileSection, lpPropValue);
  if (FAILED (hr)) {
	  MAPIFreeBuffer(lpPropValue);
	  cerr << "Error: Could not set the property of the address list search order to Custom" <<endl;
	  return 3; // can't get the prop
  }

  cout << "Configured address list search order to be Custom" <<endl;
  return S_OK;
}
//End modified by CL

// Set address list search order
STDMETHODIMP SetAddressListSearchOrder(IMAPISession &Session, const list<string> &SearchList) {
   HRESULT hr;
   LPADRBOOK lpAddrBook = NULL;
   LPVOID tempLink;
   SRowSet *NewSRowSet = NULL;

   // New SRow list of search path
   list<SRow> NewSRowList;

   // Corresponding SPropValue's for SRow.lpProps in NewSRowList
   list<SPropValue> NewSPropList;

   // Setup struct specifying MAPI fields to query
   enum {
        abPR_ENTRYID,         // Field index for ENTRYID
        abPR_DISPLAY_NAME_A,  // Field index for display name
        abNUM_COLS            // Automatically set to number of fields
   };
   static SizedSPropTagArray(abNUM_COLS, abCols) = {
        abNUM_COLS,        // Num fields to get (2)
        PR_ENTRYID,        // Get ENTRYID struct
        PR_DISPLAY_NAME_A  // Get display name
   };

   // Open address book
   hr = Session.OpenAddressBook(NULL, NULL, NULL, &lpAddrBook);
   if (FAILED(hr)) {
      cerr << "Error getting MAPI Address book." << endl;
      goto Exit;
   }

   TraceSearchPath(*lpAddrBook);

   ULONG ulObjType;
   LPMAPICONTAINER pIABRoot = NULL;
   hr = lpAddrBook->OpenEntry(0, NULL, NULL, 0, &ulObjType, (LPUNKNOWN *)&pIABRoot);
   if (FAILED(hr) || ulObjType != MAPI_ABCONT) {
      cerr << "Error opening MAPI Address book root entry." << endl;
      if (SUCCEEDED(hr)) hr = E_UNEXPECTED;
      goto Cleanup;
   }

   // Setup MAPI memory allocation link
   MAPIAllocateBuffer(0, &tempLink);

   // Query MAPI for all address lists
   LPMAPITABLE pHTable = NULL;
   hr = pIABRoot->GetHierarchyTable(CONVENIENT_DEPTH, &pHTable);
   if (FAILED(hr)) {
      cerr << "Error obtaining MAPI address list hierarchy." << endl;
      goto Cleanup;
   }

   LPSRowSet pQueryRows = NULL;
   hr = HrQueryAllRows(pHTable, (LPSPropTagArray)&abCols, NULL, NULL, 0, &pQueryRows);
   if (FAILED(hr)) {
      cerr << "Error getting MAPI address lists." << endl;
      goto Cleanup;
   }

   // Cross reference pQueryRows with SearchList for matches
   for (list<string>::const_iterator SearchListIter = SearchList.begin(); SearchListIter != SearchList.end(); SearchListIter++) {
      const string &SearchName = *SearchListIter;

      // Is SearchName in the pQueryRows list?
      for (ULONG i = 0; i < pQueryRows->cRows && pQueryRows->aRow[i].lpProps[abPR_DISPLAY_NAME_A].ulPropTag == PR_DISPLAY_NAME_A; i++) {
         SRow &QueryRow = pQueryRows->aRow[i];
         string ContainerName = QueryRow.lpProps[abPR_DISPLAY_NAME_A].Value.lpszA;

         if (ContainerName == SearchName) {
            // Found a match!
            cout << "Adding address list search path: " << SearchName << endl;

            // Build SRow/SPropValue structs
            // Assumptions: SRow contains 1 SPropValue of type SBinary
            SPropValue TmpSPropValue = { QueryRow.lpProps[0].ulPropTag, QueryRow.lpProps[0].dwAlignPad };
            NewSPropList.push_back(TmpSPropValue);
            SPropValue &NewSPropValue = NewSPropList.back();

            SRow TmpSRow = { QueryRow.ulAdrEntryPad, 1, &NewSPropValue };
            NewSRowList.push_back(TmpSRow);
            SRow &NewSRow = NewSRowList.back();

            // Safely copy binary portion of SPropValue
            hr = CopySBinary(
               NewSRow.lpProps[0].Value.bin,
               QueryRow.lpProps[0].Value.bin,
               tempLink);
            if (FAILED(hr)) {
               cerr << "Error while building MAPI data." << endl;
               goto Cleanup;
            }

            // break out of inner pQueryRows loop and continue to next in SearchList
            break;
         }
      } // for (i in pQueryRows)
   } // for (SearchList)

   // Convert NewSRowList to SRowSet
   NewSRowSet = AllocSRowSet(NewSRowList, tempLink);
   if (NewSRowSet == NULL) goto Cleanup;

   hr = lpAddrBook->SetSearchPath(0, NewSRowSet);
   if (FAILED(hr)) {
      cerr << "Error while saving address list search path" << endl;
      goto Cleanup;
   }

   TraceSearchPath(*lpAddrBook);

Cleanup:
   if (NewSRowSet) delete[] NewSRowSet;
   MAPIFreeBuffer(tempLink);
   if (lpAddrBook) lpAddrBook->Release();
Exit:
   if (FAILED(hr)) cerr << "HRESULT = 0x" << hex << hr << endl;
   return hr;
}

// Convert list<SRow> to newly allocated SRowSet
// User code is responsible for freeing the returned pointer w/ delete[].
// Assumptions: Each SRow contains 1 SPropValue of type SBinary
SRowSet *AllocSRowSet(const list<SRow> &SRowList, const LPVOID pParent) {
   HRESULT hr;

   // Calculate size of SRowSet
   size_t RowSetSize = offsetof(SRowSet, aRow) + sizeof(SRow) * SRowList.size();

   // Calculate size of referenced SPropValue's
   size_t PropValueSize = sizeof(SPropValue) * SRowList.size();

   // Allocate all memory in one block
   char *Data = new char[RowSetSize + PropValueSize];
   SRowSet *Rows = (SRowSet *)Data;
   SPropValue *PropValues = (SPropValue *)&Data[RowSetSize];

   // Populate structures
   Rows->cRows = (ULONG)SRowList.size();
   int i = 0;
   for (list<SRow>::const_iterator iter = SRowList.begin(); iter != SRowList.end(); i++, iter++) {
      const SRow &SrcRow = *iter;
      SRow &DstRow = Rows->aRow[i];

      DstRow.ulAdrEntryPad = SrcRow.ulAdrEntryPad;
      DstRow.cValues = 1;
      DstRow.lpProps = &PropValues[i];
      DstRow.lpProps->dwAlignPad = SrcRow.lpProps->dwAlignPad;
      DstRow.lpProps->ulPropTag = SrcRow.lpProps->ulPropTag;
      
      hr = CopySBinary(
         DstRow.lpProps[0].Value.bin,
         SrcRow.lpProps[0].Value.bin,
         pParent);
      if (FAILED(hr)) {
         cerr << "Error while building MAPI data." << endl;
         delete[] Data;
         return NULL;
      }
   }

   return Rows;
}

// Copy SPropValue.Value.bin buffer
// Allocate memory for destination buffer using pParent as link to memory heap
STDMETHODIMP CopySBinary(SBinary &sbDest, const SBinary &sbSrc, const LPVOID pParent) {
   HRESULT hr = S_OK;
   sbDest.cb = sbSrc.cb;

   if (sbSrc.cb) {
      if (pParent) {
         hr = MAPIAllocateMore(
            sbSrc.cb,
            pParent,
            (LPVOID *)&sbDest.lpb);
      }
      else {
         hr = MAPIAllocateBuffer(
            sbSrc.cb,
            (LPVOID*) &sbDest.lpb);
      }

      if (SUCCEEDED(hr))
         CopyMemory(sbDest.lpb, sbSrc.lpb, sbSrc.cb);
   }

   return hr;
}

string GetFilename(const char *Pathname) {
   char fname[_MAX_FNAME];
   _splitpath_s(Pathname, NULL, 0, NULL, 0, fname, sizeof(fname), NULL, 0);
   return string(fname);
}
