//
// Set MAPI Default Address List when opening the Outlook Address Book
//
// If Outlook is already open, it may need to be restarted for change to take effect.
//
// Shawn Poulson <spoulson@explodingcoder.com>, 2008.10.24
//
// Modified to work with Outlook 2010 by CL <cleitet@gmail.com> 2012-08-21
//
// Modified to check that Microsoft Outlook is installed <christoffer@jnk.dk> 2014-06-23

#include "stdafx.h"
#include <mapix.h>
#include <mapiutil.h>
// Start modified by CL

#include <INITGUID.H>

// End modified by CL

#include <string>
#include <iostream>
using namespace std;

//Start modified by CL

// See http://msdn.microsoft.com/en-us/library/office/hh204510
// PidTagAddressBookChooseDirectoryAutomatically Canonical Property
#define PR_AB_CHOOSE_DIRECTORY_AUTOMATICALLY PROP_TAG( PT_BOOLEAN, 0x3D1C)

// http://blogs.msdn.com/b/stephen_griffin/archive/2010/09/13/you-chose-wisely.aspx
// Capone profile section
// {00020D0A-0000-0000-C000-000000000046}
DEFINE_OLEGUID(IID_CAPONE_PROF, 0x00020d0a, 0, 0);

//End modified by CL

STDMETHODIMP MAPILogon(LPMAPISESSION *lpMAPISession);
void MAPILogoff(IMAPISession &Session);
STDMETHODIMP SetDefaultAddressList(IMAPISession &Session, const string &AddressList);
STDMETHODIMP ResolveAddressList(IMAPISession &Session, const string &AddressList, LPVOID pAllocLink, ULONG *cbEntry, LPENTRYID *Entry);
string GetFilename(const char *Pathname);

int main(int argc, char *argv[]) {
   HRESULT hr = S_OK;

   if (argc != 2) {
      cout << "Set MAPI default address list" << endl;
      cout << "Shawn Poulson <spoulson@explodingcoder.com>, 2008.10.24" << endl;
      // Start modified by CL

      cout << "Modified to work with Outlook2010 by CL <cleitet@gmail.com> 2012-10-16" <<endl;

      // End modified by CL

      cout << endl;
      cout << "Usage: " << GetFilename(argv[0]) << " \"Address List\"" << endl;
      cout << endl;
      cout << "Example lists:" << endl;
      cout << " All Contacts           (All Outlook contacts folders)" << endl;
      cout << " Contacts               (Default Outlook contacts)" << endl;
      cout << " Global Address List" << endl;
      cout << " All Address Lists      (All lists defined in Exchange)" << endl;
      cout << " All Users              (All Exchange users)" << endl;
      return 1;
   }

   // Start modified by CN

   HKEY subKey = nullptr;
   LONG result = RegOpenKeyEx(HKEY_LOCAL_MACHINE, TEXT("SOFTWARE\\Clients\\Mail\\Microsoft Outlook"), 0, KEY_READ, &subKey);
   if (result != ERROR_SUCCESS)
   {
	   cout << "Error: Could not find Microsoft Outlook on machine" << endl;
	   return 0;
   }

   // End modified by CN

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
      // Save SearchList
      SetDefaultAddressList(*lpSession, string(argv[1]));

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

// Set default address list by name
STDMETHODIMP SetDefaultAddressList(IMAPISession &Session, const string &AddressList) {
   HRESULT hr = S_OK;

   // Initialize memory allocation
   LPVOID pAllocLink = NULL;
   MAPIAllocateBuffer(0, &pAllocLink);

   // Resolve address list name to ENTRYID
   ULONG cbEntryID;
   LPENTRYID lpEntryID;
   hr = ResolveAddressList(Session, AddressList, pAllocLink, &cbEntryID, &lpEntryID);
   if (FAILED(hr)) {
      cerr << "Unable to resolve address list name '" << AddressList << "'." << endl;
      return hr;
   }

   // Open address book
   LPADRBOOK lpAddrBook;
   hr = Session.OpenAddressBook(NULL, NULL, NULL, &lpAddrBook);
   if (FAILED(hr)) {
      cerr << "Error getting MAPI Address book." << endl;
      goto Exit;
   }
   // Start Modified by CL
   TraceDefaultDir(*lpAddrBook);

   // Fix IID_CAPONE_PROF See note here: (http://msdn.microsoft.com/en-us/library/office/cc839797)
   // Turning off the PR_AB_CHOOSE_DIRECTORY_AUTOMATICALLY property
   // Set default address list

   cout << "Resetting choose automatically property" <<endl;

   LPPROFSECT lpProfileSection;
   hr = Session.OpenProfileSection((LPMAPIUID)&IID_CAPONE_PROF, NULL, MAPI_MODIFY  , &lpProfileSection);
   if (FAILED(hr)) {
	   cerr << "Error opening profile section" << endl;
	   goto Exit;
   }


   SPropValue lpPropValue;
   lpPropValue.ulPropTag = PR_AB_CHOOSE_DIRECTORY_AUTOMATICALLY;
   lpPropValue.Value.l = 0;

   hr = HrSetOneProp(lpProfileSection, &lpPropValue);
   if (FAILED(hr)) {
	   cerr << "Error setting property for automatic choosing to off with error" << hr << endl;
	   goto Exit;
   }

   //End modified by CL

   // Display feedback
   TraceDefaultDir(*lpAddrBook);
   cout << "Setting default address list: " << AddressList << endl;

   // Set default address list
   hr = lpAddrBook->SetDefaultDir(cbEntryID, lpEntryID);
   if (FAILED(hr)) {
      cerr << "Error setting default address list" << endl;
      goto Exit;
   }


   // Start modified by CL

   //TraceDefaultDir(*lpAddrBook);

   // End modified by CL

Exit:
   if (pAllocLink) MAPIFreeBuffer(pAllocLink);
   return hr;
}

// Resolve address list name to ENTRYID
// Memory is allocated through MAPIAllocateBuffer using pAllocLink
STDMETHODIMP ResolveAddressList(IMAPISession &Session, const string &AddressList, LPVOID pAllocLink, ULONG *cbEntry, LPENTRYID *Entry) {
   HRESULT hr = S_OK;

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
   LPADRBOOK lpAddrBook;
   hr = Session.OpenAddressBook(NULL, NULL, NULL, &lpAddrBook);
   if (FAILED(hr)) {
      cerr << "Error getting MAPI Address book." << endl;
      goto Exit;
   }

   ULONG ulObjType;
   LPMAPICONTAINER pIABRoot = NULL;
   hr = lpAddrBook->OpenEntry(0, NULL, NULL, 0, &ulObjType, (LPUNKNOWN *)&pIABRoot);
   if (FAILED(hr) || ulObjType != MAPI_ABCONT) {
      cerr << "Error opening MAPI Address book root entry." << endl;
      if (SUCCEEDED(hr)) hr = E_UNEXPECTED;
      goto Cleanup;
   }

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

   // Is AddressList in the pQueryRows list?
   for (ULONG i = 0; i < pQueryRows->cRows && pQueryRows->aRow[i].lpProps[abPR_DISPLAY_NAME_A].ulPropTag == PR_DISPLAY_NAME_A; i++) {
      SRow &QueryRow = pQueryRows->aRow[i];
      string ContainerName = QueryRow.lpProps[abPR_DISPLAY_NAME_A].Value.lpszA;

      if (ContainerName == AddressList) {
         // Found a match!
         // Build ENTRYID struct
         ULONG cbNewEntryID = QueryRow.lpProps[abPR_ENTRYID].Value.bin.cb;
         LPENTRYID lpNewEntryID;
         MAPIAllocateMore(cbNewEntryID, pAllocLink, (LPVOID *)&lpNewEntryID);
         memcpy(lpNewEntryID, QueryRow.lpProps[abPR_ENTRYID].Value.bin.lpb, cbNewEntryID);

         // Set return values
         *cbEntry = cbNewEntryID;
         *Entry = lpNewEntryID;

         // Break out
         break;
      }
   }

Cleanup:
   if (lpAddrBook) lpAddrBook->Release();

Exit:
   return hr;
}

string GetFilename(const char *Pathname) {
   char fname[_MAX_FNAME];
   _splitpath_s(Pathname, NULL, 0, NULL, 0, fname, sizeof(fname), NULL, 0);
   return string(fname);
}
