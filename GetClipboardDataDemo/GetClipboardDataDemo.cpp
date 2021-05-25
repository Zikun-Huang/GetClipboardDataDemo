// GetClipboardDataDemo.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <windows.h>
#include <ctime>

#define MAX_RETRIES 40

UINT g_rtfFormat = RegisterClipboardFormatW(L"Rich Text Format");
UINT g_htmlFormat = RegisterClipboardFormatW(L"HTML Format");

BOOL OpenClipboardWrapper();

int main()
{
	HANDLE handle = NULL;
	size_t length = 0;
	std::clock_t startTime, endTime;

	std::cout << "Please make sure that you have copied data to system Clipboard "
		<< "from Test_File.xlsx." << std::endl;
	system("pause");

	if (!OpenClipboardWrapper()) {
		std::cout << "Cannot open clipboard." << std::endl;
		system("pause");
		return 1;
	}

	startTime = clock();
	std::cout << "Fetching unicode text from system Clipboard." << std::endl;
	if (handle = GetClipboardData(CF_UNICODETEXT)) {
		std::cout << "Got unicode text from system clipboard" << std::endl;
		WCHAR* data = (WCHAR*)(GlobalLock(handle));
		if (data) {
			length = GlobalSize(handle);
			std::cout << "Unicode text data length = " << length << " bytes." << std::endl;
		}
		else {
			std::cout << "GlobalLock returned NULL for unicode text. Err code = " << GetLastError() << std::endl;
		}
		GlobalUnlock(handle);
	}
	else {
		std::cout << "Fail to get unicode text from system clipboard" << std::endl;
	}
	endTime = clock();
	std::cout << "Fetching unicode text total time : "
		<< (double)(endTime - startTime) / CLOCKS_PER_SEC
		<< "s" << std::endl << std::endl;

	handle = NULL;
	length = 0;

	startTime = clock();
	std::cout << "Fetching RTF format data from system Clipboard." << std::endl;
	if (handle = GetClipboardData(g_rtfFormat)) {
		std::cout << "Got RTF format data from system clipboard" << std::endl;
		char* data = (char*)(GlobalLock(handle));
		if (data) {
			length = GlobalSize(handle);
			std::cout << "RTF format data length = " << length << " bytes." << std::endl;
		}
		else {
			std::cout << "GlobalLock returned NULL for RTF data. Err code = " << GetLastError() << std::endl;
		}
		GlobalUnlock(handle);
	}
	else {
		std::cout << "Fail to get RTF format data from system clipboard" << std::endl;
	}
	endTime = clock();
	std::cout << "Fetching RTF format data total time : "
		<< (double)(endTime - startTime) / CLOCKS_PER_SEC
		<< "s" << std::endl << std::endl;

	handle = NULL;
	length = 0;

	startTime = clock();
	std::cout << "Fetching HTML format data from system Clipboard." << std::endl;
	if (handle = GetClipboardData(g_htmlFormat)) {
		std::cout << "Got HTML format data from system clipboard" << std::endl;
		char* data = (char*)(GlobalLock(handle));
		if (data) {
			length = GlobalSize(handle);
			std::cout << "HTML format data length = " << length << " bytes." << std::endl;
		}
		else {
			std::cout << "GlobalLock returned NULL for HTML data. Err code = " << GetLastError() << std::endl;
		}
		GlobalUnlock(handle);
	}
	else {
		std::cout << "Fail to get HTML format data from system clipboard" << std::endl;
	}
	endTime = clock();

	std::cout << "Fetching HTML format data total time : "
		<< (double)(endTime - startTime) / CLOCKS_PER_SEC
		<< "s" << std::endl << std::endl;

	CloseClipboard();
	system("pause");
	return 0;
}


/*
 *----------------------------------------------------------------------
 *
 * OpenClipboardWrapper --
 *
 *    Wrapper around the OpenClipboard Windows API.
 *    This wrapper function calls OpenClipboard repeatedly, while also
 *    calls Sleep(250) before each call. Once OpenClipboard succeeds,
 *    the loop is ended.
 *
 *----------------------------------------------------------------------
 */

BOOL
OpenClipboardWrapper()
{
	BOOL result = FALSE;
	UINT error = ERROR_SUCCESS;
	UINT retryCount = 0;

	do {
		result = OpenClipboard(NULL);
		if (result) {
			std::cout << "OpenClipboard succeeded, after retrying " << retryCount << " times." << std::endl << std::endl;
			break;
		}

		if ((error = GetLastError()) != ERROR_ACCESS_DENIED) {
			std::cout << "OpenClipboard failed with error = " << error << std::endl << std::endl;
			break;
		}

		std::cout << "OpenClipboard failed with ERROR_ACCESS_DENIED." << std::endl << std::endl;
		if (retryCount < MAX_RETRIES - 1) {
			Sleep(250);
		}
		retryCount++;
	} while (retryCount < MAX_RETRIES);

	return result;
}


// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
