#include "stdafx.h"

int main(int argc, char* argv[])
{
	HRESULT hr = S_OK;
	CComPtr <ISpVoice>	cpVoice;
	CComPtr<ISpObjectTokenCategory> cpSpCategory = NULL;
	CComPtr<IEnumSpObjectTokens> cpSpEnumTokens;

	if (FAILED(::CoInitialize(NULL)))
		return FALSE;

	cout << "initialized\n";

	cout << "creating voice\n";

	hr = cpVoice.CoCreateInstance(CLSID_SpVoice);
	if (!SUCCEEDED(hr))
		return FALSE;

	hr = cpSpCategory.CoCreateInstance(CLSID_SpObjectTokenCategory);
	if (!SUCCEEDED(hr))
		return FALSE;

	if (SUCCEEDED(hr = SpGetCategoryFromId(SPCAT_VOICES, &cpSpCategory)))
	{

		if (SUCCEEDED(hr = cpSpCategory->EnumTokens(NULL, NULL, &cpSpEnumTokens)))
		{
			CComPtr<ISpObjectToken> pSpTok;
			ULONG count, i = 0;
			hr = cpSpEnumTokens->GetCount(&count);
			while (i < count){
				hr = cpSpEnumTokens->Next(1, &pSpTok, NULL);
				i++;
				// do something with the token here; for example, set the voice
				if (SUCCEEDED(hr)){
					cpVoice->SetVoice(pSpTok);
					cpVoice->Speak(L"This is cool", SpeechVoiceSpeakFlags::SVSFlagsAsync, NULL);
					cpVoice->WaitUntilDone(-1);
					// NOTE:  IEnumSpObjectTokens::Next will *overwrite* the pointer; must manually release
					pSpTok.Release();
				}
			}
		}
	}

	//searching for voice with specific attribute
	hr = cpSpCategory->EnumTokens(L"Name=Microsoft Hazel Desktop", NULL, &cpSpEnumTokens);
	if (SUCCEEDED(hr)){
		ULONG count;
		cpSpEnumTokens->GetCount(&count);
		cout << count << endl;
	}
	//always remember to release everything at the end
	cpVoice.Release();
	cpSpCategory.Release();

	::CoUninitialize();
	return TRUE;
}