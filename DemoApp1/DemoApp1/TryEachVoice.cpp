#include "stdafx.h"

const wchar_t kAttributesKey[] = L"Attributes";
const wchar_t kGenderValue[] = L"Gender";
const wchar_t kLanguageValue[] = L"Language";

enum TtsGenderType {
	TTS_GENDER_NONE,
	TTS_GENDER_MALE,
	TTS_GENDER_FEMALE
};

int main(int argc, char* argv[])
{
	HRESULT hr = S_OK;
	CComPtr <ISpVoice>	cpVoice;
	CComPtr<IEnumSpObjectTokens> voice_tokens;

	if (FAILED(::CoInitialize(NULL)))
		return FALSE;

	cout << "initialized\n";

	//	cout << "creating voice\n";

	if (S_OK != cpVoice.CoCreateInstance(CLSID_SpVoice))
		return FALSE;

	if (S_OK != SpEnumTokens(SPCAT_VOICES, NULL, NULL, &voice_tokens))
		return FALSE;

	unsigned long voice_count, i = 0;
	hr = voice_tokens->GetCount(&voice_count);
	cout << " count " << voice_count << endl;
	for (unsigned int i = 0; i < voice_count; i++){
		//cout << i << endl;
		CComPtr<ISpObjectToken> voice_token;
		if (S_OK != voice_tokens->Next(1, &voice_token, NULL))
			return FALSE;

		// do something with the token here; for example, set the voice
		WCHAR *description;
		if (S_OK != SpGetDescription(voice_token, &description))
			return FALSE;


		CComPtr<ISpDataKey> attributes;
		if (S_OK != voice_token->OpenKey(kAttributesKey, &attributes))
			return FALSE;

		WCHAR *gender_s;
		TtsGenderType gender;
		if (S_OK == attributes->GetStringValue(kGenderValue, &gender_s)){
			if (0 == _wcsicmp(gender_s, L"male"))
				gender = TTS_GENDER_MALE;
			else if (0 == _wcsicmp(gender_s, L"female"))
				gender = TTS_GENDER_FEMALE;
		}


		WCHAR *language;
		if (S_OK != attributes->GetStringValue(kLanguageValue, &language))
			return FALSE;

		WCHAR *uri;
		voice_token->GetId(&uri);

		cout << "voice no " << i << endl;
		cout << "description " << wprintf(L"%s\n", description);
		cout << "gender " << wprintf(L"%s\n", gender_s);
		cout << "language " << wprintf(L"%s\n", language);
		cout << "uri ";
		wprintf(L"%s\n\n", uri);
		voice_token.Release();
		attributes.Release();
		// NOTE:  IEnumSpObjectTokens::Next will *overwrite* the pointer; must manually release

	}

	cpVoice.Release();
	::CoUninitialize();
	return TRUE;
}

//searching for voice with specific attribute
//hr = cpSpCategory->EnumTokens(L"Name=Microsoft Hazel Desktop", NULL, &cpSpEnumTokens);
//if (SUCCEEDED(hr)){
//ULONG count;
//cpSpEnumTokens->GetCount(&count);
//cout << count << endl;
//	}
//always remember to release everything at the end

