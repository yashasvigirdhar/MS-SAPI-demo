#include "stdafx.h"

HRESULT hr = S_OK;
CComPtr <ISpVoice>	cpVoice;
std::wstring utterance_;
int utterance_id_;
int prefix_len_;
ULONG stream_number_;
int char_position_;
bool paused_;

const wchar_t kAttributesKey[] = L"Attributes";
const wchar_t kGenderValue[] = L"Gender";
const wchar_t kLanguageValue[] = L"Language";

enum TtsGenderType {
	TTS_GENDER_NONE,
	TTS_GENDER_MALE,
	TTS_GENDER_FEMALE
};


void OnSpeechEvent() {

	SPEVENT event;
	while (S_OK == cpVoice->GetEvents(1, &event, NULL)) {
		if (event.ulStreamNum != stream_number_)
			continue;

		switch (event.eEventId) {
		case SPEI_START_INPUT_STREAM:
			std::cout << "SPEI_START_INPUT_STREAM" << endl;
			break;
		case SPEI_END_INPUT_STREAM:
			char_position_ = utterance_.size();
			std::cout << "SPEI_END_INPUT_STREAM " << char_position_ << endl;
			break;
		case SPEI_TTS_BOOKMARK:

			break;
		case SPEI_WORD_BOUNDARY:
			char_position_ = static_cast<ULONG>(event.lParam) - prefix_len_;
			std::cout << "SPEI_WORD_BOUNDARY" << endl;
			SPVOICESTATUS eventStatus;
			cpVoice->GetStatus(&eventStatus, NULL);
			ULONG start, end;
			start = eventStatus.ulInputWordPos;
			end = eventStatus.ulInputWordLen;
			std::cout << "start " << start << "  end " << start + end - 1 << "  status " << eventStatus.dwRunningState << endl;
			break;
		case SPEI_SENTENCE_BOUNDARY:
			char_position_ = static_cast<ULONG>(event.lParam) - prefix_len_;
			std::cout << "SPEI_SENTENCE_BOUNDARY" << endl;
			cpVoice->GetStatus(&eventStatus, NULL);
			start = eventStatus.ulInputSentPos;
			end = eventStatus.ulInputSentLen;
			std::cout << "start " << start << "  end " << start + end - 1 << "  status " << eventStatus.dwRunningState << endl;
			break;
		}
	}
}




static void __stdcall SpeechEventCallback(WPARAM w_param, LPARAM l_param);

HRESULT WaitAndPumpMessagesWithTimeout(HANDLE hWaitHandle, DWORD dwMilliseconds)
{
	HRESULT hr = S_OK;
	BOOL fContinue = TRUE;

	while (fContinue)
	{
		DWORD dwWaitId = ::MsgWaitForMultipleObjectsEx(1, &hWaitHandle, dwMilliseconds, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
		switch (dwWaitId)
		{
		case WAIT_OBJECT_0:
		{
							  fContinue = FALSE;
		}
			break;

		case WAIT_OBJECT_0 + 1:
		{
								  MSG Msg;
								  while (::PeekMessage(&Msg, NULL, 0, 0, PM_REMOVE))
								  {
									  ::TranslateMessage(&Msg);
									  ::DispatchMessage(&Msg);
								  }
		}
			break;

		case WAIT_TIMEOUT:
		{
							 hr = S_FALSE;
							 fContinue = FALSE;
		}
			break;
		default:// Unexpected error
		{
					fContinue = FALSE;
					hr = E_FAIL;
		}
			break;
		}
	}
	return hr;
}

int main(int argc, char* argv[])
{

	CComPtr<IEnumSpObjectTokens> voice_tokens;

	if (FAILED(::CoInitialize(NULL)))
		return FALSE;

	std::cout << "initialized\n";

	//	cout << "creating voice\n";

	if (S_OK != cpVoice.CoCreateInstance(CLSID_SpVoice))
		return FALSE;
	std::cout << "voice initialized" << endl;

	ULONGLONG event_mask =
		SPFEI(SPEI_START_INPUT_STREAM) |
		SPFEI(SPEI_TTS_BOOKMARK) |
		SPFEI(SPEI_WORD_BOUNDARY) |
		SPFEI(SPEI_SENTENCE_BOUNDARY) |
		SPFEI(SPEI_END_INPUT_STREAM);

	cpVoice->SetInterest(event_mask, event_mask);

	std::cout << "interests set" << endl;

	cpVoice->SetNotifyCallbackFunction(SpeechEventCallback, 0, 0);

	std::cout << "callback function set" << endl;

	cpVoice->Speak(L"This should work", SPF_ASYNC, &stream_number_);
	HANDLE hWait = cpVoice->SpeakCompleteEvent();
	WaitAndPumpMessagesWithTimeout(hWait, INFINITE);


	cpVoice.Release();
	::CoUninitialize();
	return TRUE;
}

void __stdcall SpeechEventCallback(WPARAM w_param, LPARAM l_param) {
	std::cout << "callback function\n";
	OnSpeechEvent();
}