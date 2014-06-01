#include "stdafx.h"

int main(int argc, char* argv[])
{
	HRESULT hr = S_OK;
	CComPtr <ISpVoice>	cpVoice;
	CComPtr <ISpStream> cpStream;
	CComPtr <IStream> cpBaseStream;
	GUID guidFormat; WAVEFORMATEX* pWavFormatEx = nullptr;
	
	if (FAILED(::CoInitialize(NULL)))
		return FALSE;
	
	cout<<"initialized\n";

	if (SUCCEEDED(hr)){
		cout<<"creating voice\n";
		hr = cpVoice.CoCreateInstance(CLSID_SpVoice);
	}

	if (SUCCEEDED(hr)){
		cout<<"voice created, initialzing istream\n";
		hr = cpStream.CoCreateInstance(CLSID_SpStream);
	}
	if (SUCCEEDED(hr))
	{
		cout<<"istream created, creating basestream\n";
		hr = CreateStreamOnHGlobal(NULL, TRUE, &cpBaseStream);
	}
	if (SUCCEEDED(hr))
	{
		cout<<"basestream created, setting format\n";
		hr = SpConvertStreamFormatEnum(SPSF_22kHz16BitMono, &guidFormat,
			&pWavFormatEx);
	}
	if (SUCCEEDED(hr))
	{
		cout<<"format set,setting the basestream\n";
		hr = cpStream->SetBaseStream(cpBaseStream, guidFormat,
			pWavFormatEx);
		cpBaseStream.Release();
	}
	if (SUCCEEDED(hr))
	{
		cout<<"basestream set, setting cpvoice output\n";
		hr = cpVoice->SetOutput(cpStream, TRUE);
		if (SUCCEEDED(hr)){
			cout<<"output set, speaking\n";
			SpeechVoiceSpeakFlags my_Spflag = SpeechVoiceSpeakFlags::SVSFlagsAsync; // declaring and initializing Speech Voice Flags
			hr = cpVoice->Speak(L"I need to spend more and more time on this", my_Spflag, NULL);
			cpVoice->WaitUntilDone(-1);
		}
	}
	if (SUCCEEDED(hr)){
		cout<<"spoken correctly\n";

		/*To verify that the data has been written correctly, uncomment this, you should hear the voice. 
			cpVoice->SetOutput(NULL, FALSE);
			cpVoice->SpeakStream(cpStream, SPF_DEFAULT, NULL);
			*/

		//After SAPI writes the stream, the stream position is at the end, so we need to set it to the beginning.
		_LARGE_INTEGER a = { 0 };
		hr = cpStream->Seek(a, STREAM_SEEK_SET, NULL);
		
		//get the base istream from the ispstream 
		IStream *pIstream;
		cpStream->GetBaseStream(&pIstream);

		//calculate the size that is to be read
		STATSTG stats;
		pIstream->Stat(&stats, STATFLAG_NONAME);
		
		ULONG sSize = stats.cbSize.QuadPart;	//size of the data to be read
		cout << "size : " << stats.cbSize.QuadPart << endl;
		
		ULONG bytesRead;	//	this will tell the number of bytes that have been read
		char *pBuffer = new char[sSize];	//buffer to read the data
		
		//read the data into the buffer
		pIstream->Read(pBuffer, sSize, &bytesRead);
		
		cout << "bytesRead : "<< bytesRead << endl;
		
		/*uncomment the following to print the contents of the buffer
			cout << "following data read \n";
			for (int i = 0; i < sSize; i++)
				cout << pBuffer[i] << " ";
			cout << endl;
		*/
	
	}
	
	//don't forget to release everything
	cpStream.Release();
	cpVoice.Release();

	::CoUninitialize();
	return TRUE;
}