// MQTextToSpeech.cpp
#define SPDLOG_ACTIVE_LEVEL SPDLOG_LEVEL_TRACE
#define SPDLOG_EOL ""
#include <mq/Plugin.h>
#include <mq/base/WString.h>
#include "contrib/knightlog/knightlog.h"

#include <sapi.h>
#include <extras/wil/Constants.h>
#include <wil/result.h>
#include <wil/registry.h>
#include <imgui/ImGuiUtils.h>

#include "contrib/sphelper_stub.h"

PreSetup("MQTextToSpeech");
PLUGIN_VERSION(1.0);

constexpr const wchar_t* registry_key = LR"(HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices)";
constexpr const char* registry_tokens_key                 = R"(SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens)";

int speech_volume = 100;
int speech_speed_modifier = 0;
std::string current_voice = "Microsoft David";

#ifdef __WIN64 
constexpr const char* ini_voice_setting = "Voice";
#else
constexpr const char* ini_voice_setting = "Voice_32bit";
#endif

std::vector<std::string> voices;

HRESULT voice_com = -1;
ISpVoice* voice_ptr = nullptr;
ISpObjectToken* current_voice_idx = nullptr;

KnightLog* knightlog = nullptr;

using unique_couninitialize_call = wil::unique_call<decltype(&::CoUninitialize), ::CoUninitialize>;
unique_couninitialize_call cleanup_com;

void RePopulateVoices()
{
	voices.clear();
	HKEY hRegKey;
	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, registry_tokens_key, 0, KEY_READ, &hRegKey) == ERROR_SUCCESS)
	{
		DWORD dwSubkeyCount = 0;
		if (RegQueryInfoKey(hRegKey, nullptr, nullptr, nullptr, &dwSubkeyCount, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS)
		{
			for (DWORD i = 0; i < dwSubkeyCount; ++i)
			{
				char buffer[wil::max_registry_key_name_length + 1] = { 0 };
				DWORD bufferSize = wil::max_registry_key_name_length;
				if (RegEnumKeyEx(hRegKey, i, buffer, &bufferSize, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS && buffer[0] != '\0')
				{
					const std::string registry_subkey = fmt::format("{}\\{}\\Attributes", registry_tokens_key, buffer);
					
					HKEY hRegSubKey;
					if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, registry_subkey.c_str(), 0, KEY_READ, &hRegSubKey) == ERROR_SUCCESS)
					{
						char data[wil::max_registry_value_name_length + 1] = { 0 };
						DWORD dataSize = wil::max_registry_value_name_length;
						if (RegGetValue(hRegSubKey, nullptr, "Name", RRF_RT_REG_SZ, nullptr, data, &dataSize) == ERROR_SUCCESS)
						{
							voices.emplace_back(data);
						}
						RegCloseKey(hRegSubKey);
					}
				}
			}
		}
		RegCloseKey(hRegKey);
	}
}

bool SetVoice(const std::string& voice)
{
	if (voices.empty())
	{
		SPDLOG_ERROR("Could not set voice, no voices found.");
	}
	else
	{
		bool voice_set = false;
		if (std::find(voices.begin(), voices.end(), voice) != voices.end())
		{
			current_voice = voice;
			voice_set = true;
		}
		else
		{
			for (const auto& this_voice : voices)
			{
				if (ci_find_substr(this_voice, voice) != -1)
				{
					current_voice = this_voice;
					voice_set = true;
					break;
				}
			}
		}

		if (!voice_set)
		{
			SPDLOG_WARN(R"(No voice matching "{}" was found.  Defaulting to "{}".)", voice, voices[1]);
			current_voice = voices[1];
		}

		if (current_voice_idx)
		{
			current_voice_idx->Release();
			current_voice_idx = nullptr;
		}
		const std::wstring attributes = L"Name=" + utf8_to_wstring(current_voice);
		SpFindBestToken(registry_key, attributes.c_str(), L"", &current_voice_idx);
		if(current_voice_idx)
		{
			WritePrivateProfileString("Settings", ini_voice_setting, current_voice, INIFileName);
			return true;
		}

		SPDLOG_ERROR(R"(Could not get index for voice "{}".)", current_voice);
	}

	return false;
}

void LoadSettings()
{
	speech_volume = GetPrivateProfileInt("Settings", "Volume", speech_volume, INIFileName);
	speech_speed_modifier = GetPrivateProfileInt("Settings", "Speed Modifier", speech_speed_modifier, INIFileName);
	SetVoice(GetPrivateProfileString("Settings", ini_voice_setting, current_voice, INIFileName));
}

void ShowHelp()
{
	// Help is critical so it always shows.
	SPDLOG_CRITICAL("\a#00E0E0Usage: \n"
	                "    \ao/tts loglevel|reload|say|sayxml|speed|voice|volume [argument]\ax\n"
	                "Available Verbs:\n"
	                "    \aoLogLevel\ax - Set the log level (default is 'info')\n"
	                "    \aoReload\ax - Reload the settings from the ini file\n"
	                "    \aoSay\ax - Speak a line of text\n"
	                "    \aoSayXML\ax - Speak a line of text using XML formatting for pattern changes\n"
	                "    \aoSpeed\ax - Set the speech speed (-10 to 10)\n"
	                "    \aoVoice\ax - Change the voice (must be installed)\n"
	                "    \aoVolume\ax - Set the volume of speech (0 to 100)\n"
	                "  \ayNOTE: All settings are available in the settings menu.\ax\n"
	                "Example:\n"
	                "    \ao/tts say Some people juggle geese.\ax\n"
	                " \n"
	                "TLO and Members:\n"
	                "    \ao${TTS.Voice}\ax -- Which voice is being used\n"
	                "    \ao${TTS.Volume}\ax -- What volume is currently set\n"
	                "    \ao${TTS.Speed}\ax -- What speech speed is currently set\n"
	                "Example:\n"
	                "    \ao/tts say The current voice is ${TTS.Voice}.\ax\ax\n");
}

void Say(const char* words, bool usexml = false)
{
	SPDLOG_DEBUG("Saying: {}", words);
	voice_ptr->SetRate(speech_speed_modifier);
	voice_ptr->SetVolume(static_cast<unsigned short>(speech_volume));
	if (current_voice_idx)
	{
		voice_ptr->SetVoice(current_voice_idx);
	}
	const std::wstring wstr_words = utf8_to_wstring(words);
	int flags = SPF_DEFAULT | SPF_PURGEBEFORESPEAK | SPF_ASYNC;
	if (usexml)
	{
		flags |= SPF_IS_XML;
	}

	voice_ptr->Speak(wstr_words.c_str(), flags, nullptr);
}

PLUGIN_API void commandTTS(SPAWNINFO* pSpawn, char* szLine)
{
	if (szLine[0] == '\0')
	{
		ShowHelp();
	}
	else
	{
		char szArg1[MAX_STRING] = { 0 };
		GetArg(szArg1, szLine, 1);
		// Items with no second argument
		if (ci_equals(szArg1, "reload"))
		{
			RePopulateVoices();
			LoadSettings();
			SPDLOG_INFO("Settings Reloaded...");
		}
		else if (ci_equals(szArg1, "help"))
		{
			ShowHelp();
		}
		else {
			char szArg2[MAX_STRING] = { 0 };
			GetMaybeQuotedArg(szArg2, MAX_STRING, szLine, 2);
			if (szArg2[0] == '\0')
			{
				SPDLOG_ERROR("Additional arguments required :: /tts {}", szLine);
				ShowHelp();
			}
			else if (ci_equals(szArg1, "loglevel"))
			{
				if (knightlog->SetLogLevel(szArg2))
				{
					SPDLOG_CRITICAL("Log level set to: {}", knightlog->GetLogLevel());
				}
				else
				{
					SPDLOG_CRITICAL("Log level could not be set, invalid loglevel: {}", szArg2);
				}
			}
			else if(ci_equals(szArg1, "say"))
			{
				Say(szArg2);
			}
			else if(ci_equals(szArg1, "sayxml"))
			{
				Say(szArg2, true);
			}
			else if(ci_equals(szArg1, "speed"))
			{
				const int speed_setting = GetIntFromString(szArg2, -11);
				if (speed_setting < -10 || speed_setting > 10)
				{
					SPDLOG_ERROR("Invalid speed setting, must be between -10 and 10, you entered {} ", szArg2);
				}
				else
				{
					speech_speed_modifier = speed_setting;
					WritePrivateProfileInt("Settings", "Speed Modifier", speech_speed_modifier, INIFileName);
					SPDLOG_ERROR("Speed is now {}", speech_speed_modifier);
				}
			}
			else if(ci_equals(szArg1, "voice"))
			{
				if(SetVoice(szArg2))
				{
					SPDLOG_INFO("Voice is now {}", current_voice);
				}
				else
				{
					SPDLOG_ERROR("Could not set voice to {}", szArg2);
				}
			}
			else if(ci_equals(szArg1, "volume"))
			{
				const int volume_setting = GetIntFromString(szArg2, -1);
				if (volume_setting < 0 || volume_setting > 100)
				{
					SPDLOG_ERROR("Invalid volume setting, must be between 0 and 100: ", szArg2);
				}
				else
				{
					speech_volume = volume_setting;
					WritePrivateProfileInt("Settings", "Volume", speech_volume, INIFileName);
					SPDLOG_INFO("Volume is now {}", speech_volume);
				}
			}
			else
			{
				SPDLOG_ERROR("Invalid verb: {}", szArg1);
			}
		}
	}
}

void TTSImGuiSettingsPanel()
{
	ImGui::Text("Speech Volume");
	ImGui::SameLine();
	mq::imgui::HelpMarker("Volume 0 to 100. Adjust the volume of speech.\n\nINISetting: Volume");
	if (ImGui::SliderInt("##Speech Volume", &speech_volume, 0, 100))
	{
		WritePrivateProfileInt("Settings", "Volume", speech_volume, INIFileName);
	}
	ImGui::Separator();

	ImGui::Text("Speech Speed Modifier");
	ImGui::SameLine();
	mq::imgui::HelpMarker("Speed -10 to 10. Default is 0, others either speed up or slow down the speech.\n\nINISetting: Speed Modifier");
	if (ImGui::SliderInt("##Speech Speed Modifier", &speech_speed_modifier, -10, 10))
	{
		WritePrivateProfileInt("Settings", "Speed Modifier", speech_speed_modifier, INIFileName);
	}
	ImGui::Separator();

	ImGui::Text("Voice to Use");
	ImGui::SameLine();
	mq::imgui::HelpMarker("Select the voice that should be used for text to speech playback");
	if (ImGui::BeginCombo("##Voice", current_voice.c_str(), ImGuiComboFlags_HeightSmall))
	{
		for (const auto& voice_name : voices)
		{
			bool is_selected = voice_name == current_voice;
			if (ImGui::Selectable(voice_name.c_str(), is_selected))
			{
				if (!SetVoice(voice_name))
					is_selected = false;
			}

			if (is_selected)
				ImGui::SetItemDefaultFocus();
			
		}

		ImGui::EndCombo();
	}
	ImGui::SameLine();
	if (ImGui::Button("Refresh Voices"))
	{
		RePopulateVoices();
	}
	ImGui::SameLine();
	mq::imgui::HelpMarker("Repopulate the voices that are available on this system.");
	ImGui::Separator();

	ImGui::Text("Test Settings Here:");
	ImGui::SameLine();
	mq::imgui::HelpMarker("Enter a string here and then press the Test button to test it.");
	static char szCommand[MAX_STRING] = { 0 };
	ImGui::InputTextWithHint("##Command", "Some people juggle geese.", szCommand, MAX_STRING);
	ImGui::SameLine();
	if (ImGui::Button("Test"))
	{
		if (szCommand[0] == '\0')
		{
			Say(R"(You <pitch middle="10">really should enter</pitch> something to say if you want to test.  This box also supports <spell>XML</spell>)", true);
		}
		else
		{
			Say(szCommand, true);
		}
	}
}

class MQTTSType : public MQ2Type {
	public:
		enum class TTSMembers {
			Voice,
			Volume,
			Speed,
		};

		MQTTSType() : MQ2Type("TTS") {
			ScopedTypeMember(TTSMembers, Voice);
			ScopedTypeMember(TTSMembers, Volume);
			ScopedTypeMember(TTSMembers, Speed);
		}

		bool GetMember(MQVarPtr VarPtr, const char* Member, char* Index, MQTypeVar& Dest) override {
			MQTypeMember* pMember = FindMember(Member);
			if (!pMember) {
				return false;
			}

			switch (static_cast<TTSMembers>(pMember->ID))
			{
			case TTSMembers::Voice:
				Dest.Type = mq::datatypes::pStringType;
				Dest.Ptr = &current_voice[0];
				return true;
			case TTSMembers::Volume:
				Dest.Type = mq::datatypes::pIntType;
				Dest.Set(speech_volume);
				return true;
			case TTSMembers::Speed:
				Dest.Type = mq::datatypes::pIntType;
				Dest.Set(speech_speed_modifier);
				return true;
			}

			return false;
		}
};
MQTTSType *pTTSType = nullptr;

bool dataTTS(const char* szIndex, MQTypeVar& Dest)
{
	Dest.DWord = 0;
	Dest.Type = pTTSType;
	return true;
}

/**
 * @fn ShutdownPlugin
 *
 * This is called once when the plugin has been asked to shutdown.  The plugin has
 * not actually shut down until this completes.
 */
PLUGIN_API void ShutdownPlugin()
{
	DebugSpewAlways("MQTextToSpeech::Shutting down");
	if(SUCCEEDED(voice_com))
	{
		if (voice_ptr)
		{
			voice_ptr->Release();
			voice_ptr = nullptr;
		}

		if (current_voice_idx)
		{
			current_voice_idx->Release();
			current_voice_idx = nullptr;
		}

		RemoveCommand("/tts");

		RemoveMQ2Data("TTS");
		delete pTTSType;

		RemoveSettingsPanel("plugins/TextToSpeech");

		delete knightlog;
	}
}

/**
 * @fn InitializePlugin
 *
 * This is called once on plugin initialization and can be considered the startup
 * routine for the plugin.
 */
PLUGIN_API void InitializePlugin()
{
	DebugSpewAlways("MQTextToSpeech::Initializing version %f", MQ2Version);

	knightlog = new KnightLog("\a#7BACDD[\a#B93CF6%n\ax]\ax %^%L\ax \a#808080::\ax %^%v\ax");
	UNUSED(knightlog->SetColorByLevel(
		{
					{spdlog::level::trace,    "#FF00FF"},
					{spdlog::level::debug,    "#FF8C00"},
					{spdlog::level::info,     "#00E0E0"},
					{spdlog::level::warn,     "#FFD700"},
					{spdlog::level::err,      "#F22613"},
					{spdlog::level::critical, "#F22613"}
				}));

	if (SUCCEEDED(CoInitialize(nullptr)))
	{
		voice_com = CoCreateInstance(CLSID_SpVoice, nullptr, CLSCTX_ALL, IID_ISpVoice, reinterpret_cast<void**>(&voice_ptr));
		if(SUCCEEDED(voice_com))
		{
			RePopulateVoices();
			LoadSettings();

			AddCommand("/tts", commandTTS);

			AddMQ2Data("TTS", dataTTS);
			pTTSType = new MQTTSType;

			AddSettingsPanel("plugins/TextToSpeech", TTSImGuiSettingsPanel);
		}
		else
		{
			SPDLOG_CRITICAL("Could not initialize Voice COM");
			ShutdownPlugin();
		}
	}
	else
	{
		SPDLOG_CRITICAL("Could not initialize COM");
		ShutdownPlugin();
	}
}