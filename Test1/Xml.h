#pragma once

#include <windows.h>
#include <string>
#include <vector>

bool SaveSelectedDeviceId(UINT deviceId);
bool LoadSelectedDeviceId(UINT& deviceId);
bool SaveSelectedInput2DeviceId(UINT deviceId);
bool LoadSelectedInput2DeviceId(UINT& deviceId);
bool SaveSelectedOutputDeviceId(UINT deviceId);
bool LoadSelectedOutputDeviceId(UINT& deviceId);
bool SaveMidiRouterEnabled(bool enabled);
bool LoadMidiRouterEnabled(bool& enabled);
bool SaveCloseSettingsOnDisconnect(bool enabled);
bool LoadCloseSettingsOnDisconnect(bool& enabled);
bool SaveShowDebugConsole(bool enabled);
bool LoadShowDebugConsole(bool& enabled);
void RefreshSettingsFile();

// Force a re-read of the standby organ names from the Hauptwerk config.
// Call this when Hauptwerk is launched or detected running so the cache
// is refreshed with up-to-date data.
void ReloadStandbyOrgans();

// Returns the standby organ names (1..8) read from Settings.xml.
// The vector always has 8 entries; unused slots are empty strings.
std::vector<std::wstring> LoadStandbyOrganNames();

// Detect / configure Hauptwerk installation folders.
// On first run, checks the default install path or shows a folder picker.
// Reads FileLocations to extract UserData, SampleSets and WorkingFiles paths.
// Returns true if paths were successfully configured.
bool InitHauptwerkPaths();

// Load the saved Hauptwerk application root folder from Settings.xml.
bool LoadHauptwerkAppPath(std::wstring& path);
