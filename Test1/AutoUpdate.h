#pragma once

#include <windows.h>
#include <string>

// Check GitHub for a newer release.
// On success, fills newVersion (e.g. "0.2.0") and downloadUrl (asset URL).
// Returns true when a version newer than APP_VERSION is available.
bool CheckForUpdate(std::wstring& newVersion, std::wstring& downloadUrl);

// Download the installer from downloadUrl to a temp file and launch it.
// Returns true if the installer was started successfully.
bool DownloadAndInstallUpdate(const std::wstring& downloadUrl);

// Interactive flow: check → prompt user → download & install → exit app.
// Shows message boxes to communicate progress/results.
void CheckForUpdateInteractive(HWND hParent);
