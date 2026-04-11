#pragma once

#include <endpointvolume.h>

void Mixer_SetCallbackWindow(HWND hwndCallback);
MMRESULT Mixer_ToggleMute(void);
MMRESULT Mixer_SetVolume(int Increment);
MMRESULT Mixer_ToggleBassBoost(void);
MMRESULT Mixer_SetBass(int Increment);
MMRESULT Mixer_SetTreble(int Increment);

void Mixer_Shutdown();
void Mixer_DeviceChange(WPARAM wParam, LPARAM lParam);
void Mixer_ControlChange(WPARAM wParam, LPARAM lParam);
void Mixer_MMDeviceChange(void);

// default step size is 4% of max volume.
#define MIXER_DEFAULT_STEP        ((int)(65535/25))

#define MMHID_VOLUME_CONTROL    0
#define MMHID_BASS_CONTROL      1
#define MMHID_TREBLE_CONTROL    2
#define MMHID_BALANCE_CONTROL   3
#define MMHID_MUTE_CONTROL      4
#define MMHID_LOUDNESS_CONTROL  5
#define MMHID_BASSBOOST_CONTROL 6
#define MMHID_NUM_CONTROLS      7

class CSystemMixer
{
public:
    CSystemMixer(HWND hwndCallback);
    LONG Release();

    void ForwardWindowMessage(UINT uMsg, WPARAM wParam, LPARAM lParam);

    HRESULT ToggleMute();
    HRESULT AdjustVolume(int direction); // up/down
    MMRESULT ToggleBassBoost();
    MMRESULT AdjustBass(int direction);
    MMRESULT AdjustTreble(int increment);

private:
    void _RefreshMixCache(const DWORD* padwVolume);
    MMRESULT _GetVolume(DWORD* padwVolume);
    MMRESULT _GetDefaultMixerID(int* pid);
    BOOL _GetDestLine();
    void _GetLineControls();
    BOOL _Open();
    void _Close();
    void _Refresh();
    BOOL _CheckMissing();
    HRESULT _CreateVolumeObject(IAudioEndpointVolume** ppVolume);

private:
    LONG                _cRef;
    HMIXER              _hMixer;
    HWND                _hwndCallback;
    WCHAR*              _pszDeviceInterface;
    double*             _pdblCacheMix;
    DWORD*              _pdwLastVolume;
    MIXERLINEW          _mxlDst;
    DWORD               _rgdwControlType[MMHID_NUM_CONTROLS];
    BOOL                _rgfControlPresent[MMHID_NUM_CONTROLS];
    MIXERCONTROLW       _rgmxctrl[MMHID_NUM_CONTROLS];
    BOOL                _fMixerStartup;
    BOOL                _fMixerPresent;
    UINT                _uWinMM_DeviceChange;
    HDEVNOTIFY          _hdevnotify;
};
