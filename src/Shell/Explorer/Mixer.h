#pragma once

#include <endpointvolume.h>

class CSystemMixer
{
public:
    CSystemMixer(HWND hwndCallback);

    ULONG Release();

    void ForwardWindowMessage(UINT uMsg, WPARAM wParam, LPARAM lParam);

    HRESULT ToggleMute();
    MMRESULT ToggleBassBoost();
    HRESULT AdjustVolume(BOOL bUp);
    MMRESULT AdjustBass(BOOL bUp);
    MMRESULT AdjustTreble(BOOL bUp);

private:
    ~CSystemMixer();

    HRESULT _CreateVolumeObject(IAudioEndpointVolume** ppVolume);

    BOOL _Open();
    void _Close();
    void _Refresh();
    BOOL _CheckMissing();
    MMRESULT _GetVolume(DWORD* padwVolume);
    MMRESULT _GetDefaultMixerID(int* pid);
    void _RefreshMixCache(const DWORD* pdwVolume);
    BOOL _GetDestLine();
    void _GetLineControls();

    enum
    {
        MMHID_VOLUME_CONTROL        = 0,
        MMHID_BASS_CONTROL          = 1,
        MMHID_TREBLE_CONTROL        = 2,
        MMHID_BALANCE_CONTROL       = 3,
        MMHID_MUTE_CONTROL          = 4,
        MMHID_LOUDNESS_CONTROL      = 5,
        MMHID_BASSBOOST_CONTROL     = 6,
        MMHID_NUM_CONTROLS          = 7,
    };

    LONG                _cRef;
    HMIXER              _hMixer;
    HWND                _hwndCallback;
    WCHAR*              _pszDeviceInterface;
    double*             _pdblCacheMix;
    DWORD*              _pdwLastVolume;
    MIXERLINEW          _MixerLine;
    DWORD               _rgControlType[MMHID_NUM_CONTROLS];
    BOOL                _rgControlPresent[MMHID_NUM_CONTROLS];
    MIXERCONTROLW       _rgControl[MMHID_NUM_CONTROLS];
    BOOL                _fMixerStartup;
    BOOL                _fMixerPresent;
    UINT                _uMsgMMDeviceChanged;
    HDEVNOTIFY          _hdevnotify;
};
