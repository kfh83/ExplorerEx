#include "pch.h"

#include "mixer.h"
#include <mmdeviceapi.h>

CSystemMixer::CSystemMixer(HWND hwndCallback)
    : _cRef(1)
    , _hMixer(nullptr)
    , _hwndCallback(hwndCallback)
    , _pszDeviceInterface(nullptr)
    , _pdblCacheMix(nullptr)
    , _pdwLastVolume(nullptr)
    , _fMixerStartup(TRUE)
    , _fMixerPresent(FALSE)
{
    _rgControlType[MMHID_VOLUME_CONTROL]      = MIXERCONTROL_CONTROLTYPE_VOLUME;
    _rgControlType[MMHID_BASS_CONTROL]        = MIXERCONTROL_CONTROLTYPE_BASS;
    _rgControlType[MMHID_TREBLE_CONTROL]      = MIXERCONTROL_CONTROLTYPE_TREBLE;
    _rgControlType[MMHID_BALANCE_CONTROL]     = MIXERCONTROL_CONTROLTYPE_PAN;
    _rgControlType[MMHID_MUTE_CONTROL]        = MIXERCONTROL_CONTROLTYPE_MUTE;
    _rgControlType[MMHID_LOUDNESS_CONTROL]    = MIXERCONTROL_CONTROLTYPE_LOUDNESS;
    _rgControlType[MMHID_BASSBOOST_CONTROL]   = MIXERCONTROL_CONTROLTYPE_BASS_BOOST;

    _rgControlPresent[MMHID_VOLUME_CONTROL]    = FALSE;
    _rgControlPresent[MMHID_BASS_CONTROL]      = FALSE;
    _rgControlPresent[MMHID_TREBLE_CONTROL]    = FALSE;
    _rgControlPresent[MMHID_BALANCE_CONTROL]   = FALSE;
    _rgControlPresent[MMHID_MUTE_CONTROL]      = FALSE;
    _rgControlPresent[MMHID_LOUDNESS_CONTROL]  = FALSE;
    _rgControlPresent[MMHID_BASSBOOST_CONTROL] = FALSE;

    _uMsgMMDeviceChanged = RegisterWindowMessageW(L"winmm_devicechange");
}

ULONG CSystemMixer::Release()
{
    _ASSERTE(_cRef != 0); // 96
    LONG cRef = InterlockedDecrement(&_cRef);
    if (cRef == 0)
    {
        delete this;
    }
    return cRef;
}

void CSystemMixer::ForwardWindowMessage(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    DWORD* pdwVolume;
    DWORD* pdwLastVolume;

    DEV_BROADCAST_DEVICEINTERFACE* dbdi = reinterpret_cast<DEV_BROADCAST_DEVICEINTERFACE*>(lParam);

    if (uMsg == 537)
    {
        if (!_pszDeviceInterface)
        {
            return;
        }
        if (wParam != 32769)
        {
            if (wParam == 32770)
            {
                if (dbdi->dbcc_devicetype == 5 && !lstrcmpiW(dbdi->dbcc_name, _pszDeviceInterface))
                {
                    goto LABEL_23;
                }
                return;
            }
            if (wParam != 32771)
                return;
        }
        if (dbdi->dbcc_devicetype == 5 && !lstrcmpiW(dbdi->dbcc_name, _pszDeviceInterface))
        {
            _Close();
        }
        return;
    }

    if (uMsg == 977)
    {
        if (_hMixer == (HMIXER)wParam && lParam == _rgControl[0].dwControlID)
        {
            pdwVolume = (DWORD*)LocalAlloc(LPTR, _MixerLine.cChannels * sizeof(DWORD));
            if (pdwVolume)
            {
                if (!_GetVolume(pdwVolume))
                {
                    pdwLastVolume = _pdwLastVolume;
                    if (!pdwLastVolume || memcmp(pdwLastVolume, pdwVolume, _MixerLine.cChannels * sizeof(DWORD)))
                    {
                        _RefreshMixCache(pdwVolume);
                    }
                }
                LocalFree(pdwVolume);
            }
        }
    }
    else if (uMsg == _uMsgMMDeviceChanged)
    {
    LABEL_23:
        _Refresh();
    }
}

HRESULT CSystemMixer::ToggleMute()
{
    IAudioEndpointVolume* pVolume = nullptr;
    HRESULT hr = _CreateVolumeObject(&pVolume);
    if (SUCCEEDED(hr))
    {
        _ASSERTE(nullptr != pVolume); // 214

        BOOL fMute;
        hr = pVolume->GetMute(&fMute);
        if (SUCCEEDED(hr))
        {
            hr = pVolume->SetMute(!fMute, nullptr);
        }
        pVolume->Release();
    }
    return hr;
}

MMRESULT CSystemMixer::ToggleBassBoost()
{
    DWORD fEnabled = 0;
    if (_CheckMissing())
        return MMSYSERR_NODRIVER;

    if (!_rgControlPresent[6])
        return MMSYSERR_NOERROR;

    MIXERCONTROLDETAILS mxcd;
    mxcd.cMultipleItems = 0;
    mxcd.dwControlID = _rgControl[6].dwControlID;
    mxcd.paDetails = &fEnabled;
    mxcd.cbStruct = 24;
    mxcd.cChannels = 1;
    mxcd.cbDetails = 4;
    MMRESULT mmr = mixerGetControlDetailsW((HMIXEROBJ)_hMixer, &mxcd, 0x80000000);
    if (mmr == 0)
    {
        fEnabled = fEnabled == 0;
        mmr = mixerSetControlDetails((HMIXEROBJ)_hMixer, &mxcd, 0x80000000);
    }
    return mmr;
}

HRESULT CSystemMixer::AdjustVolume(BOOL bUp)
{
    IAudioEndpointVolume* pVolume = nullptr;
    HRESULT hr = _CreateVolumeObject(&pVolume);
    if (SUCCEEDED(hr))
    {
        _ASSERTE(nullptr != pVolume); // 301

        if (bUp)
        {
            hr = pVolume->VolumeStepUp(nullptr);
        }
        else
        {
            hr = pVolume->VolumeStepDown(nullptr);
        }
        pVolume->Release();
    }
    return hr;
}

MMRESULT CSystemMixer::AdjustBass(BOOL bUp)
{
    LONG lLevel;

    if (_CheckMissing())
        return MMSYSERR_NODRIVER;

    MMRESULT mmr = MMSYSERR_NOERROR;

    if (_rgControlPresent[1])
    {
        lLevel = 0;

        MIXERCONTROLDETAILS mxcd;
        mxcd.cMultipleItems = 0;
        mxcd.paDetails = &lLevel;
        mxcd.cbStruct = 0x18;
        mxcd.dwControlID = _rgControl[1].dwControlID;
        mxcd.cChannels = 1;
        mxcd.cbDetails = 4;
        mmr = mixerGetControlDetailsW((HMIXEROBJ)_hMixer, &mxcd, 0x80000000);
        if (!mmr)
        {
            lLevel += bUp != 0 ? 2621 : -2621;

            lLevel = std::min<LONG>(65535, lLevel); // @MOD Don't use macro
            lLevel = std::max<LONG>(0, lLevel); // @MOD Don't use macro

            mxcd.paDetails = &lLevel;
            mxcd.cChannels = 1;
            mxcd.cMultipleItems = 0;
            mxcd.cbDetails = 4;
            mmr = mixerSetControlDetails((HMIXEROBJ)_hMixer, &mxcd, 0x80000000);
        }
    }
    return mmr;
}

MMRESULT CSystemMixer::AdjustTreble(BOOL bUp)
{
    if (_CheckMissing())
        return MMSYSERR_NODRIVER;

    MMRESULT mmr = MMSYSERR_NOERROR;

    LONG lLevel = 0;

    if (_rgControlPresent[MMHID_TREBLE_CONTROL])
    {
        MIXERCONTROLDETAILS mxcd;
        mxcd.cbStruct = sizeof(mxcd);
        mxcd.dwControlID = _rgControl[MMHID_TREBLE_CONTROL].dwControlID;

        mxcd.cChannels = 1;
        mxcd.cMultipleItems = 0;
        mxcd.cbDetails = sizeof(lLevel);
        mxcd.paDetails = static_cast<void*>(&lLevel);
        mmr = mixerGetControlDetailsW(reinterpret_cast<HMIXEROBJ>(_hMixer), &mxcd, 0x80000000);
        if (mmr == MMSYSERR_NOERROR)
        {
            lLevel += bUp != 0 ? 2621 : -2621;
            lLevel = std::min<LONG>(65535, lLevel); // @MOD Don't use macro
            lLevel = std::max<LONG>(0, lLevel);     // @MOD Don't use macro

            mxcd.cChannels = 1;
            mxcd.cMultipleItems = 0;
            mxcd.cbDetails = sizeof(lLevel);
            mxcd.paDetails = static_cast<void*>(&lLevel);
            mmr = mixerSetControlDetails(reinterpret_cast<HMIXEROBJ>(_hMixer), &mxcd, 0x80000000);
        }
    }

    return mmr;
}

CSystemMixer::~CSystemMixer()
{
    _Close();
}

HRESULT CSystemMixer::_CreateVolumeObject(IAudioEndpointVolume** ppVolume)
{
    ASSERT(nullptr != ppVolume); // 173
    *ppVolume = nullptr;

    IMMDeviceEnumerator* pEnum = nullptr;
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pEnum));
    if (SUCCEEDED(hr))
    {
        IMMDevice* pDevice = nullptr;
        hr = pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
        if (SUCCEEDED(hr))
        {
            IAudioEndpointVolume* pVolume = nullptr;
            hr = pDevice->Activate(
                __uuidof(IAudioEndpointVolume),
                CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER,
                nullptr, reinterpret_cast<void**>(&pVolume));
            if (SUCCEEDED(hr))
            {
                *ppVolume = pVolume;
                (*ppVolume)->AddRef();
                pVolume->Release();
            }
            pDevice->Release();
        }
        pEnum->Release();
    }

    return hr;
}

BOOL CSystemMixer::_Open()
{
    DEV_BROADCAST_HANDLE dbh;

    DWORD_PTR cbDeviceInterface = 0;

    _ASSERTE((nullptr == _hMixer)
        && (nullptr == _pdblCacheMix)
        && (nullptr == _pdblCacheMix)
        && (nullptr == _pszDeviceInterface)); // 481

    int mixerID;
    if (!_GetDefaultMixerID(&mixerID))
    {
        if (!mixerOpen(&_hMixer, mixerID, (DWORD_PTR)_hwndCallback, 0, 0x10000u))
        {
            HWND hwndCallback = _hwndCallback;
            if (hwndCallback)
            {
                memset(&dbh.dbch_devicetype, 0, 0x28u);
                dbh.dbch_size = 0x2C;
                dbh.dbch_devicetype = DBT_DEVTYP_HANDLE;
                dbh.dbch_handle = *(HANDLE*)&_hMixer;
                _hdevnotify = RegisterDeviceNotificationW(hwndCallback, &dbh, DEVICE_NOTIFY_WINDOW_HANDLE);
            }

            if (_GetDestLine())
            {
                _GetLineControls();
                cbDeviceInterface = 0;
                if (!mixerMessage((HMIXER)mixerID, 0x80Du, (DWORD_PTR)&cbDeviceInterface, 0) && cbDeviceInterface)
                {
                    WCHAR* pwstrDeviceInterface = (WCHAR*)LocalAlloc(0x40u, cbDeviceInterface);
                    if (pwstrDeviceInterface)
                    {
                        if (mixerMessage((HMIXER)mixerID, 0x80Cu, (DWORD_PTR)pwstrDeviceInterface, cbDeviceInterface))
                        {
                            LocalFree(pwstrDeviceInterface);
                        }
                        else
                        {
                            _pszDeviceInterface = pwstrDeviceInterface;
                        }
                    }
                }
                return 1;
            }
            else
            {
                mixerClose((HMIXER)*(HANDLE*)&_hMixer);
                *(HANDLE*)&_hMixer = nullptr;
            }
        }
    }

    return cbDeviceInterface;
}

void CSystemMixer::_Close()
{
    LocalFree(_pszDeviceInterface);
    _pszDeviceInterface = nullptr;

    LocalFree(_pdblCacheMix);
    _pdblCacheMix = nullptr;

    LocalFree(_pdwLastVolume);
    _pdwLastVolume = nullptr;

    if (_hMixer)
    {
        _ASSERTE(MMSYSERR_NOERROR == mixerClose(_hMixer)); // 552
        _hMixer = nullptr;
    }
    if (_hdevnotify)
    {
        UnregisterDeviceNotification(_hdevnotify);
        _hdevnotify = nullptr;
    }
}

void CSystemMixer::_Refresh()
{
    _Close();
    _fMixerPresent = _Open();
}

BOOL CSystemMixer::_CheckMissing()
{
    if (_fMixerStartup)
    {
        _fMixerStartup = FALSE;
        _Refresh();
    }
    return !_fMixerPresent;
}

MMRESULT CSystemMixer::_GetVolume(DWORD* padwVolume)
{
    if (!_rgControlPresent[MMHID_VOLUME_CONTROL])
        return MIXERR_INVALCONTROL;

    MIXERCONTROLDETAILS mxcd;
    mxcd.cbStruct = sizeof(mxcd);
    mxcd.dwControlID = _rgControl[MMHID_VOLUME_CONTROL].dwControlID;
    mxcd.cChannels = _MixerLine.cChannels;
    mxcd.cMultipleItems = 0;
    mxcd.cbDetails = sizeof(DWORD);
    mxcd.paDetails = static_cast<void*>(padwVolume);
    return mixerGetControlDetailsW(
        reinterpret_cast<HMIXEROBJ>(_hMixer), &mxcd, MIXER_OBJECTF_HANDLE | MIXER_GETCONTROLDETAILSF_VALUE);
}

MMRESULT CSystemMixer::_GetDefaultMixerID(int* pid)
{
    MMRESULT mmr = MMSYSERR_NODRIVER;

    if (waveOutGetNumDevs() != 0)
    {
        UINT uWaveID;
        DWORD dwFlags;
        mmr = waveOutMessage((HWAVEOUT)WAVE_MAPPER, 0x2015, (DWORD_PTR)&uWaveID, (DWORD_PTR)&dwFlags);
        if (mmr == MMSYSERR_NOERROR)
        {
            if (uWaveID != -1)
            {
                UINT uMxID;
                mmr = mixerGetID((HMIXEROBJ)uWaveID, &uMxID, 0x10000000);
                if (mmr == MMSYSERR_NOERROR)
                {
                    *pid = uMxID;
                }
            }
            else
            {
                mmr = MMSYSERR_NODRIVER;
            }
        }
    }

    return mmr;
}

void CSystemMixer::_RefreshMixCache(const DWORD* pdwVolume)
{
    DWORD dwMaxVol = 0;

    if (pdwVolume && _MixerLine.cChannels)
    {
        if (!_pdblCacheMix)
        {
            _pdblCacheMix = static_cast<double*>(LocalAlloc(LPTR, _MixerLine.cChannels * sizeof(double)));
        }
        if (_pdblCacheMix)
        {
            for (UINT uiIndx = 0; uiIndx < _MixerLine.cChannels; ++uiIndx)
            {
                dwMaxVol = std::max<DWORD>(dwMaxVol, pdwVolume[uiIndx]); // @MOD Don't use macro
            }

            for (UINT uiIndx = 0; uiIndx < _MixerLine.cChannels; ++uiIndx)
            {
                DWORD dwVolume = pdwVolume[uiIndx];

                if (dwMaxVol == dwVolume)
                {
                    _pdblCacheMix[uiIndx] = 1.0;
                }
                else
                {
                    _pdblCacheMix[uiIndx] = static_cast<double>(dwVolume) / static_cast<double>(dwMaxVol);
                }
            }
        }
    }
}

BOOL CSystemMixer::_GetDestLine()
{
    _MixerLine.cbStruct = sizeof(_MixerLine);
    _MixerLine.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_SPEAKERS;
    return mixerGetLineInfoW(
        reinterpret_cast<HMIXEROBJ>(_hMixer), &_MixerLine, MIXER_GETLINEINFOF_COMPONENTTYPE | MIXER_OBJECTF_HANDLE) == 0;
}

void CSystemMixer::_GetLineControls()
{
    MIXERLINECONTROLSW mxlc;

    for (int i = 0; i < ARRAYSIZE(_rgControl); ++i)
    {
        mxlc.cbStruct = sizeof(mxlc);
        mxlc.dwLineID = _MixerLine.dwLineID;
        mxlc.dwControlID = _rgControlType[i];
        mxlc.cControls = 1;
        mxlc.cbmxctrl = sizeof(MIXERCONTROLW);
        mxlc.pamxctrl = &_rgControl[i];

        _rgControlPresent[i] = mixerGetLineControlsW(reinterpret_cast<HMIXEROBJ>(_hMixer), &mxlc, 0x80000002) == 0;
    }
}
