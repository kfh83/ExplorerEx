#include "pch.h"
#include "runtask.h"
#include "debug.h"
#include "shundoc.h"

#define SUPERCLASS  

// #define TF_RUNTASK  TF_GENERAL
#define TF_RUNTASK  0
// #define TF_RUNTASKV TF_CUSTOM1     // verbose version
#define TF_RUNTASKV 0


// constructor
CRunnableTask::CRunnableTask(DWORD dwFlags)
{
    _lState = IRTIR_TASK_NOT_RUNNING;
    _dwFlags = dwFlags;

    _fAbort = 0;

	_hklKeyboard = GetKeyboardLayout(0);

#ifdef DEBUG
    _dwTaskID = GetTickCount();

    TraceMsg(TF_RUNTASK, "CRunnableTask (%#lx): creating task", _dwTaskID);
#endif

    _cRef = 1;
}


// destructor
CRunnableTask::~CRunnableTask()
{
    //DEBUG_CODE(TraceMsg(TF_RUNTASK, "CRunnableTask (%#lx): deleting task", _dwTaskID); )
}


STDMETHODIMP CRunnableTask::QueryInterface(REFIID riid, LPVOID *ppvObj)
{
    static const QITAB qit[] =
    {
        QITABENT(CRunnableTask, IRunnableTask),         // IID_IRunnableTask
        QITABENT(CRunnableTask, IQueryContinue),        // IID_IQueryContinue
        QITABENT(CRunnableTask, IServiceProvider),      // IID_IServiceProvider
        { 0 },
    };
    return QISearch(this, qit, riid, ppvObj);
}

STDMETHODIMP_(ULONG) CRunnableTask::AddRef()
{
    //_AssertMsgW(this->_cRef != 0, L"RefCount problem.");
    return InterlockedIncrement(&_cRef);
}


STDMETHODIMP_(ULONG) CRunnableTask::Release()
{
	ULONG cRef = InterlockedDecrement(&_cRef);
    if (0 == cRef)
    {
        delete this;
	}
	return cRef;
}


/*----------------------------------------------------------
Purpose: IRunnableTask::Run method

         This does a lot of the state-related work, and then
         calls the derived-class's RunRT() method.

*/
STDMETHODIMP CRunnableTask::Run(void)
{
#ifdef DEAD_CODE
    HRESULT hr = E_FAIL;

    // Are we already running?
    if (_lState == IRTIR_TASK_RUNNING)
    {
        // Yes; nothing to do 
        hr = S_FALSE;
    }
    else if (_lState == IRTIR_TASK_PENDING)
    {
        hr = E_FAIL;
    }
    else if (_lState == IRTIR_TASK_NOT_RUNNING)
    {
        // Say we're running
        LONG lRes = InterlockedExchange(&_lState, IRTIR_TASK_RUNNING);
        if (lRes == IRTIR_TASK_PENDING)
        {
            _lState = IRTIR_TASK_FINISHED;
            return NOERROR;
        }

        if (_lState == IRTIR_TASK_RUNNING)
        {
            // Prepare to run 
            //DEBUG_CODE(TraceMsg(TF_RUNTASKV, "CRunnableTask (%#lx): initialize to run", _dwTaskID); )

            hr = RunInitRT();

            ASSERT(E_PENDING != hr);
        }

        if (SUCCEEDED(hr))
        {
            if (_lState == IRTIR_TASK_RUNNING)
            {
                // Continue to do the work
                hr = InternalResumeRT();
            }
            else if (_lState == IRTIR_TASK_SUSPENDED)
            {
                // it is possible that RunInitRT took a little longer to complete and our state changed
                // from running to suspended with _hDone signaled, which would cause us to not call
                // internal resume.  We simulate internal resume here
                if (_hDone)
                    ResetEvent(_hDone);
                hr = E_PENDING;
            }
        }

        if (FAILED(hr) && E_PENDING != hr)
        {
            //DEBUG_CODE(TraceMsg(TF_WARNING, "CRunnableTask (%#lx): task failed to run: %#lx", _dwTaskID, hr); )
        }

        // Are we finished?
        if (_lState != IRTIR_TASK_SUSPENDED || hr != E_PENDING)
        {
            // Yes
            _lState = IRTIR_TASK_FINISHED;
        }
    }

    return hr;
#else
    HRESULT hr = E_FAIL;

    if (_lState == IRTIR_TASK_RUNNING)
        return S_FALSE;
    
    if (_lState == IRTIR_TASK_NOT_RUNNING)
    {
        LONG lRes = InterlockedExchange(&_lState, IRTIR_TASK_RUNNING);
        if (lRes == IRTIR_TASK_PENDING)
        {
            _lState = IRTIR_TASK_FINISHED;
            return 0;
        }

        if (_lState == IRTIR_TASK_RUNNING)
        {
            //CcshellDebugMsgW(0, "CRunnableTask (%#lx): initialize to run", _dwTaskID);
            HKL hklKeyboard = 0;
            if ((_dwFlags & 4) != 0)
            {
                hklKeyboard = GetKeyboardLayout(0);
                ActivateKeyboardLayout(_hklKeyboard, 0);
            }

            hr = RunInitRT();

            if ((_dwFlags & 4) != 0)
                ActivateKeyboardLayout(hklKeyboard, 0);

            ASSERT(E_PENDING != hr); // 118
        }

        if (hr >= 0)
        {
            if (_lState == IRTIR_TASK_RUNNING)
            {
                HKL hklKeyboard = 0;
                if ((_dwFlags & 4) != 0)
                {
                    hklKeyboard = GetKeyboardLayout(0);
                    ActivateKeyboardLayout(_hklKeyboard, 0);
                }

                hr = InternalResumeRT();

                if ((_dwFlags & 4) != 0)
                    ActivateKeyboardLayout(hklKeyboard, 0);
            }
            else if (_lState == IRTIR_TASK_SUSPENDED)
            {
                _fAbort = FALSE;
                hr = E_PENDING;
            }
        }

        //if (hr < 0 && hr != 0x8000000A)
        //    CcshellDebugMsgW(1, "CRunnableTask (%#lx): task failed to run: %#lx", _dwTaskID, hr);

        if (_lState != IRTIR_TASK_SUSPENDED || hr != E_PENDING)
        {
            _lState = IRTIR_TASK_FINISHED;
        }
    }
    return hr;
#endif
}


/*----------------------------------------------------------
Purpose: IRunnableTask::Kill method

*/
STDMETHODIMP CRunnableTask::Kill(BOOL fWait)
{
#ifdef DEAD_CODE
    if (!(_dwFlags & RTF_SUPPORTKILLSUSPEND))
        return E_NOTIMPL;

    if (_lState != IRTIR_TASK_RUNNING)
        return S_FALSE;

    //DEBUG_CODE(TraceMsg(TF_RUNTASKV, "CRunnableTask (%#lx): killing task", _dwTaskID); )

    LONG lRes = InterlockedExchange(&_lState, IRTIR_TASK_PENDING);
    if (lRes == IRTIR_TASK_FINISHED)
    {
        //DEBUG_CODE(TraceMsg(TF_RUNTASKV, "CRunnableTask (%#lx): task already finished", _dwTaskID); )

        _lState = lRes;
    }
    else
    {
		_fAbort = TRUE;
    }

    return KillRT(fWait);
#else
    if ((this->_dwFlags & 2) == 0)
        return E_NOTIMPL;

    HRESULT hr = S_FALSE;

    if (_lState == 1)
    {
        //CcshellDebugMsgW(0, "CRunnableTask (%#lx): killing task", this->_dwTaskID);
        if (InterlockedExchange(&_lState, 3) == 4)
        {
            //CcshellDebugMsgW(0, "CRunnableTask (%#lx): task already finished", this->_dwTaskID);
            _lState = 4;
        }
        else
        {
            _fAbort = 1;
        }
        return KillRT(fWait);
    }
    return hr;
#endif
}


/*----------------------------------------------------------
Purpose: IRunnableTask::Suspend method

*/
STDMETHODIMP CRunnableTask::Suspend(void)
{
#ifdef DEAD_CODE
    if (!(_dwFlags & RTF_SUPPORTKILLSUSPEND))
        return E_NOTIMPL;

    if (_lState != IRTIR_TASK_RUNNING)
        return E_FAIL;

    //DEBUG_CODE(TraceMsg(TF_RUNTASKV, "CRunnableTask (%#lx): suspending task", _dwTaskID); )

    LONG lRes = InterlockedExchange(&_lState, IRTIR_TASK_SUSPENDED);

    if (IRTIR_TASK_FINISHED == lRes)
    {
        // we finished before we could suspend
        //DEBUG_CODE(TraceMsg(TF_RUNTASKV, "CRunnableTask (%#lx): task already finished", _dwTaskID); )

        _lState = lRes;
        return S_OK;
    }
    else
    {
        _fAbort = TRUE;
        return SuspendRT();
    }
#else
    if ((this->_dwFlags & 1) == 0)
        return 0x80004001;
    if (this->_lState != 1)
        return 0x80004005;

    //CcshellDebugMsgW(0, "CRunnableTask (%#lx): suspending task", this->_dwTaskID);
    if (InterlockedExchange(&this->_lState, 2) == 4)
    {
        //CcshellDebugMsgW(0, "CRunnableTask (%#lx): task already finished", this->_dwTaskID);
        this->_lState = 4;
        return 0;
    }
    else
    {
        this->_fAbort = 1;
        return SuspendRT();
    }
#endif
}


/*----------------------------------------------------------
Purpose: IRunnableTask::Resume method

*/
STDMETHODIMP CRunnableTask::Resume(void)
{
    if (_lState != IRTIR_TASK_SUSPENDED)
        return E_FAIL;

    //DEBUG_CODE(TraceMsg(TF_RUNTASKV, "CRunnableTask (%#lx): resuming task", _dwTaskID); )

    _lState = IRTIR_TASK_RUNNING;

	_fAbort = FALSE;

    return ResumeRT();
}


/*----------------------------------------------------------
Purpose: IRunnableTask::IsRunning method

*/
STDMETHODIMP_(ULONG) CRunnableTask::IsRunning(void)
{
    return _lState;
}

/*----------------------------------------------------------
Purpose: IServiceProvider::QueryService method
*/
STDMETHODIMP CRunnableTask::QueryService(REFGUID guidService, REFIID riid, void **ppvObject)
{
    *ppvObject = NULL;

    HRESULT hr = E_FAIL;
    if (IsEqualGUID(guidService, IID_IQueryContinue))
    {
        return QueryInterface(riid, ppvObject);
    }
    return hr;
}

/*----------------------------------------------------------
Purpose: IQueryContinue::QueryContinue method
*/
STDMETHODIMP CRunnableTask::QueryContinue(void)
{
    return ShouldContinue();
}