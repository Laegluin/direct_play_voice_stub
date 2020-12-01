#![allow(non_snake_case)]

use com::interfaces::IUnknown;
use com::production::Class;
use com::sys::{
    CLASS_E_CLASSNOTAVAILABLE, CLSID, E_INVALIDARG, GUID, HRESULT, IID, SELFREG_E_CLASS, S_OK,
};
use std::ffi::OsString;
use std::io;
use std::mem;
use std::os::windows::ffi::OsStringExt;
use std::ptr;
use std::sync::Mutex;
use winapi::ctypes::c_void;
use winapi::shared::minwindef::{BOOL, DWORD, HINSTANCE, PDWORD, TRUE};
use winapi::shared::ntdef::{LONG, LPWSTR, PVOID};
use winapi::shared::windef::HWND;
use winapi::shared::winerror::ERROR_SUCCESS;
use winapi::um::dsound::{LPDIRECTSOUND, LPDIRECTSOUNDBUFFER};
use winapi::um::errhandlingapi::{GetLastError, SetLastError};
use winapi::um::libloaderapi::GetModuleFileNameW;
use winapi::um::unknwnbase::LPUNKNOWN;
use winreg::enums::HKEY_LOCAL_MACHINE;
use winreg::RegKey;

// {B9F3EB85-B781-4AC1-8D90-93A05EE37D7D}
const DIRECT_PLAY_VOICE_CLIENT_CLSID: CLSID = CLSID {
    data1: 0xB9F3EB85,
    data2: 0xB781,
    data3: 0x4AC1,
    data4: [0x8D, 0x90, 0x93, 0xA0, 0x5E, 0xE3, 0x7D, 0x7D],
};

// {D3F5B8E6-9B78-4A4C-94EA-CA2397B663D3}
const DIRECT_PLAY_VOICE_SERVER_CLSID: CLSID = CLSID {
    data1: 0xD3F5B8E6,
    data2: 0x9B78,
    data3: 0x4A4C,
    data4: [0x94, 0xEA, 0xCA, 0x23, 0x97, 0xB6, 0x63, 0xD3],
};

// {0F0F094B-B01C-4091-A14D-DD0CD807711A}
const DIRECT_PLAY_VOICE_TEST_CLSID: CLSID = CLSID {
    data1: 0x0F0F094B,
    data2: 0xB01C,
    data3: 0x4091,
    data4: [0xA1, 0x4D, 0xDD, 0x0C, 0xD8, 0x07, 0x71, 0x1A],
};

static mut HMODULE: HINSTANCE = ptr::null_mut();

#[no_mangle]
unsafe extern "stdcall" fn DllMain(
    hinstance: HINSTANCE,
    fdw_reason: u32,
    _reserved: *mut c_void,
) -> BOOL {
    const DLL_PROCESS_ATTACH: u32 = 1;

    if fdw_reason == DLL_PROCESS_ATTACH {
        HMODULE = hinstance;
    }

    TRUE
}

#[no_mangle]
unsafe extern "stdcall" fn DllGetClassObject(
    class_id: *const CLSID,
    iid: *const IID,
    result: *mut *mut c_void,
) -> HRESULT {
    assert!(!class_id.is_null());

    if *class_id == DIRECT_PLAY_VOICE_CLIENT_CLSID {
        let instance = <DirectPlayVoiceClient as Class>::Factory::allocate();
        instance.QueryInterface(&*iid, result)
    } else if *class_id == DIRECT_PLAY_VOICE_SERVER_CLSID {
        let instance = <DirectPlayVoiceServer as Class>::Factory::allocate();
        instance.QueryInterface(&*iid, result)
    } else if *class_id == DIRECT_PLAY_VOICE_TEST_CLSID {
        let instance = <DirectPlayVoiceTest as Class>::Factory::allocate();
        instance.QueryInterface(&*iid, result)
    } else {
        CLASS_E_CLASSNOTAVAILABLE
    }
}

#[no_mangle]
extern "stdcall" fn DllRegisterServer() -> HRESULT {
    fn register() -> Result<(), io::Error> {
        let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);
        let dll_path = get_dll_path()?;

        let register_class = |name: &str, class_id: &CLSID| -> Result<(), io::Error> {
            let (key, _) =
                hklm.create_subkey(format!("Software\\Classes\\CLSID\\{{{}}}", class_id))?;
            key.set_value("", &name)?;

            let (in_proc, _) = key.create_subkey("InProcServer32")?;
            in_proc.set_value("", &dll_path.as_os_str())?;
            in_proc.set_value("ThreadingModel", &"Both")?;

            Ok(())
        };

        register_class("DirectPlayVoiceClient", &DIRECT_PLAY_VOICE_CLIENT_CLSID)?;
        register_class("DirectPlayVoiceServer", &DIRECT_PLAY_VOICE_SERVER_CLSID)?;
        register_class("DirectPlayVoiceTest", &DIRECT_PLAY_VOICE_TEST_CLSID)?;

        Ok(())
    }

    match register() {
        Ok(_) => S_OK,
        Err(ref why) => {
            let _ = DllUnregisterServer();
            eprintln!("error: {}", why);
            SELFREG_E_CLASS
        }
    }
}

#[no_mangle]
extern "stdcall" fn DllUnregisterServer() -> HRESULT {
    fn unregister() -> Result<(), io::Error> {
        let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);

        let unregister_class = |class_id: &CLSID| -> Result<(), io::Error> {
            match hklm.delete_subkey_all(format!("Software\\Classes\\CLSID\\{{{}}}", class_id)) {
                Ok(_) => Ok(()),
                Err(ref err) if err.kind() == io::ErrorKind::NotFound => Ok(()),
                err => err,
            }
        };

        unregister_class(&DIRECT_PLAY_VOICE_CLIENT_CLSID)?;
        unregister_class(&DIRECT_PLAY_VOICE_SERVER_CLSID)?;
        unregister_class(&DIRECT_PLAY_VOICE_TEST_CLSID)?;

        Ok(())
    }

    match unregister() {
        Ok(_) => S_OK,
        Err(ref why) => {
            eprintln!("error: {}", why);
            SELFREG_E_CLASS
        }
    }
}

fn get_dll_path() -> Result<OsString, io::Error> {
    unsafe {
        assert!(!HMODULE.is_null());

        let mut path = [0; 1024];
        SetLastError(ERROR_SUCCESS);
        let len = GetModuleFileNameW(HMODULE, path.as_mut_ptr(), path.len() as u32);

        let err = GetLastError();
        if err != ERROR_SUCCESS {
            return Err(io::Error::from_raw_os_error(err as i32));
        }

        Ok(OsString::from_wide(&path[..len as usize]))
    }
}

#[repr(C)]
#[derive(Clone)]
pub struct DVSOUNDDEVICECONFIG {
    dwSize: DWORD,
    dwFlags: DWORD,
    guidPlaybackDevice: GUID,
    lpdsPlaybackDevice: LPDIRECTSOUND,
    guidCaptureDevice: GUID,
    lpdsCaptureDevice: *mut c_void, // LPDIRECTSOUNDCAPTURE
    hwndAppWindow: HWND,
    lpdsMainBuffer: LPDIRECTSOUNDBUFFER,
    dwMainBufferFlags: DWORD,
    dwMainBufferPriority: DWORD,
}

#[repr(C)]
#[derive(Clone)]
pub struct DVCLIENTCONFIG {
    dwSize: DWORD,
    dwFlags: DWORD,
    lRecordVolume: LONG,
    lPlaybackVolume: LONG,
    dwThreshold: DWORD,
    dwBufferQuality: DWORD,
    dwBufferAggressiveness: DWORD,
    dwNotifyPeriod: DWORD,
}

#[repr(C)]
#[derive(Clone)]
pub struct DVSESSIONDESC {
    dwSize: DWORD,
    dwFlags: DWORD,
    dwSessionType: DWORD,
    guidCT: GUID,
    dwBufferQuality: DWORD,
    dwBufferAggressiveness: DWORD,
}

const DVSESSION_NOHOSTMIGRATION: DWORD = 0x00000001;
const DVSESSIONTYPE_PEER: DWORD = 0x00000001;
const DVBUFFERQUALITY_DEFAULT: DWORD = 0x00000000;
const DVBUFFERAGGRESSIVENESS_DEFAULT: DWORD = 0x00000000;

#[repr(C)]
#[derive(Clone)]
pub struct DVCOMPRESSIONINFO {
    dwSize: DWORD,
    guidType: GUID,
    lpszName: LPWSTR,
    lpszDescription: LPWSTR,
    dwFlags: DWORD,
    dwMaxBitsPerSecond: DWORD,
}

// MS-PCM 64 kbit/s (null-terminated, UTF-16)
const PCM_NAME: [u16; 17] = [
    0x004d, 0x0053, 0x002d, 0x0050, 0x0043, 0x004d, 0x0020, 0x0036, 0x0034, 0x0020, 0x006b, 0x0062,
    0x0069, 0x0074, 0x002f, 0x0073, 0x0000,
];

const PCM: DVCOMPRESSIONINFO = DVCOMPRESSIONINFO {
    dwSize: mem::size_of::<DVCOMPRESSIONINFO>() as u32 + PCM_NAME.len() as u32 * 2,
    guidType: GUID {
        data1: 0x8de12fd4,
        data2: 0x7cb3,
        data3: 0x48ce,
        data4: [0xa7, 0xe8, 0x9c, 0x47, 0xa2, 0x2e, 0x8a, 0xc5],
    },
    lpszName: ptr::null_mut(),
    lpszDescription: ptr::null_mut(),
    dwFlags: 0,
    dwMaxBitsPerSecond: 64000,
};

#[repr(C)]
pub struct DVCAPS {
    dwSize: DWORD,
    dwFlags: DWORD,
}

type DVID = DWORD;
type PDVID = *mut DWORD;
type IDirectSound3DBuffer = *mut c_void;
type LPDIRECTSOUND3DBUFFER = *mut IDirectSound3DBuffer;

const DVFLAGS_QUERYONLY: DWORD = 0x00000002;
const FACDPV: u32 = 0x15;

const fn make_hresult(sev: u32, fac: u32, code: u32) -> HRESULT {
    ((sev << 31) | (fac << 16) | code) as i32
}

const fn make_dv_hresult(code: u32) -> HRESULT {
    make_hresult(1, FACDPV, code)
}

const DV_OK: HRESULT = S_OK;
const DV_FULLDUPLEX: HRESULT = make_hresult(0, FACDPV, 0x0005);
const DVERR_INVALIDPARAM: HRESULT = E_INVALIDARG;
const DVERR_NOTHOSTING: HRESULT = make_dv_hresult(0x017B);
const DVERR_BUFFERTOOSMALL: HRESULT = make_dv_hresult(0x001E);
const DVERR_NOTCONNECTED: HRESULT = make_dv_hresult(0x016B);
const DVERR_ALREADYBUFFERED: HRESULT = make_dv_hresult(0x0178);

com::interfaces! {
    #[uuid("1DFDC8EA-BCF7-41D6-B295-AB64B3B23306")]
    pub unsafe interface IDirectPlayVoiceClient: IUnknown {
        fn Initialize(&self, pVoid: LPUNKNOWN, pMessageHandler: *mut c_void, pUserContext: PVOID, pdwMessageMask: PDWORD, dwMessageMaskElements: DWORD) -> HRESULT;

        fn Connect(&self, pSoundDeviceConfig: *mut DVSOUNDDEVICECONFIG, pdvClientConfig: *mut DVCLIENTCONFIG, dwFlags: DWORD) -> HRESULT;

        fn Disconnect(&self, dwFlags: DWORD) -> HRESULT;

        fn GetSessionDesc(&self, pvSessionDesc: *mut DVSESSIONDESC) -> HRESULT;

        fn GetClientConfig(&self, pClientConfig: *mut DVCLIENTCONFIG) -> HRESULT;

        fn SetClientConfig(&self, pClientConfig: *mut DVCLIENTCONFIG) -> HRESULT;

        fn GetCaps(&self, pDVCaps: *mut DVCAPS) -> HRESULT;

        fn GetCompressionTypes(&self, pData: PVOID, pdwDataSize: PDWORD, pdwNumElements: PDWORD, dwFlags: DWORD) -> HRESULT;

        fn SetTransmitTargets(&self, pdvIDTargets: PDVID, dwNumTargets: DWORD, dwFlags: DWORD) -> HRESULT;

        fn GetTransmitTargets(&self, pdvIDTargets: PDVID, dwNumTargets: PDWORD, dwFlags: DWORD) -> HRESULT;

        fn Create3DSoundBuffer(&self, dvID: DVID, lpdsSourceBuffer: LPDIRECTSOUNDBUFFER, dwPriority: DWORD, dwFlags: DWORD, lpUserBuffer: LPDIRECTSOUND3DBUFFER) -> HRESULT;

        fn Delete3DSoundBuffer(&self, dvID: DVID, lpUserBuffer: LPDIRECTSOUND3DBUFFER) -> HRESULT;

        fn SetNotifyMask(&self, pdwMessageMask: PDWORD, dwMessageMaskElements: DWORD) -> HRESULT;

        fn GetSoundDeviceConfig(&self, pSoundDeviceConfig: *mut DVSOUNDDEVICECONFIG, pdwSize: PDWORD) -> HRESULT;
    }
}

com::interfaces! {
    #[uuid("FAA1C173-0468-43b6-8A2A-EA8A4F2076C9")]
    pub unsafe interface IDirectPlayVoiceServer: IUnknown {
        fn Initialize(&self, pVoid: LPUNKNOWN, pMessageHandler: *mut c_void, pUserContext: PVOID, pdwMessageMask: PDWORD, dwMessageMaskElements: DWORD) -> HRESULT;

        fn StartSession(&self, pSessionDesc: *mut DVSESSIONDESC, dwFlags: DWORD) -> HRESULT;

        fn StopSession(&self, dwFlags: DWORD) -> HRESULT;

        fn GetSessionDesc(&self, pvSessionDesc: *mut DVSESSIONDESC) -> HRESULT;

        fn SetSessionDesc(&self, pvSessionDesc: *mut DVSESSIONDESC) -> HRESULT;

        fn GetCaps(&self, pDVCaps: *mut DVCAPS) -> HRESULT;

        fn GetCompressionTypes(&self, pData: PVOID, pdwDataSize: PDWORD, pdwNumElements: PDWORD, dwFlags: DWORD) -> HRESULT;

        fn SetTransmitTargets(&self, dvSource: DVID, pdvIDTargets: PDVID, dwNumTargets: DWORD, dwFlags: DWORD) -> HRESULT;

        fn GetTransmitTargets(&self, dvSource: DVID, pdvIDTargets: PDVID, pdwNumTargets: PDWORD, dwFlags: DWORD) -> HRESULT;

        fn SetNotifyMask(&self, pdwMessageMask: PDWORD, dwMessageMaskElements: DWORD) -> HRESULT;
    }
}

com::interfaces! {
    #[uuid("D26AF734-208B-41da-8224-E0CE79810BE1")]
    pub unsafe interface IDirectPlayVoiceTest: IUnknown {
        fn CheckAudioSetup(&self, pguidPlaybackDevice: *const GUID, pguidCaptureDevice: *const GUID, hwndParent: HWND, dwFlags: DWORD) -> HRESULT;
    }
}

com::class! {
    pub class DirectPlayVoiceClient: IDirectPlayVoiceClient {
        client_config: Mutex<Option<DVCLIENTCONFIG>>,
        sound_device_config: Mutex<Option<DVSOUNDDEVICECONFIG>>,
    }

    impl IDirectPlayVoiceClient for DirectPlayVoiceClient {
        fn Initialize(&self, _pVoid: LPUNKNOWN, _pMessageHandler: *mut c_void, _pUserContext: PVOID, _pdwMessageMask: PDWORD, _dwMessageMaskElements: DWORD) -> HRESULT {
            DV_OK
        }

        unsafe fn Connect(&self, pSoundDeviceConfig: *mut DVSOUNDDEVICECONFIG, pdvClientConfig: *mut DVCLIENTCONFIG, _dwFlags: DWORD) -> HRESULT {
            *self.client_config.lock().unwrap() = Some((*pdvClientConfig).clone());
            *self.sound_device_config.lock().unwrap() = Some((*pSoundDeviceConfig).clone());
            DV_OK
        }

        fn Disconnect(&self, _dwFlags: DWORD) -> HRESULT {
            DV_OK
        }

        unsafe fn GetSessionDesc(&self, pvSessionDesc: *mut DVSESSIONDESC) -> HRESULT {
            *pvSessionDesc = DVSESSIONDESC {
                dwSize: mem::size_of::<DVSESSIONDESC>() as u32,
                dwFlags: DVSESSION_NOHOSTMIGRATION,
                dwSessionType: DVSESSIONTYPE_PEER,
                guidCT: PCM.guidType,
                dwBufferQuality: DVBUFFERQUALITY_DEFAULT,
                dwBufferAggressiveness: DVBUFFERAGGRESSIVENESS_DEFAULT,
            };

            DV_OK
        }

        unsafe fn GetClientConfig(&self, pClientConfig: *mut DVCLIENTCONFIG) -> HRESULT {
            match *self.client_config.lock().unwrap() {
                Some(ref client_config) => {
                    if (*pClientConfig).dwSize != mem::size_of::<DVCLIENTCONFIG>() as u32 {
                        return DVERR_INVALIDPARAM;
                    }

                    *pClientConfig = client_config.clone();
                    DV_OK
                },
                None => DVERR_NOTCONNECTED,
            }
        }

        unsafe fn SetClientConfig(&self, pClientConfig: *mut DVCLIENTCONFIG) -> HRESULT {
            if (*pClientConfig).dwSize != mem::size_of::<DVCLIENTCONFIG>() as u32 {
                return DVERR_INVALIDPARAM;
            }

            *self.client_config.lock().unwrap() = Some((*pClientConfig).clone());
            DV_OK
        }

        unsafe fn GetCaps(&self, pDVCaps: *mut DVCAPS) -> HRESULT {
            *pDVCaps = DVCAPS {
                dwSize: mem::size_of::<DVCAPS>() as u32,
                dwFlags: 0,
            };

            DV_OK
        }

        unsafe fn GetCompressionTypes(&self, pData: PVOID, pdwDataSize: PDWORD, pdwNumElements: PDWORD, dwFlags: DWORD) -> HRESULT {
            GetCompressionTypes(pData, pdwDataSize, pdwNumElements, dwFlags)
        }

        fn SetTransmitTargets(&self, _pdvIDTargets: PDVID, _dwNumTargets: DWORD, _dwFlags: DWORD) -> HRESULT {
            DV_OK
        }

        unsafe fn GetTransmitTargets(&self, _pdvIDTargets: PDVID, pdwNumTargets: PDWORD, _dwFlags: DWORD) -> HRESULT {
            *pdwNumTargets = 0;
            DV_OK
        }

        fn Create3DSoundBuffer(&self, _dvID: DVID, _lpdsSourceBuffer: LPDIRECTSOUNDBUFFER, _dwPriority: DWORD, _dwFlags: DWORD, _lpUserBuffer: LPDIRECTSOUND3DBUFFER) -> HRESULT {
            DVERR_ALREADYBUFFERED
        }

        fn Delete3DSoundBuffer(&self, _dvID: DVID, _lpUserBuffer: LPDIRECTSOUND3DBUFFER) -> HRESULT {
            DV_OK
        }

        fn SetNotifyMask(&self, _pdwMessageMask: PDWORD, _dwMessageMaskElements: DWORD) -> HRESULT {
            DV_OK
        }

        unsafe fn GetSoundDeviceConfig(&self, pSoundDeviceConfig: *mut DVSOUNDDEVICECONFIG, _pdwSize: PDWORD) -> HRESULT {
            match *self.sound_device_config.lock().unwrap() {
                Some(ref sound_device_config) => {
                    if (*pSoundDeviceConfig).dwSize != mem::size_of::<DVSOUNDDEVICECONFIG>() as u32 {
                        return DVERR_INVALIDPARAM;
                    }

                    *pSoundDeviceConfig = sound_device_config.clone();
                    DV_OK
                },
                None => DVERR_NOTCONNECTED,
            }
        }
    }
}

com::class! {
    pub class DirectPlayVoiceServer: IDirectPlayVoiceServer {
        session_desc: Mutex<Option<DVSESSIONDESC>>,
    }

    impl IDirectPlayVoiceServer for DirectPlayVoiceServer {
        fn Initialize(&self, _pVoid: LPUNKNOWN, _pMessageHandler: *mut c_void, _pUserContext: PVOID, _pdwMessageMask: PDWORD, _dwMessageMaskElements: DWORD) -> HRESULT {
            DV_OK
        }

        unsafe fn StartSession(&self, pSessionDesc: *mut DVSESSIONDESC, _dwFlags: DWORD) -> HRESULT {
            *self.session_desc.lock().unwrap() = Some((*pSessionDesc).clone());
            DV_OK
        }

        fn StopSession(&self, _dwFlags: DWORD) -> HRESULT {
            DV_OK
        }

        unsafe fn GetSessionDesc(&self, pvSessionDesc: *mut DVSESSIONDESC) -> HRESULT {
            match *self.session_desc.lock().unwrap() {
                Some(ref session_desc) => {
                    if (*pvSessionDesc).dwSize != mem::size_of::<DVSESSIONDESC>() as u32 {
                        return DVERR_INVALIDPARAM;
                    }

                    *pvSessionDesc = session_desc.clone();
                    DV_OK
                },
                None => DVERR_NOTHOSTING,
            }
        }

        unsafe fn SetSessionDesc(&self, pvSessionDesc: *mut DVSESSIONDESC) -> HRESULT {
            if (*pvSessionDesc).dwSize != mem::size_of::<DVSESSIONDESC>() as u32 {
                return DVERR_INVALIDPARAM;
            }

            *self.session_desc.lock().unwrap() = Some((*pvSessionDesc).clone());
            DV_OK
        }

        unsafe fn GetCaps(&self, pDVCaps: *mut DVCAPS) -> HRESULT {
            *pDVCaps = DVCAPS {
                dwSize: mem::size_of::<DVCAPS>() as u32,
                dwFlags: 0,
            };

            DV_OK
        }

        unsafe fn GetCompressionTypes(&self, pData: PVOID, pdwDataSize: PDWORD, pdwNumElements: PDWORD, dwFlags: DWORD) -> HRESULT {
            GetCompressionTypes(pData, pdwDataSize, pdwNumElements, dwFlags)
        }

        fn SetTransmitTargets(&self, _dvSource: DVID, _pdvIDTargets: PDVID, _dwNumTargets: DWORD, _dwFlags: DWORD) -> HRESULT {
            DV_OK
        }

        unsafe fn GetTransmitTargets(&self, _dvSource: DVID, _pdvIDTargets: PDVID, pdwNumTargets: PDWORD, _dwFlags: DWORD) -> HRESULT {
            *pdwNumTargets = 0;
            DV_OK
        }

        fn SetNotifyMask(&self, _pdwMessageMask: PDWORD, _dwMessageMaskElements: DWORD) -> HRESULT {
            DV_OK
        }
    }
}

com::class! {
    pub class DirectPlayVoiceTest: IDirectPlayVoiceTest {}

    impl IDirectPlayVoiceTest for DirectPlayVoiceTest {
        fn CheckAudioSetup(&self, _pguidPlaybackDevice: *const GUID, _pguidCaptureDevice: *const GUID, _hwndParent: HWND, dwFlags: DWORD) -> HRESULT {
            if dwFlags & DVFLAGS_QUERYONLY > 0 {
                DV_FULLDUPLEX
            } else {
                DV_OK
            }
        }
    }
}

unsafe fn GetCompressionTypes(
    pData: PVOID,
    pdwDataSize: PDWORD,
    pdwNumElements: PDWORD,
    _dwFlags: DWORD,
) -> HRESULT {
    *pdwNumElements = 1;

    if *pdwDataSize < PCM.dwSize {
        *pdwDataSize = PCM.dwSize;
        return DVERR_BUFFERTOOSMALL;
    }

    let name_start = (pData as *mut DVCOMPRESSIONINFO).offset(1) as *mut u16;
    ptr::copy(PCM_NAME.as_ptr(), name_start, PCM_NAME.len());

    let mut pcm_with_names = PCM.clone();
    pcm_with_names.lpszName = name_start;
    pcm_with_names.lpszDescription = name_start;

    *(pData as *mut DVCOMPRESSIONINFO) = pcm_with_names;
    DV_OK
}
