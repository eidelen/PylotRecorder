import argparse
import csv
import ctypes
from ctypes import wintypes
import json
import sys
import time
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple

import win32api
import win32con
import win32gui


user32 = ctypes.windll.user32
hid = ctypes.windll.hid

HRAWINPUT = wintypes.HANDLE

WM_INPUT = 0x00FF
RID_INPUT = 0x10000003
RIDEV_INPUTSINK = 0x00000100
RIM_TYPEHID = 2

HIDP_STATUS_SUCCESS = 0x00110000
HIDP_REPORT_TYPE_INPUT = 0

USAGE_PAGE_GENERIC_DESKTOP = 0x01

USAGE_X = 0x30
USAGE_Y = 0x31
USAGE_Z = 0x32
USAGE_RX = 0x33
USAGE_RY = 0x34
USAGE_RZ = 0x35
USAGE_SLIDER = 0x36
USAGE_DIAL = 0x37
USAGE_WHEEL = 0x38
USAGE_HAT = 0x39

AXIS_USAGE_NAMES = {
    USAGE_X: "x",
    USAGE_Y: "y",
    USAGE_Z: "z",
    USAGE_RX: "rx",
    USAGE_RY: "ry",
    USAGE_RZ: "rz",
    USAGE_SLIDER: "slider",
    USAGE_DIAL: "dial",
    USAGE_WHEEL: "wheel",
    USAGE_HAT: "hat",
}


class RAWINPUTDEVICELIST(ctypes.Structure):
    _fields_ = [
        ("hDevice", wintypes.HANDLE),
        ("dwType", wintypes.DWORD),
    ]


class RAWINPUTDEVICE(ctypes.Structure):
    _fields_ = [
        ("usUsagePage", wintypes.USHORT),
        ("usUsage", wintypes.USHORT),
        ("dwFlags", wintypes.DWORD),
        ("hwndTarget", wintypes.HWND),
    ]


class RAWINPUTHEADER(ctypes.Structure):
    _fields_ = [
        ("dwType", wintypes.DWORD),
        ("dwSize", wintypes.DWORD),
        ("hDevice", wintypes.HANDLE),
        ("wParam", wintypes.WPARAM),
    ]


class RAWHID(ctypes.Structure):
    _fields_ = [
        ("dwSizeHid", wintypes.DWORD),
        ("dwCount", wintypes.DWORD),
        ("bRawData", wintypes.BYTE * 1),
    ]


class RID_DEVICE_INFO_HID(ctypes.Structure):
    _fields_ = [
        ("dwVendorId", wintypes.DWORD),
        ("dwProductId", wintypes.DWORD),
        ("dwVersionNumber", wintypes.DWORD),
        ("usUsagePage", wintypes.USHORT),
        ("usUsage", wintypes.USHORT),
    ]


class RID_DEVICE_INFO(ctypes.Structure):
    class _U(ctypes.Union):
        _fields_ = [("hid", RID_DEVICE_INFO_HID)]

    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("dwType", wintypes.DWORD),
        ("u", _U),
    ]


class HIDP_CAPS(ctypes.Structure):
    _fields_ = [
        ("Usage", wintypes.USHORT),
        ("UsagePage", wintypes.USHORT),
        ("InputReportByteLength", wintypes.USHORT),
        ("OutputReportByteLength", wintypes.USHORT),
        ("FeatureReportByteLength", wintypes.USHORT),
        ("Reserved", wintypes.USHORT * 17),
        ("NumberLinkCollectionNodes", wintypes.USHORT),
        ("NumberInputButtonCaps", wintypes.USHORT),
        ("NumberInputValueCaps", wintypes.USHORT),
        ("NumberInputDataIndices", wintypes.USHORT),
        ("NumberOutputButtonCaps", wintypes.USHORT),
        ("NumberOutputValueCaps", wintypes.USHORT),
        ("NumberOutputDataIndices", wintypes.USHORT),
        ("NumberFeatureButtonCaps", wintypes.USHORT),
        ("NumberFeatureValueCaps", wintypes.USHORT),
        ("NumberFeatureDataIndices", wintypes.USHORT),
    ]


class HIDP_BUTTON_CAPS_RANGE(ctypes.Structure):
    _fields_ = [
        ("UsageMin", wintypes.USHORT),
        ("UsageMax", wintypes.USHORT),
        ("StringMin", wintypes.USHORT),
        ("StringMax", wintypes.USHORT),
        ("DesignatorMin", wintypes.USHORT),
        ("DesignatorMax", wintypes.USHORT),
        ("DataIndexMin", wintypes.USHORT),
        ("DataIndexMax", wintypes.USHORT),
    ]


class HIDP_BUTTON_CAPS_NOTRANGE(ctypes.Structure):
    _fields_ = [
        ("Usage", wintypes.USHORT),
        ("StringIndex", wintypes.USHORT),
        ("DesignatorIndex", wintypes.USHORT),
        ("DataIndex", wintypes.USHORT),
        ("Reserved1", wintypes.USHORT),
        ("Reserved2", wintypes.USHORT),
        ("Reserved3", wintypes.USHORT),
        ("Reserved4", wintypes.USHORT),
    ]


class HIDP_BUTTON_CAPS_UNION(ctypes.Union):
    _fields_ = [
        ("Range", HIDP_BUTTON_CAPS_RANGE),
        ("NotRange", HIDP_BUTTON_CAPS_NOTRANGE),
    ]


class HIDP_BUTTON_CAPS(ctypes.Structure):
    _fields_ = [
        ("UsagePage", wintypes.USHORT),
        ("ReportID", ctypes.c_ubyte),
        ("IsAlias", ctypes.c_ubyte),
        ("BitField", wintypes.USHORT),
        ("LinkCollection", wintypes.USHORT),
        ("LinkUsage", wintypes.USHORT),
        ("LinkUsagePage", wintypes.USHORT),
        ("IsRange", ctypes.c_ubyte),
        ("IsStringRange", ctypes.c_ubyte),
        ("IsDesignatorRange", ctypes.c_ubyte),
        ("IsAbsolute", ctypes.c_ubyte),
        ("Reserved", wintypes.ULONG * 10),
        ("u", HIDP_BUTTON_CAPS_UNION),
    ]


class HIDP_VALUE_CAPS_RANGE(ctypes.Structure):
    _fields_ = [
        ("UsageMin", wintypes.USHORT),
        ("UsageMax", wintypes.USHORT),
        ("StringMin", wintypes.USHORT),
        ("StringMax", wintypes.USHORT),
        ("DesignatorMin", wintypes.USHORT),
        ("DesignatorMax", wintypes.USHORT),
        ("DataIndexMin", wintypes.USHORT),
        ("DataIndexMax", wintypes.USHORT),
    ]


class HIDP_VALUE_CAPS_NOTRANGE(ctypes.Structure):
    _fields_ = [
        ("Usage", wintypes.USHORT),
        ("StringIndex", wintypes.USHORT),
        ("DesignatorIndex", wintypes.USHORT),
        ("DataIndex", wintypes.USHORT),
        ("Reserved1", wintypes.USHORT),
        ("Reserved2", wintypes.USHORT),
        ("Reserved3", wintypes.USHORT),
        ("Reserved4", wintypes.USHORT),
    ]


class HIDP_VALUE_CAPS_UNION(ctypes.Union):
    _fields_ = [
        ("Range", HIDP_VALUE_CAPS_RANGE),
        ("NotRange", HIDP_VALUE_CAPS_NOTRANGE),
    ]


class HIDP_VALUE_CAPS(ctypes.Structure):
    _fields_ = [
        ("UsagePage", wintypes.USHORT),
        ("ReportID", ctypes.c_ubyte),
        ("IsAlias", ctypes.c_ubyte),
        ("BitField", wintypes.USHORT),
        ("LinkCollection", wintypes.USHORT),
        ("LinkUsage", wintypes.USHORT),
        ("LinkUsagePage", wintypes.USHORT),
        ("IsRange", ctypes.c_ubyte),
        ("IsStringRange", ctypes.c_ubyte),
        ("IsDesignatorRange", ctypes.c_ubyte),
        ("IsAbsolute", ctypes.c_ubyte),
        ("HasNull", ctypes.c_ubyte),
        ("Reserved", ctypes.c_ubyte),
        ("BitSize", wintypes.USHORT),
        ("ReportCount", wintypes.USHORT),
        ("Reserved2", wintypes.USHORT * 5),
        ("UnitsExp", wintypes.ULONG),
        ("Units", wintypes.ULONG),
        ("LogicalMin", wintypes.LONG),
        ("LogicalMax", wintypes.LONG),
        ("PhysicalMin", wintypes.LONG),
        ("PhysicalMax", wintypes.LONG),
        ("u", HIDP_VALUE_CAPS_UNION),
    ]


def _check_hidp(status: int, label: str) -> bool:
    if status != HIDP_STATUS_SUCCESS:
        return False
    return True


def get_raw_input_device_list() -> List[RAWINPUTDEVICELIST]:
    device_count = wintypes.UINT(0)
    res = user32.GetRawInputDeviceList(
        None,
        ctypes.byref(device_count),
        ctypes.sizeof(RAWINPUTDEVICELIST),
    )
    if res != 0:
        raise ctypes.WinError()
    if device_count.value == 0:
        return []
    array_type = RAWINPUTDEVICELIST * device_count.value
    device_list = array_type()
    res = user32.GetRawInputDeviceList(
        device_list,
        ctypes.byref(device_count),
        ctypes.sizeof(RAWINPUTDEVICELIST),
    )
    if res == wintypes.UINT(-1).value:
        raise ctypes.WinError()
    return list(device_list)


def get_device_name(handle: wintypes.HANDLE) -> str:
    size = wintypes.UINT(0)
    res = user32.GetRawInputDeviceInfoW(
        handle, 0x20000007, None, ctypes.byref(size)
    )
    if res == wintypes.UINT(-1).value:
        return ""
    buffer = ctypes.create_unicode_buffer(size.value + 1)
    res = user32.GetRawInputDeviceInfoW(
        handle, 0x20000007, buffer, ctypes.byref(size)
    )
    if res == wintypes.UINT(-1).value:
        return ""
    return buffer.value


def get_device_info(handle: wintypes.HANDLE) -> Optional[RID_DEVICE_INFO]:
    info = RID_DEVICE_INFO()
    info.cbSize = ctypes.sizeof(RID_DEVICE_INFO)
    size = wintypes.UINT(info.cbSize)
    res = user32.GetRawInputDeviceInfoW(
        handle, 0x2000000B, ctypes.byref(info), ctypes.byref(size)
    )
    if res == wintypes.UINT(-1).value:
        return None
    return info


def get_preparsed_data(handle: wintypes.HANDLE) -> Optional[ctypes.Array]:
    size = wintypes.UINT(0)
    res = user32.GetRawInputDeviceInfoW(
        handle, 0x20000005, None, ctypes.byref(size)
    )
    if res == wintypes.UINT(-1).value or size.value == 0:
        return None
    buffer = ctypes.create_string_buffer(size.value)
    res = user32.GetRawInputDeviceInfoW(
        handle, 0x20000005, buffer, ctypes.byref(size)
    )
    if res == wintypes.UINT(-1).value:
        return None
    return buffer


def parse_hid_caps(
    preparsed: ctypes.Array,
) -> Optional[Tuple[HIDP_CAPS, List[HIDP_VALUE_CAPS], List[HIDP_BUTTON_CAPS]]]:
    caps = HIDP_CAPS()
    status = hid.HidP_GetCaps(preparsed, ctypes.byref(caps))
    if not _check_hidp(status, "HidP_GetCaps"):
        return None

    value_caps_count = wintypes.USHORT(caps.NumberInputValueCaps)
    button_caps_count = wintypes.USHORT(caps.NumberInputButtonCaps)

    value_caps = (HIDP_VALUE_CAPS * caps.NumberInputValueCaps)()
    button_caps = (HIDP_BUTTON_CAPS * caps.NumberInputButtonCaps)()

    status = hid.HidP_GetValueCaps(
        HIDP_REPORT_TYPE_INPUT,
        value_caps,
        ctypes.byref(value_caps_count),
        preparsed,
    )
    if not _check_hidp(status, "HidP_GetValueCaps"):
        value_caps = []
    else:
        value_caps = list(value_caps)[: value_caps_count.value]

    status = hid.HidP_GetButtonCaps(
        HIDP_REPORT_TYPE_INPUT,
        button_caps,
        ctypes.byref(button_caps_count),
        preparsed,
    )
    if not _check_hidp(status, "HidP_GetButtonCaps"):
        button_caps = []
    else:
        button_caps = list(button_caps)[: button_caps_count.value]

    return caps, value_caps, button_caps


def normalize_value(value: int, logical_min: int, logical_max: int) -> Optional[float]:
    if logical_max == logical_min:
        return None
    span = float(logical_max - logical_min)
    if logical_min < 0:
        return ((value - logical_min) / span) * 2.0 - 1.0
    return (value - logical_min) / span


def hid_get_usage_value(
    preparsed: ctypes.Array,
    report_buf: ctypes.Array,
    report_len: int,
    usage_page: int,
    usage: int,
) -> Optional[int]:
    value = wintypes.ULONG(0)
    status = hid.HidP_GetUsageValue(
        HIDP_REPORT_TYPE_INPUT,
        wintypes.USHORT(usage_page),
        0,
        wintypes.USHORT(usage),
        ctypes.byref(value),
        preparsed,
        ctypes.cast(report_buf, ctypes.c_char_p),
        wintypes.ULONG(report_len),
    )
    if not _check_hidp(status, "HidP_GetUsageValue"):
        return None
    return int(value.value)


def hid_get_usages(
    preparsed: ctypes.Array,
    report_buf: ctypes.Array,
    report_len: int,
    usage_page: int,
    link_collection: int,
    max_usages: int,
) -> List[int]:
    usages = (wintypes.USHORT * max_usages)()
    usage_count = wintypes.ULONG(max_usages)
    status = hid.HidP_GetUsages(
        HIDP_REPORT_TYPE_INPUT,
        wintypes.USHORT(usage_page),
        wintypes.USHORT(link_collection),
        usages,
        ctypes.byref(usage_count),
        preparsed,
        ctypes.cast(report_buf, ctypes.c_char_p),
        wintypes.ULONG(report_len),
    )
    if not _check_hidp(status, "HidP_GetUsages"):
        return []
    return [int(usages[i]) for i in range(usage_count.value)]


def decode_hid_report(
    report: bytes,
    preparsed: ctypes.Array,
    value_caps: List[HIDP_VALUE_CAPS],
    button_caps: List[HIDP_BUTTON_CAPS],
) -> Tuple[Dict[str, Dict[str, Any]], List[int]]:
    # Add device-specific decoding here if you want richer mappings later.
    axes: Dict[str, Dict[str, Any]] = {}
    buttons: List[int] = []
    report_buf = ctypes.create_string_buffer(report, len(report))

    for cap in value_caps:
        usage_page = int(cap.UsagePage)
        if usage_page != USAGE_PAGE_GENERIC_DESKTOP:
            continue

        usages: List[int] = []
        if cap.IsRange:
            usages = list(
                range(int(cap.u.Range.UsageMin), int(cap.u.Range.UsageMax) + 1)
            )
        else:
            usages = [int(cap.u.NotRange.Usage)]

        for usage in usages:
            if usage not in AXIS_USAGE_NAMES:
                continue
            value = hid_get_usage_value(
                preparsed, report_buf, len(report), usage_page, usage
            )
            if value is None:
                continue
            norm = normalize_value(value, int(cap.LogicalMin), int(cap.LogicalMax))
            axes[AXIS_USAGE_NAMES[usage]] = {
                "raw": value,
                "norm": norm,
                "min": int(cap.LogicalMin),
                "max": int(cap.LogicalMax),
            }

    for cap in button_caps:
        usage_page = int(cap.UsagePage)
        max_usages = 0
        link = int(cap.LinkCollection)
        if cap.IsRange:
            max_usages = int(cap.u.Range.UsageMax) - int(cap.u.Range.UsageMin) + 1
        else:
            max_usages = 1
        if max_usages <= 0:
            continue
        pressed = hid_get_usages(
            preparsed, report_buf, len(report), usage_page, link, max_usages
        )
        buttons.extend(pressed)

    buttons = sorted(set(buttons))
    return axes, buttons


def parse_report_bytes(data: bytes, report_size: int, count: int) -> List[bytes]:
    reports = []
    for idx in range(count):
        start = idx * report_size
        end = start + report_size
        if end <= len(data):
            reports.append(data[start:end])
    return reports


class DeviceState:
    def __init__(self, handle: wintypes.HANDLE):
        self.handle = handle
        self.name = get_device_name(handle)
        self.info = get_device_info(handle)
        self.preparsed = get_preparsed_data(handle)
        self.caps: Optional[HIDP_CAPS] = None
        self.value_caps: List[HIDP_VALUE_CAPS] = []
        self.button_caps: List[HIDP_BUTTON_CAPS] = []
        if self.preparsed is not None:
            parsed = parse_hid_caps(self.preparsed)
            if parsed is not None:
                self.caps, self.value_caps, self.button_caps = parsed

    @property
    def usage_page(self) -> Optional[int]:
        if self.info is None:
            return None
        return int(self.info.u.hid.usUsagePage)

    @property
    def usage(self) -> Optional[int]:
        if self.info is None:
            return None
        return int(self.info.u.hid.usUsage)


class EventWriter:
    def __init__(self, path: str, fmt: str):
        self.path = path
        self.fmt = fmt
        self.file = open(path, "w", newline="", encoding="utf-8")
        self.csv_writer: Optional[csv.DictWriter] = None
        if fmt == "csv":
            fieldnames = [
                "timestamp_iso",
                "timestamp_ms",
                "device_handle",
                "device_name",
                "usage_page",
                "usage",
                "report_size",
                "report_hex",
                "axes_json",
                "buttons_json",
            ]
            self.csv_writer = csv.DictWriter(self.file, fieldnames=fieldnames)
            self.csv_writer.writeheader()

    def write_event(self, row: Dict[str, Any]) -> None:
        if self.fmt == "jsonl":
            self.file.write(json.dumps(row, separators=(",", ":")) + "\n")
        else:
            if self.csv_writer is None:
                return
            self.csv_writer.writerow(row)

    def close(self) -> None:
        try:
            self.file.flush()
        finally:
            self.file.close()


class RawInputLogger:
    def __init__(self, writer: EventWriter, device_filter: Optional[str], echo: bool):
        self.writer = writer
        self.device_filter = device_filter.lower() if device_filter else None
        self.device_cache: Dict[int, DeviceState] = {}
        self.echo = echo

    def _match_filter(self, name: str) -> bool:
        if not self.device_filter:
            return True
        return self.device_filter in name.lower()

    def get_device_state(self, handle: wintypes.HANDLE) -> Optional[DeviceState]:
        handle_value = int(ctypes.cast(handle, ctypes.c_void_p).value or 0)
        if handle_value in self.device_cache:
            return self.device_cache[handle_value]
        state = DeviceState(handle)
        if not self._match_filter(state.name):
            return None
        self.device_cache[handle_value] = state
        return state

    def handle_wm_input(self, lparam: int) -> None:
        data_size = wintypes.UINT(0)
        res = user32.GetRawInputData(
            HRAWINPUT(lparam),
            RID_INPUT,
            None,
            ctypes.byref(data_size),
            ctypes.sizeof(RAWINPUTHEADER),
        )
        if res == wintypes.UINT(-1).value or data_size.value == 0:
            return

        buffer = (ctypes.c_ubyte * data_size.value)()
        res = user32.GetRawInputData(
            HRAWINPUT(lparam),
            RID_INPUT,
            buffer,
            ctypes.byref(data_size),
            ctypes.sizeof(RAWINPUTHEADER),
        )
        if res == wintypes.UINT(-1).value:
            return

        header = RAWINPUTHEADER.from_buffer_copy(buffer)
        if header.dwType != RIM_TYPEHID:
            return

        rahid_offset = ctypes.sizeof(RAWINPUTHEADER)
        rahid = RAWHID.from_buffer_copy(buffer, rahid_offset)
        data_offset = rahid_offset + 8
        raw_bytes = bytes(buffer[data_offset : data_offset + rahid.dwSizeHid * rahid.dwCount])
        reports = parse_report_bytes(raw_bytes, int(rahid.dwSizeHid), int(rahid.dwCount))

        device_state = self.get_device_state(header.hDevice)
        if device_state is None:
            return

        timestamp = time.time()
        ts_iso = datetime.fromtimestamp(timestamp, tz=timezone.utc).isoformat()
        ts_ms = int(timestamp * 1000)
        device_handle_str = hex(int(ctypes.cast(header.hDevice, ctypes.c_void_p).value or 0))

        for report in reports:
            axes: Dict[str, Dict[str, Any]] = {}
            buttons: List[int] = []
            if device_state.preparsed is not None and device_state.value_caps:
                try:
                    axes, buttons = decode_hid_report(
                        report,
                        device_state.preparsed,
                        device_state.value_caps,
                        device_state.button_caps,
                    )
                except Exception:
                    axes = {}
                    buttons = []

            row = {
                "timestamp_iso": ts_iso,
                "timestamp_ms": ts_ms,
                "device_handle": device_handle_str,
                "device_name": device_state.name,
                "usage_page": device_state.usage_page,
                "usage": device_state.usage,
                "report_size": len(report),
                "report_hex": report.hex(),
                "axes_json": json.dumps(axes, separators=(",", ":")),
                "buttons_json": json.dumps(buttons, separators=(",", ":")),
            }
            self.writer.write_event(row)
            if self.echo:
                axes_summary = {
                    key: value.get("norm") for key, value in axes.items()
                }
                print(
                    f"{ts_ms} {device_handle_str} {len(report)} "
                    f"axes={axes_summary} buttons={buttons}"
                )


def list_devices() -> None:
    devices = get_raw_input_device_list()
    if not devices:
        print("No raw input devices found.")
        return
    print("Raw input devices:")
    for entry in devices:
        name = get_device_name(entry.hDevice)
        info = get_device_info(entry.hDevice)
        usage_page = None
        usage = None
        if info is not None:
            usage_page = info.u.hid.usUsagePage
            usage = info.u.hid.usUsage
        handle_value = int(ctypes.cast(entry.hDevice, ctypes.c_void_p).value or 0)
        print(
            f"- handle=0x{handle_value:x} type={entry.dwType} "
            f"usage_page={usage_page} usage={usage} name={name}"
        )


def register_raw_input(hwnd: int) -> None:
    devices = (RAWINPUTDEVICE * 2)()
    devices[0].usUsagePage = USAGE_PAGE_GENERIC_DESKTOP
    devices[0].usUsage = 0x04
    devices[0].dwFlags = RIDEV_INPUTSINK
    devices[0].hwndTarget = hwnd
    devices[1].usUsagePage = USAGE_PAGE_GENERIC_DESKTOP
    devices[1].usUsage = 0x05
    devices[1].dwFlags = RIDEV_INPUTSINK
    devices[1].hwndTarget = hwnd

    if not user32.RegisterRawInputDevices(devices, 2, ctypes.sizeof(RAWINPUTDEVICE)):
        raise ctypes.WinError()


def create_message_window(wnd_proc):
    h_instance = win32api.GetModuleHandle(None)
    class_name = "RawInputHiddenWindow"

    wnd_class = win32gui.WNDCLASS()
    wnd_class.lpfnWndProc = wnd_proc
    wnd_class.hInstance = h_instance
    wnd_class.lpszClassName = class_name
    try:
        win32gui.RegisterClass(wnd_class)
    except win32gui.error:
        pass

    hwnd = win32gui.CreateWindowEx(
        0,
        class_name,
        "RawInputHiddenWindow",
        0,
        0,
        0,
        0,
        0,
        win32con.HWND_MESSAGE,
        0,
        h_instance,
        None,
    )
    return hwnd


def main() -> int:
    parser = argparse.ArgumentParser(description="Raw Input joystick/gamepad logger")
    parser.add_argument("--out", default="rawinput_log.csv", help="Output file path")
    parser.add_argument(
        "--format",
        choices=["csv", "jsonl"],
        default="csv",
        help="Output format",
    )
    parser.add_argument(
        "--device",
        default=None,
        help="Filter device by substring (e.g. VID_044F or Thrustmaster)",
    )
    parser.add_argument(
        "--print",
        action="store_true",
        help="Print a short summary of each input event to the console",
    )
    args = parser.parse_args()

    list_devices()

    writer = EventWriter(args.out, args.format)
    logger = RawInputLogger(writer, args.device, args.print)

    def wnd_proc(hwnd, msg, wparam, lparam):
        if msg == WM_INPUT:
            logger.handle_wm_input(lparam)
            return 0
        if msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
            return 0
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    hwnd = create_message_window(wnd_proc)
    register_raw_input(hwnd)

    print("Logging WM_INPUT (Ctrl+C to stop).")
    try:
        while True:
            bret, msg = win32gui.GetMessage(None, 0, 0)
            if bret == 0:
                break
            win32gui.TranslateMessage(msg)
            win32gui.DispatchMessage(msg)
    except KeyboardInterrupt:
        win32gui.PostQuitMessage(0)
    finally:
        writer.close()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
