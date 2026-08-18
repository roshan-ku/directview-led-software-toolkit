#!/usr/bin/env python3
"""
Validation Report Generator for Direct View LED Software Toolkit (dvledtx).

Runs functional validation tests against the dvledtx binary, captures logs,
and generates an Excel workbook with one sheet per use-case category.
Frame Transmission sheets are split by pixel format and session count.

Usage:
    python3 generate_validation_report.py [--binary PATH] [--video PATH] [--output PATH]
                                          [--skip-tx] [--results PATH]
"""

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from collections import defaultdict
from datetime import datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required. Install with: pip install openpyxl")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
SUPPORTED_FORMATS = ["yuv422p10le", "yuv420", "yuv444p10le", "gbrp10le",
                     "yuv422p12le", "yuv444p12le", "gbrp12le"]
TEN_BIT_FORMATS = ["yuv422p10le", "yuv444p10le", "gbrp10le"]
TWELVE_BIT_FORMATS = ["yuv422p12le", "yuv444p12le", "gbrp12le"]
SESSION_COUNTS = [1, 3]
MAX_WIDTH = 3840
MAX_HEIGHT = 2160
SUPPORTED_FPS = [25, 30, 50, 60]
SUPPORTED_INPUT_MODES = ["file", "screen_capture"]
BINARY_NAME = "dvledtx"
VERSION = "1.0.0"

# Resolution matrix for FT tests
RESOLUTIONS = [
    {"name": "1080p", "width": 1920, "height": 1080},
    {"name": "2K",    "width": 2560, "height": 1440},
    {"name": "4K",    "width": 3840, "height": 2160},
]
FT_FPS_LIST = [30, 60]

# How long each TX/FT test streams for, in seconds (dvledtx --test-time).
# 30s is enough to validate correctness; raise it via --ft-test-time for soak
# runs where you want stable throughput/FPS averages.
DEFAULT_TX_TEST_TIME = 30
# Extra seconds allowed on top of --test-time for MTL/DPDK init and teardown
# before the subprocess is force-killed.
TX_TIMEOUT_MARGIN = 30

# Map config fmt names to folder names for video lookup
FMT_TO_FOLDER = {
    "yuv422p10le": "yuv422p10le",
    "yuv420":      "yuv420p",
    "yuv444p10le": "yuv444p10le",
    "gbrp10le":    "gbrp10le",
    "yuv422p12le": "yuv422p12le",
    "yuv444p12le": "yuv444p12le",
    "gbrp12le":    "gbrp12le",
}

# Map config fmt to the actual filename token
FMT_TO_FILETOK = {
    "yuv422p10le": "yuv422p10le",
    "yuv420":      "yuv420p",
    "yuv444p10le": "yuv444p10le",
    "gbrp10le":    "gbrp10le",
    "yuv422p12le": "yuv422p12le",
    "yuv444p12le": "yuv444p12le",
    "gbrp12le":    "gbrp12le",
}

# Map resolution name to filename token
RES_TO_FILETOK = {
    "1080p": "1080p",
    "2K":    "2k",
    "4K":    "4k",
}

# Bits-per-pixel for bandwidth calculation (uncompressed ST2110-20)
FMT_BPP = {
    "yuv422p10le": 20,   # 4:2:2 10-bit
    "yuv420":      12,   # 4:2:0 8-bit
    "yuv444p10le": 30,   # 4:4:4 10-bit
    "gbrp10le":    30,   # RGB 10-bit (same as 4:4:4)
    "yuv422p12le": 24,   # 4:2:2 12-bit
    "yuv444p12le": 36,   # 4:4:4 12-bit
    "gbrp12le":    36,   # RGB 12-bit
}

# Bit depth per pixel format (used for the 10-bit vs 12-bit coverage matrix)
FMT_BIT_DEPTH = {
    "yuv422p10le": 10,
    "yuv420":       8,
    "yuv444p10le": 10,
    "gbrp10le":    10,
    "yuv422p12le": 12,
    "yuv444p12le": 12,
    "gbrp12le":    12,
}

# Expected ST2110-20 wire (transport) format per pixel format.
# mtl_tx.c maps AV_PIX_FMT_* -> ST20_FMT_*; the MTL log prints either the
# ST20_FMT_* enum name or the RFC4175 pixel-group name, so both are accepted.
EXPECTED_TRANSPORT = {
    "yuv422p10le": r"(YUV_?422_?10|422RFC4175PG2BE10)",
    "yuv420":      r"(YUV_?420_?8|420(CUSTOM|PLANAR)8)",
    "yuv444p10le": r"(YUV_?444_?10|444RFC4175PG4BE10)",
    "gbrp10le":    r"(RGB_?10|GBR.*10|444RFC4175PG4BE10)",
    "yuv422p12le": r"(YUV_?422_?12|422RFC4175PG2BE12)",
    "yuv444p12le": r"(YUV_?444_?12|444RFC4175PG4BE12)",
    "gbrp12le":    r"(RGB_?12|GBR.*12|444RFC4175PG4BE12)",
}

# Scaling test matrix: (source WxH) -> (scaled WxH)
SCALE_CASES = [
    ("downscale 4K->1080p", 3840, 2160, 1920, 1080),
    ("upscale 1080p->4K",   1920, 1080, 3840, 2160),
    ("downscale 4K->2K",    3840, 2160, 2560, 1440),
    ("upscale 1080p->2K",   1920, 1080, 2560, 1440),
]

# DPDK-bound VFs (vfio-pci), not the kernel-bound PF. Check with:
#   dpdk-devbind.py --status | grep vfio-pci
DEFAULT_BDF = "0000:03:10.1"
DEFAULT_BDF2 = "0000:03:10.3"
DEFAULT_SIP = "192.168.50.29"
DEFAULT_DIP = "239.168.85.20"
DEFAULT_SCREEN_INPUT = ":0.0+0,0"

CATEGORY_MAP = {
    "CP": "Config Parsing",
    "CV": "Config Validation",
    "CL": "Command Line",
    "LF": "Log File",
    "SEC": "Security",
    "MN": "Multi-NIC",
    "SL": "Scaling",
    "SC": "Screen Capture",
    "PTP": "PTP Timing",
    "FD": "Frame Delivery",
    "FT": "Frame Transmission",
    "FH": "Frame Handler",
    "SM": "Session Manager",
    "MT": "MTL TX",
}

# Matches tests that exercise more than one tiled TX session, either by TC-ID
# convention ("FT-P03-3ses-...") or by description ("3 tiled sessions", ...).
MULTI_SESSION_RE = re.compile(r"\dses\b|tiled|[2-8][\s-]*(?:tiled\s*)?sessions?\b", re.I)

# Feature coverage map: feature name -> TC-ID prefixes / matchers used to
# collect the tests that exercise it (rendered on the Summary sheet).
FEATURE_COVERAGE = [
    ("10-bit formats",       lambda t: any(f in t["tc_id"] for f in TEN_BIT_FORMATS)),
    ("12-bit formats",       lambda t: any(f in t["tc_id"] for f in TWELVE_BIT_FORMATS)),
    ("8-bit (yuv420)",       lambda t: "yuv420" in t["tc_id"]),
    ("Scaling",              lambda t: t["tc_id"].startswith("SL-") or "SCALE" in t["tc_id"].upper()),
    ("Screen capture",       lambda t: t["tc_id"].startswith("SC-") or "SCREEN" in t["tc_id"].upper()),
    ("Multiple NIC",         lambda t: t["tc_id"].startswith("MN-")),
    ("Multi-session (tiled)",
     lambda t: bool(MULTI_SESSION_RE.search(t["tc_id"] + " " + t.get("description", "")))),
    ("PTP timing",           lambda t: t["tc_id"].startswith("PTP-")),
    ("Frame transmission",   lambda t: t["tc_id"].startswith("FT-")),
]

# Excel styling
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
PASS_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
FAIL_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
SKIP_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
SUMMARY_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SUMMARY_FONT = Font(name="Calibri", bold=True, size=11)
TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="1F4E79")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


# ---------------------------------------------------------------------------
# Test infrastructure
# ---------------------------------------------------------------------------
class TestRunner:
    """Runs dvledtx binary with generated configs and captures output."""

    def __init__(self, binary_path, video_path, tmp_dir, tx_test_time=DEFAULT_TX_TEST_TIME):
        self.binary = binary_path
        self.video = video_path
        self.tmp_dir = tmp_dir
        self.tx_test_time = tx_test_time
        self.results = []

    # Keys that live in the "tx_video" block of the JSON schema (config_reader.c
    # reads fps/fmt/scale_* from "tx_video", everything else from "video").
    TX_VIDEO_KEYS = ("fps", "fmt", "scale_width", "scale_height")

    @classmethod
    def _normalize_key(cls, key_path):
        """Route legacy 'video.fps'-style overrides to the 'tx_video' block."""
        parts = key_path.split(".")
        if len(parts) == 2 and parts[0] == "video" and parts[1] in cls.TX_VIDEO_KEYS:
            return "tx_video." + parts[1]
        return key_path

    def _make_config(self, overrides=None, sessions=None, no_log_file=False,
                     log_file=None, no_video=False, raw_json=None,
                     interfaces=None, input_mode=None, screen_input=None,
                     scale=None, ptp=None):
        """Create a temporary JSON config file.

        Schema (see src/util/config_reader.c):
          interfaces[] : name / sip / dip (+ optional nic_index)
          video        : width / height / input_mode / screen_input / tx_url
          tx_video     : scale_width / scale_height / fps / fmt
          ptp          : enable / pi / unicast (optional)
          tx_sessions[]: nic_index / udp_port / payload_type / crop
        """
        if raw_json is not None:
            path = os.path.join(self.tmp_dir, f"txval_{id(raw_json) & 0xFFFFFFFF:08x}.json")
            with open(path, "w") as f:
                f.write(raw_json)
            return path

        config = {
            "interfaces": interfaces or [{
                "name": DEFAULT_BDF,
                "sip": DEFAULT_SIP,
                "dip": DEFAULT_DIP,
            }],
            "video": {
                "width": MAX_WIDTH,
                "height": MAX_HEIGHT,
            },
            "tx_video": {
                "fps": 30,
                "fmt": "yuv422p10le",
            },
            "tx_sessions": sessions or [{
                "udp_port": 20000,
                "payload_type": 96,
                "crop": {"x": 0, "y": 0, "w": MAX_WIDTH, "h": MAX_HEIGHT},
            }],
        }

        if input_mode is not None:
            config["video"]["input_mode"] = input_mode
        if screen_input is not None:
            config["video"]["screen_input"] = screen_input
        if scale is not None:
            config["tx_video"]["scale_width"] = scale[0]
            config["tx_video"]["scale_height"] = scale[1]
        if ptp is not None:
            config["ptp"] = ptp

        if not no_video and self.video:
            config["video"]["tx_url"] = self.video

        if not no_log_file:
            if log_file is not None:
                config["log_file"] = log_file
            elif not no_video:
                config["log_file"] = os.path.join(self.tmp_dir, "dvledtx_test.log")

        if overrides:
            for key_path, value in overrides.items():
                parts = self._normalize_key(key_path).split(".")
                obj = config
                for p in parts[:-1]:
                    if p.isdigit():
                        obj = obj[int(p)]
                    else:
                        obj = obj.setdefault(p, {})
                last = parts[-1]
                if last.isdigit():
                    obj[int(last)] = value
                else:
                    obj[last] = value

        path = os.path.join(self.tmp_dir, f"txval_{len(self.results):04d}.json")
        with open(path, "w") as f:
            json.dump(config, f, indent=2)
        return path

    def _run(self, args, timeout=5):
        """Run dvledtx with given args and capture output."""
        try:
            proc = subprocess.run(
                args, capture_output=True, text=True, timeout=timeout,
                env={**os.environ, "TERM": "dumb"},
            )
            return proc.returncode, proc.stdout + proc.stderr
        except subprocess.TimeoutExpired as e:
            stdout = (e.stdout or b"").decode("utf-8", errors="replace")
            stderr = (e.stderr or b"").decode("utf-8", errors="replace")
            return -9, stdout + stderr + "\n[TIMEOUT after {}s]".format(timeout)

    def run_config_test(self, tc_id, description, test_type, config_path,
                        expected_pattern, expect_pass=False, timeout=5,
                        extra_args=None):
        """Run a test with a config file and check output pattern."""
        args = [self.binary, "--config", config_path]
        if extra_args:
            args.extend(extra_args)

        t0 = time.monotonic()
        exit_code, output = self._run(args, timeout=timeout)
        duration_ms = (time.monotonic() - t0) * 1000

        # If output is empty/minimal, check the log file from the config
        # (stdout/stderr may have been redirected to the log file by dvledtx)
        if not output.strip() or (expected_pattern and not re.search(expected_pattern, output, re.IGNORECASE)):
            try:
                with open(config_path, "r") as f:
                    cfg_data = json.load(f)
                log_file = cfg_data.get("log_file", "")
                if log_file and os.path.isfile(log_file):
                    with open(log_file, "r", errors="replace") as lf:
                        output += "\n" + lf.read()
            except (json.JSONDecodeError, OSError, KeyError):
                pass

        pattern_found = bool(re.search(expected_pattern, output, re.IGNORECASE)) if expected_pattern else True

        if expect_pass:
            status = "PASS" if pattern_found else "FAIL"
        else:
            status = "PASS" if exit_code != 0 and pattern_found else "FAIL"

        result = {
            "tc_id": tc_id,
            "description": description,
            "test_type": test_type,
            "status": status,
            "log_output": output[:3000],
            "exit_code": exit_code,
            "matched_pattern": expected_pattern or "",
            "duration_ms": duration_ms,
        }
        self.results.append(result)
        sym = "+" if status == "PASS" else "X"
        print(f"  [{sym}] {tc_id}: {description} [{status}]")
        return result

    def run_cli_test(self, tc_id, description, test_type, args,
                     expected_pattern, expect_pass=False, timeout=5):
        """Run a CLI test with arbitrary args."""
        t0 = time.monotonic()
        exit_code, output = self._run(args, timeout=timeout)
        duration_ms = (time.monotonic() - t0) * 1000

        pattern_found = bool(re.search(expected_pattern, output, re.IGNORECASE)) if expected_pattern else True

        if expect_pass:
            status = "PASS" if (exit_code == 0 or exit_code == -9) and pattern_found else "FAIL"
        else:
            status = "PASS" if exit_code != 0 and pattern_found else "FAIL"

        result = {
            "tc_id": tc_id,
            "description": description,
            "test_type": test_type,
            "status": status,
            "log_output": output[:3000],
            "exit_code": exit_code,
            "matched_pattern": expected_pattern or "",
            "duration_ms": duration_ms,
        }
        self.results.append(result)
        sym = "+" if status == "PASS" else "X"
        print(f"  [{sym}] {tc_id}: {description} [{status}]")
        return result

    def run_tx_test(self, tc_id, description, fmt, session_count, timeout=50):
        """Run a frame transmission test with specific format and session count."""
        return self.run_tx_test_full(tc_id, description, fmt, session_count,
                                     MAX_WIDTH, MAX_HEIGHT, 30, self.video, timeout)

    def run_tx_test_full(self, tc_id, description, fmt, session_count,
                         width, height, fps, video_path, timeout=50,
                         scale=None, interfaces=None, input_mode=None,
                         screen_input=None):
        """Run a frame transmission test with full control over resolution/fps/video.

        scale        -- (scale_w, scale_h) to exercise the scaling path
        interfaces   -- list of interface dicts to exercise multiple NICs
        input_mode   -- "file" (default) or "screen_capture"
        screen_input -- x11grab source string when input_mode=screen_capture
        """
        # Crop rectangles are validated against the *effective* (scaled) size
        out_w, out_h = (scale if scale else (width, height))
        nic_count = len(interfaces) if interfaces else 1

        if session_count == 1:
            sessions = [{
                "nic_index": 0, "udp_port": 20000, "payload_type": 96,
                "crop": {"x": 0, "y": 0, "w": out_w, "h": out_h},
            }]
        else:
            strip_w = (out_w // session_count) & ~1  # keep even for YUV
            sessions = []
            for i in range(session_count):
                sessions.append({
                    "nic_index": i % nic_count,
                    "udp_port": 20000 + i * 2, "payload_type": 96,
                    "crop": {"x": i * strip_w, "y": 0, "w": strip_w, "h": out_h},
                })

        # Use a dedicated log file in CWD (dvledtx only allows /var/log/ or CWD)
        safe_id = re.sub(r"[^A-Za-z0-9_.-]", "_", tc_id)
        log_file = os.path.join(os.getcwd(), f"ft_{safe_id}.log")

        overrides = {"tx_video.fmt": fmt, "video.width": width,
                     "video.height": height, "tx_video.fps": fps}
        if input_mode == "screen_capture":
            overrides["video.screen_input"] = screen_input or DEFAULT_SCREEN_INPUT
        else:
            overrides["video.tx_url"] = video_path

        config = self._make_config(
            overrides=overrides,
            sessions=sessions,
            interfaces=interfaces,
            input_mode=input_mode,
            scale=scale,
            no_video=True,  # source is set via overrides
            log_file=log_file,
        )
        args = [self.binary, "--config", config,
                "--test-time", str(self.tx_test_time)]
        # Per-call timeouts assume the default 30s run; when --ft-test-time is
        # raised for soak runs the subprocess budget must grow with it.
        timeout = max(timeout, self.tx_test_time + TX_TIMEOUT_MARGIN)
        t0 = time.monotonic()
        exit_code, output = self._run(args, timeout=timeout)
        duration_ms = (time.monotonic() - t0) * 1000

        # Read log file content (dvledtx writes detailed MTL info there)
        log_content = ""
        if os.path.isfile(log_file):
            try:
                with open(log_file, "r", errors="replace") as f:
                    log_content = f.read()
            except Exception:
                pass
            try:
                os.unlink(log_file)
            except OSError:
                pass

        # Combine stdout/stderr and log file for full analysis
        full_output = output + "\n" + log_content

        # Check for successful transmission indicators in log
        started = bool(re.search(r"dvledtx started successfully", full_output))
        has_fps = bool(re.search(r"fps [\d.]+ frames \d+", full_output))
        has_transport = bool(re.search(r"transport fmt", full_output))

        status = "PASS" if (started and (has_fps or has_transport)) else "FAIL"

        # Extract throughput (Mbps) from MTL TX_VIDEO_SESSION lines
        throughput_vals = re.findall(
            r"TX_VIDEO_SESSION\(\d+,\d+\):\s+throughput\s+([\d.]+)\s+Mb/s", full_output)
        throughput_mbps = float(throughput_vals[-1]) if throughput_vals else 0.0

        # Extract actual FPS from MTL
        fps_vals = re.findall(
            r"TX_VIDEO_SESSION\(\d+,\d+:\S+\):\s+fps\s+([\d.]+)", full_output)
        actual_fps = float(fps_vals[-1]) if fps_vals else 0.0

        # Calculate theoretical bandwidth from the transmitted (effective) size
        bpp = FMT_BPP.get(fmt, 20)
        theoretical_bw = (out_w * out_h * fps * bpp) / 1e6  # Mbps

        # Verify the ST2110-20 wire format matches the configured pixel format
        # (this is what distinguishes a real 10-bit from a real 12-bit stream).
        expected_transport = EXPECTED_TRANSPORT.get(fmt, "")
        transport_ok = ""
        if expected_transport:
            m = re.search(r"transport fmt\s+(\S+)|output fmt:\s*(\S+)", full_output)
            if m:
                transport_ok = "YES" if re.search(expected_transport, full_output,
                                                  re.IGNORECASE) else "NO"
        if transport_ok == "NO":
            status = "FAIL"

        # Build condensed log: keep key lines for format/frame extraction
        key_patterns = [
            "dvledtx initializ", "Config loaded", "Video:",
            "st20_get_converter", "st20p_tx_create", "ffmpeg_tx opened",
            "opened '", "Session Manager", "dvledtx started",
            "thread stopped", "shared thread started", "TX_VIDEO_SESSION",
            "dvledtx shutdown", "Stopping", "Shutdown reason",
            "NIC[", "x11grab", "scale",
        ]
        key_lines = []
        for line in full_output.split("\n"):
            if any(pat in line for pat in key_patterns):
                key_lines.append(line.strip())
        condensed = "\n".join(key_lines)

        result = {
            "tc_id": tc_id,
            "description": description,
            "test_type": "Functional",
            "status": status,
            "log_output": condensed[:15000],
            "exit_code": exit_code,
            "matched_pattern": "dvledtx started + fps/frames",
            "duration_ms": duration_ms,
            "throughput_mbps": throughput_mbps,
            "actual_fps": actual_fps,
            "theoretical_bw_mbps": theoretical_bw,
            "resolution": f"{width}x{height}",
            "tx_resolution": f"{out_w}x{out_h}",
            "scaled": "YES" if scale else "NO",
            "configured_fps": fps,
            "pixel_format": fmt,
            "bit_depth": FMT_BIT_DEPTH.get(fmt, ""),
            "session_count": session_count,
            "nic_count": nic_count,
            "input_mode": input_mode or "file",
            "expected_transport": expected_transport,
            "transport_ok": transport_ok,
            "source_video": os.path.basename(video_path) if video_path else "",
            "source_fmt": _video_src_fmt(video_path),
        }
        self.results.append(result)
        sym = "+" if status == "PASS" else "X"
        bw_str = f" [{throughput_mbps:.0f} Mbps, {actual_fps:.1f} fps]" if throughput_mbps else ""
        print(f"  [{sym}] {tc_id}: {description} [{status}]{bw_str}")
        return result


# ---------------------------------------------------------------------------
# Test case definitions
# ---------------------------------------------------------------------------
def run_config_parsing_tests(runner):
    """CP: Config Parsing tests."""
    print("\n=== Config Parsing (CP) ===")

    nonexistent = os.path.join(os.getcwd(), "dvledtx_no_such_file_xyz_99.json")
    runner.run_config_test("CP-N01", "Config file does not exist", "Negative",
        nonexistent, r"Cannot open config file")

    empty_path = os.path.join(runner.tmp_dir, "empty.json")
    open(empty_path, "w").close()
    runner.run_config_test("CP-N02", "Config file is empty (0 bytes)", "Negative",
        empty_path, r"parse|Failed")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"}}')
    runner.run_config_test("CP-N03", "Missing tx_sessions key", "Negative",
        cfg, r"tx_sessions.*not found")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[]}')
    runner.run_config_test("CP-N04", "Empty tx_sessions array", "Negative",
        cfg, r"No tx_sessions found")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"payload_type":96,"crop":{"x":0,"y":0,"w":1920,"h":1080}}]}')
    runner.run_config_test("CP-N05", "Missing udp_port in session", "Negative",
        cfg, r"udp_port not set or invalid")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":0,"payload_type":96,"crop":{"x":0,"y":0,"w":1920,"h":1080}}]}')
    runner.run_config_test("CP-N06", "udp_port = 0", "Negative",
        cfg, r"udp_port not set or invalid")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":70000,"payload_type":96,"crop":{"x":0,"y":0,"w":1920,"h":1080}}]}')
    runner.run_config_test("CP-N07", "udp_port = 70000 (> 65535)", "Negative",
        cfg, r"udp_port not set or invalid")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":-1,"payload_type":96,"crop":{"x":0,"y":0,"w":1920,"h":1080}}]}')
    runner.run_config_test("CP-N08", "udp_port = -1 (negative)", "Negative",
        cfg, r"udp_port not set or invalid")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080},"tx_video":{"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":20000,"crop":{"x":0,"y":0,"w":1920,"h":1080}}]}')
    runner.run_config_test("CP-P04", "payload_type omitted defaults to 96", "Positive",
        cfg, r"pt=96|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080},"tx_video":{"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":20000,"payload_type":0,"crop":{"x":0,"y":0,"w":1920,"h":1080}}]}')
    runner.run_config_test("CP-P05", "payload_type = 0 falls back to default 96", "Positive",
        cfg, r"pt=96|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080},"tx_video":{"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":20000,"payload_type":256,"crop":{"x":0,"y":0,"w":1920,"h":1080}}]}')
    runner.run_config_test("CP-P06", "payload_type = 256 (>255) falls back to default 96",
        "Positive", cfg, r"pt=96|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":20000,"payload_type":96}]}')
    runner.run_config_test("CP-N12", "Missing crop object in session", "Negative",
        cfg, r"crop.*required")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":20000,"payload_type":96,"crop":{"y":0,"w":1920,"h":1080}}]}')
    runner.run_config_test("CP-N13", "crop.x missing", "Negative",
        cfg, r"crop.*x.*required")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":20000,"payload_type":96,"crop":{"x":0,"w":1920,"h":1080}}]}')
    runner.run_config_test("CP-N14", "crop.y missing", "Negative",
        cfg, r"crop.*y.*required")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":20000,"payload_type":96,"crop":{"x":0,"y":0,"w":0,"h":1080}}]}')
    runner.run_config_test("CP-N15", "crop.w = 0", "Negative",
        cfg, r"crop.*w.*required|crop values must be positive")

    cfg = runner._make_config(raw_json='{"interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":20000,"payload_type":96,"crop":{"x":0,"y":0,"w":1920,"h":0}}]}')
    runner.run_config_test("CP-N16", "crop.h = 0", "Negative",
        cfg, r"crop.*h.*required|crop values must be positive")

    sessions_9 = [{"udp_port": 20000 + i * 2, "payload_type": 96,
                   "crop": {"x": 0, "y": 0, "w": 3840, "h": 2160}} for i in range(9)]
    cfg = runner._make_config(sessions=sessions_9, no_video=True)
    runner.run_config_test("CP-N17", "More than 8 sessions (capped at 8)", "Positive",
        cfg, r"Session [78]|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(no_video=True)
    runner.run_config_test("CP-P01", "Parse valid 1-session config", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    sessions_3 = [
        {"udp_port": 20000, "payload_type": 96, "crop": {"x": 0, "y": 0, "w": 1280, "h": 2160}},
        {"udp_port": 20002, "payload_type": 96, "crop": {"x": 1280, "y": 0, "w": 1280, "h": 2160}},
        {"udp_port": 20004, "payload_type": 96, "crop": {"x": 2560, "y": 0, "w": 1280, "h": 2160}},
    ]
    cfg = runner._make_config(sessions=sessions_3, no_video=True)
    runner.run_config_test("CP-P02", "Parse valid 3-session config", "Positive",
        cfg, r"Session [012]|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(log_file=os.path.join(runner.tmp_dir, "test_log.log"), no_video=True)
    runner.run_config_test("CP-P03", "Parse config with log_file", "Positive",
        cfg, r"initializ|TIMEOUT", expect_pass=True)


def run_config_validation_tests(runner):
    """CV: Config Validation tests."""
    print("\n=== Config Validation (CV) ===")

    cfg = runner._make_config(overrides={"interfaces.0.name": ""}, no_video=True)
    runner.run_config_test("CV-N01", "Empty interface_name", "Negative",
        cfg, r"interfaces\[0\]\.name is required")

    cfg = runner._make_config(overrides={"interfaces.0.dip": ""}, no_video=True)
    runner.run_config_test("CV-N02", "Empty interface DIP", "Negative",
        cfg, r"interfaces\[0\]\.dip is required")

    cfg = runner._make_config(overrides={"interfaces.0.sip": "not.an.ip"}, no_video=True)
    runner.run_config_test("CV-N03", "Invalid SIP format", "Negative",
        cfg, r"Invalid source IP address")

    cfg = runner._make_config(overrides={"interfaces.0.dip": "abc.def.ghi.jkl"}, no_video=True)
    runner.run_config_test("CV-N04", "Invalid DIP format", "Negative",
        cfg, r"Invalid destination IP address")

    cfg = runner._make_config(overrides={"interfaces.0.dip": "192.168.1.1"}, no_video=True)
    runner.run_config_test("CV-N05", "Unicast DIP (not multicast)", "Negative",
        cfg, r"not a valid multicast address")

    cfg = runner._make_config(overrides={"interfaces.0.dip": "240.0.0.1"}, no_video=True)
    runner.run_config_test("CV-N06", "DIP above multicast range (240.x)", "Negative",
        cfg, r"not a valid multicast")

    cfg = runner._make_config(overrides={"interfaces.0.name": "eth0"}, no_video=True)
    runner.run_config_test("CV-N07", "Invalid PCI BDF format (eth0)", "Negative",
        cfg, r"Invalid PCI BDF format")

    cfg = runner._make_config(overrides={"interfaces.0.name": "ZZZZ:ZZ:ZZ.Z"}, no_video=True)
    runner.run_config_test("CV-N08", "Invalid PCI BDF (non-hex chars)", "Negative",
        cfg, r"Invalid PCI BDF format")

    cfg = runner._make_config(overrides={"video.width": 0}, no_video=True)
    runner.run_config_test("CV-N09", "Width = 0", "Negative",
        cfg, r"width/height must be non-zero")

    cfg = runner._make_config(overrides={"video.height": 0}, no_video=True)
    runner.run_config_test("CV-N10", "Height = 0", "Negative",
        cfg, r"width/height must be non-zero")

    cfg = runner._make_config(overrides={"video.width": 4000}, no_video=True)
    runner.run_config_test("CV-N11", "Width > 3840", "Negative",
        cfg, r"exceeds maximum 3840x2160")

    cfg = runner._make_config(overrides={"video.height": 2200}, no_video=True)
    runner.run_config_test("CV-N12", "Height > 2160", "Negative",
        cfg, r"exceeds maximum 3840x2160")

    cfg = runner._make_config(overrides={"video.width": 7680, "video.height": 4320}, no_video=True)
    runner.run_config_test("CV-N13", "7680x4320 exceeds 3840x2160 cap", "Negative",
        cfg, r"exceeds maximum 3840x2160")

    cfg = runner._make_config(overrides={"video.width": 4096, "video.height": 2160}, no_video=True)
    runner.run_config_test("CV-N14", "4096x2160 exceeds 3840x2160 cap", "Negative",
        cfg, r"exceeds maximum 3840x2160")

    cfg = runner._make_config(overrides={"video.width": 1919}, no_video=True)
    runner.run_config_test("CV-N15", "Odd width (1919)", "Negative",
        cfg, r"width.*must be even")

    cfg = runner._make_config(overrides={"video.fps": 24}, no_video=True)
    runner.run_config_test("CV-N16", "Unsupported FPS (24)", "Negative",
        cfg, r"unsupported fps 24")

    cfg = runner._make_config(overrides={"video.fps": 0}, no_video=True)
    runner.run_config_test("CV-N17", "Unsupported FPS (0)", "Negative",
        cfg, r"unsupported fps 0")

    cfg = runner._make_config(overrides={"video.fmt": "rgb24"}, no_video=True)
    runner.run_config_test("CV-N18", "Unsupported pixel format (rgb24)", "Negative",
        cfg, r"unsupported pixel format.*rgb24")

    cfg_data = {"interfaces": [{"name": DEFAULT_BDF, "sip": DEFAULT_SIP, "dip": DEFAULT_DIP}],
                "video": {"width": 1920, "height": 1080,
                          "tx_url": "/nonexistent/video.mp4"},
                "tx_video": {"fps": 30, "fmt": "yuv422p10le"},
                "tx_sessions": [{"udp_port": 20000, "payload_type": 96,
                                 "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}]}
    cfg = os.path.join(runner.tmp_dir, "cv_n19.json")
    with open(cfg, "w") as f:
        json.dump(cfg_data, f)
    runner.run_config_test("CV-N19", "tx_url file does not exist", "Negative",
        cfg, r"video source file not found")

    sessions = [{"udp_port": 80, "payload_type": 96,
                 "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-N20", "UDP port < 1024 (privileged)", "Negative",
        cfg, r"udp_port.*privileged range")

    sessions = [{"udp_port": 20000, "payload_type": 95,
                 "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-N21", "Payload type < 96", "Negative",
        cfg, r"payload_type.*out of range")

    sessions = [{"udp_port": 20000, "payload_type": 128,
                 "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-N22", "Payload type > 127", "Negative",
        cfg, r"payload_type.*out of range")

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 100, "y": 0, "w": 3840, "h": 2160}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-N23", "crop_x + crop_w > width", "Negative",
        cfg, r"crop.*exceeds.*width")

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 0, "y": 100, "w": 3840, "h": 2160}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-N24", "crop_y + crop_h > height", "Negative",
        cfg, r"crop.*exceeds.*height")

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 0, "y": 0, "w": 641, "h": 2160}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-N25", "Odd crop width (641)", "Negative",
        cfg, r"crop width.*must be even")

    sessions = [
        {"udp_port": 20000, "payload_type": 96, "crop": {"x": 0, "y": 0, "w": 1920, "h": 2160}},
        {"udp_port": 20000, "payload_type": 96, "crop": {"x": 1920, "y": 0, "w": 1920, "h": 2160}},
    ]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-N26", "Duplicate UDP ports", "Negative",
        cfg, r"duplicate udp_port")

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 1, "y": 0, "w": 1918, "h": 2160}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-N27", "crop_x misaligned for yuv422", "Negative",
        cfg, r"crop_x.*must be a multiple of 2")

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 2147483647, "y": 0, "w": 2, "h": 2160}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-N28", "Crop overflow (uint32 cast)", "Negative",
        cfg, r"crop.*exceeds")

    # Positive tests
    cfg = runner._make_config(no_video=True)
    runner.run_config_test("CV-P01", "Valid config passes validation", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(overrides={"interfaces.0.sip": ""}, no_video=True)
    runner.run_config_test("CV-P02", "SIP field empty (DHCP mode)", "Positive",
        cfg, r"Config loaded|DHCP|initializ", expect_pass=True)

    cfg = runner._make_config(no_video=True)
    runner.run_config_test("CV-P03", "DIP in 239.x.x.x (admin-scoped)", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(overrides={"interfaces.0.dip": "224.0.1.1"}, no_video=True)
    runner.run_config_test("CV-P04", "DIP in 224.x.x.x (with warning)", "Positive",
        cfg, r"outside the administratively-scoped|Config loaded|initializ", expect_pass=True)

    for fps in SUPPORTED_FPS:
        cfg = runner._make_config(overrides={"video.fps": fps}, no_video=True)
        runner.run_config_test(f"CV-P05-{fps}fps", f"Supported FPS = {fps}", "Positive",
            cfg, r"Config loaded|initializ", expect_pass=True)

    for fmt in SUPPORTED_FORMATS:
        cfg = runner._make_config(overrides={"video.fmt": fmt}, no_video=True)
        runner.run_config_test(f"CV-P06-{fmt}", f"Supported format: {fmt}", "Positive",
            cfg, r"Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(overrides={"video.width": 3840, "video.height": 2160}, no_video=True)
    runner.run_config_test("CV-P07", "Maximum resolution 3840x2160", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    sessions = [{"udp_port": 1024, "payload_type": 96,
                 "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-P08", "UDP port = 1024 (minimum)", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    sessions = [{"udp_port": 65535, "payload_type": 96,
                 "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-P09", "UDP port = 65535 (maximum)", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-P10", "Payload type = 96 (boundary)", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    sessions = [{"udp_port": 20000, "payload_type": 127,
                 "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-P11", "Payload type = 127 (boundary)", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    sessions = [
        {"udp_port": 20000, "payload_type": 96, "crop": {"x": 0, "y": 0, "w": 1280, "h": 2160}},
        {"udp_port": 20002, "payload_type": 96, "crop": {"x": 1280, "y": 0, "w": 1280, "h": 2160}},
        {"udp_port": 20004, "payload_type": 96, "crop": {"x": 2560, "y": 0, "w": 1280, "h": 2160}},
    ]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("CV-P12", "3 tiled sessions non-overlapping", "Positive",
        cfg, r"Session [012]|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(no_video=True)
    runner.run_config_test("CV-P13", "Valid PCI BDF format", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    for mode in SUPPORTED_INPUT_MODES:
        kwargs = {"input_mode": mode}
        if mode == "screen_capture":
            kwargs["screen_input"] = DEFAULT_SCREEN_INPUT
        cfg = runner._make_config(no_video=True, **kwargs)
        runner.run_config_test(f"CV-P19-{mode}", f"Supported input_mode: {mode}", "Positive",
            cfg, r"mode=" + mode + r"|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(overrides={"tx_video.fmt": ""}, no_video=True)
    runner.run_config_test("CV-N29", "Empty fmt string rejected", "Negative",
        cfg, r"Unsupported pixel format")

    cfg = runner._make_config(raw_json=json.dumps({
        "interfaces": [{"name": DEFAULT_BDF, "sip": DEFAULT_SIP, "dip": DEFAULT_DIP}],
        "video": {"width": 1920, "height": 1080},
        "tx_sessions": [{"udp_port": 20000, "payload_type": 96,
                         "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}],
    }))
    runner.run_config_test("CV-N30", "Missing tx_video block (no fps/fmt)", "Negative",
        cfg, r"unsupported fps 0")


def run_scaling_tests(runner, skip_tx=False, video_dir=None):
    """SL: Scaling (tx_video.scale_width / scale_height) tests."""
    print("\n=== Scaling (SL) ===")

    # --- Negative: invalid scale configuration ---
    cfg = runner._make_config(overrides={"tx_video.scale_width": 2560}, no_video=True)
    runner.run_config_test("SL-N01", "scale_width only (no scale_height) rejected", "Negative",
        cfg, r"scale_width and scale_height must both be set")

    cfg = runner._make_config(overrides={"tx_video.scale_height": 1440}, no_video=True)
    runner.run_config_test("SL-N02", "scale_height only (no scale_width) rejected", "Negative",
        cfg, r"scale_width and scale_height must both be set")

    cfg = runner._make_config(scale=(4096, 2160), no_video=True)
    runner.run_config_test("SL-N03", "scale 4096x2160 exceeds 3840x2160 rejected", "Negative",
        cfg, r"scale resolution.*exceeds maximum 3840x2160")

    cfg = runner._make_config(scale=(7680, 4320), no_video=True)
    runner.run_config_test("SL-N04", "scale 7680x4320 (8K) rejected", "Negative",
        cfg, r"scale resolution.*exceeds maximum 3840x2160")

    cfg = runner._make_config(scale=(1921, 1080), no_video=True)
    runner.run_config_test("SL-N05", "odd scale_width rejected (yuv422p10le)", "Negative",
        cfg, r"scale_width.*must be a multiple of 2")

    cfg = runner._make_config(scale=(1920, 1081),
        overrides={"tx_video.fmt": "yuv420"}, no_video=True)
    runner.run_config_test("SL-N06", "odd scale_height rejected (yuv420 4:2:0)", "Negative",
        cfg, r"scale_height.*must be a multiple of 2")

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 0, "y": 0, "w": 3840, "h": 2160}}]
    cfg = runner._make_config(scale=(1920, 1080), sessions=sessions, no_video=True)
    runner.run_config_test("SL-N07", "crop exceeds scaled width rejected", "Negative",
        cfg, r"crop x=\d+ \+ w=\d+ = \d+ exceeds effective width")

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 0, "y": 0, "w": 1920, "h": 2160}}]
    cfg = runner._make_config(scale=(1920, 1080), sessions=sessions, no_video=True)
    runner.run_config_test("SL-N08", "crop exceeds scaled height rejected", "Negative",
        cfg, r"crop y=\d+ \+ h=\d+ = \d+ exceeds effective height")

    # --- Positive: valid up/down scaling for every supported format ---
    for i, (label, sw, sh, dw, dh) in enumerate(SCALE_CASES, 1):
        sessions = [{"udp_port": 20000, "payload_type": 96,
                     "crop": {"x": 0, "y": 0, "w": dw, "h": dh}}]
        cfg = runner._make_config(
            overrides={"video.width": sw, "video.height": sh},
            scale=(dw, dh), sessions=sessions, no_video=True)
        runner.run_config_test(f"SL-P{i:02d}", f"Valid {label}", "Positive",
            cfg, r"scale \d+x\d+|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(scale=(3840, 2160),
        overrides={"video.width": 1920, "video.height": 1080}, no_video=True)
    runner.run_config_test("SL-P05", "Scale to maximum 3840x2160 passes", "Positive",
        cfg, r"scale 3840x2160|Config loaded|initializ", expect_pass=True)

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 0, "y": 0, "w": 1280, "h": 720}}]
    cfg = runner._make_config(scale=(1920, 1080), sessions=sessions, no_video=True)
    runner.run_config_test("SL-P06", "Crop within scaled dimensions passes", "Positive",
        cfg, r"scale 1920x1080|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(overrides={"tx_video.scale_width": 0,
                                         "tx_video.scale_height": 0}, no_video=True)
    runner.run_config_test("SL-P07", "No scaling (0,0) passes", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    # Scaled output tiled across 3 sessions (scaling + multi-session combined)
    sessions = [
        {"udp_port": 20000 + i * 2, "payload_type": 96,
         "crop": {"x": i * 640, "y": 0, "w": 640, "h": 1080}} for i in range(3)
    ]
    cfg = runner._make_config(scale=(1920, 1080), sessions=sessions, no_video=True)
    runner.run_config_test("SL-P08", "Scaled output tiled across 3 sessions", "Positive",
        cfg, r"scale 1920x1080|Session [012]|Config loaded|initializ", expect_pass=True)

    # Scaling with every supported pixel format (10-bit and 12-bit)
    for fmt in SUPPORTED_FORMATS:
        sessions = [{"udp_port": 20000, "payload_type": 96,
                     "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}]
        cfg = runner._make_config(overrides={"tx_video.fmt": fmt},
            scale=(1920, 1080), sessions=sessions, no_video=True)
        runner.run_config_test(f"SL-P09-{fmt}", f"Scale 4K->1080p with {fmt}", "Positive",
            cfg, r"scale 1920x1080|Config loaded|initializ", expect_pass=True)

    # --- Functional: actually transmit a scaled stream ---
    if skip_tx or not (runner.video or (video_dir and os.path.isdir(video_dir))):
        print("  [SKIP] SL TX tests skipped (--skip-tx or no video file)")
        return

    for fmt in ["yuv422p10le", "yuv422p12le"]:
        video = _find_video(video_dir, fmt, "4K", 30) if video_dir else None
        if video is None and fmt == "yuv422p10le":
            video = runner.video
        if not video:
            print(f"  [SKIP] SL TX ({fmt}): no 4K source video")
            continue
        runner.run_tx_test_full(
            f"SL-TX01-{fmt}", f"TX scaled 4K->1080p ({fmt})", fmt, 1,
            3840, 2160, 30, video, timeout=50, scale=(1920, 1080))
        runner.run_tx_test_full(
            f"SL-TX02-3ses-{fmt}", f"TX scaled 4K->1080p, 3 tiled sessions ({fmt})",
            fmt, 3, 3840, 2160, 30, video, timeout=60, scale=(1920, 1080))


def run_screen_capture_tests(runner, skip_tx=False):
    """SC: Screen capture (input_mode=screen_capture / x11grab) tests."""
    print("\n=== Screen Capture (SC) ===")

    cfg = runner._make_config(input_mode="screen_capture",
        screen_input=DEFAULT_SCREEN_INPUT, no_video=True)
    runner.run_config_test("SC-P01", "input_mode=screen_capture with screen_input", "Positive",
        cfg, r"mode=screen_capture|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(input_mode="screen_capture", no_video=True)
    runner.run_config_test("SC-P02", "screen_input omitted defaults to :0.0+0,0", "Positive",
        cfg, r"source=:0\.0\+0,0|mode=screen_capture|Config loaded|initializ",
        expect_pass=True)

    cfg = runner._make_config(input_mode="screen_capture", screen_input=":99.0+0,0",
        overrides={"video.width": 1920, "video.height": 1080},
        sessions=[{"udp_port": 20000, "payload_type": 96,
                   "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}], no_video=True)
    runner.run_config_test("SC-P03", "Virtual display :99.0 screen_input accepted", "Positive",
        cfg, r"source=:99\.0|mode=screen_capture|Config loaded|initializ", expect_pass=True)

    sessions = [
        {"udp_port": 20000 + i * 2, "payload_type": 96,
         "crop": {"x": i * 640, "y": 0, "w": 640, "h": 1080}} for i in range(3)
    ]
    cfg = runner._make_config(input_mode="screen_capture", screen_input=":0.0+0,0",
        overrides={"video.width": 1920, "video.height": 1080},
        sessions=sessions, no_video=True)
    runner.run_config_test("SC-P04", "Screen capture tiled across 3 sessions", "Positive",
        cfg, r"mode=screen_capture|Session [012]|Config loaded|initializ", expect_pass=True)

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 0, "y": 0, "w": 1280, "h": 720}}]
    cfg = runner._make_config(input_mode="screen_capture", screen_input=":0.0+0,0",
        overrides={"video.width": 1920, "video.height": 1080},
        scale=(1280, 720), sessions=sessions, no_video=True)
    runner.run_config_test("SC-P05", "Screen capture + scaling 1080p->720p", "Positive",
        cfg, r"scale 1280x720|mode=screen_capture|Config loaded|initializ", expect_pass=True)

    # tx_url is ignored in screen-capture mode (no "file not found" error)
    cfg = runner._make_config(input_mode="screen_capture", screen_input=":0.0+0,0",
        overrides={"video.tx_url": "/nonexistent/video.mp4"}, no_video=True)
    runner.run_config_test("SC-P06", "tx_url ignored when input_mode=screen_capture",
        "Positive", cfg, r"mode=screen_capture|Config loaded|initializ", expect_pass=True)

    for fmt in SUPPORTED_FORMATS:
        cfg = runner._make_config(input_mode="screen_capture", screen_input=":0.0+0,0",
            overrides={"tx_video.fmt": fmt}, no_video=True)
        runner.run_config_test(f"SC-P07-{fmt}", f"Screen capture with {fmt}", "Positive",
            cfg, r"mode=screen_capture|Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(input_mode="camera", no_video=True)
    runner.run_config_test("SC-N01", "Unsupported input_mode 'camera' rejected", "Negative",
        cfg, r"unsupported input_mode 'camera'")

    cfg = runner._make_config(input_mode="SCREEN_CAPTURE", no_video=True)
    runner.run_config_test("SC-N02", "input_mode is case-sensitive (SCREEN_CAPTURE)", "Negative",
        cfg, r"unsupported input_mode")

    cfg = runner._make_config(input_mode="screen_capture", screen_input=":0.0+0,0",
        sessions=[{"udp_port": 20000, "payload_type": 96,
                   "crop": {"x": 0, "y": 0, "w": 4096, "h": 2160}}], no_video=True)
    runner.run_config_test("SC-N03", "Screen capture crop exceeds capture size", "Negative",
        cfg, r"exceeds effective width|exceeds maximum")

    # --- Functional: capture the real X display (needs DISPLAY + x11grab) ---
    display = os.environ.get("DISPLAY", "")
    if skip_tx or not display:
        print("  [SKIP] SC TX test (no DISPLAY or --skip-tx)")
        return
    runner.run_tx_test_full(
        "SC-TX01", f"TX screen capture from {display} (1080p yuv422p10le)",
        "yuv422p10le", 1, 1920, 1080, 30, None, timeout=50,
        input_mode="screen_capture", screen_input=f"{display}+0,0")


def run_ptp_tests(runner):
    """PTP: PTP timing block tests."""
    print("\n=== PTP Timing (PTP) ===")

    cfg = runner._make_config(no_video=True)
    runner.run_config_test("PTP-P01", "PTP disabled by default (TSC pacing)", "Positive",
        cfg, r"PTP disabled", expect_pass=True)

    cfg = runner._make_config(ptp={"enable": True, "pi": False, "unicast": False},
        no_video=True)
    runner.run_config_test("PTP-P02", "ptp.enable=true enables built-in PTP client",
        "Positive", cfg, r"PTP enabled: pi_controller=0 unicast_delay_req=0",
        expect_pass=True)

    cfg = runner._make_config(ptp={"enable": True, "pi": True, "unicast": False},
        no_video=True)
    runner.run_config_test("PTP-P03", "ptp.pi=true selects PI controller", "Positive",
        cfg, r"PTP enabled: pi_controller=1", expect_pass=True)

    cfg = runner._make_config(ptp={"enable": True, "pi": False, "unicast": True},
        no_video=True)
    runner.run_config_test("PTP-P04", "ptp.unicast=true selects unicast DELAY_REQ",
        "Positive", cfg, r"unicast_delay_req=1", expect_pass=True)

    cfg = runner._make_config(ptp={"enable": False}, no_video=True)
    runner.run_config_test("PTP-P05", "ptp.enable=false keeps PTP disabled", "Positive",
        cfg, r"PTP disabled", expect_pass=True)


def run_command_line_tests(runner):
    """CL: Command Line tests."""
    print("\n=== Command Line (CL) ===")

    cfg = runner._make_config(no_video=True)
    runner.run_cli_test("CL-P01", "dvledtx --config valid_config.json", "Positive",
        [runner.binary, "--config", cfg], r"initializ", expect_pass=True)

    cfg = runner._make_config(no_video=True)
    runner.run_cli_test("CL-P02", "dvledtx -C short option", "Positive",
        [runner.binary, "-C", cfg], r"initializ", expect_pass=True)

    runner.run_cli_test("CL-P03", "dvledtx -v prints version", "Positive",
        [runner.binary, "-v"], VERSION, expect_pass=True)

    runner.run_cli_test("CL-P04", "dvledtx --version prints version", "Positive",
        [runner.binary, "--version"], VERSION, expect_pass=True)

    runner.run_cli_test("CL-N01", "No arguments -> usage error", "Negative",
        [runner.binary], r"Usage")

    runner.run_cli_test("CL-N02", "Unknown option -> error", "Negative",
        [runner.binary, "--unknown-xyz"], r"Usage|unrecognized")

    runner.run_cli_test("CL-N03", "--config without argument", "Negative",
        [runner.binary, "--config"], r"Usage|requires an argument")

    nonexistent_cfg = os.path.join(os.getcwd(), "dvledtx_no_such_file_xyz.json")
    runner.run_cli_test("CL-N04", "--config nonexistent file", "Negative",
        [runner.binary, "--config", nonexistent_cfg],
        r"Cannot open config file|Failed")

    cfg = runner._make_config(no_video=True)
    runner.run_cli_test("CL-N05", "--test-time invalid value", "Negative",
        [runner.binary, "--config", cfg, "--test-time", "abc"], r"Invalid --test-time")

    cfg = runner._make_config(no_video=True)
    runner.run_cli_test("CL-N06", "--test-time 0 (out of range)", "Negative",
        [runner.binary, "--config", cfg, "--test-time", "0"], r"Invalid --test-time")

    real_cfg = runner._make_config(no_video=True)
    link_path = os.path.join(runner.tmp_dir, "symlink_config.json")
    try:
        os.symlink(real_cfg, link_path)
        runner.run_cli_test("CL-N07", "Symlinked config file rejected", "Negative",
            [runner.binary, "--config", link_path], r"symbolic link.*rejected")
    except OSError:
        pass


def run_log_file_tests(runner):
    """LF: Log File tests."""
    print("\n=== Log File (LF) ===")

    cfg = runner._make_config(log_file="/etc/cron.d/evil", no_video=True)
    runner.run_config_test("LF-N01", "Log path traversal /etc/cron.d/evil", "Negative",
        cfg, r"Log file path.*rejected|not under allowed", expect_pass=True)

    cfg = runner._make_config(log_file="/etc/passwd", no_video=True)
    runner.run_config_test("LF-N02", "Log path to /etc/passwd", "Negative",
        cfg, r"Log file path.*rejected|not under allowed", expect_pass=True)

    cfg = runner._make_config(log_file="../../../etc/shadow", no_video=True)
    runner.run_config_test("LF-N03", "Log path with .. traversal", "Negative",
        cfg, r"Log file path.*rejected|not under allowed", expect_pass=True)

    cfg = runner._make_config(log_file="/nonexistent_dir_xyz/app.log", no_video=True)
    runner.run_config_test("LF-N04", "Log file in nonexistent directory", "Negative",
        cfg, r"rejected|not under allowed|Could not open", expect_pass=True)

    cfg = runner._make_config(log_file="", no_video=True)
    runner.run_config_test("LF-N05", "Empty log_file path", "Positive",
        cfg, r"initializ", expect_pass=True)

    cfg = runner._make_config(log_file=os.path.join(runner.tmp_dir, "test.log"), no_video=True)
    runner.run_config_test("LF-P01", "Log file in CWD subdirectory", "Positive",
        cfg, r"initializ|TIMEOUT", expect_pass=True)

    cfg = runner._make_config(no_log_file=True, no_video=True)
    runner.run_config_test("LF-P02", "No log_file in config (console)", "Positive",
        cfg, r"initializ", expect_pass=True)

    cwd_log = os.path.join(os.getcwd(), "dvledtx_test.log")
    cfg = runner._make_config(log_file=cwd_log, no_video=True)
    runner.run_config_test("LF-P03", "Log file in CWD", "Positive",
        cfg, r"initializ|TIMEOUT", expect_pass=True)
    try:
        os.unlink("dvledtx_test.log")
    except OSError:
        pass


def run_security_tests(runner):
    """SEC: Security tests."""
    print("\n=== Security (SEC) ===")

    cfg = runner._make_config(log_file="/etc/cron.d/evil", no_video=True)
    runner.run_config_test("SEC-N01", "Log path /etc/cron.d/evil rejected", "Negative",
        cfg, r"Log file path.*rejected", expect_pass=True)

    cfg = runner._make_config(log_file="../../../etc/shadow", no_video=True)
    runner.run_config_test("SEC-N02", "Log path with .. rejected", "Negative",
        cfg, r"Log file path.*rejected", expect_pass=True)

    cfg = runner._make_config(overrides={"interfaces.0.dip": "10.0.0.1"}, no_video=True)
    runner.run_config_test("SEC-N03", "Unicast DIP rejected", "Negative",
        cfg, r"not a valid multicast address")

    cfg = runner._make_config(overrides={"interfaces.0.name": "ZZZZ:ZZ:ZZ.Z"}, no_video=True)
    runner.run_config_test("SEC-N04", "Invalid BDF format rejected", "Negative",
        cfg, r"Invalid PCI BDF format")

    sessions = [{"udp_port": 20000, "payload_type": 96,
                 "crop": {"x": 2147483647, "y": 0, "w": 2, "h": 2160}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("SEC-N05", "Crop overflow prevented", "Negative",
        cfg, r"crop.*exceeds")

    real_cfg = runner._make_config(no_video=True)
    link_path = os.path.join(runner.tmp_dir, "sec_symlink.json")
    try:
        os.symlink(real_cfg, link_path)
        runner.run_config_test("SEC-N06", "Symlinked config rejected", "Negative",
            link_path, r"symbolic link.*rejected")
    except OSError:
        pass

    cfg = runner._make_config(overrides={"video.width": 4096, "video.height": 2160}, no_video=True)
    runner.run_config_test("SEC-N07", "4096x2160 exceeds 3840x2160 cap", "Negative",
        cfg, r"exceeds maximum 3840x2160")

    sessions = [{"udp_port": 80, "payload_type": 96,
                 "crop": {"x": 0, "y": 0, "w": 3840, "h": 2160}}]
    cfg = runner._make_config(sessions=sessions, no_video=True)
    runner.run_config_test("SEC-N08", "Privileged UDP port rejected", "Negative",
        cfg, r"udp_port.*privileged")

    cfg = runner._make_config(no_video=True)
    runner.run_config_test("SEC-P01", "Multicast DIP enforced", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    cfg = runner._make_config(no_video=True)
    runner.run_config_test("SEC-P02", "PCI BDF format validated", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    escape_cfg = os.path.join(runner.tmp_dir, "escape_test.json")
    with open(escape_cfg, "w") as f:
        f.write('{"log_file":"path\\\\\\":/etc/passwd","interfaces":[{"name":"0000:06:00.0","sip":"192.168.50.29","dip":"239.168.85.20"}],"video":{"width":1920,"height":1080,"fps":30,"fmt":"yuv422p10le"},"tx_sessions":[{"udp_port":20000,"payload_type":96,"crop":{"x":0,"y":0,"w":1920,"h":1080}}]}')
    runner.run_config_test("SEC-P03", "JSON escape injection handled", "Positive",
        escape_cfg, r"Log file path.*rejected|initializ", expect_pass=True)


def run_multi_nic_tests(runner):
    """MN: Multiple NIC (nic_index / interfaces[]) tests."""
    print("\n=== Multi-NIC (MN) ===")

    two_nics = [
        {"name": DEFAULT_BDF,     "sip": DEFAULT_SIP,     "dip": DEFAULT_DIP},
        {"name": "0000:07:00.0",  "sip": "192.168.50.30", "dip": "239.168.85.21"},
    ]
    six_nics = [
        {"name": "0000:03:10.1", "sip": "192.168.50.29", "dip": "239.168.85.20"},
        {"name": "0000:03:10.3", "sip": "192.168.50.30", "dip": "239.168.85.21"},
        {"name": "0000:03:10.5", "sip": "192.168.50.31", "dip": "239.168.85.22"},
        {"name": "0000:03:10.7", "sip": "192.168.50.32", "dip": "239.168.85.23"},
        {"name": "0000:03:11.1", "sip": "192.168.50.33", "dip": "239.168.85.24"},
        {"name": "0000:03:11.3", "sip": "192.168.50.34", "dip": "239.168.85.25"},
    ]

    # --- Positive: parse configs that spread sessions across multiple NICs ---
    sessions_2nic = [
        {"nic_index": 0, "udp_port": 20000, "payload_type": 96,
         "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}},
        {"nic_index": 1, "udp_port": 20000, "payload_type": 96,
         "crop": {"x": 1920, "y": 0, "w": 1920, "h": 1080}},
    ]
    cfg = runner._make_config(interfaces=two_nics, sessions=sessions_2nic,
        overrides={"video.width": 3840, "video.height": 1080}, no_video=True)
    runner.run_config_test("MN-P01", "Parse config with 2 NICs, 1 session each", "Positive",
        cfg, r"2 NIC\(s\)|Config loaded|initializ", expect_pass=True)

    sessions_6nic = [
        {"nic_index": i, "udp_port": 20000 + i * 2, "payload_type": 96,
         "crop": {"x": i * 320, "y": 0, "w": 320, "h": 1080}}
        for i in range(6)
    ]
    cfg = runner._make_config(interfaces=six_nics, sessions=sessions_6nic,
        overrides={"video.width": 1920, "video.height": 1080}, no_video=True)
    runner.run_config_test("MN-P02", "Parse config with 6 NICs (matches production multi-nic layout)",
        "Positive", cfg, r"6 NIC\(s\)|Config loaded|initializ", expect_pass=True)

    sessions_same_nic = [
        {"udp_port": 20000, "payload_type": 96, "crop": {"x": 0, "y": 0, "w": 960, "h": 1080}},
        {"udp_port": 20002, "payload_type": 96, "crop": {"x": 960, "y": 0, "w": 960, "h": 1080}},
    ]
    cfg = runner._make_config(sessions=sessions_same_nic,
        overrides={"video.width": 1920, "video.height": 1080}, no_video=True)
    runner.run_config_test("MN-P03", "nic_index omitted defaults to NIC 0 (backward compat)",
        "Positive", cfg, r"Config loaded|initializ", expect_pass=True)

    sessions_dup_port_diff_nic = [
        {"nic_index": 0, "udp_port": 20000, "payload_type": 96,
         "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}},
        {"nic_index": 1, "udp_port": 20000, "payload_type": 96,
         "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}},
    ]
    cfg = runner._make_config(interfaces=two_nics, sessions=sessions_dup_port_diff_nic,
        overrides={"video.width": 1920, "video.height": 1080}, no_video=True)
    runner.run_config_test("MN-P04", "Same udp_port allowed on different NICs", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    sessions_last_nic = [
        {"nic_index": 1, "udp_port": 20000, "payload_type": 96,
         "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}},
    ]
    cfg = runner._make_config(interfaces=two_nics, sessions=sessions_last_nic,
        overrides={"video.width": 1920, "video.height": 1080}, no_video=True)
    runner.run_config_test("MN-P05", "nic_index at upper boundary (last interface)", "Positive",
        cfg, r"Config loaded|initializ", expect_pass=True)

    # --- Negative: invalid nic_index / interface configuration ---
    # Built with an explicit tx_video block (scale_width/scale_height/fps/fmt)
    # so validation reaches the nic_index/session checks instead of bailing
    # out early on missing fps/fmt (those only live under "tx_video").
    n01_json = json.dumps({
        "interfaces": [{"name": DEFAULT_BDF, "sip": DEFAULT_SIP, "dip": DEFAULT_DIP}],
        "video": {"width": 1920, "height": 1080},
        "tx_video": {"scale_width": 1920, "scale_height": 1080, "fps": 30, "fmt": "yuv422p10le"},
        "tx_sessions": [{"nic_index": 5, "udp_port": 20000, "payload_type": 96,
                         "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}],
    })
    cfg = runner._make_config(raw_json=n01_json)
    runner.run_config_test("MN-N01", "nic_index out of range (>= nic_count)", "Negative",
        cfg, r"nic_index \d+ is out of range")

    n02_json = json.dumps({
        "interfaces": [i for i in two_nics],
        "video": {"width": 3840, "height": 1080},
        "tx_video": {"scale_width": 3840, "scale_height": 1080, "fps": 30, "fmt": "yuv422p10le"},
        "tx_sessions": [
            {"nic_index": 0, "udp_port": 20000, "payload_type": 96,
             "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}},
            {"nic_index": 0, "udp_port": 20000, "payload_type": 96,
             "crop": {"x": 1920, "y": 0, "w": 1920, "h": 1080}},
        ],
    })
    cfg = runner._make_config(raw_json=n02_json)
    runner.run_config_test("MN-N02", "Duplicate udp_port on same NIC rejected", "Negative",
        cfg, r"duplicate udp_port")

    bad_interfaces = [
        {"name": DEFAULT_BDF,    "sip": DEFAULT_SIP,     "dip": DEFAULT_DIP},
        {"name": "not-a-bdf",    "sip": "192.168.50.30", "dip": "239.168.85.21"},
    ]
    sessions_two = [
        {"nic_index": 0, "udp_port": 20000, "payload_type": 96,
         "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}},
    ]
    cfg = runner._make_config(interfaces=bad_interfaces, sessions=sessions_two, no_video=True)
    runner.run_config_test("MN-N03", "Invalid PCI BDF on second interface[1]", "Negative",
        cfg, r"Invalid PCI BDF format")

    bad_dip_interfaces = [
        {"name": DEFAULT_BDF,    "sip": DEFAULT_SIP,     "dip": DEFAULT_DIP},
        {"name": "0000:07:00.0", "sip": "192.168.50.30", "dip": "10.0.0.1"},
    ]
    cfg = runner._make_config(interfaces=bad_dip_interfaces, sessions=sessions_two, no_video=True)
    runner.run_config_test("MN-N04", "Unicast DIP on second interface[1] rejected", "Negative",
        cfg, r"not a valid multicast address")

    sessions_negative_nic = [{"nic_index": -1, "udp_port": 20000, "payload_type": 96,
                              "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}]
    cfg = runner._make_config(sessions=sessions_negative_nic, no_video=True)
    runner.run_config_test("MN-P06", "Negative nic_index defaults to NIC 0 (parser clamps)",
        "Positive", cfg, r"Config loaded|initializ", expect_pass=True)

    eight_nics = [
        {"nic_index": i, "name": f"0000:03:1{i}.0", "sip": f"192.168.50.{29 + i}",
         "dip": f"239.168.85.{20 + i}"} for i in range(8)
    ]
    sessions_8nic = [
        {"nic_index": i, "udp_port": 20000 + i * 2, "payload_type": 96,
         "crop": {"x": i * 240, "y": 0, "w": 240, "h": 1080}} for i in range(8)
    ]
    cfg = runner._make_config(interfaces=eight_nics, sessions=sessions_8nic,
        overrides={"video.width": 1920, "video.height": 1080}, no_video=True)
    runner.run_config_test("MN-P07", "Parse config with 8 NICs (MTL maximum)", "Positive",
        cfg, r"8 NIC\(s\)|Config loaded|initializ", expect_pass=True)

    # Multiple NICs combined with scaling (scaled output tiled per NIC)
    sessions_scaled = [
        {"nic_index": i, "udp_port": 20000, "payload_type": 96,
         "crop": {"x": i * 960, "y": 0, "w": 960, "h": 1080}} for i in range(2)
    ]
    cfg = runner._make_config(interfaces=two_nics, sessions=sessions_scaled,
        scale=(1920, 1080), no_video=True)
    runner.run_config_test("MN-P08", "2 NICs with scaled (4K->1080p) tiled output",
        "Positive", cfg, r"scale 1920x1080|2 NIC\(s\)|Config loaded|initializ",
        expect_pass=True)

    # Multiple NICs with 12-bit transport
    cfg = runner._make_config(interfaces=two_nics, sessions=sessions_2nic,
        overrides={"video.width": 3840, "video.height": 1080,
                   "tx_video.fmt": "yuv422p12le"}, no_video=True)
    runner.run_config_test("MN-P09", "2 NICs with 12-bit format (yuv422p12le)", "Positive",
        cfg, r"2 NIC\(s\)|Config loaded|initializ", expect_pass=True)

    mismatched = [
        {"nic_index": 0, "name": DEFAULT_BDF,    "sip": DEFAULT_SIP,     "dip": DEFAULT_DIP},
        {"nic_index": 5, "name": "0000:07:00.0", "sip": "192.168.50.30", "dip": "239.168.85.21"},
    ]
    cfg = runner._make_config(interfaces=mismatched, sessions=sessions_two, no_video=True)
    runner.run_config_test("MN-N05", "interfaces[1].nic_index != array position rejected",
        "Negative", cfg, r"nic_index 5 does not match its array position")

    cfg = runner._make_config(raw_json=json.dumps({
        "video": {"width": 1920, "height": 1080},
        "tx_video": {"fps": 30, "fmt": "yuv422p10le"},
        "tx_sessions": [{"udp_port": 20000, "payload_type": 96,
                         "crop": {"x": 0, "y": 0, "w": 1920, "h": 1080}}],
    }))
    runner.run_config_test("MN-N06", "Missing interfaces[] array rejected", "Negative",
        cfg, r"'interfaces' array is empty or missing")


def run_multi_nic_tx_tests(runner, bdf_list, skip_tx=False):
    """MN: functional multi-NIC transmission (requires 2+ DPDK-bound NICs)."""
    if skip_tx or not runner.video or len(bdf_list) < 2:
        print("  [SKIP] MN TX test (needs --bdf2 with a second DPDK-bound NIC)")
        return
    interfaces = [
        {"name": bdf_list[0], "sip": DEFAULT_SIP,     "dip": DEFAULT_DIP},
        {"name": bdf_list[1], "sip": "192.168.50.30", "dip": "239.168.85.21"},
    ]
    runner.run_tx_test_full(
        "MN-TX01-2ses", "TX 2 tiled sessions across 2 NICs (yuv422p10le)",
        "yuv422p10le", 2, 1920, 1080, 30, runner.video, timeout=60,
        interfaces=interfaces)
    runner.run_tx_test_full(
        "MN-TX02-2ses-yuv422p12le", "TX 2 tiled sessions across 2 NICs (yuv422p12le)",
        "yuv422p12le", 2, 1920, 1080, 30, runner.video, timeout=60,
        interfaces=interfaces)


def run_frame_delivery_tests(runner):
    """FD: Frame Delivery tests."""
    print("\n=== Frame Delivery (FD) ===")

    if runner.video:
        cfg = runner._make_config(no_log_file=True)
        runner.run_config_test("FD-P01", "Open valid MP4 file", "Positive",
            cfg, r"Config loaded|initializ|opened", expect_pass=True, timeout=10,
            extra_args=["--test-time", "3"])
    else:
        print("  [SKIP] FD-P01: No video file available")

    no_video_file = os.path.join(runner.tmp_dir, "audio_only.mp4")
    with open(no_video_file, "wb") as f:
        f.write(b"\x00" * 1024)
    cfg = runner._make_config(overrides={"video.tx_url": no_video_file}, no_log_file=True)
    runner.run_config_test("FD-N01", "Open file with no video stream", "Negative",
        cfg, r"no video stream|Failed|error", timeout=10)

    big_yuv = os.path.join(runner.tmp_dir, "too_big.yuv")
    with open(big_yuv, "wb") as f:
        f.truncate(2 * 1024 * 1024 * 1024 + 1)
    cfg = runner._make_config(overrides={"video.tx_url": big_yuv}, no_log_file=True)
    runner.run_config_test("FD-N02", "Raw YUV file exceeds 2GB", "Negative",
        cfg, r"size|2GB|Failed|error", timeout=10)
    os.unlink(big_yuv)

    devzero_link = os.path.join(runner.tmp_dir, "devzero_link.yuv")
    try:
        os.symlink("/dev/zero", devzero_link)
        cfg = runner._make_config(overrides={"video.tx_url": devzero_link}, no_log_file=True)
        runner.run_config_test("FD-N03", "Raw YUV symlink to /dev/zero", "Negative",
            cfg, r"fstat|symlink|Failed|not a regular file|error|TIMEOUT", timeout=10)
    except OSError:
        pass

    dir_path = os.path.join(runner.tmp_dir, "is_a_directory")
    os.makedirs(dir_path, exist_ok=True)
    cfg = runner._make_config(overrides={"video.tx_url": dir_path}, no_log_file=True)
    runner.run_config_test("FD-N04", "Raw YUV path is a directory", "Negative",
        cfg, r"Is a directory|Failed|error", timeout=10)


def run_frame_tx_tests(runner, skip_tx=False, video_dir=None):
    """FT: Frame Transmission tests per format, resolution, and fps."""
    print("\n=== Frame Transmission (FT) ===")

    has_dir = bool(video_dir) and os.path.isdir(video_dir)
    if skip_tx or not (runner.video or has_dir):
        print("  [SKIP] TX tests skipped (--skip-tx or no video file)")
        return

    # If video_dir provided, run full matrix (format × resolution × fps)
    if has_dir:
        run_frame_tx_matrix(runner, video_dir)
    else:
        # Legacy mode: single video, all formats, max resolution
        for fmt in SUPPORTED_FORMATS:
            depth = FMT_BIT_DEPTH.get(fmt, "?")
            runner.run_tx_test(f"FT-P01-{fmt}", f"TX single session ({depth}-bit {fmt})",
                               fmt, 1, timeout=50)
            runner.run_tx_test(f"FT-P02-3ses-{fmt}", f"TX 3 tiled sessions ({depth}-bit {fmt})",
                               fmt, 3, timeout=50)
            runner.run_tx_test_full(
                f"FT-P03-SCALE-{fmt}", f"TX scaled 4K->1080p ({depth}-bit {fmt})",
                fmt, 1, MAX_WIDTH, MAX_HEIGHT, 30, runner.video,
                timeout=50, scale=(1920, 1080))


def _find_video(video_dir, fmt, res_name, fps, allow_any=True):
    """Find a source video for a given transport format, resolution, and fps.

    Layout: <video_dir>/<fmt folder>/<name>_<res>_<fmt>_<fps>fps_*.<ext>
    Falls back to a recursive search over video_dir when the per-format
    folder layout is not used.

    The source pixel format does NOT have to match the configured transport
    format: the decoder builds a single swscale context that converts (and
    optionally scales) from the decoded format to tx_video.fmt.  So when no
    format-matched clip exists, any clip at the same resolution/fps is used
    (allow_any=True), which is what lets one gbrp12le source drive the whole
    7-format matrix.
    """
    if not video_dir or not os.path.isdir(video_dir):
        return None
    folder = FMT_TO_FOLDER.get(fmt, fmt)
    filetok = FMT_TO_FILETOK.get(fmt, fmt)
    restok = RES_TO_FILETOK.get(res_name, res_name.lower())
    exts = (".mp4", ".mov", ".mkv", ".yuv")

    def _match(fname, require_fmt=True):
        low = fname.lower()
        return (restok in low and (filetok in low if require_fmt else True)
                and f"{fps}fps" in low and low.endswith(exts))

    base_dir = os.path.join(video_dir, folder)
    if os.path.isdir(base_dir):
        for fname in sorted(os.listdir(base_dir)):
            if _match(fname):
                return os.path.join(base_dir, fname)

    # Fallback: recursive search (handles flat or differently-named folders)
    for root, _dirs, files in os.walk(video_dir):
        for fname in sorted(files):
            if _match(fname):
                return os.path.join(root, fname)

    # Last resort: any source at this resolution/fps, regardless of its
    # pixel format — swscale converts it to the configured transport format.
    if allow_any:
        for root, _dirs, files in os.walk(video_dir):
            for fname in sorted(files):
                if _match(fname, require_fmt=False):
                    return os.path.join(root, fname)
    return None


def _video_src_fmt(path):
    """Best-effort source pixel format from the filename token."""
    if not path:
        return ""
    low = os.path.basename(path).lower()
    for fmt, tok in sorted(FMT_TO_FILETOK.items(), key=lambda kv: -len(kv[1])):
        if tok in low:
            return tok
    return ""


def _any_video(video_dir):
    """Pick a generic source video from video_dir (prefers 4K then 1080p @30fps).

    Used as the default --video when only --video-dir is given, so the TX
    suites that need a single generic source (MN-TX, legacy FT) still run.
    """
    if not video_dir or not os.path.isdir(video_dir):
        return None
    exts = (".mp4", ".mov", ".mkv", ".yuv")
    found = []
    for root, _dirs, files in os.walk(video_dir):
        for fname in sorted(files):
            if fname.lower().endswith(exts):
                found.append(os.path.join(root, fname))
    if not found:
        return None
    for res_tok in ("4k", "1080p"):
        for path in found:
            low = os.path.basename(path).lower()
            if res_tok in low and "30fps" in low:
                return path
    return found[0]


def _src_note(video, fmt):
    """Describe the source clip when its pixel format differs from the
    transport format (the decoder converts via swscale)."""
    src = _video_src_fmt(video)
    tok = FMT_TO_FILETOK.get(fmt, fmt)
    return f" [src {src}]" if src and src != tok else ""


def run_frame_tx_matrix(runner, video_dir):
    """Run FT tests for all formats (8/10/12-bit) × resolutions × fps.

    The source clip does not need to be in the transport pixel format: the
    decoder converts source -> tx_video.fmt with swscale, so a single
    gbrp12le clip can drive every format in the matrix.
    """
    test_num = 0
    for fmt in SUPPORTED_FORMATS:
        depth = FMT_BIT_DEPTH.get(fmt, "?")
        for res in RESOLUTIONS:
            for fps in FT_FPS_LIST:
                video = _find_video(video_dir, fmt, res["name"], fps)
                if not video:
                    print(f"  [SKIP] No {res['name']}/{fps}fps source video for {fmt}")
                    continue
                test_num += 1
                tc_id = f"FT-{test_num:03d}-{fmt}-{res['name']}-{fps}fps"
                desc = (f"TX {depth}-bit {fmt} {res['name']} "
                        f"({res['width']}x{res['height']}) @{fps}fps"
                        f"{_src_note(video, fmt)}")
                runner.run_tx_test_full(
                    tc_id, desc, fmt, 1,
                    res["width"], res["height"], fps, video, timeout=50)

    # Multi-session (tiled) transmission at 1080p for every format
    for fmt in SUPPORTED_FORMATS:
        depth = FMT_BIT_DEPTH.get(fmt, "?")
        video = _find_video(video_dir, fmt, "1080p", 30)
        if not video:
            print(f"  [SKIP] No 1080p/30fps source video for {fmt} (3-session test)")
            continue
        test_num += 1
        runner.run_tx_test_full(
            f"FT-{test_num:03d}-3ses-{fmt}",
            f"TX {depth}-bit {fmt} 1080p @30fps, 3 tiled sessions"
            f"{_src_note(video, fmt)}",
            fmt, 3, 1920, 1080, 30, video, timeout=60)

    # Scaled transmission for every format (source 4K -> 1080p on the wire)
    for fmt in SUPPORTED_FORMATS:
        video = _find_video(video_dir, fmt, "4K", 30)
        if not video:
            print(f"  [SKIP] No 4K/30fps source video for {fmt} (scaling test)")
            continue
        test_num += 1
        runner.run_tx_test_full(
            f"FT-{test_num:03d}-SCALE-{fmt}",
            f"TX {fmt} scaled 4K->1080p @30fps{_src_note(video, fmt)}",
            fmt, 1, 3840, 2160, 30, video, timeout=50, scale=(1920, 1080))


# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------
def get_category(tc_id):
    match = re.match(r"^([A-Z]+)", tc_id)
    return match.group(1) if match else "OTHER"


def _sanitize_for_excel(text):
    """Remove characters illegal in Excel worksheets (openpyxl) and ANSI codes."""
    if not text:
        return text
    # Strip ANSI escape sequences
    text = re.sub(r'\x1b\[[0-9;]*m', '', text)
    # Strip control characters illegal in Excel
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)


def extract_log_summary(log_output, max_len=200):
    if not log_output:
        return ""
    # Strip ANSI codes before processing
    log_output = re.sub(r'\x1b\[[0-9;]*m', '', log_output)
    lines = [ln.strip() for ln in log_output.strip().split("\n") if ln.strip()]
    # For FT tests, prefer frames/started/session lines
    for line in lines:
        if any(kw in line for kw in ["sent", "frames", "started successfully", "Session Manager"]):
            cleaned = re.sub(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d+ \[\w+ *\] ", "", line)
            return _sanitize_for_excel(cleaned[:max_len])
    for line in lines:
        if any(kw in line for kw in ["WARN", "ERROR", "fail", "reject", "invalid", "exceeds"]):
            cleaned = re.sub(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d+ \[\w+ *\] ", "", line)
            return _sanitize_for_excel(cleaned[:max_len])
    if lines:
        cleaned = re.sub(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d+ \[\w+ *\] ", "", lines[0])
        return _sanitize_for_excel(cleaned[:max_len])
    return ""


def extract_frames_from_log(log_output):
    """Extract transmission info from dvledtx log output."""
    info = {}

    # MTL st20_get_converter line: "input fmt: YUV422PLANAR10LE, output fmt: YUV422RFC4175PG2BE10"
    # This gives the actual ST2110 wire format names
    m = re.search(r"st20_get_converter.*input fmt:\s*(\S+?),\s*output fmt:\s*(\S+)", log_output)
    if m:
        info["input_fmt"] = m.group(1)
        info["transport_fmt"] = m.group(2)

    # MTL st20p_tx_create line: "transport fmt ST20_FMT_YUV_422_10BIT, input fmt: YUV422PLANAR10LE"
    if "transport_fmt" not in info:
        m = re.search(r"st20p_tx_create.*transport fmt\s+(\S+?),\s*input fmt:\s*(\S+?),", log_output)
        if m:
            info["transport_fmt"] = m.group(1)
            if "input_fmt" not in info:
                info["input_fmt"] = m.group(2)

    # Fallback: input fmt from any MTL line
    if "input_fmt" not in info:
        m = re.search(r"input fmt:\s*(\S+)", log_output)
        if m:
            info["input_fmt"] = m.group(1).rstrip(",")

    # Fallback: transport fmt from any MTL line
    if "transport_fmt" not in info:
        m = re.search(r"transport fmt\s+(\S+)", log_output)
        if m:
            info["transport_fmt"] = m.group(1).rstrip(",")

    # App-level input/output format: "opened '...' Codec=h264 1920x1080 yuv420p -> 1920x1080 yuv422p10le"
    m = re.search(r"opened\s+'[^']+'\s+Codec=(\S+)\s+\d+x\d+\s+(\S+)\s+->\s+\d+x\d+\s+(\S+)", log_output)
    if m:
        if "input_fmt" not in info:
            info["input_fmt"] = m.group(2)
        info["codec"] = m.group(1)
        info["src_pix_fmt"] = m.group(2)
        info["output_fmt"] = m.group(3)

    # App-level transport: "ffmpeg_tx opened (1920x1080 yuv422p10le @ 30fps) -> 239.168.85.20:20000"
    m = re.search(r"ffmpeg_tx opened \((\d+)x(\d+)\s+(\S+)\s+@\s+(\d+)fps\)", log_output)
    if m:
        if "transport_fmt" not in info:
            info["transport_fmt"] = m.group(3)
        info["tx_width"] = int(m.group(1))
        info["tx_height"] = int(m.group(2))
        info["configured_fps"] = int(m.group(4))

    # Frames sent from app log: "thread stopped, sent 304 frames"
    sent_frames = re.findall(r"thread stopped, sent (\d+) frames", log_output)
    if sent_frames:
        per_session = [int(x) for x in sent_frames]
        info["frames_sent"] = per_session[0]  # first session
        info["total_frames"] = sum(per_session)
        info["session_count"] = len(per_session)
        info["per_session_frames"] = ",".join(str(x) for x in per_session)

    # FPS and frames from MTL TX_VIDEO_SESSION lines (fallback):
    # "TX_VIDEO_SESSION(0,0:st20p_ffmpge): fps 29.900123 frames 299"
    if "frames_sent" not in info:
        fps_frames = re.findall(r"TX_VIDEO_SESSION\(\d+,(\d+):\S+\):\s+fps\s+([\d.]+)\s+frames\s+(\d+)", log_output)
        if fps_frames:
            last_sessions = {}
            for session_id, fps_val, frame_count in fps_frames:
                last_sessions[session_id] = int(frame_count)
            info["frames_sent"] = max(last_sessions.values())
            info["total_frames"] = sum(last_sessions.values())
            info["session_count"] = len(last_sessions)
            info["per_session_frames"] = ",".join(str(v) for v in last_sessions.values())

    # FPS from TX_VIDEO_SESSION
    fps_match = re.findall(r"TX_VIDEO_SESSION\(\d+,\d+:\S+\):\s+fps\s+([\d.]+)", log_output)
    if fps_match:
        info["fps"] = float(fps_match[-1])

    # Video resolution and format from app log, scaled variant first:
    # "Video: 3840x2160 -> scale 1920x1080 30fps yuv422p10le mode=file source=..."
    m = re.search(r"Video:\s+(\d+)x(\d+)\s+->\s+scale\s+(\d+)x(\d+)\s+(\d+)fps\s+(\S+)", log_output)
    if m:
        info["src_width"] = int(m.group(1))
        info["src_height"] = int(m.group(2))
        info["scale_width"] = int(m.group(3))
        info["scale_height"] = int(m.group(4))
        info["scaled"] = "YES"
        info["configured_fps"] = int(m.group(5))
        info.setdefault("fps", int(m.group(5)))
        info["fmt"] = m.group(6)
    else:
        # "Video: 1920x1080 30fps yuv422p10le mode=file source=..."
        m = re.search(r"Video:\s+(\d+)x(\d+)\s+(\d+)fps\s+(\S+)", log_output)
        if m:
            info["src_width"] = int(m.group(1))
            info["src_height"] = int(m.group(2))
            if "configured_fps" not in info:
                info["configured_fps"] = int(m.group(3))
            if "fps" not in info:
                info["fps"] = int(m.group(3))
            info["fmt"] = m.group(4)

    # Input mode / source from the same log line
    m = re.search(r"mode=(\S+)\s+source=(\S+)", log_output)
    if m:
        info["input_mode"] = m.group(1)
        info["source"] = m.group(2)

    # NIC count from "Config loaded: <file> (N NIC(s), M session(s))"
    m = re.search(r"Config loaded:.*\((\d+) NIC\(s\), (\d+) session\(s\)\)", log_output)
    if m:
        info["nic_count"] = int(m.group(1))
        info["config_sessions"] = int(m.group(2))

    # Use tx resolution (per-session crop) as primary, fallback to source
    if "tx_width" in info:
        info["width"] = info["tx_width"]
        info["height"] = info["tx_height"]
    elif "src_width" in info:
        info["width"] = info["src_width"]
        info["height"] = info["src_height"]

    # Fallback: check for format name in output
    if "fmt" not in info:
        for f in SUPPORTED_FORMATS:
            if f in log_output:
                info["fmt"] = f
                break

    # Fallback resolution from "Video: 1920x1080, ST20P sessions: N"
    if "width" not in info:
        m = re.search(r"Video:\s+(\d+)x(\d+)", log_output)
        if m:
            info["width"] = int(m.group(1))
            info["height"] = int(m.group(2))

    # ST20P sessions count
    m = re.search(r"ST20P sessions:\s+(\d+)", log_output)
    if m:
        info["st20p_sessions"] = int(m.group(1))

    return info


def style_header_row(ws, row, num_cols):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def style_data_cell(ws, row, col, value=None):
    cell = ws.cell(row=row, column=col)
    if value is not None:
        cell.value = _sanitize_for_excel(value) if isinstance(value, str) else value
    cell.border = THIN_BORDER
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    return cell


def write_summary_sheet(ws, results, timestamp):
    ws.sheet_properties.tabColor = "1F4E79"
    ws.merge_cells("A1:G1")
    ws["A1"].value = "DV-LED Software Toolkit (dvledtx) - Validation Report"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="center")

    ws["A3"] = "Report Generated:"
    ws["B3"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws["A4"] = "Test Timestamp:"
    ws["B4"] = timestamp
    ws["A5"] = "Binary:"
    ws["B5"] = BINARY_NAME
    ws["A6"] = "Version:"
    ws["B6"] = VERSION
    ws["A7"] = "Max Resolution:"
    ws["B7"] = f"{MAX_WIDTH}x{MAX_HEIGHT}"
    ws["A8"] = "Supported Formats:"
    ws["B8"] = ", ".join(SUPPORTED_FORMATS)
    ws["A9"] = "Supported FPS:"
    ws["B9"] = ", ".join(str(f) for f in SUPPORTED_FPS)
    ws["A10"] = "Input Modes:"
    ws["B10"] = ", ".join(SUPPORTED_INPUT_MODES)
    for r in range(3, 11):
        ws.cell(row=r, column=1).font = Font(bold=True)

    headers = ["Category", "Description", "Total", "Passed", "Failed", "Skipped", "Pass Rate"]
    row = 12
    for ci, h in enumerate(headers, 1):
        ws.cell(row=row, column=ci, value=h)
    style_header_row(ws, row, len(headers))

    cat_stats = defaultdict(lambda: {"total": 0, "passed": 0, "failed": 0, "skipped": 0})
    for tc in results:
        cat = get_category(tc["tc_id"])
        cat_stats[cat]["total"] += 1
        if tc.get("status", "").upper() == "PASS":
            cat_stats[cat]["passed"] += 1
        elif tc.get("status", "").upper() == "FAIL":
            cat_stats[cat]["failed"] += 1
        else:
            cat_stats[cat]["skipped"] += 1

    row = 13
    all_cats = list(CATEGORY_MAP.keys())
    for cat in sorted(cat_stats.keys(), key=lambda c: all_cats.index(c) if c in all_cats else 99):
        stats = cat_stats[cat]
        desc = CATEGORY_MAP.get(cat, cat)
        rate = f"{stats['passed'] / stats['total'] * 100:.1f}%" if stats["total"] > 0 else "N/A"
        style_data_cell(ws, row, 1, cat)
        style_data_cell(ws, row, 2, desc)
        style_data_cell(ws, row, 3, stats["total"])
        style_data_cell(ws, row, 4, stats["passed"])
        style_data_cell(ws, row, 5, stats["failed"])
        style_data_cell(ws, row, 6, stats["skipped"])
        cell = style_data_cell(ws, row, 7, rate)
        cell.fill = PASS_FILL if stats["failed"] == 0 and stats["passed"] == stats["total"] else FAIL_FILL
        row += 1

    total_all = sum(s["total"] for s in cat_stats.values())
    pass_all = sum(s["passed"] for s in cat_stats.values())
    fail_all = sum(s["failed"] for s in cat_stats.values())
    skip_all = sum(s["skipped"] for s in cat_stats.values())
    rate_all = f"{pass_all / total_all * 100:.1f}%" if total_all > 0 else "N/A"
    for ci, val in enumerate(["TOTAL", "", total_all, pass_all, fail_all, skip_all, rate_all], 1):
        cell = style_data_cell(ws, row, ci, val)
        cell.font = SUMMARY_FONT
        cell.fill = SUMMARY_FILL

    # --- Feature coverage section ---
    row += 3
    ws.cell(row=row, column=1, value="Feature Coverage").font = TITLE_FONT
    row += 1
    feat_headers = ["Feature", "Tests", "Passed", "Failed", "Pass Rate", "Covered", ""]
    for ci, h in enumerate(feat_headers, 1):
        ws.cell(row=row, column=ci, value=h)
    style_header_row(ws, row, len(feat_headers))
    row += 1
    for name, matcher in FEATURE_COVERAGE:
        matched = [t for t in results if matcher(t)]
        passed = sum(1 for t in matched if t.get("status", "").upper() == "PASS")
        failed = sum(1 for t in matched if t.get("status", "").upper() == "FAIL")
        rate = f"{passed / len(matched) * 100:.1f}%" if matched else "N/A"
        style_data_cell(ws, row, 1, name)
        style_data_cell(ws, row, 2, len(matched))
        style_data_cell(ws, row, 3, passed)
        style_data_cell(ws, row, 4, failed)
        style_data_cell(ws, row, 5, rate)
        cov = style_data_cell(ws, row, 6, "YES" if matched else "NOT TESTED")
        cov.fill = (PASS_FILL if matched and failed == 0
                    else FAIL_FILL if matched else SKIP_FILL)
        style_data_cell(ws, row, 7, "")
        row += 1

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 35
    for c in "CDEFG":
        ws.column_dimensions[c].width = 12


def write_category_sheet(ws, category, description, tests):
    ws.sheet_properties.tabColor = "2E75B6"
    ws.merge_cells("A1:H1")
    ws["A1"].value = f"{category} - {description}"
    ws["A1"].font = TITLE_FONT

    headers = ["TC ID", "Description", "Test Type", "Status",
               "Duration (ms)", "Key Log Message", "Matched Pattern"]
    row = 3
    for ci, h in enumerate(headers, 1):
        ws.cell(row=row, column=ci, value=h)
    style_header_row(ws, row, len(headers))

    row = 4
    for tc in tests:
        style_data_cell(ws, row, 1, tc["tc_id"])
        style_data_cell(ws, row, 2, tc["description"])
        style_data_cell(ws, row, 3, tc.get("test_type", ""))
        status_cell = style_data_cell(ws, row, 4, tc["status"])
        style_data_cell(ws, row, 5, f"{tc.get('duration_ms', 0):.2f}")
        style_data_cell(ws, row, 6, extract_log_summary(tc.get("log_output", "")))
        style_data_cell(ws, row, 7, tc.get("matched_pattern", ""))
        status_cell.fill = PASS_FILL if tc["status"] == "PASS" else FAIL_FILL if tc["status"] == "FAIL" else SKIP_FILL
        row += 1

    passed = sum(1 for t in tests if t["status"] == "PASS")
    total = len(tests)
    row += 1
    cell = style_data_cell(ws, row, 1, "SUMMARY")
    cell.font = SUMMARY_FONT
    cell.fill = SUMMARY_FILL
    cell = style_data_cell(ws, row, 2, f"{passed}/{total} passed ({passed/total*100:.0f}%)" if total else "No tests")
    cell.font = SUMMARY_FONT
    cell.fill = SUMMARY_FILL
    for ci in range(3, 8):
        style_data_cell(ws, row, ci, "").fill = SUMMARY_FILL

    widths = [14, 42, 12, 10, 14, 50, 30]
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w


def extract_ft_key_log(log_output, max_len=500):
    """Extract thread started and TX_VIDEO_SESSION lines for FT Key Log column."""
    if not log_output:
        return ""
    log_output = re.sub(r'\x1b\[[0-9;]*m', '', log_output)

    def _clean(text):
        text = re.sub(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d+\s*\[\w+\s*\]\s*", "", text)
        return re.sub(r"^MTL:\s*\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},\s*", "MTL: ", text)

    lines, fallback = [], []
    for ln in log_output.split("\n"):
        stripped = ln.strip()
        if not stripped:
            continue
        if "shared thread started" in stripped or "TX_VIDEO_SESSION" in stripped:
            lines.append(_clean(stripped))
        elif ("transport fmt" in stripped or "sent " in stripped
                or stripped.startswith("Video:")):
            fallback.append(_clean(stripped))
    result = "\n".join(lines or fallback)
    return _sanitize_for_excel(result[:max_len])


def write_frame_tx_sheet_combined(ws, tests):
    """Write a single combined FT sheet for all formats and sessions."""
    ws.sheet_properties.tabColor = "00B050"
    ws.merge_cells("A1:X1")
    ws["A1"].value = "Frame Transmission - All Formats (8/10/12-bit, scaling, screen capture, multi-NIC)"
    ws["A1"].font = TITLE_FONT

    headers = ["TC ID", "Description", "Source Fmt", "Pixel Format", "Bit Depth",
               "Input Mode", "Source Res", "TX Res", "Scaled", "Sessions", "NICs",
               "Configured FPS", "Status", "Input Fmt", "Transport Fmt",
               "Transport OK", "Actual FPS", "Throughput (Mbps)",
               "Theoretical BW (Mbps)", "NIC Limit (Mbps)", "Frames Sent",
               "Duration (ms)", "Source Video", "Key Log"]
    row = 3
    for ci, h in enumerate(headers, 1):
        ws.cell(row=row, column=ci, value=h)
    style_header_row(ws, row, len(headers))

    NIC_LIMIT = 2500.0  # Intel I225 = 2.5 Gbps

    row = 4
    for tc in tests:
        info = extract_frames_from_log(tc.get("log_output", ""))
        # Determine pixel format from tc or result
        pix_fmt = tc.get("pixel_format", "")
        if not pix_fmt:
            for f in SUPPORTED_FORMATS:
                if f in tc["tc_id"]:
                    pix_fmt = f
                    break

        resolution = tc.get("resolution", "")
        if not resolution and info.get("width"):
            resolution = f"{info['width']}x{info['height']}"
        tx_resolution = tc.get("tx_resolution", "") or resolution

        configured_fps = tc.get("configured_fps", "")
        actual_fps = tc.get("actual_fps", 0.0)
        if not actual_fps:
            actual_fps = info.get("fps", 0.0)
        throughput = tc.get("throughput_mbps", 0.0)
        theoretical_bw = tc.get("theoretical_bw_mbps", 0.0)

        style_data_cell(ws, row, 1, tc["tc_id"])
        style_data_cell(ws, row, 2, tc["description"])
        src_fmt = tc.get("source_fmt", "") or info.get("src_pix_fmt", "")
        src_cell = style_data_cell(ws, row, 3, src_fmt)
        if src_fmt and pix_fmt and src_fmt != FMT_TO_FILETOK.get(pix_fmt, pix_fmt):
            src_cell.fill = SKIP_FILL  # converted by swscale
        style_data_cell(ws, row, 4, pix_fmt)
        style_data_cell(ws, row, 5, tc.get("bit_depth", FMT_BIT_DEPTH.get(pix_fmt, "")))
        style_data_cell(ws, row, 6, tc.get("input_mode", info.get("input_mode", "")))
        style_data_cell(ws, row, 7, resolution)
        style_data_cell(ws, row, 8, tx_resolution)
        style_data_cell(ws, row, 9, tc.get("scaled", info.get("scaled", "NO")))
        style_data_cell(ws, row, 10, tc.get("session_count", info.get("session_count", "")))
        style_data_cell(ws, row, 11, tc.get("nic_count", info.get("nic_count", "")))
        style_data_cell(ws, row, 12, configured_fps)
        status_cell = style_data_cell(ws, row, 13, tc["status"])
        style_data_cell(ws, row, 14, info.get("input_fmt", ""))
        style_data_cell(ws, row, 15, info.get("transport_fmt", ""))
        transport_cell = style_data_cell(ws, row, 16, tc.get("transport_ok", ""))
        if tc.get("transport_ok") == "YES":
            transport_cell.fill = PASS_FILL
        elif tc.get("transport_ok") == "NO":
            transport_cell.fill = FAIL_FILL
        fps_str = f"{actual_fps:.2f}" if isinstance(actual_fps, float) and actual_fps else ""
        style_data_cell(ws, row, 17, fps_str)
        style_data_cell(ws, row, 18, f"{throughput:.2f}" if throughput else "")
        style_data_cell(ws, row, 19, f"{theoretical_bw:.0f}" if theoretical_bw else "")
        nic_cell = style_data_cell(ws, row, 20, f"{NIC_LIMIT:.0f}")
        if theoretical_bw > NIC_LIMIT:
            nic_cell.fill = FAIL_FILL
        style_data_cell(ws, row, 21, info.get("frames_sent", ""))
        style_data_cell(ws, row, 22, f"{tc.get('duration_ms', 0):.2f}")
        style_data_cell(ws, row, 23, tc.get("source_video", ""))
        style_data_cell(ws, row, 24, extract_ft_key_log(tc.get("log_output", "")))
        status_cell.fill = PASS_FILL if tc["status"] == "PASS" else FAIL_FILL if tc["status"] == "FAIL" else SKIP_FILL
        row += 1

    passed = sum(1 for t in tests if t["status"] == "PASS")
    total = len(tests)
    row += 1
    cell = style_data_cell(ws, row, 1, "RESULT")
    cell.font = SUMMARY_FONT
    cell.fill = SUMMARY_FILL
    cell = style_data_cell(ws, row, 2, f"{passed}/{total} PASS" if total else "No tests")
    cell.font = SUMMARY_FONT
    cell.fill = PASS_FILL if passed == total and total > 0 else FAIL_FILL
    for ci in range(3, len(headers) + 1):
        style_data_cell(ws, row, ci, "").fill = SUMMARY_FILL

    widths = [30, 52, 12, 14, 10, 14, 14, 14, 8, 10, 8, 12, 10, 22, 28, 12, 10,
              14, 16, 14, 12, 12, 34, 60]
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w


def write_frame_tx_sheet(ws, fmt, session_count, tests):
    ws.sheet_properties.tabColor = "00B050"
    label = "1-Session" if session_count == 1 else f"{session_count}-Session (Tiled)"
    ws.merge_cells("A1:K1")
    ws["A1"].value = f"Frame Transmission - {fmt} / {label}"
    ws["A1"].font = TITLE_FONT
    ws["A3"] = "Pixel Format:"
    ws["B3"] = fmt
    ws["A4"] = "Sessions:"
    ws["B4"] = session_count
    ws["A5"] = "Max Resolution:"
    ws["B5"] = f"{MAX_WIDTH}x{MAX_HEIGHT}"
    for r in range(3, 6):
        ws.cell(row=r, column=1).font = Font(bold=True)

    headers = ["TC ID", "Description", "Status", "Input Fmt", "Transport Fmt",
               "Resolution", "FPS", "Frames Sent", "Total Frames", "Duration (ms)",
               "Key Log"]
    row = 7
    for ci, h in enumerate(headers, 1):
        ws.cell(row=row, column=ci, value=h)
    style_header_row(ws, row, len(headers))

    row = 8
    for tc in tests:
        info = extract_frames_from_log(tc.get("log_output", ""))
        style_data_cell(ws, row, 1, tc["tc_id"])
        style_data_cell(ws, row, 2, tc["description"])
        status_cell = style_data_cell(ws, row, 3, tc["status"])
        style_data_cell(ws, row, 4, info.get("input_fmt", ""))
        style_data_cell(ws, row, 5, info.get("transport_fmt", ""))
        style_data_cell(ws, row, 6, f"{info.get('width', '')}x{info.get('height', '')}" if info.get("width") else "")
        fps_val = info.get("fps", "")
        if isinstance(fps_val, float):
            fps_val = f"{fps_val:.2f}"
        style_data_cell(ws, row, 7, fps_val)
        style_data_cell(ws, row, 8, info.get("frames_sent", ""))
        style_data_cell(ws, row, 9, info.get("total_frames", ""))
        style_data_cell(ws, row, 10, f"{tc.get('duration_ms', 0):.2f}")
        style_data_cell(ws, row, 11, extract_log_summary(tc.get("log_output", "")))
        status_cell.fill = PASS_FILL if tc["status"] == "PASS" else FAIL_FILL if tc["status"] == "FAIL" else SKIP_FILL
        row += 1

    passed = sum(1 for t in tests if t["status"] == "PASS")
    total = len(tests)
    row += 1
    cell = style_data_cell(ws, row, 1, "RESULT")
    cell.font = SUMMARY_FONT
    cell.fill = SUMMARY_FILL
    cell = style_data_cell(ws, row, 2, f"{passed}/{total} PASS" if total else "No tests")
    cell.font = SUMMARY_FONT
    cell.fill = PASS_FILL if passed == total and total > 0 else FAIL_FILL
    for ci in range(3, len(headers) + 1):
        style_data_cell(ws, row, ci, "").fill = SUMMARY_FILL

    for ci in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 16
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions[get_column_letter(len(headers))].width = 42


def write_format_summary_sheet(ws, results):
    ws.sheet_properties.tabColor = "FF6600"
    ws.merge_cells("A1:L1")
    ws["A1"].value = "Format x Session Validation Matrix (8/10/12-bit)"
    ws["A1"].font = TITLE_FONT
    headers = ["Pixel Format", "Bit Depth", "Sessions", "TC ID", "Status", "Input Fmt",
               "Transport Fmt", "Transport OK", "Scaled", "FPS", "Frames Sent", "Resolution"]
    row = 3
    for ci, h in enumerate(headers, 1):
        ws.cell(row=row, column=ci, value=h)
    style_header_row(ws, row, len(headers))

    ft_tests = [tc for tc in results if get_category(tc["tc_id"]) == "FT"]
    row = 4
    for fmt in SUPPORTED_FORMATS:
        depth = FMT_BIT_DEPTH.get(fmt, "")
        for ses in SESSION_COUNTS:
            matching = [tc for tc in ft_tests if fmt in tc["tc_id"]
                        and (("3ses" in tc["tc_id"]) == (ses == 3))]
            if not matching:
                style_data_cell(ws, row, 1, fmt)
                style_data_cell(ws, row, 2, depth)
                style_data_cell(ws, row, 3, ses)
                style_data_cell(ws, row, 4, "-")
                style_data_cell(ws, row, 5, "N/A").fill = SKIP_FILL
                for ci in range(6, len(headers) + 1):
                    style_data_cell(ws, row, ci, "-")
                row += 1
                continue
            for tc in matching:
                info = extract_frames_from_log(tc.get("log_output", ""))
                style_data_cell(ws, row, 1, fmt)
                style_data_cell(ws, row, 2, depth)
                style_data_cell(ws, row, 3, ses)
                style_data_cell(ws, row, 4, tc["tc_id"])
                sc = style_data_cell(ws, row, 5, tc["status"])
                sc.fill = PASS_FILL if tc["status"] == "PASS" else FAIL_FILL
                style_data_cell(ws, row, 6, info.get("input_fmt", ""))
                style_data_cell(ws, row, 7, info.get("transport_fmt", ""))
                tcell = style_data_cell(ws, row, 8, tc.get("transport_ok", ""))
                if tc.get("transport_ok") == "YES":
                    tcell.fill = PASS_FILL
                elif tc.get("transport_ok") == "NO":
                    tcell.fill = FAIL_FILL
                style_data_cell(ws, row, 9, tc.get("scaled", info.get("scaled", "NO")))
                fps_val = info.get("fps", "")
                if isinstance(fps_val, float):
                    fps_val = f"{fps_val:.2f}"
                style_data_cell(ws, row, 10, fps_val)
                style_data_cell(ws, row, 11, info.get("total_frames") or info.get("frames_sent") or "")
                res = tc.get("tx_resolution") or (
                    f"{info['width']}x{info['height']}" if info.get("width") else "")
                style_data_cell(ws, row, 12, res)
                row += 1

    widths = [18, 10, 10, 30, 10, 24, 30, 12, 8, 10, 14, 14]
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w


def generate_excel_report(results, timestamp, output_path):
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = "Summary"
    write_summary_sheet(ws_summary, results, timestamp)
    print("  Created sheet: Summary")

    ws_matrix = wb.create_sheet("Format x Session")
    write_format_summary_sheet(ws_matrix, results)
    print("  Created sheet: Format x Session")

    cat_tests = defaultdict(list)
    for tc in results:
        cat = get_category(tc["tc_id"])
        cat_tests[cat].append(tc)

    for cat in [c for c in CATEGORY_MAP if c != "FT"]:
        if cat not in cat_tests:
            continue
        desc = CATEGORY_MAP.get(cat, cat)
        name = f"{cat} - {desc}"[:31]
        ws = wb.create_sheet(name)
        write_category_sheet(ws, cat, desc, cat_tests[cat])
        print(f"  Created sheet: {name}")

    all_ft = cat_tests.get("FT", [])
    if all_ft:
        ws = wb.create_sheet("FT - Frame Transmission")
        write_frame_tx_sheet_combined(ws, all_ft)
        print("  Created sheet: FT - Frame Transmission")

    wb.save(output_path)
    passed = sum(1 for t in results if t["status"] == "PASS")
    failed = sum(1 for t in results if t["status"] == "FAIL")
    print(f"\nValidation report saved to: {output_path}")
    print(f"  Total: {len(results)} | Passed: {passed} | Failed: {failed} | Sheets: {len(wb.sheetnames)}")


def main():
    global DEFAULT_BDF

    parser = argparse.ArgumentParser(description="DV-LED Toolkit validation report generator")
    parser.add_argument("--binary", "-b", default=None, help="Path to dvledtx binary")
    parser.add_argument("--video", "-V", default=None, help="Path to test video (MP4)")
    parser.add_argument("--output", "-o", default=None, help="Output Excel path")
    parser.add_argument("--skip-tx", action="store_true", help="Skip TX tests (need HW)")
    parser.add_argument("--ft-only", action="store_true", help="Run only Frame Transmission (FT) tests")
    parser.add_argument("--no-ft", action="store_true", help="Run all tests EXCEPT Frame Transmission")
    parser.add_argument("--video-dir", default="/home/intel/workspace/sample",
                        help="Directory with per-format video folders")
    parser.add_argument("--bdf", default=DEFAULT_BDF,
                        help=f"PCI BDF of the primary DPDK-bound NIC (default {DEFAULT_BDF})")
    parser.add_argument("--bdf2", default=None,
                        help="PCI BDF of a second DPDK-bound NIC (enables multi-NIC TX tests)")
    parser.add_argument("--ft-test-time", type=int, default=DEFAULT_TX_TEST_TIME,
                        metavar="SEC",
                        help=f"Seconds each TX/FT test streams for "
                             f"(default {DEFAULT_TX_TEST_TIME}); raise for soak runs")
    parser.add_argument("--results", "-r", default=None, help="JSON results path")
    parser.add_argument("--load-only", action="store_true", help="Load JSON, generate Excel only")
    args = parser.parse_args()

    DEFAULT_BDF = args.bdf
    bdf_list = [args.bdf] + ([args.bdf2] if args.bdf2 else [])

    script_dir = os.path.dirname(os.path.abspath(__file__))

    binary = args.binary
    if not binary:
        for c in [os.path.join(script_dir, "build", "dvledtx"), os.path.join(script_dir, "dvledtx")]:
            if os.path.isfile(c) and os.access(c, os.X_OK):
                binary = c
                break
    if not binary and not args.load_only:
        sys.exit("Error: dvledtx binary not found. Use --binary to specify path.")

    video = args.video
    if not video:
        for c in [os.path.join(script_dir, "bbb_sunflower_1080p_30fps_normal.mp4"),
                   "/home/intel/workspace/sample/ball_4k_yuv420p_30fps_5min.mp4",
                   "/home/intel/workspace/sample/ball_4k_gbrp10le_30fps_5min.mp4"]:
            if os.path.isfile(c):
                video = c
                break
    if not video:
        video = _any_video(args.video_dir)

    output_path = args.output or os.path.join(script_dir, "validation_report.xlsx")
    results_path = args.results or os.path.join(script_dir, "validation_results.json")

    if args.load_only:
        if not os.path.isfile(results_path):
            sys.exit(f"Error: {results_path} not found")
        with open(results_path) as f:
            data = json.load(f)
        print(f"Loaded {len(data['results'])} results from {results_path}")
        generate_excel_report(data["results"], data.get("timestamp", "N/A"), output_path)
        return

    print(f"Binary:  {binary}")
    print(f"Video:   {video or '(none — TX tests skipped)'}")
    print(f"Video dir: {args.video_dir}")
    print(f"NIC BDF: {', '.join(bdf_list)}")
    print(f"TX time: {args.ft_test_time}s per test")
    print(f"Output:  {output_path}")
    print(f"Results: {results_path}")
    if args.ft_only:
        print(f"Mode:    FT-only (Frame Transmission matrix)")
    elif args.no_ft:
        print(f"Mode:    No-FT (skip Frame Transmission)")

    tmp_dir = os.path.join(os.getcwd(), ".dvledtx_val_tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    try:
        runner = TestRunner(binary, video, tmp_dir, tx_test_time=args.ft_test_time)
        if args.ft_only:
            run_frame_tx_tests(runner, skip_tx=args.skip_tx, video_dir=args.video_dir)
        elif args.no_ft:
            run_config_parsing_tests(runner)
            run_config_validation_tests(runner)
            run_command_line_tests(runner)
            run_log_file_tests(runner)
            run_security_tests(runner)
            run_multi_nic_tests(runner)
            run_scaling_tests(runner, skip_tx=True)
            run_screen_capture_tests(runner, skip_tx=True)
            run_ptp_tests(runner)
            run_frame_delivery_tests(runner)
        else:
            run_config_parsing_tests(runner)
            run_config_validation_tests(runner)
            run_command_line_tests(runner)
            run_log_file_tests(runner)
            run_security_tests(runner)
            run_multi_nic_tests(runner)
            run_multi_nic_tx_tests(runner, bdf_list, skip_tx=args.skip_tx)
            run_scaling_tests(runner, skip_tx=args.skip_tx, video_dir=args.video_dir)
            run_screen_capture_tests(runner, skip_tx=args.skip_tx)
            run_ptp_tests(runner)
            run_frame_delivery_tests(runner)
            run_frame_tx_tests(runner, skip_tx=args.skip_tx, video_dir=args.video_dir)
        results = runner.results
        timestamp = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    finally:
        import shutil
        shutil.rmtree(tmp_dir, ignore_errors=True)

    data = {
        "timestamp": timestamp,
        "binary": binary,
        "version": VERSION,
        "max_resolution": f"{MAX_WIDTH}x{MAX_HEIGHT}",
        "total": len(results),
        "passed": sum(1 for r in results if r["status"] == "PASS"),
        "failed": sum(1 for r in results if r["status"] == "FAIL"),
        "results": results,
    }
    with open(results_path, "w") as f:
        json.dump(data, f, indent=2)
    print(f"\nJSON results saved to: {results_path}")

    print("\nGenerating Excel report...")
    generate_excel_report(results, timestamp, output_path)


if __name__ == "__main__":
    main()
