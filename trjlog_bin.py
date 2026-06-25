# -*- coding: utf-8 -*-
"""
trjlog_bin.py  -  pure-Python TrueBeam Trajectory Log (.bin) parser.

Implements the Varian "TrueBeam Trajectory Log File Specification" (P1012906-005-E,
log v3.0/4.0) directly, with no pylinac / scipy / scikit-image dependency — so it
freezes into a small, robust click-button Windows exe and runs on locked-down PCs.

Validated byte-for-byte against pylinac on real LA1 logs (see validate_vs_pylinac.py).

Extracts ALL recorded data: header, metadata (incl. Plan UID), every axis
(expected + actual) at 50 Hz, MLC leaves, beam holds, subbeam timing — plus the
standard QA metrics (MLC RMS, axis errors, beam-on time, holds, duration).

Only dependency: numpy (lightweight, freezes cleanly).
"""
from __future__ import annotations

import struct
from dataclasses import dataclass, field

import numpy as np

# Axis enumeration -> name (Trajectory Log spec, Axis Data Structure table)
AXIS_NAMES = {
    0: "collimator", 1: "gantry", 2: "y1", 3: "y2", 4: "x1", 5: "x2",
    6: "couch_vrt", 7: "couch_lng", 8: "couch_lat", 9: "couch_rtn",
    10: "couch_pit", 11: "couch_rol", 40: "mu", 41: "beam_hold",
    42: "control_point", 50: "mlc",
    60: "target_position", 61: "tracking_target", 62: "tracking_base",
    63: "tracking_phase", 64: "tracking_conformity_index",
}
MLC_MODELS = {0: "NDS 80", 2: "NDS 120 (Millennium)", 3: "NDS 120 HD"}
# Metadata label (as written in the .bin) -> tidy key
_META_KEYS = {
    "Patient ID": "patient_id", "Plan Name": "plan_name", "Plan UID": "plan_uid",
    "Original MU": "mu_planned", "Remaining MU": "mu_remaining",
    "Energy": "energy", "BeamName": "beam_name",
}


@dataclass
class Subbeam:
    cp: int
    mu: float
    rad_time: float          # expected irradiation time (s)
    seq: int
    name: str


@dataclass
class TrajectoryLog:
    path: str
    signature: str
    version: str
    sampling_interval_ms: int
    num_axes: int
    axis_enum: list
    samples_per_axis: list
    axis_scale: int          # 1=Machine, 2=Modified IEC 61217
    num_subbeams: int
    is_truncated: bool
    num_snapshots: int
    mlc_model: int
    metadata: dict
    subbeams: list
    # axis_name -> (expected, actual) float arrays; mlc -> (exp[N,122], act[N,122])
    axes: dict = field(default_factory=dict, repr=False)
    crc_ok: bool | None = None

    # ---- convenience ----
    @property
    def mlc_model_name(self) -> str:
        return MLC_MODELS.get(self.mlc_model, "model %s" % self.mlc_model)

    @property
    def duration_s(self) -> float:
        return self.num_snapshots * self.sampling_interval_ms / 1000.0

    def difference(self, axis: str):
        exp, act = self.axes[axis]
        return act - exp


# ----------------------------------------------------------------------
def _read_metadata(block: bytes) -> dict:
    """Metadata is a tab-separated, CRLF-delimited text block (UTF-8)."""
    text = block.split(b"\x00", 1)[0].decode("utf-8", "replace")
    out = {}
    for line in text.replace("\r\n", "\n").split("\n"):
        if "\t" in line:
            label, _, val = line.partition("\t")
            key = _META_KEYS.get(label.rstrip(":").strip())
            if key:
                out[key] = val.strip()
    # numeric coercions
    for k in ("mu_planned", "mu_remaining"):
        if k in out:
            try:
                out[k] = float(out[k])
            except ValueError:
                pass
    return out


def _crc16_ccitt(data: bytes, seed: int = 0xFFFF) -> int:
    crc = seed
    for byte in data:
        crc ^= byte << 8
        for _ in range(8):
            crc = ((crc << 1) ^ 0x1021) & 0xFFFF if (crc & 0x8000) else (crc << 1) & 0xFFFF
    return crc


def load_log(path: str, *, read_axes: bool = True, check_crc: bool = False) -> TrajectoryLog:
    with open(path, "rb") as fh:
        raw = fh.read()

    signature = raw[0:16].decode("latin-1").rstrip("\x00")
    version = raw[16:32].decode("latin-1").rstrip("\x00")
    header_size, sampling, num_axes = struct.unpack_from("<iii", raw, 32)
    off = 44
    axis_enum = list(struct.unpack_from("<%di" % num_axes, raw, off)); off += 4 * num_axes
    samples = list(struct.unpack_from("<%di" % num_axes, raw, off)); off += 4 * num_axes
    axis_scale, num_subbeams, is_trunc, num_snap, mlc_model = struct.unpack_from("<iiiii", raw, off)
    off += 20
    metadata = _read_metadata(raw[off:off + 745])

    log = TrajectoryLog(
        path=path, signature=signature, version=version,
        sampling_interval_ms=sampling, num_axes=num_axes, axis_enum=axis_enum,
        samples_per_axis=samples, axis_scale=axis_scale, num_subbeams=num_subbeams,
        is_truncated=bool(is_trunc), num_snapshots=num_snap, mlc_model=mlc_model,
        metadata=metadata, subbeams=[],
    )

    # ---- subbeams (start at fixed header end = 1024) ----
    pos = header_size  # 1024
    for _ in range(num_subbeams):
        cp, mu, rad_time, seq = struct.unpack_from("<iffi", raw, pos)
        name = raw[pos + 16:pos + 16 + 512].split(b"\x00", 1)[0].decode("latin-1", "replace")
        log.subbeams.append(Subbeam(cp, mu, rad_time, seq, name))
        pos += 560

    if not read_axes:
        return log

    # ---- axis data: num_snap snapshots, each sum(samples)*2 float32 ----
    floats_per_snap = sum(samples) * 2
    n_floats = floats_per_snap * num_snap
    arr = np.frombuffer(raw, dtype="<f4", count=n_floats, offset=pos)
    arr = arr.reshape(num_snap, floats_per_snap)

    col = 0
    for aid, ns in zip(axis_enum, samples):
        width = ns * 2
        block = arr[:, col:col + width]           # (num_snap, ns*2)
        exp = block[:, 0::2]                       # expected (num_snap, ns)
        act = block[:, 1::2]                       # actual   (num_snap, ns)
        name = AXIS_NAMES.get(aid, "axis_%d" % aid)
        if ns == 1:
            log.axes[name] = (exp[:, 0].copy(), act[:, 0].copy())
        else:
            log.axes[name] = (exp.copy(), act.copy())  # e.g. mlc (num_snap, 122)
        col += width

    if check_crc:
        stored = struct.unpack_from("<H", raw, pos + n_floats * 4)[0]
        log.crc_ok = (_crc16_ccitt(raw[:pos + n_floats * 4]) == stored)
    return log


# ----------------------------------------------------------------------
def mlc_leaf_arrays(log: TrajectoryLog):
    """Return (expected, actual) MLC LEAF arrays (cm), dropping the 2 carriages.

    Snapshot MLC order (spec p.11): [Carriage A, Carriage B, bank-A leaves 1..60,
    bank-B leaves 1..60] = num_leaves + 2. So samples 0,1 are carriages; 2.. are
    the leaves. Verified vs pylinac (identical leaf positions).
    """
    exp, act = log.axes["mlc"]                      # (num_snap, n_samples)
    num_leaves = exp.shape[1] - 2
    return exp[:, 2:2 + num_leaves], act[:, 2:2 + num_leaves]


def beam_on_mask(log: TrajectoryLog):
    """Boolean per-snapshot mask of beam-ON samples (Beam Hold == Normal/0).

    Mirrors pylinac's exclude_beam_off so MLC statistics match its reference."""
    if "beam_hold" not in log.axes:
        return None
    _, bh = log.axes["beam_hold"]
    return bh.astype(int) == 0


def mlc_error_series(log: TrajectoryLog):
    """Per-snapshot MLC leaf error for the per-log graph.

    Returns (time_s, mean_err_mm, max_err_mm) over the whole delivery (no beam-off
    masking, so the user sees where in the beam any deviation occurred)."""
    exp, act = mlc_leaf_arrays(log)
    err = np.abs(act - exp) * 10.0                      # cm -> mm
    t = np.arange(err.shape[0]) * log.sampling_interval_ms / 1000.0
    return t, err.mean(axis=1), err.max(axis=1)


def qa_metrics(log: TrajectoryLog) -> dict:
    """Standard delivery-QA metrics, all from the raw axis data (no pylinac)."""
    m: dict = {
        "filename": log.path.replace("\\", "/").split("/")[-1],
        "patient_id": log.metadata.get("patient_id", ""),
        "plan_name": log.metadata.get("plan_name", ""),
        "beam_name": log.metadata.get("beam_name", ""),
        "plan_uid": log.metadata.get("plan_uid", ""),
        "energy": log.metadata.get("energy", ""),
        "mlc_model": log.mlc_model_name,
        "num_snapshots": log.num_snapshots,
        "sampling_interval_ms": log.sampling_interval_ms,
        "duration_s": round(log.duration_s, 2),
        "num_subbeams": log.num_subbeams,
        "is_truncated": log.is_truncated,
        "mu_planned": log.metadata.get("mu_planned", ""),
    }
    # delivered MU = last actual MU sample
    mu_deliv = None
    if "mu" in log.axes:
        _, mu_act = log.axes["mu"]
        mu_deliv = round(float(mu_act[-1]), 3) if mu_act.size else None
        m["mu_delivered"] = mu_deliv
    # imaging beam (kV/CBCT/setup): no MLC modulation, no dose -> blank MLC metrics
    energy = str(log.metadata.get("energy", "")).lower()
    is_imaging = energy.endswith("k") or (mu_deliv is not None and mu_deliv < 0.01)
    m["is_imaging"] = is_imaging
    # MU discrepancy (dose-relevant): delivered vs planned, %.
    mu_plan = log.metadata.get("mu_planned")
    if (not is_imaging and isinstance(mu_plan, (int, float)) and mu_plan > 0.01
            and mu_deliv is not None):
        m["mu_diff_pct"] = round((mu_deliv - mu_plan) / mu_plan * 100.0, 3)
    else:
        m["mu_diff_pct"] = ""
    # MLC leaf errors — actual vs expected. Log stores cm; report mm (×10, clinical).
    # Restrict to beam-ON snapshots to match pylinac's exclude_beam_off reference.
    if "mlc" in log.axes and not is_imaging:
        exp, act = mlc_leaf_arrays(log)
        mask = beam_on_mask(log)
        if mask is not None and mask.any():
            exp, act = exp[mask], act[mask]
        err = np.abs(act - exp)
        rms = np.sqrt(np.mean((act - exp) ** 2, axis=0))     # per-leaf RMS over time
        m["mlc_rms_avg_mm"] = round(float(np.mean(rms)) * 10, 4)   # cm->mm
        m["mlc_rms_max_mm"] = round(float(np.max(rms)) * 10, 4)
        m["mlc_err_p95_mm"] = round(float(np.percentile(err, 95)) * 10, 4)
        m["mlc_err_max_mm"] = round(float(np.max(err)) * 10, 4)
    # mechanical axis max errors
    for ax, key, scale in [("gantry", "max_gantry_error_deg", 1),
                           ("collimator", "max_collimator_error_deg", 1)]:
        if ax in log.axes:
            m[key] = round(float(np.max(np.abs(log.difference(ax)))), 4)
    jaw_err = []
    for j in ("x1", "x2", "y1", "y2"):
        if j in log.axes:
            jaw_err.append(float(np.max(np.abs(log.difference(j)))))
    if jaw_err:
        m["max_jaw_error_cm"] = round(max(jaw_err), 4)
    # beam holds (Beam Hold axis == 2 == Hold; gating/transitions)
    if "beam_hold" in log.axes:
        _, bh = log.axes["beam_hold"]
        holds = (bh.astype(int) == 2)
        # count hold "events" (rising edges)
        m["num_beam_holds"] = int(np.sum((holds[1:]) & (~holds[:-1]))) + int(holds[0])
        m["beam_hold_fraction"] = round(float(np.mean(holds)), 4)
    # beam-on time = expected irradiation time of subbeams (radTime)
    m["beam_on_s"] = round(sum(sb.rad_time for sb in log.subbeams), 2)
    return m
