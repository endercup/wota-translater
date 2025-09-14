# -*- coding: utf-8 -*-
"""
wota-translater v5.6.2
- 多段保留 (第X段進入/退出) 支援
- 保留指定區段並剪去其餘區段
- 支援單一數字秒數、mm:ss、hh:mm:ss
- 0:00 例外 (首段 start == 0 且使用者未填或填0，則不淡入)
- Peak 正規化至 -1 dBFS
- 保留模糊配對、錯誤記錄、進度列與彩色輸出
- E1 預設已移除（第X段欄位專用）
"""

VERSION = "5.6.2"
TARGET_PEAK = -1.0     # dBFS
INPUT_FOLDER = "songs"
OUTPUT_FOLDER = "output"
TRACKLIST = "tracklist.xlsx"
SUPPORTED_EXT = [".mp3", ".m4a", ".wav", ".flac"]
FUZZY_TH = 0.4

import os
import re
import sys
import unicodedata
import traceback
import ctypes
from difflib import get_close_matches
import pandas as pd
from pydub import AudioSegment
import imageio_ffmpeg as ffmpeg

# 強制指定 ffmpeg 與 ffprobe 路徑（避免使用者需安裝 ffmpeg）
AudioSegment.converter = ffmpeg.get_ffmpeg_exe()
AudioSegment.ffprobe = ffmpeg.get_ffprobe_exe()

from colorama import init as color_init, Fore, Style
from tqdm import tqdm
from openpyxl import load_workbook

color_init(autoreset=True)

# --- small helpers ---
def cprint(msg, col=Fore.WHITE):
    print(col + msg + Style.RESET_ALL)

def rprint(msg):
    cprint(msg, Fore.RED)

def final_popup(msg):
    """Windows MessageBox uses CRLF for line breaks; wrap in try/except for cross-platform safety."""
    try:
        ctypes.windll.user32.MessageBoxW(None, msg, "wota-translater " + VERSION, 0x40)
    except Exception:
        # fallback: nothing (we already printed to console)
        pass

def write_txt(path, lines):
    if not lines:
        # ensure empty file removed when no errors? (leave absent)
        return
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

# --- time parsing ---
def to_ms(t: str) -> int:
    """
    Parse time formats:
    - "20" -> 20 seconds
    - "0:20" -> mm:ss
    - "1:02:03" -> hh:mm:ss
    Raises ValueError on invalid.
    """
    if t is None:
        raise ValueError("時間為空")
    s = str(t).strip()
    if s == "":
        raise ValueError("時間為空")
    parts = s.split(":")
    try:
        if len(parts) == 1:
            return int(float(parts[0]) * 1000)
        if len(parts) == 2:
            m, sec = map(float, parts)
            return int((m * 60 + sec) * 1000)
        if len(parts) == 3:
            h, m, sec = map(float, parts)
            return int((h * 3600 + m * 60 + sec) * 1000)
    except Exception:
        raise ValueError(f"無效時間格式: {t}")
    raise ValueError(f"無效時間格式: {t}")

# --- chinese numerals (basic) ---
CHINESE_NUM_MAP = {
    "零":0,"〇":0,"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9,
    "十":10,"百":100
}
def chinese_to_int(s: str) -> int:
    s = str(s).strip()
    if s == "":
        raise ValueError("empty")
    # arabic digits?
    if re.fullmatch(r"\d+", s):
        return int(s)
    # single char common
    if s in CHINESE_NUM_MAP and CHINESE_NUM_MAP[s] < 10:
        return CHINESE_NUM_MAP[s]
    # handle hundred (rare)
    if "百" in s:
        parts = s.split("百")
        hundreds = CHINESE_NUM_MAP.get(parts[0], 0) if parts[0] else 1
        rest = parts[1] if len(parts) > 1 else ""
        tens = 0; units = 0
        if "十" in rest:
            idx = rest.index("十")
            tens_char = rest[:idx]
            tens = CHINESE_NUM_MAP.get(tens_char, 1) if tens_char else 1
            units_char = rest[idx+1:]
            if units_char:
                units = CHINESE_NUM_MAP.get(units_char, 0)
        else:
            if rest:
                units = CHINESE_NUM_MAP.get(rest, 0)
        return hundreds*100 + tens*10 + units
    # ten patterns
    if "十" in s:
        parts = s.split("十")
        if parts[0] == "":
            tens = 1
        else:
            tens = CHINESE_NUM_MAP.get(parts[0], 0)
        units = 0
        if len(parts) > 1 and parts[1] != "":
            units = CHINESE_NUM_MAP.get(parts[1], 0)
        return tens*10 + units
    # fallback: map characters
    val = 0
    for ch in s:
        if ch in CHINESE_NUM_MAP:
            val = val*10 + CHINESE_NUM_MAP[ch]
        else:
            raise ValueError(f"無法解析數字: {s}")
    return val

# --- header parsing ---
def normalize_header(h: str) -> str:
    return re.sub(r"\s+", "", str(h)).replace("（","(").replace("）",")")

def find_segment_pairs(columns):
    """
    Scan header columns and return list of (index, enter_col, exit_col) sorted by index.
    Recognizes:
      - 第X段進入 / 第X段退出
      - 第X進入 / 第X退出
    X can be Chinese numerals or Arabic digits.
    """
    pairs = {}
    for col in columns:
        name = normalize_header(col)
        m = re.match(r"第(.+?)段(進入|退出)$", name)
        if not m:
            m = re.match(r"第(.+?)(進入|退出)$", name)
        if m:
            xraw, typ = m.group(1), m.group(2)
            try:
                idx = chinese_to_int(xraw)
            except Exception:
                digits = re.sub(r"\D", "", xraw)
                if digits:
                    idx = int(digits)
                else:
                    continue
            pairs.setdefault(idx, {})[typ] = col
    out = []
    for k in sorted(pairs.keys()):
        ent = pairs[k].get("進入") or pairs[k].get("Enter") or pairs[k].get("enter")
        ex = pairs[k].get("退出") or pairs[k].get("Exit") or pairs[k].get("exit")
        if ent and ex:
            out.append((k, ent, ex))
    return out

# --- sanitize filenames & subtitles removal ---
def sanitize_filename(name: str) -> str:
    name = unicodedata.normalize("NFKC", name)
    # remove bracketed subtitles
    name = re.sub(r"[\[\(（【].*?[\]\)）】]", " ", name)
    # remove common subtitles
    name = re.sub(r"(?i)\b(MV|Official|Lyric[s]?|Audio|HD|Live|Karaoke|完整版|官方|Full)\b", " ", name)
    # remove illegal filesystem chars
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = re.sub(r"[_\s]+", " ", name).strip()
    return name or "untitled"

def list_block(title, lines):
    if not lines:
        return ""
    return f"{title} {len(lines)} 首\n  • " + "\n  • ".join(lines)

# ----------------- main -----------------
def main():
    # console intro (no popup)
    intro_lines = [
        f"🎵 wota-translater {VERSION} — 快速剪輯 御宅藝 用副歌檔案",
        "• 依 Excel 清單批次裁切 songs/ → output/ 320 kbps MP3",
        "• 支援多段保留：開頭進入/退出、第X段進入/退出",
        "• 0:00 例外：首段從0開始時不做淡入",
        "• 時間格式：ss, mm:ss, hh:mm:ss (單數字=秒)",
        f"• Peak 正規化：{TARGET_PEAK} dBFS",
        "• 防呆：模糊比對、重複檢查、錯誤分類記錄"
    ]
    cprint("\n".join(intro_lines), Fore.CYAN)

    # prepare output dir and clear previous outputs
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    for f in os.listdir(OUTPUT_FOLDER):
        try:
            os.remove(os.path.join(OUTPUT_FOLDER, f))
        except Exception:
            pass

    # remove old logs
    logs = ["duplicate_titles.txt","duplicate_matches.txt","invalid_time.txt",
            "processing_errors.txt","unmatched_titles.txt","unmatched_audio.txt"]
    for l in logs:
        try:
            if os.path.exists(l):
                os.remove(l)
        except Exception:
            pass

    # read excel
    try:
        df = pd.read_excel(TRACKLIST, dtype=str)
    except Exception as e:
        rprint(f"無法讀取 {TRACKLIST}: {e}")
        sys.exit(1)

    # compatibility: if first row labelled '開頭' (older sheets), drop it
    if len(df) > 0 and str(df.iloc[0,0]).strip() == "開頭":
        df = df.drop(index=0).reset_index(drop=True)

    if "歌名" not in df.columns:
        rprint('Excel 缺少「歌名」欄')
        sys.exit(1)

    # prepare audio_files
    audio_files = {}
    if os.path.exists(INPUT_FOLDER):
        for fn in os.listdir(INPUT_FOLDER):
            name, ext = os.path.splitext(fn)
            if ext.lower() in SUPPORTED_EXT:
                audio_files[name] = os.path.join(INPUT_FOLDER, fn)
    if not audio_files:
        rprint("songs/ 內無支援音檔")
        sys.exit(1)

    # duplicate titles
    dup_titles = df[df.duplicated('歌名', keep=False)]['歌名'].unique().tolist()
    if dup_titles:
        write_txt('duplicate_titles.txt', dup_titles)
        rprint(f"表格內重複歌名：{len(dup_titles)} 首，詳見 duplicate_titles.txt")

    # unmatched detection
    titles = df['歌名'].astype(str).tolist()
    matched_audio = set()
    unmatched_titles = []
    for t in titles:
        m = get_close_matches(t, audio_files.keys(), n=1, cutoff=FUZZY_TH)
        if m:
            matched_audio.update(m)
        else:
            unmatched_titles.append(t)
    unmatched_audio = [k for k in audio_files if k not in matched_audio]
    write_txt('unmatched_titles.txt', unmatched_titles)
    write_txt('unmatched_audio.txt', unmatched_audio)
    if unmatched_titles:
        rprint(list_block('未配對歌名', unmatched_titles))
    if unmatched_audio:
        rprint(list_block('未配對音檔', unmatched_audio))

    # dynamic header parsing
    seg_pairs = find_segment_pairs(df.columns.tolist())

    audio_used = {}
    ok = 0
    ng = 0
    dup_match_err = []
    time_err = []
    process_err = []

    total = len(df)
    for idx, row in enumerate(tqdm(df.itertuples(index=False), total=total, ncols=80, colour='yellow', desc='Processing'), 1):
        row_series = df.iloc[idx-1]
        title = str(row_series.get('歌名') or "").strip()
        tqdm.write(Fore.YELLOW + f"[{idx}/{total}] {title} 處理中…" + Style.RESET_ALL)

        def read_cell(col):
            if col not in df.columns:
                return None
            val = row_series.get(col)
            if pd.isna(val):
                return None
            s = str(val).strip()
            return s if s != "" else None

        # collect segments: each as (start_ms, end_ms, ent_col, exit_col, raw_ent_value)
        segments = []

        # head: support '開頭進入' or '開頭進入(選填)'
        head_ent_col = None
        if '開頭進入' in df.columns:
            head_ent_col = '開頭進入'
        elif '開頭進入(選填)' in df.columns:
            head_ent_col = '開頭進入(選填)'
        if head_ent_col and '開頭退出' in df.columns:
            ent = read_cell(head_ent_col)
            ex = read_cell('開頭退出')
            if ent is None and ex is not None:
                # unspecified ent -> treat as 0
                try:
                    segments.append((0, to_ms(ex), head_ent_col, '開頭退出', ent))
                except Exception:
                    time_err.append(f"{title}: 開頭退出格式錯誤 {ex}")
            elif ent is not None and ex is not None:
                try:
                    segments.append((to_ms(ent), to_ms(ex), head_ent_col, '開頭退出', ent))
                except Exception:
                    time_err.append(f"{title}: 開頭時間格式錯誤 {ent} / {ex}")

        # dynamic parsed pairs (第X段進入/退出)
        for (num, ent_col, ex_col) in seg_pairs:
            ent = read_cell(ent_col)
            ex = read_cell(ex_col)
            if ent and ex:
                try:
                    segments.append((to_ms(ent), to_ms(ex), ent_col, ex_col, ent))
                except Exception:
                    time_err.append(f"{title}: 第{num}段時間格式錯誤 {ent} / {ex}")
            elif ent is None and ex is not None:
                # if user only supplied exit, treat enter as 0
                try:
                    segments.append((0, to_ms(ex), ent_col, ex_col, ent))
                except Exception:
                    time_err.append(f"{title}: 第{num}段退出格式錯誤 {ex}")

        # if no segments found
        if not segments:
            msg = f"{title}: 未找到任何有效保留段 (跳過)"
            process_err.append(msg)
            rprint(msg)
            ng += 1
            continue

        # sort and validate segments
        segments.sort(key=lambda x: x[0])
        invalid = False
        for i, (s_ms, e_ms, sc, ec, raw_ent) in enumerate(segments):
            if s_ms >= e_ms:
                time_err.append(f"{title}: 段落時間錯誤 {s_ms} >= {e_ms}")
                invalid = True
                break
            if i > 0:
                prev_end = segments[i-1][1]
                if s_ms < prev_end:
                    time_err.append(f"{title}: 段落重疊或未排序 {s_ms} < prev_end {prev_end}")
                    invalid = True
                    break
        if invalid:
            rprint(f"{title}: 段位驗證失敗，已記錄")
            ng += 1
            continue

        # fuzzy match audio
        m = get_close_matches(title, audio_files.keys(), n=1, cutoff=FUZZY_TH)
        if not m:
            msg = f"找不到相符音檔: {title}"
            process_err.append(msg)
            rprint(msg)
            ng += 1
            continue
        audio_key = m[0]
        if audio_key in audio_used:
            msg = f"音檔重複配對: {audio_key}: {audio_used[audio_key]} & {title}"
            dup_match_err.append(msg)
            rprint(msg)
            ng += 1
            continue
        audio_used[audio_key] = title

        try:
            a = AudioSegment.from_file(audio_files[audio_key])
            audio_len = len(a)
            # check segment ends within audio
            for s_ms, e_ms, sc, ec, raw_ent in segments:
                if e_ms > audio_len:
                    raise ValueError(f"{title}: 段落結束 {e_ms/1000:.2f}s 超過音檔長度 {audio_len/1000:.2f}s")
            # extract pieces
            pieces = []
            for idx_seg, (s_ms, e_ms, sc, ec, raw_ent) in enumerate(segments):
                seg = a[s_ms:e_ms]
                seg_len = len(seg)
                apply_fade_in = True
                # 0:00 exception for first segment
                if idx_seg == 0 and s_ms == 0 and (raw_ent is None or str(raw_ent).strip() in ("0", "0:00")):
                    apply_fade_in = False
                if apply_fade_in and seg_len > 0:
                    seg = seg.fade_in(min(2000, seg_len))
                if seg_len > 1000:
                    tail_len = min(2000, seg_len)
                    seg = seg[:-1000].fade_out(min(2000, tail_len))
                pieces.append(seg)
            # concatenate
            result = AudioSegment.silent(duration=0)
            for p in pieces:
                result += p
            # peak normalize
            try:
                change = TARGET_PEAK - result.max_dBFS
                result = result.apply_gain(change)
            except Exception:
                # silent or problematic file -> leave as is
                pass
            out_name = sanitize_filename(title) + ".mp3"
            out_path = os.path.join(OUTPUT_FOLDER, out_name)
            result.export(out_path, format="mp3", bitrate="320k")
            ok += 1
        except Exception as ex:
            msg = f"{title}: 處理失敗: {ex}"
            process_err.append(msg)
            rprint(msg)
            ng += 1
            continue

    # end for

    # write logs
    write_txt('duplicate_matches.txt', dup_match_err)
    write_txt('invalid_time.txt', time_err)
    write_txt('processing_errors.txt', process_err)

    # prepare summary and detail blocks (use CRLF for messagebox)
    summary = [f"✔ 成功：{ok}", f"✘ 失敗：{ng}"]
    detail_blocks = []
    for title_label, lst in [
        ("重複歌名", dup_titles if 'dup_titles' in locals() else []),
        ("重複配對", dup_match_err),
        ("時間錯誤", time_err),
        ("處理錯誤", process_err),
        ("未配對歌名", unmatched_titles),
        ("未配對音檔", unmatched_audio),
    ]:
        if lst:
            summary.append(f"{title_label} {len(lst)} 首")
            detail_blocks.append(f"{title_label} {len(lst)} 首\r\n  • " + "\r\n  • ".join(lst))

    # console print (green success, red others)
    for line in summary:
        if line.startswith("✔"):
            cprint(line, Fore.GREEN)
        else:
            cprint(line, Fore.RED)

    # final popup (CRLF)
    popup_text = "\r\n".join(summary)
    if detail_blocks:
        popup_text += "\r\n\r\n" + "\r\n\r\n".join(detail_blocks)
    final_popup(popup_text)


if __name__ == "__main__":
    main()