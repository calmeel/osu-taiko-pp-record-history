import pandas as pd
from pathlib import Path
import html
import math

from datetime import date
TODAY = date.today()

# ==== 設定 ====
DATA_FILE = Path("data.xlsx")
OUTPUT_FILE = Path("index.html")

# 列名（data.xlsx 側）
COL_DATE = "DATE"
COL_DAYS = "Days Maintained"
COL_SECTION_LINK = "SECTION LINK"
COL_PP = "PP"
COL_PP_DIFF = "PP DIFF"
COL_FLAG = "flag"
COL_PLAYER = "PLAYER"
COL_PLAYER_LINK = "PLAYER LINK"
COL_JACKET = "Jacket"
COL_MAP = "MAP"
COL_MAP_LINK = "MAP link"
COL_SR = "SR"
COL_ACC = "ACC"
COL_OSU_ICON = "osu"
COL_OSU_LINK = "osu link"
COL_YT_ICON = "youtube"
COL_YT_LINK = "youtube link"
COL_RD_ICON = "reddit"
COL_RD_LINK = "reddit link"
COL_X_ICON = "X"
COL_X_LINK = "X link"
COL_REPLAY = "Replay"
COL_REMARKS = "Remarks"
MOD_COLS = ["MOD1", "MOD2", "MOD3", "MOD4", "MOD5", "MOD6", "MOD7"]

# ==== HTML テンプレ（ヘッダー部分） ====
HTML_HEAD = """<!DOCTYPE html>
<html lang="ja">
<meta name="description" content="Historical record of osu!taiko top performance points (pp) plays, manually collected by Vanity8.">
<head>
  <meta charset="UTF-8">
  <!-- スマホでもPC版レイアウトで表示させる -->
  <meta name="viewport" content="width=1280">
  <title>osu!taiko pp record history</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>

<h1>osu!taiko pp record history (created by <a href="https://osu.ppy.sh/users/12029122" target="_blank" rel="noopener noreferrer">Vanity8</a>)</h1>
<p class="subtitle">
  I manually collect records using web archives, osu! comments, and all kinds of social media.<br>
  If you notice any missing information or have replay data, please let me know on my Discord (id: vanity8) and I'll add it.
</p>

<div class="table-wrapper">
<table class="pp-table">
  <thead>
    <tr>
      <th class="col-date" scope="col">Date</th>
      <th class="col-days" scope="col">Days maintained</th>
      <th class="icon-col"> </th>
      <th class="col-player" scope="col">Player</th>
      <th class="icon-col"> </th>
      <th class="col-map" scope="col">Map</th>
      <th class="col-sr" scope="col">SR</th>
      <th>Mod</th>
      <th class="col-acc" scope="col">Acc</th>
      <th class="col-pp" scope="col">PP</th>
      <th class="col-pp"> </th>
      <th class="icon-col">osu!</th>
      <th class="icon-col">YouTube</th>
      <th class="icon-col">reddit</th>
      <th class="icon-col">X</th>
      <th class="col-replay">Replay</th>
      <th class="col-remarks" scope="col">Remarks</th>
    </tr>
  </thead>
  <tbody>
"""

HTML_GITHUB_LINK = """
<div class="footer">
  <a href="https://github.com/calmeel/osu-taiko-pp-record-history" target="_blank" rel="noopener noreferrer">
    View this project on GitHub
  </a>
</div>
"""

HTML_FOOT = """  </tbody>
</table>
</div>

</body>
</html>
"""


def fmt_date(val):
    """DATE列を 'YYYY-MM-DD' 文字列に整形"""
    if pd.isna(val):
        return ""
    if hasattr(val, "strftime"):
        return val.strftime("%Y-%m-%d")
    return str(val)

def fmt_int(val):
    if pd.isna(val):
        return ""
    try:
        return str(int(round(float(val))))
    except Exception:
        return str(val)

def compute_days_display(row) -> str:
    """
    Days maintained の最終表示文字列を返す。
    仕様:
      - COL_DAYS が "-"   -> "-" （suffixなし）
      - COL_DAYS に数値   -> その値を優先して "N day(s)"
      - COL_DAYS が空欄   -> COL_DATE と TODAY から自動計算して "N day(s)"
      - 計算失敗          -> ""
    """
    print("[DEBUG IN]", repr(row.get(COL_DATE)), repr(row.get(COL_DAYS)))
    days_raw = row.get(COL_DAYS)
    date_val = row.get(COL_DATE)

    # 1) "-" はそのまま
    if isinstance(days_raw, str) and days_raw.strip() == "-":
        return "-"

    # 2) まず「Excel側に値が入っているか」を見る
    has_explicit_days = False
    n_days_from_excel = None

    if days_raw is not None:
        # pandas の NaN 対策
        if isinstance(days_raw, float) and math.isnan(days_raw):
            has_explicit_days = False
        else:
            s = str(days_raw).strip()
            if s != "":
                try:
                    n_days_from_excel = int(s)
                    has_explicit_days = True
                except Exception:
                    has_explicit_days = False

    # 3) Excel に日数が入っていればそれを優先
    if has_explicit_days and n_days_from_excel is not None:
        n = n_days_from_excel
        if n == 1:
            return f"{n} day"
        else:
            return f"{n} days"

    # 4) ここまで来たら「空欄なので日付から自動計算」モード
    try:
        if date_val is None:
            print("[DEBUG ERROR] date_val is None, days_raw =", repr(days_raw))
            return ""

        # pandas の NaN / 空文字対応
        if isinstance(date_val, float) and math.isnan(date_val):
            print("[DEBUG ERROR] date_val is NaN, days_raw =", repr(days_raw))
            return ""

        sdate = str(date_val).strip()
        if not sdate:
            print("[DEBUG ERROR] sdate is empty, days_raw =", repr(days_raw))
            return ""

        # "2025/11/8" でも、Excel日付シリアルでも pd.to_datetime が解釈してくれる想定
        d = pd.to_datetime(sdate).date()
        print("[DEBUG PARSE]", repr(sdate), "->", d)

        delta = (TODAY - d).days
        print("[DEBUG DELTA]", delta)

        if delta < 0:
            # 未来日付は表示しない
            print("[DEBUG NEGATIVE]", delta, "(future date?)")
            return ""

        n = int(delta)
        if n == 1:
            return f"{n} day"
        else:
            return f"{n} days"

    except Exception as e:
        print("[DEBUG ERROR] exception in auto-calc:", repr(e),
              " raw date_val=", repr(date_val), " days_raw=", repr(days_raw))
        return ""

def fmt_acc(val):
    """0.991234 → '99.12%' などに整形"""
    if pd.isna(val):
        return ""
    try:
        return f"{float(val) * 100:.2f}%"
    except Exception:
        return str(val)


def build_link_cell(text, url, td_class):
    """テキスト＋URL から <td> を生成（URL が NaN ならリンク無し／空）"""
    text = "" if pd.isna(text) else str(text)
    url = "" if pd.isna(url) else str(url)

    if not text and not url:
        return f'      <td class="{td_class}"></td>\n'

    esc_text = html.escape(text)

    if url:
        esc_url = html.escape(url, quote=True)
        return f'      <td class="{td_class}"><a href="{esc_url}">{esc_text}</a></td>\n'
    else:
        return f'      <td class="{td_class}">{esc_text}</td>\n'


def build_icon_link_cell(icon_filename, url, td_class, img_folder):
    """
    アイコン + URL から <td> を生成。
    url が NaN → 完全に空のセル。
    """
    if pd.isna(url):
        return f'      <td class="{td_class}"></td>\n'

    url = str(url)
    esc_url = html.escape(url, quote=True)

    if pd.isna(icon_filename) or not str(icon_filename):
        # URL はあるが icon が無い → テキストリンクにする
        return f'      <td class="{td_class}"><a href="{esc_url}">link</a></td>\n'

    icon_filename = str(icon_filename)
    esc_icon = html.escape(f"{img_folder}/{icon_filename}", quote=True)
    alt_name = html.escape(icon_filename.rsplit(".", 1)[0])

    return (
        f'      <td class="{td_class}">'
        f'<a href="{esc_url}"><img src="{esc_icon}" alt="{alt_name}" class="ref-icon"></a>'
        f"</td>\n"
    )


def build_mod_cell(row):
    """MOD1〜MOD7 から MOD アイコン群の <td> を作る"""
    imgs = []
    for col in MOD_COLS:
        if col in row and not pd.isna(row[col]) and str(row[col]):
            fname = str(row[col])
            src = html.escape(f"icons_mod/{fname}", quote=True)
            alt = html.escape(fname.rsplit(".", 1)[0])
            imgs.append(f'<img src="{src}" alt="{alt}" class="mod-icon">')

    if not imgs:
        return "      <td></td>\n"

    inner = " ".join(imgs)
    return f"      <td>{inner}</td>\n"


def build_flag_cell(flag_filename):
    if pd.isna(flag_filename) or not str(flag_filename):
        return '      <td class="icon-col"></td>\n'
    src = html.escape(f"icons_flag/{flag_filename}", quote=True)
    return f'      <td class="icon-col"><img src="{src}" alt="" class="flag-icon"></td>\n'


def build_jacket_cell(jacket_filename):
    if pd.isna(jacket_filename) or not str(jacket_filename):
        return '      <td class="icon-col"></td>\n'
    src = html.escape(f"jackets/{jacket_filename}", quote=True)
    return f'      <td class="icon-col"><img src="{src}" alt="jacket" class="jacket"></td>\n'

def fmt_sr(val):
    """SR を小数第2位まで表示"""
    if pd.isna(val):
        return ""
    try:
        return f"{float(val):.2f}"
    except Exception:
        return str(val)
    
def render_sr_pill(row):
    sr = row.get(COL_SR)
    sr_str = fmt_sr(sr)

    # 空の場合は empty cell
    if not sr_str:
        return ""

    # Excel 側で色指定（任意）
    bg = row.get("SR_BG", "").strip() if row.get("SR_BG") else ""
    fg = row.get("SR_FG", "").strip() if row.get("SR_FG") else ""

    style = []
    if bg:
        style.append(f"background-color:{bg}")
    if fg:
        style.append(f"color:{fg}")

    style_attr = f' style="{";".join(style)}"' if style else ""

    return f'<span class="sr-pill"{style_attr}>★ {sr_str}</span>'

def render_sr_cell(row) -> str:
    """SR 列の <td> HTML を生成する."""

    # SR が空なら空セル
    sr = row.get("SR")
    if sr is None or sr == "":
        return "<td class='sr'></td>"

    try:
        sr_value = float(sr)
    except (TypeError, ValueError):
        # 数値化できない場合はそのまま出す（トラブルシュート用）
        label = f"★ {sr}"
        return f"<td class='sr'><span class='sr-pill'>{label}</span></td>"

    # ★ 6.85 の形式（小数第2位まで固定）
    label = f"★ {sr_value:.2f}"

    # Excel 側で指定した色（任意）
    sr_bg = row.get("SR_BG", "") or ""
    sr_fg = row.get("SR_FG", "") or ""

    style_parts = []
    if isinstance(sr_bg, str) and sr_bg.strip():
        style_parts.append(f"background-color:{sr_bg.strip()}")
    if isinstance(sr_fg, str) and sr_fg.strip():
        style_parts.append(f"color:{sr_fg.strip()}")

    style_attr = ""
    if style_parts:
        style_attr = f' style="{";".join(style_parts)}"'

    return f"<td class='sr'><span class='sr-pill'{style_attr}>{label}</span></td>"

def main():
    df = pd.read_excel(DATA_FILE)

    rows_html = []

    for _, row in df.iterrows():
        date_val = row.get(COL_DATE)

        # ---- セクション見出し行 (SRv2 / SRv3 / Majimanji... deletion など) ----
        # DATE が「非日付っぽい文字列」（先頭が数字でない）なら区切り行として扱う
        if isinstance(date_val, str) and not date_val[:1].isdigit():
            text = html.escape(date_val)
            sec_url = None
            if COL_SECTION_LINK in df.columns:
                sec_url = row.get(COL_SECTION_LINK)

            if pd.notna(sec_url) and str(sec_url):
                esc_url = html.escape(str(sec_url), quote=True)
                cell_html = (
                    f'<a href="{esc_url}" target="_blank" rel="noopener noreferrer">'
                    f'{text}'
                    f'</a>'
                )
            else:
                cell_html = text

            # 列数はヘッダーの列数（ここでは 14 列）に合わせる
            rows_html.append(
                f'    <tr class="section-row"><td colspan="17">{cell_html}</td></tr>\n'
            )
            continue

        # DATE が空で、PP / PLAYER / MAP も全部空なら完全な空行としてスキップ
        if pd.isna(date_val) and all(
            pd.isna(row.get(c, None)) for c in [COL_PP, COL_PLAYER, COL_MAP]
        ):
            continue

        # ---- 通常データ行 ----
        date_str = fmt_date(date_val)

        acc_str = fmt_acc(row.get(COL_ACC))

        player = row.get(COL_PLAYER)
        player_url = row.get(COL_PLAYER_LINK)

        map_text = row.get(COL_MAP)
        map_url = row.get(COL_MAP_LINK)

        flag_filename = row.get(COL_FLAG)
        jacket_filename = row.get(COL_JACKET)

        osu_icon = row.get(COL_OSU_ICON)
        osu_url = row.get(COL_OSU_LINK)
        yt_icon = row.get(COL_YT_ICON)
        yt_url = row.get(COL_YT_LINK)
        rd_icon = row.get(COL_RD_ICON)
        rd_url = row.get(COL_RD_LINK)
        x_icon = row.get(COL_X_ICON)
        x_url = row.get(COL_X_LINK)

        remarks = row.get(COL_REMARKS)

        # Days maintained
        days_display = compute_days_display(row)

        # ---- SR ----
        sr_html = render_sr_pill(row)

        # ---- PP ----
        raw_pp = row.get(COL_PP)
        pp_main_text = ""
        if not pd.isna(raw_pp):
            try:
                pp_numeric = float(raw_pp)
                pp_main_text = f"{int(round(pp_numeric)):,} pp"
            except Exception:
                pp_main_text = str(raw_pp)

        # ---- PP DIFF ----
        pp_diff_raw = row.get(COL_PP_DIFF) if COL_PP_DIFF in df.columns else None
        pp_diff_text = ""
        pp_diff_class = "ppdiff-zero"  # デフォルトはグレー扱い

        if not pd.isna(pp_diff_raw) and str(pp_diff_raw).strip():
            try:
                d = float(str(pp_diff_raw))
                sign = "+" if d >= 0 else "-"
                d_int = int(round(abs(d)))
                pp_diff_text = f"({sign}{d_int})"

                if d > 0:
                    pp_diff_class = "ppdiff-pos"
                elif d < 0:
                    pp_diff_class = "ppdiff-neg"
                else:
                    pp_diff_class = "ppdiff-zero"

            except Exception:
                # 数値として扱えない → グレー扱い
                pp_diff_text = str(pp_diff_raw)
                pp_diff_class = "ppdiff-zero"
        else:
            pp_diff_text = ""
            pp_diff_class = "ppdiff-zero"

        # ここから行の組み立て
        row_parts = []
        row_parts.append("    <tr>\n")

        # Date
        row_parts.append(f'      <td class="col-date">{html.escape(date_str)}</td>\n')

        # Days
        row_parts.append(
            f'      <td class="col-days">{html.escape(days_display)}</td>\n'
        )

        # Flag
        row_parts.append(build_flag_cell(flag_filename))

        # Player
        row_parts.append(build_link_cell(player, player_url, "col-player"))

        # Jacket
        row_parts.append(build_jacket_cell(jacket_filename))

        # Map
        row_parts.append(build_link_cell(map_text, map_url, "col-map"))

        # SR
        row_parts.append(f'      <td class="col-sr">{sr_html}</td>\n')

        # MODs
        row_parts.append(build_mod_cell(row))

        # ACC
        row_parts.append(f'      <td class="col-acc">{html.escape(acc_str)}</td>\n')

        # PP（1,749 pp (+83)）
        row_parts.append(f'      <td class="col-pp">{html.escape(pp_main_text)}</td>\n')

        # PP diff
        row_parts.append(
            f'      <td class="col-ppdiff {pp_diff_class}">{html.escape(pp_diff_text)}</td>\n'
        )

        # osu / YouTube / reddit / X
        row_parts.append(
            build_icon_link_cell(osu_icon, osu_url, "icon-col", "ref_icons")
        )
        row_parts.append(
            build_icon_link_cell(yt_icon, yt_url, "icon-col", "ref_icons")
        )
        row_parts.append(
            build_icon_link_cell(rd_icon, rd_url, "icon-col", "ref_icons")
        )
        row_parts.append(
            build_icon_link_cell(x_icon, x_url, "icon-col", "ref_icons")
        )

        # Replay
        REPLAYS_DIR = Path("replays")

        # --- Replay 列 ---
        raw_replay = row.get(COL_REPLAY)

        # 1) 本当に有効な文字列かどうかを判定して "クリーンな値" にする
        replay_val = ""

        if raw_replay is None:
            replay_val = ""
        elif isinstance(raw_replay, float):
            # pandas の NaN 対応（空セル → NaN になっていることが多い）
            if math.isnan(raw_replay):
                replay_val = ""
            else:
                replay_val = str(raw_replay).strip()
        else:
            replay_val = str(raw_replay).strip()

        # 2) クリーンな値が空なら → 完全に空セル
        if not replay_val:
            row_parts.append('      <td class="col-replay"></td>\n')
        else:
            # ファイル存在チェック（任意だけど入れておくと安全）
            replay_path = REPLAYS_DIR / replay_val

            if replay_path.exists():
                href = html.escape(str(replay_path).replace("\\", "/"), quote=True)
                row_parts.append(
                    f'      <td class="col-replay">'
                    f'<a class="replay-link" href="{href}" download>Replay</a>'
                    f'</td>\n'
                )
            else:
                # ファイルがない場合：ボタンを出さずに空欄 & コンソールに警告
                print(f"[WARN] Replay file not found: {replay_path}")
                row_parts.append('      <td class="col-replay"></td>\n')

        # Remarks（空欄なら空文字）
        if pd.isna(remarks) or str(remarks).strip() == "":
            remarks_text = ""
        else:
            remarks_text = str(remarks)

        row_parts.append(
            f'      <td class="col-remarks">{html.escape(remarks_text)}</td>\n'
        )

        row_parts.append("    </tr>\n")

        rows_html.append("".join(row_parts))

    html_output = HTML_HEAD + "".join(rows_html) + HTML_GITHUB_LINK + HTML_FOOT
    OUTPUT_FILE.write_text(html_output, encoding="utf-8")
    print(f"生成完了: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
