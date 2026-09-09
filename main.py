import json
import base64
import math
import threading
import urllib.request
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog, simpledialog
import csv
import re
import urllib.parse
import os
import zipfile
import subprocess
import sys
import webbrowser
from io import BytesIO
from datetime import datetime, timedelta, timezone
try:
    from bs4 import BeautifulSoup
    BS4_OK = True
except Exception:
    BeautifulSoup = None
    BS4_OK = False

try:
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    OPENPYXL_OK = True
except Exception:
    OPENPYXL_OK = False
    Workbook = XLImage = Font = PatternFill = Alignment = Border = Side = DataValidation = None

from pathlib import Path
try:
    from PIL import Image, ImageTk
    PIL_OK = True
except Exception:
    PIL_OK = False

# ============================================================
# SinKA PvP - V114 ELITE 90+ TEAM ENGINE
# ============================================================

RANKING_URL = "https://raw.githubusercontent.com/pvpoke/pvpoke/master/src/data/rankings/all/overall/rankings-{cp}.json"
POKEMON_URL = "https://raw.githubusercontent.com/pvpoke/pvpoke/master/src/data/gamemaster/pokemon.json"
TR_MOVES_URL = "https://raw.githubusercontent.com/WatWowMap/pogo-data-api/main/data/v1/translations/tr/moves.json"
DITTOBASE_MAX_ATTACKERS_URL = "https://www.dittobase.com/pokemon-go/best-attackers/max-battles/overall"
MEGA_RANKING_URL = "https://raw.githubusercontent.com/pvpoke/pvpoke/master/src/data/rankings/mega/overall/rankings-10000.json"
POGO_MOVES_URL = "https://raw.githubusercontent.com/WatWowMap/pogo-data-api/main/data/v1/moves.json"
RAID_SOURCE_URL = "https://leekduck.com/events/"
RAID_IMAGE_SOURCE_URL = "https://leekduck.com/raid-bosses/"

# Dashboard colors (global so background refresh callbacks can use them too)
WHITE = "#ffffff"
TEXT = "#152b4d"
MUTED = "#6d7e98"
BLUE = "#1d65da"
GREEN = "#008f58"
BORDER = "#d9e2ef"

LEAGUES = {
    "Great League": 1500,
    "Ultra League": 2500,
    "Master League": 10000,
    "Mega": 10000,
    "Gigamax": None,
}

CPM_VALUES = [
0.094,0.135137432,0.16639787,0.192650919,0.21573247,0.236572661,
0.25572005,0.273530381,0.29024988,0.306057377,0.3210876,0.335445036,
0.34921268,0.362457751,0.37523559,0.387592406,0.39956728,0.411193551,
0.42250001,0.432926419,0.44310755,0.453059958,0.46279839,0.472336083,
0.48168495,0.4908558,0.49985844,0.508701765,0.51739395,0.525942511,
0.53435433,0.542635767,0.55079269,0.558830576,0.56675452,0.574569153,
0.58227891,0.589887917,0.59740001,0.604818814,0.61215729,0.619399365,
0.62656713,0.633644533,0.64065295,0.647576426,0.65443563,0.661214806,
0.667934,0.674577537,0.68116492,0.687680648,0.69414365,0.700538673,
0.70688421,0.713164996,0.71939909,0.725571552,0.7317,0.734741009,
0.73776948,0.740785574,0.74378943,0.746781211,0.74976104,0.752729087,
0.75568551,0.758630378,0.76156384,0.764486065,0.76739717,0.770297266,
0.7731865,0.776064962,0.77893275,0.781790055,0.78463697,0.787473578,
0.79030001,0.79280395,0.79530001,0.79780391,0.80030001,0.80280389,
0.80530001,0.80780390,0.81030001,0.81280390,0.81530001,0.81780390,
0.82030001,0.82280390,0.82530001,0.82780390,0.83030001,0.83280390,
0.83530003,0.83780375,0.84030003,0.84280373,0.84530002
]
CPM = {1.0 + i * 0.5: v for i, v in enumerate(CPM_VALUES)}
LEVELS = list(CPM.keys())

FALLBACK_TR = {
    "ROLLOUT": "Yuvarlanma",
    "BODY_SLAM": "Vücut Çarpması",
    "SHADOW_BALL": "Gölge Topu",
    "EARTHQUAKE": "Deprem",
    "HYPER_BEAM": "Hiper Işın",
    "SOLAR_BEAM": "Güneş Işını",
}

def _decode_json_response(raw_bytes, url=""):
    """JSON indirirken BOM/XSSI ve arka arkaya gelen JSON bloklarını güvenli biçimde işler."""
    text = raw_bytes.decode("utf-8-sig", errors="replace").strip()
    # Bazı sunucular JSON'un başına XSSI koruma öneki koyabilir.
    for prefix in (")]}\'\n", ")]}'\n", ")]}'", "while(1);\n"):
        if text.startswith(prefix):
            text = text[len(prefix):].lstrip()
            break

    decoder = json.JSONDecoder()
    try:
        value, end = decoder.raw_decode(text)
    except json.JSONDecodeError:
        # Normal hata mesajını koru; hangi URL'nin sorun çıkardığı da görünsün.
        raise ValueError(f"JSON okunamadı: {url}\n{ text[:160] }")

    # 'Extra data' durumunda ikinci JSON bloğu da varsa birleştir.
    rest = text[end:].strip()
    if rest:
        try:
            second, end2 = decoder.raw_decode(rest)
            if not rest[end2:].strip():
                if isinstance(value, list) and isinstance(second, list):
                    value = value + second
                elif isinstance(value, dict) and isinstance(second, dict):
                    value = {**value, **second}
        except json.JSONDecodeError:
            # Sondaki bozuk/HTML içerik veri kaynağını tamamen düşürmesin.
            pass
    return value

def download_json(url):
    req = urllib.request.Request(url, headers={
        "User-Agent": "SinKa-PvP-100/2.0",
        "Accept": "application/json,text/plain,*/*",
    })
    with urllib.request.urlopen(req, timeout=30) as response:
        return _decode_json_response(response.read(), url)

def build_move_translator(pogo_moves, tr_moves):
    proto_to_id = {}
    for m in pogo_moves:
        proto = m.get("proto")
        mid = m.get("moveId")
        if proto is not None and mid is not None:
            proto_to_id[str(proto).upper()] = int(mid)

    def translate(move_id):
        mid = proto_to_id.get(str(move_id).upper())
        if mid is not None:
            value = tr_moves.get("move_" + str(mid))
            if value and not value.startswith("<<"):
                return value
        return FALLBACK_TR.get(str(move_id).upper(), str(move_id).replace("_", " ").title())

    return translate

def calc_cp(base, ivs, level):
    cpm = CPM[level]
    atk = (base["atk"] + ivs[0]) * cpm
    defense = (base["def"] + ivs[1]) * cpm
    stamina = (base["hp"] + ivs[2]) * cpm
    return math.floor(atk * math.sqrt(defense) * math.sqrt(stamina) / 10)

def stat_product(base, ivs, level):
    cpm = CPM[level]
    atk = (base["atk"] + ivs[0]) * cpm
    defense = (base["def"] + ivs[1]) * cpm
    hp = math.floor((base["hp"] + ivs[2]) * cpm)
    return atk * defense * hp

def rank1_iv(base, cp_limit):
    best_key = None
    best_ivs = (15, 15, 15)
    best_level = 50.0
    best_cp = 10

    if cp_limit >= 10000:
        level = 50.0
        return (15, 15, 15), level, calc_cp(base, (15, 15, 15), level)

    for atk_iv in range(16):
        for def_iv in range(16):
            for hp_iv in range(16):
                ivs = (atk_iv, def_iv, hp_iv)

                lo, hi = 0, len(LEVELS) - 1
                legal = -1
                while lo <= hi:
                    mid = (lo + hi) // 2
                    lvl = LEVELS[mid]
                    if calc_cp(base, ivs, lvl) <= cp_limit:
                        legal = mid
                        lo = mid + 1
                    else:
                        hi = mid - 1

                if legal < 0:
                    continue

                level = LEVELS[legal]
                cp = calc_cp(base, ivs, level)
                product = stat_product(base, ivs, level)

                cpm = CPM[level]
                eff_def = (base["def"] + def_iv) * cpm
                eff_hp = math.floor((base["hp"] + hp_iv) * cpm)

                key = (product, cp, eff_def, eff_hp, -atk_iv)

                if best_key is None or key > best_key:
                    best_key = key
                    best_ivs = ivs
                    best_level = level
                    best_cp = cp

    return best_ivs, best_level, best_cp

def is_shadow(item):
    sid = item.get("speciesId", "")
    return sid.endswith("_shadow") or "(Shadow)" in item.get("speciesName", "")

def is_gigantamax(item):
    sid = item.get("speciesId", "").lower()
    sname = item.get("speciesName", "").lower()
    return "gigantamax" in sid or "gmax" in sid or "gigantamax" in sname or "g-max" in sname


# ============================================================
# SİNKA TURN-BY-TURN BATTLE ENGINE
# ============================================================
# Bu motor PvPoke'nin açık Game Master / battle mekaniklerinden esinlenen,
# SinKA içinde bağımsız çalışan bir Python simülatörüdür. Amaç: ranking
# rating'ini tekrar puanlamak yerine gerçek move/energy/turn/HP/shield akışını
# takım motoruna dahil etmek. Tam PvPoke JS kaynak kodu portu değildir.

PVPoke_MOVES_URL = "https://raw.githubusercontent.com/pvpoke/pvpoke/master/src/data/gamemaster/moves.json"


class SinkaBattleEngine:
    """SinKA için deterministik, turn tabanlı PvP savaş simülatörü.

    Desteklenen çekirdek mekanikler:
      - 500 ms turn mantığı / fast move turn süreleri
      - Energy 0-100
      - Charged Move maliyeti
      - CMP (Attack karşılaştırması + eşitlikte deterministic tie-break)
      - Shield
      - STAB + tip effectiveness
      - Shadow Attack/Defense çarpanları
      - Charged Move buff/debuff'ları ve deterministic proc meter
      - Mimikyu Disguise özel koruması
      - Bait / no-bait karar politikaları
      - Başlangıç energy avantajı

    Takım motoru bunu 1v1 çekirdek sonuçları olarak kullanır; 3v3 alignment
    ve switch değerlendirmesi üst katmanda yapılır.
    """

    STAB = 1.200000047316284
    SHADOW_ATK = 1.2
    SHADOW_DEF = 0.83333331
    BONUS = 1.2999999523162842
    SUPER = 1.600000023841858
    RESISTED = 0.625
    DOUBLE_RESISTED = 0.390625

    TYPE_TRAITS = {
        "normal": ({"fighting"}, {"ghost"}),
        "fighting": ({"flying", "psychic", "fairy"}, {"rock", "bug", "dark"}),
        "flying": ({"rock", "electric", "ice"}, {"fighting", "bug", "grass"}),
        "poison": ({"ground", "psychic"}, {"fighting", "poison", "bug", "fairy", "grass"}),
        "ground": ({"water", "grass", "ice"}, {"poison", "rock"}),
        "rock": ({"fighting", "ground", "steel", "water", "grass"}, {"normal", "flying", "poison", "fire"}),
        "bug": ({"flying", "rock", "fire"}, {"fighting", "ground", "grass"}),
        "ghost": ({"ghost", "dark"}, {"poison", "bug"}),
        "steel": ({"fighting", "ground", "fire"}, {"normal", "flying", "rock", "bug", "steel", "grass", "psychic", "ice", "dragon", "fairy"}),
        "fire": ({"ground", "rock", "water"}, {"bug", "steel", "fire", "grass", "ice", "fairy"}),
        "water": ({"grass", "electric"}, {"steel", "fire", "water", "ice"}),
        "grass": ({"flying", "poison", "bug", "fire", "ice"}, {"ground", "water", "grass", "electric"}),
        "electric": ({"ground"}, {"flying", "steel", "electric"}),
        "psychic": ({"bug", "ghost", "dark"}, {"fighting", "psychic"}),
        "ice": ({"fighting", "fire", "steel", "rock"}, {"ice"}),
        "dragon": ({"dragon", "ice", "fairy"}, {"fire", "water", "grass", "electric"}),
        "dark": ({"fighting", "fairy", "bug"}, {"ghost", "dark"}),
        "fairy": ({"poison", "steel"}, {"fighting", "bug", "dark"}),
    }
    TYPE_IMMUNITIES = {
        "normal": {"ghost"}, "flying": {"ground"}, "ground": {"electric"},
        "ghost": {"normal", "fighting"}, "dark": {"psychic"}, "steel": {"poison"},
        "fairy": {"dragon"},
    }

    def __init__(self, moves=None, pokemon_data=None):
        self.moves = moves or {}
        self.pokemon_data = pokemon_data or {}
        self._result_cache = {}
        self._mon_cache = {}
        self._team_result_cache = {}

    @staticmethod
    def _num(value, default=0.0):
        try:
            return float(value)
        except Exception:
            return float(default)

    def _move(self, move_id):
        if not move_id:
            return None
        m = self.moves.get(str(move_id))
        if not isinstance(m, dict):
            return None
        return m

    def _types(self, mon):
        return tuple(t for t in (mon.get("types") or []) if t and t != "none")

    def effectiveness(self, move_type, defender_types):
        mult = 1.0
        mt = str(move_type or "").lower()
        for target in defender_types:
            target = str(target).lower()
            if mt in self.TYPE_IMMUNITIES.get(target, set()):
                mult *= self.DOUBLE_RESISTED
                continue
            weak, resist = self.TYPE_TRAITS.get(target, (set(), set()))
            if mt in weak:
                mult *= self.SUPER
            elif mt in resist:
                mult *= self.RESISTED
        return mult

    def _rank_moveset(self, p):
        moveset = p.get("moveset") or p.get("moves") or []
        if isinstance(moveset, dict):
            moveset = list(moveset.values())
        if not isinstance(moveset, (list, tuple)):
            moveset = [moveset]
        moveset = [str(x) for x in moveset if x]
        if len(moveset) >= 3:
            return moveset[0], moveset[1:3]
        sid = str(p.get("speciesId") or "")
        gm = self.pokemon_data.get(sid) or self.pokemon_data.get(sid.replace("_shadow", "")) or {}
        fast = list(gm.get("fastMoves") or [])
        charged = list(gm.get("chargedMoves") or [])
        return (moveset[0] if moveset else (fast[0] if fast else None),
                (moveset[1:3] if len(moveset) > 1 else charged[:2]))

    def make_mon(self, p):
        sid = str(p.get("speciesId") or p.get("speciesName") or p.get("name") or "")
        moveset = p.get("moveset") or p.get("moves") or []
        if isinstance(moveset, dict):
            moveset = tuple(moveset.values())
        elif isinstance(moveset, (list, tuple)):
            moveset = tuple(moveset)
        else:
            moveset = (moveset,) if moveset else ()
        stats = p.get("stats") or {}
        cache_key = (sid, tuple(str(x) for x in moveset),
                     str(stats.get("atk", stats.get("attack", ""))),
                     str(stats.get("def", stats.get("defense", ""))),
                     str(stats.get("hp", stats.get("stamina", ""))))
        cached = self._mon_cache.get(cache_key)
        if cached is not None:
            return cached
        gm = self.pokemon_data.get(sid) or self.pokemon_data.get(sid.replace("_shadow", "")) or {}
        types = list(p.get("types") or gm.get("types") or [])
        stats = p.get("stats") or {}
        base = gm.get("baseStats") or {}
        atk = self._num(stats.get("atk", stats.get("attack", base.get("atk", 0))))
        defense = self._num(stats.get("def", stats.get("defense", base.get("def", 0))))
        hp = self._num(stats.get("hp", stats.get("stamina", base.get("hp", 0))))
        fast_id, charged_ids = self._rank_moveset(p)
        fast = self._move(fast_id)
        charged = [self._move(x) for x in charged_ids]
        charged = [x for x in charged if x]
        if not fast or not charged:
            return None
        shadow = str(sid).endswith("_shadow") or "(Shadow)" in str(p.get("speciesName") or "")
        # Ranking stats are already effective stats at the ranking IV/level.
        # Shadow multiplier is applied here because ranking stats don't include it.
        result = {
            "sid": sid, "name": str(p.get("speciesName") or p.get("name") or sid),
            "types": tuple(types), "atk": atk, "def": defense, "hp": max(1.0, hp),
            "shadow": shadow, "fast": fast, "charged": charged,
        }
        self._mon_cache[cache_key] = result
        return result

    def _stage_mult(self, stage):
        stage = max(-4, min(4, int(stage)))
        if stage >= 0:
            return (4 + stage) / 4.0
        return 4.0 / (4 - stage)

    def _effective_atk(self, state):
        return state["atk"] * self._stage_mult(state["atk_stage"]) * (self.SHADOW_ATK if state["shadow"] else 1.0)

    def _effective_def(self, state):
        return state["def"] * self._stage_mult(state["def_stage"]) * (self.SHADOW_DEF if state["shadow"] else 1.0)

    def _damage(self, attacker, defender, move):
        eff = self.effectiveness(move.get("type"), defender["types"])
        stab = self.STAB if str(move.get("type")) in attacker["types"] else 1.0
        power = self._num(move.get("power"), 0)
        if str(move.get("damageMethod")) == "percentMaxHP":
            return max(1, math.floor((power / 100.0) * defender["hp_max"]) + 1)
        value = power * stab * (self._effective_atk(attacker) / max(1.0, self._effective_def(defender))) * eff * 0.5 * self.BONUS
        return max(1, math.floor(value) + 1)

    def _charge_candidates(self, state, opponent, bait_mode):
        available = [m for m in state["charged"] if state["energy"] >= self._num(m.get("energy"), 100)]
        if not available:
            return []
        # Lethal always wins over bait.
        lethal = [m for m in available if self._damage(state, opponent, m) >= opponent["hp"]]
        if lethal:
            return sorted(lethal, key=lambda m: (self._num(m.get("energy"), 100), -self._damage(state, opponent, m)))
        if bait_mode and len(available) > 1:
            # Cheapest reasonable move is the bait candidate; keep a heavier move as fallback.
            return sorted(available, key=lambda m: (self._num(m.get("energy"), 100), -self._damage(state, opponent, m)))
        return sorted(available, key=lambda m: (-self._damage(state, opponent, m), self._num(m.get("energy"), 100)))

    def _should_shield(self, defender, attacker, move, bait_mode):
        if defender["shields"] <= 0:
            return False
        damage = self._damage(attacker, defender, move)
        energy = self._num(move.get("energy"), 100)
        if damage >= defender["hp"]:
            return True
        if bait_mode:
            # Aggressive shield policy: cheap moves can draw shields, while large nukes
            # are almost always shielded when they threaten ~30%+ HP.
            return (energy <= 45 and damage >= defender["hp"] * 0.18) or damage >= defender["hp"] * 0.30
        return damage >= defender["hp"] * 0.38 or energy >= 60 and damage >= defender["hp"] * 0.25

    def _apply_buffs(self, attacker, defender, move, meters):
        buffs = move.get("buffs")
        if not buffs:
            return
        try:
            chance = float(move.get("buffApplyChance", 1) or 0)
        except Exception:
            chance = 0.0
        key = (attacker["sid"], str(move.get("moveId")))
        meters[key] = meters.get(key, 0.0) + chance
        if meters[key] < 1.0:
            return
        meters[key] -= 1.0
        target = attacker if move.get("buffTarget") == "self" else defender
        vals = list(buffs) + [0, 0]
        target["atk_stage"] = max(-4, min(4, target["atk_stage"] + int(vals[0] or 0)))
        target["def_stage"] = max(-4, min(4, target["def_stage"] + int(vals[1] or 0)))

    def simulate(self, p1, p2, shields=(2, 2), energy_delta=0, bait_mode=True, max_turns=240):
        """Tek 1v1 battle döndürür. Result score: p1 perspektifinden 0-100."""
        m1, m2 = self.make_mon(p1), self.make_mon(p2)
        if not m1 or not m2:
            return None
        s = []
        for i, mon in enumerate((m1, m2)):
            energy = 0.0
            if i == 0:
                energy = max(0.0, min(100.0, float(energy_delta)))
            state = dict(mon)
            state.update({"hp_max": mon["hp"], "hp": mon["hp"], "energy": energy,
                          "shields": int(shields[i]), "atk_stage": 0, "def_stage": 0,
                          "cooldown": 0, "busted": False})
            s.append(state)
        # Symmetric negative energy when requested.
        if energy_delta < 0:
            s[0]["energy"] = 0.0
            s[1]["energy"] = min(100.0, -float(energy_delta))

        turn = 0
        cmp_count = 0
        bait_attempts = 0
        bait_successes = 0
        buff_meters = {}
        fast_count = [0, 0]
        charge_count = [0, 0]

        while s[0]["hp"] > 0 and s[1]["hp"] > 0 and turn < max_turns:
            turn += 1
            ready = []
            for i in (0, 1):
                s[i]["cooldown"] = max(0, int(s[i]["cooldown"]) - 1)
                ready.append(s[i]["cooldown"] == 0)

            actions = []
            for i, j in ((0, 1), (1, 0)):
                if not ready[i] or s[i]["hp"] <= 0:
                    actions.append(None)
                    continue
                choices = self._charge_candidates(s[i], s[j], bait_mode)
                if choices:
                    move = choices[0]
                    # Bait only if there is a cheaper move and a materially stronger second move.
                    if bait_mode and len(choices) > 1:
                        cheapest, strongest = choices[0], choices[-1]
                        if self._num(cheapest.get("energy"), 100) <= 45 and self._damage(s[i], s[j], strongest) > self._damage(s[i], s[j], cheapest) * 1.22:
                            move = cheapest
                            bait_attempts += 1
                    actions.append(("charged", move))
                else:
                    actions.append(("fast", s[i]["fast"]))

            # Charged moves act first. If both are charged, CMP decides.
            order = [0, 1]
            if actions[0] and actions[1] and actions[0][0] == actions[1][0] == "charged":
                atk0 = self._effective_atk(s[0]); atk1 = self._effective_atk(s[1])
                if abs(atk0 - atk1) < 1e-9:
                    cmp_count += 1
                    # deterministic species-id tie break keeps the engine reproducible.
                    order = [0, 1] if s[0]["sid"] <= s[1]["sid"] else [1, 0]
                else:
                    order = [0, 1] if atk0 > atk1 else [1, 0]
            elif actions[0] and actions[1] and actions[0][0] == "charged":
                order = [0, 1]
            elif actions[0] and actions[1] and actions[1][0] == "charged":
                order = [1, 0]

            for i in order:
                j = 1 - i
                action = actions[i]
                if not action or s[i]["hp"] <= 0 or s[j]["hp"] <= 0:
                    continue
                kind, move = action
                if kind == "fast":
                    dmg = self._damage(s[i], s[j], move)
                    s[j]["hp"] -= dmg
                    s[i]["energy"] = min(100.0, s[i]["energy"] + self._num(move.get("energyGain"), 0))
                    s[i]["cooldown"] = max(0, int(move.get("turns") or round(self._num(move.get("cooldown"), 500) / 500.0)) - 1)
                    fast_count[i] += 1
                else:
                    cost = self._num(move.get("energy"), 100)
                    if s[i]["energy"] < cost:
                        continue
                    s[i]["energy"] -= cost
                    charge_count[i] += 1
                    use_shield = self._should_shield(s[j], s[i], move, bait_mode)
                    if use_shield:
                        s[j]["shields"] -= 1
                        damage = 1
                        if bait_mode and cost <= 45:
                            bait_successes += 1
                    else:
                        damage = self._damage(s[i], s[j], move)
                        # Mimikyu Disguise: first unshielded charged hit is protected,
                        # then its defense drops one stage permanently.
                        if s[j]["sid"] in ("mimikyu", "mimikyu_busted") and not s[j]["busted"]:
                            damage = 1
                            s[j]["busted"] = True
                            s[j]["def_stage"] = max(-4, s[j]["def_stage"] - 1)
                        else:
                            self._apply_buffs(s[i], s[j], move, buff_meters)
                    s[j]["hp"] -= damage
                    s[i]["cooldown"] = 0

                    if s[j]["hp"] <= 0:
                        break

            # If a fast move was selected but a same-turn charged move KO'd it,
            # the later fast action naturally gets skipped above.

        hp0 = max(0.0, s[0]["hp"])
        hp1 = max(0.0, s[1]["hp"])
        if hp0 > 0 and hp1 <= 0:
            outcome = 1.0
        elif hp1 > 0 and hp0 <= 0:
            outcome = 0.0
        else:
            # Time limit / tie: remaining HP + energy + shields decide.
            v0 = hp0 / s[0]["hp_max"] + 0.004 * s[0]["energy"] + 0.08 * s[0]["shields"]
            v1 = hp1 / s[1]["hp_max"] + 0.004 * s[1]["energy"] + 0.08 * s[1]["shields"]
            outcome = 1.0 if v0 > v1 else (0.0 if v1 > v0 else 0.5)

        margin = (hp0 / s[0]["hp_max"]) - (hp1 / s[1]["hp_max"])
        score = 50.0 + 50.0 * (outcome - 0.5) + 18.0 * margin
        score = max(0.0, min(100.0, score))
        return {
            "score": score,
            "winner": 0 if outcome == 1 else (1 if outcome == 0 else None),
            "turns": turn,
            "hp": [hp0, hp1],
            "hp_pct": [100.0 * hp0 / s[0]["hp_max"], 100.0 * hp1 / s[1]["hp_max"]],
            "energy": [s[0]["energy"], s[1]["energy"]],
            "shields": [s[0]["shields"], s[1]["shields"]],
            "cmp": cmp_count,
            "bait_attempts": bait_attempts,
            "bait_successes": bait_successes,
            "fast_moves": fast_count,
            "charged_moves": charge_count,
        }

    def simulate_3v3(self, team_a, team_b, lead_a=0, lead_b=0, shields=2,
                     rating_lookup=None, bait_mode=True, max_turns=720):
        """Stateful 3v3 team battle. Not a verbatim PvPoke JS port."""
        if not isinstance(team_a, (list, tuple)) or not isinstance(team_b, (list, tuple)) or len(team_a) != 3 or len(team_b) != 3:
            return None
        cache_key = (tuple(str(p.get("speciesId") or p.get("speciesName") or p.get("name") or "") for p in team_a),
                     tuple(str(p.get("speciesId") or p.get("speciesName") or p.get("name") or "") for p in team_b),
                     int(lead_a), int(lead_b), int(shields), bool(bait_mode), int(max_turns))
        cached = self._team_result_cache.get(cache_key)
        if cached is not None:
            return cached

        def make_state(p):
            mon = self.make_mon(p)
            if not mon:
                return None
            st = dict(mon)
            st.update({"hp_max":mon["hp"],"hp":mon["hp"],"energy":0.0,"shields":int(shields),
                       "atk_stage":0,"def_stage":0,"cooldown":0,"busted":False,"fainted":False})
            return st

        a = [make_state(p) for p in team_a]
        b = [make_state(p) for p in team_b]
        if any(x is None for x in a+b):
            return None
        active = [max(0,min(2,int(lead_a))), max(0,min(2,int(lead_b)))]
        used_switches = [0,0]
        switch_lock = [0,0]
        buff_meters = {}
        cmp_count = bait_attempts = bait_successes = switch_count = farm_events = 0
        turns = 0

        def sid(st): return str(st.get("sid") or "")
        def rating(attacker, defender):
            if rating_lookup:
                try:
                    v = rating_lookup(sid(attacker), sid(defender))
                    if v is not None: return float(v)
                except Exception: pass
            vals=[]
            for mv in [attacker.get("fast")]+list(attacker.get("charged") or []):
                if mv: vals.append(self.effectiveness(mv.get("type"), defender.get("types",())))
            avg=sum(vals)/max(1,len(vals))
            return 500.0+(avg-1.0)*220.0
        def alive(pool): return [i for i,st in enumerate(pool) if st["hp"]>0]
        def best_switch(side):
            own=a if side==0 else b
            opp=b[active[1]] if side==0 else a[active[0]]
            choices=[i for i in alive(own) if i!=active[side]]
            if not choices: return None,-999.0
            return max(((rating(own[i],opp),i) for i in choices), key=lambda x:x[0])[1::-1]
        def maybe_switch(side):
            if used_switches[side]>=2 or switch_lock[side]>0 or turns<3: return False
            own=a if side==0 else b; opp=b[active[1]] if side==0 else a[active[0]]
            cur=own[active[side]]
            if cur["hp"]<=0: return False
            cur_r=rating(cur,opp); idx,best_r=best_switch(side)
            if idx is None: return False
            hp_pct=cur["hp"]/max(1.0,cur["hp_max"])
            trigger=(cur_r<450 and best_r>=cur_r+45 and hp_pct<0.82) or (hp_pct<0.30 and best_r>=cur_r+25)
            if not trigger: return False
            active[side]=idx; used_switches[side]+=1; switch_lock[side]=1
            return True

        while turns<max_turns and alive(a) and alive(b):
            turns+=1
            switch_lock[0]=max(0,switch_lock[0]-1); switch_lock[1]=max(0,switch_lock[1]-1)
            for side,own in ((0,a),(1,b)):
                if own[active[side]]["hp"]<=0:
                    own[active[side]]["fainted"]=True
                    idx,_=best_switch(side)
                    if idx is not None: active[side]=idx; switch_count+=1; switch_lock[side]=1
            hp0=a[active[0]]["hp"]/max(1.0,a[active[0]]["hp_max"]); hp1=b[active[1]]["hp"]/max(1.0,b[active[1]]["hp_max"])
            switched=False
            for side in ((0,1) if hp0<=hp1 else (1,0)):
                if maybe_switch(side): switch_count+=1; switched=True
            if switched: continue

            sa,sb=a[active[0]],b[active[1]]
            sa["cooldown"]=max(0,int(sa["cooldown"])-1); sb["cooldown"]=max(0,int(sb["cooldown"])-1)
            actions=[None,None]
            for side,(actor,target) in enumerate(((sa,sb),(sb,sa))):
                if actor["cooldown"]!=0: continue
                choices=self._charge_candidates(actor,target,bait_mode)
                if choices:
                    move=choices[0]
                    if bait_mode and len(choices)>1:
                        cheap,strong=choices[0],choices[-1]
                        if self._num(cheap.get("energy"),100)<=45 and self._damage(actor,target,strong)>self._damage(actor,target,cheap)*1.22:
                            move=cheap; bait_attempts+=1
                    actions[side]=("charged",move)
                else: actions[side]=("fast",actor["fast"])
            order=[0,1]
            if actions[0] and actions[1] and actions[0][0]==actions[1][0]=="charged":
                aa,ab=self._effective_atk(sa),self._effective_atk(sb)
                if abs(aa-ab)<1e-9: cmp_count+=1; order=[0,1] if sid(sa)<=sid(sb) else [1,0]
                else: order=[0,1] if aa>ab else [1,0]
            elif actions[1] and actions[1][0]=="charged": order=[1,0]

            for side in order:
                own=a if side==0 else b; opp=b if side==0 else a
                actor=own[active[side]]; target=opp[active[1-side]]; action=actions[side]
                if not action or actor["hp"]<=0 or target["hp"]<=0: continue
                kind,move=action
                if kind=="fast":
                    target["hp"]-=self._damage(actor,target,move)
                    actor["energy"]=min(100.0,actor["energy"]+self._num(move.get("energyGain"),0))
                    actor["cooldown"]=max(0,int(move.get("turns") or round(self._num(move.get("cooldown"),500)/500.0))-1)
                    if actor["energy"]>=80: farm_events+=1
                else:
                    cost=self._num(move.get("energy"),100)
                    if actor["energy"]<cost: continue
                    actor["energy"]-=cost
                    use_shield=self._should_shield(target,actor,move,bait_mode)
                    if use_shield:
                        target["shields"]-=1; damage=1
                        if bait_mode and cost<=45: bait_successes+=1
                    else:
                        damage=self._damage(actor,target,move)
                        if target["sid"] in ("mimikyu","mimikyu_busted") and not target["busted"]:
                            damage=1; target["busted"]=True; target["def_stage"]=max(-4,target["def_stage"]-1)
                        else: self._apply_buffs(actor,target,move,buff_meters)
                    target["hp"]-=damage; actor["cooldown"]=0
                    if target["hp"]<=0:
                        target["fainted"]=True
                        idx,_=best_switch(1-side)
                        if idx is not None: active[1-side]=idx; switch_count+=1; switch_lock[1-side]=1
                        else: break

        alive_a,alive_b=alive(a),alive(b)
        def team_value(pool): return sum(st["hp"]/max(1.0,st["hp_max"])+0.0025*st["energy"]+0.05*st["shields"] for st in pool if st["hp"]>0)
        if alive_a and not alive_b: outcome=1.0
        elif alive_b and not alive_a: outcome=0.0
        else:
            va,vb=team_value(a),team_value(b); outcome=1.0 if va>vb else (0.0 if vb>va else 0.5)
        va,vb=team_value(a),team_value(b)
        score=max(0.0,min(100.0,50.0+50.0*(outcome-0.5)+8.0*(va-vb)))
        result = {"score":score,"winner":0 if outcome==1 else (1 if outcome==0 else None),"turns":turns,
                "alive":[len(alive_a),len(alive_b)],"team_value":[va,vb],"cmp":cmp_count,
                "bait_attempts":bait_attempts,"bait_successes":bait_successes,"switches":switch_count,
                "farm_events":farm_events,"active":list(active),
                "hp_pct":[[round(max(0.0,st["hp"])/max(1.0,st["hp_max"])*100,1) for st in a],
                           [round(max(0.0,st["hp"])/max(1.0,st["hp_max"])*100,1) for st in b]]}
        self._team_result_cache[cache_key] = result
        return result

    def scenario(self, p1, p2, shields=(1, 1), energy_delta=0, bait=True):
        key = (str(p1.get("speciesId")), str(p2.get("speciesId")), int(shields[0]), int(shields[1]), int(energy_delta), bool(bait))
        if key not in self._result_cache:
            self._result_cache[key] = self.simulate(p1, p2, shields=shields, energy_delta=energy_delta, bait_mode=bait)
        return self._result_cache[key]


# ============================================================
# YAKINDAKİ EVENTLER - SinKa entegrasyonu
# ============================================================
FALLBACK_EVENTS = [
    {
        "title": "Gible Community Day Classic",
        "date": "12.09.2026",
        "time": "14:00–17:00",
        "evolution": "Gible → Gabite → Garchomp",
        "attack": "Earth Power",
        "detail": "Gabite'ı Garchomp'a 14:00–21:00 arasında evrimleştirirsen Earth Power öğrenir.",
        "deadline": "12.09.2026 21:00",
        "url": "https://pokemongo.com/news/communitydayclassic-gible-september-2026",
    },
    {
        "title": "Staraptor Super Mega Raid Day",
        "date": "19.09.2026",
        "time": "14:00–17:00",
        "evolution": "Staraptor → Mega Staraptor",
        "attack": "Brave Bird+",
        "detail": "Mega Staraptor Mega Evolved durumdayken ilave Charged Attack özelliği.",
        "deadline": "Etkinlik süresi",
        "url": "https://pokemongo.com/news/staraptor-super-mega-raid-day-2026",
    },
    {
        "title": "Phantump Catch Mastery",
        "date": "26.09.2026",
        "time": "10:00–20:00",
        "evolution": "Phantump → Trevenant",
        "attack": "Özel evrim hareketi belirtilmedi",
        "detail": "Phantump odaklı Catch Mastery etkinliği.",
        "deadline": "26.09.2026 20:00",
        "url": "https://pokemongo.com/news/phantump-catch-mastery-september-2026",
    },
]


POGO_EVENT_IMAGE_IDS = {
    "Bulbasaur": 1, "Ivysaur": 2, "Venusaur": 3,
    "Charmander": 4, "Charmeleon": 5, "Charizard": 6,
    "Squirtle": 7, "Wartortle": 8, "Blastoise": 9,
    "Weedle": 13, "Kakuna": 14, "Beedrill": 15,
    "Pikachu": 25, "Raichu": 26,
    "Machop": 66, "Machoke": 67, "Machamp": 68,
    "Gastly": 92, "Haunter": 93, "Gengar": 94,
    "Houndour": 228, "Houndoom": 229,
    "Lapras": 131, "Eevee": 133,
    "Gible": 443, "Gabite": 444, "Garchomp": 445,
    "Staraptor": 398, "Phantump": 708, "Trevenant": 709,
}

POGO_EVENT_IMAGE_CACHE = Path.home() / "PokemonGO_Event_Images"
POGO_EVENT_IMAGE_CACHE.mkdir(exist_ok=True)

def _event_image_url(name):
    pid = POGO_EVENT_IMAGE_IDS.get(str(name).strip())
    if not pid:
        return None
    return f"https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/other/official-artwork/{pid}.png"

def _event_load_image(name, size=(125, 125)):
    if not PIL_OK:
        return None
    safe = re.sub(r"[^A-Za-z0-9_-]", "_", str(name))
    path = POGO_EVENT_IMAGE_CACHE / f"{safe}.png"
    try:
        if not path.exists():
            url = _event_image_url(name)
            if not url:
                return None
            req = urllib.request.Request(url, headers={"User-Agent": "SinKa-PvP-100"})
            with urllib.request.urlopen(req, timeout=12) as r:
                path.write_bytes(r.read())
        img = Image.open(path).convert("RGBA")
        img.thumbnail(size, Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception:
        return None

def _event_pokemon_names(event):
    names = []
    raw_groups = [event.get("pokemon", []), event.get("features", []), event.get("spawns", []),
                  event.get("raids", []), event.get("research", []), event.get("shiny", []),
                  event.get("moves", [])]
    for group in raw_groups:
        for value in (group or []):
            if isinstance(value, dict):
                value = value.get("name") or value.get("pokemon") or ""
            for part in re.split(r"→|->|,", str(value)):
                name = part.strip()
                # Mega Pokémon adlarında temel tür adını da dene.
                base = re.sub(r"^(?:Mega|G-Max|Gigantamax)\s+", "", name, flags=re.I).strip()
                for candidate in (name, base):
                    if candidate and candidate not in names:
                        names.append(candidate)
    return names[:20]


EVENT_TITLE_TR = {
    "Mega Squads": "Mega Takımları",
    "Gible Community Day Classic": "Gible Topluluk Günü Klasiği",
    "Staraptor Super Mega Raid Day": "Staraptor Süper Mega Baskın Günü",
}

EVENT_CATEGORY_TR = {
    "community day": "Topluluk Günü",
    "community day classic": "Topluluk Günü Klasiği",
    "raid day": "Baskın Günü",
    "mega raid": "Mega Baskın",
    "pokemon spotlight hour": "Pokémon Vitrin Saati",
    "spotlight hour": "Vitrin Saati",
}

TR_MONTHS = {1:"Ocak",2:"Şubat",3:"Mart",4:"Nisan",5:"Mayıs",6:"Haziran",7:"Temmuz",8:"Ağustos",9:"Eylül",10:"Ekim",11:"Kasım",12:"Aralık"}

def _tr_event_title(title):
    s = str(title or "Etkinlik")
    if s in EVENT_TITLE_TR:
        return EVENT_TITLE_TR[s]
    replacements = [
        ("Community Day Classic", "Topluluk Günü Klasiği"),
        ("Community Day", "Topluluk Günü"),
        ("Super Mega Raid Day", "Süper Mega Baskın Günü"),
        ("Mega Raid Day", "Mega Baskın Günü"),
        ("Raid Day", "Baskın Günü"),
        ("Spotlight Hour", "Vitrin Saati"),
        ("Evolution Event", "Evrim Etkinliği"),
    ]
    for a,b in replacements:
        s = s.replace(a,b)
    return s

def _tr_category(value):
    s = str(value or "")
    low = s.lower()
    for a,b in EVENT_CATEGORY_TR.items():
        if a in low:
            return b
    return s

def _tr_date(value):
    s = str(value or "")
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})", s)
    if m:
        y,mo,d,h,mi = map(int,m.groups())
        return f"{d} {TR_MONTHS.get(mo, str(mo))} {y} • {h:02d}:{mi:02d}"
    return s

MOVE_TR = {
    "Earth Power": "Earth Power (Güç Toprağı)",
    "Fell Stinger+": "Fell Stinger+ (İğne Sokması+)",
    "Dark Pulse+": "Dark Pulse+ (Karanlık Darbe+)",
    "Hydro Cannon": "Hydro Cannon (Hidro Topu)",
    "Blast Burn": "Blast Burn (Yakıcı Patlama)",
    "Frenzy Plant": "Frenzy Plant (Çılgın Bitki)",
    "Meteor Mash": "Meteor Mash (Meteor Ezmesi)",
    "Psychic": "Psychic (Psişik)",
    "Rock Slide": "Rock Slide (Kaya Kayması)",
}

def _tr_common(value):
    s = str(value or "")
    replacements = [
        ("Trainer Battles", "Antrenör Savaşları"),
        ("Gyms and raids", "Spor Salonları ve Baskınlar"),
        ("Gyms & raids", "Spor Salonları ve Baskınlar"),
        ("Charged Attack", "Yüklü Saldırı"),
        ("Fast Attack", "Hızlı Saldırı"),
        ("Featured Attack", "Öne Çıkan Saldırı"),
        ("Special Move", "Özel Hareket"),
        ("Exclusive Attack", "Özel Saldırı"),
        ("beginning of event", "etkinlik başlangıcından"),
        ("until", "kadar"),
        ("local time", "yerel saat"),
        ("power", "güç"),
        ("Shiny", "Parlak"),
        ("Bonus", "Bonus"),
        ("Catch XP", "Yakalama XP"),
        ("Increased", "Artırılmış"),
        ("Spawns", "Çıkışlar"),
        ("Raid", "Baskın"),
        ("Community Day", "Topluluk Günü"),
        ("Classic", "Klasiği"),
        ("Super Mega Raid Day", "Süper Mega Baskın Günü"),
        ("Mega Raid", "Mega Baskın"),
        ("Mega Squads", "Mega Takımları"),
        ("Evolve", "Evrimleştir"),
        ("evolve", "evrimleştir"),
        ("Pokémon featured", "öne çıkan Pokémon"),
        ("featured", "öne çıkan"),
        ("Featured", "Öne Çıkan"),
        ("Learn more about additional", "Ek ayrıntılar hakkında daha fazla bilgi"),
        ("while they are Mega Evolved", "Mega Evrimliyken"),
        ("from the beginning of event", "etkinlik başlangıcından itibaren"),
        ("from the beginning of the event", "etkinlik başlangıcından itibaren"),
        ("at 9:00 p.m.", "saat 21:00'de"),
        ("at 9:00 PM", "saat 21:00'de"),
        ("local time", "yerel saat"),
        ("until September", "Eylül tarihine kadar"),
        ("will have a very high chance", "çok yüksek ihtimalle çıkacak"),
        ("very high chance of appearing", "çıkma ihtimali çok yüksek"),
        ("appearing at PokéStops", "PokéStop'larda çıkma"),
        ("PokéStops", "PokéStop'lar"),
        ("PokéStop", "PokéStop"),
        ("during the event", "etkinlik sırasında"),
        ("during this event", "bu etkinlik sırasında"),
        ("will be featured", "öne çıkacak"),
        ("is featured", "öne çıkıyor"),
        ("featured Pokémon", "öne çıkan Pokémon"),
        ("knows the Charged Attack", "Yüklü Saldırıyı bilen"),
        ("knows the", "öğrenmiş olan"),
        ("Charged Attack", "Yüklü Saldırı"),
        ("Fast Attack", "Hızlı Saldırı"),
        ("Trainer Battles", "Antrenör Savaşları"),
        ("Gyms and raids", "Spor Salonları ve Baskınlar"),
        ("Gyms & raids", "Spor Salonları ve Baskınlar"),
        ("Lure Module Bonus", "Yem Modülü Bonusu"),
        ("Catch XP", "Yakalama XP'si"),
        ("Increased Shiny Odds", "Artırılmış Parlak Pokémon ihtimali"),
        ("Increased Spawns", "Artırılmış Pokémon çıkışı"),
        ("1-hour Lures", "1 saatlik Yem Modülleri"),
        ("3-hour Incense", "3 saatlik Tütsüler"),
        ("3x Catch XP", "3 kat Yakalama XP'si"),
        ("Incense Spawns", "Tütsüden çıkan Pokémonlar"),
        ("Shiny available", "Parlak mevcut"),
        ("available", "mevcut"),
        ("Evolve", "Evrimleştir"),
        ("evolve", "evrimleştir"),
        ("until", "kadar"),
        ("beginning", "başlangıç"),
        ("event", "etkinlik"),
    ]
    for a,b in replacements:
        s = s.replace(a,b)
    return s


POGO_COORDINATES_AUTOMATION_NOTE = "Satır bazlı arama: Pokémon adı + mevcut lig + PvP + Top Rank 1 + Search"

POGO_COORDINATES_SEARCH_URL = "https://pogocoordinates.com/search"

def _find_chrome_executable():
    """Find a normal Chrome/Edge executable on Windows."""
    candidates = [
        os.environ.get("PROGRAMFILES", "") + r"\Google\Chrome\Application\chrome.exe",
        os.environ.get("PROGRAMFILES(X86)", "") + r"\Google\Chrome\Application\chrome.exe",
        os.environ.get("LOCALAPPDATA", "") + r"\Google\Chrome\Application\chrome.exe",
        os.environ.get("PROGRAMFILES", "") + r"\Microsoft\Edge\Application\msedge.exe",
        os.environ.get("PROGRAMFILES(X86)", "") + r"\Microsoft\Edge\Application\msedge.exe",
        os.environ.get("LOCALAPPDATA", "") + r"\Microsoft\Edge\Application\msedge.exe",
    ]
    for p in candidates:
        if p and Path(p).exists():
            return p
    return None


def _pogo_search_name(pokemon_name):
    """PogoCoordinates için isim normalizasyonu.
    'Alolan' form adını 'Alola' olarak gönderir; uygulamadaki görünen isim değişmez.
    """
    name = str(pokemon_name or "").strip()
    replacements = {
        "(Alolan)": "(Alola)",
        "(Alolan Form)": "(Alola)",
        " Alolan": " Alola",
    }
    for old, new in replacements.items():
        name = name.replace(old, new)
    return name

def _run_pogo_form_automation(pokemon_name, league_name):
    """Open PogoCoordinates and fill/submit its actual search form."""
    try:
        pokemon_name = _pogo_search_name(pokemon_name)
        from playwright.sync_api import sync_playwright

        chrome = _find_chrome_executable()
        if not chrome:
            raise RuntimeError("Google Chrome veya Microsoft Edge bulunamadı.")

        league = str(league_name or "").lower()
        if "great" in league:
            league_text = "Great"
        elif "ultra" in league:
            league_text = "Ultra"
        elif "master" in league:
            league_text = "Master"
            # V69_GIGAMAX_BUTTON
            try:
                _gigamax_btn = tk.Button(parent, text="Gigamax", command=lambda: show_gigamax_inline())
                _gigamax_btn.pack(side="left", padx=4)
            except Exception:
                pass

        else:
            league_text = "Great"

        user_dir = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "SinKA_PvP" / "browser_profile"
        user_dir.mkdir(parents=True, exist_ok=True)

        with sync_playwright() as pw:
            browser = pw.chromium.launch_persistent_context(
                str(user_dir),
                executable_path=chrome,
                headless=False,
                args=["--start-maximized"],
            )
            page = browser.pages[0] if browser.pages else browser.new_page()
            page.goto(POGO_COORDINATES_SEARCH_URL, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(1800)

            result = page.evaluate("""
            ({pokemon, league}) => {
                const norm = s => (s || '').toString().trim().toLowerCase()
                    .replace(/ı/g,'i').replace(/ş/g,'s').replace(/ğ/g,'g')
                    .replace(/ü/g,'u').replace(/ö/g,'o').replace(/ç/g,'c');

                const textOf = el => {
                    let t = '';
                    if (el.labels) t += ' ' + Array.from(el.labels).map(x=>x.innerText).join(' ');
                    if (el.getAttribute('aria-label')) t += ' ' + el.getAttribute('aria-label');
                    if (el.getAttribute('placeholder')) t += ' ' + el.getAttribute('placeholder');
                    if (el.name) t += ' ' + el.name;
                    if (el.id) t += ' ' + el.id;
                    if (el.parentElement) t += ' ' + (el.parentElement.innerText || '').slice(0,220);
                    return norm(t);
                };

                const controls = Array.from(document.querySelectorAll('input, select, textarea'));
                let poke = controls.find(e => {
                    const t = textOf(e);
                    return t.includes('pokemon') || t.includes('pokémon') || t.includes('name');
                });
                let stat = controls.find(e => {
                    const t = textOf(e);
                    return t.includes('stat') || t.includes('ranking type');
                });
                let rank = controls.find(e => {
                    const t = textOf(e);
                    return t.includes('top rank') || t.includes('rank') || t.includes('limit');
                });
                let leagueCtl = controls.find(e => {
                    const t = textOf(e);
                    return t.includes('league') || t.includes('lig');
                });

                const selects = Array.from(document.querySelectorAll('select'));
                if (!stat) stat = selects.find(s => Array.from(s.options).some(o => norm(o.text).includes('pvp')));
                if (!rank) rank = selects.find(s => Array.from(s.options).some(o => ['1','top 1','rank 1'].includes(norm(o.text))));
                if (!leagueCtl) leagueCtl = selects.find(s => {
                    const opts = Array.from(s.options).map(o=>norm(o.text));
                    return opts.some(x=>x.includes('great')) && opts.some(x=>x.includes('ultra'));
                });

                const setValue = (el, value) => {
                    if (!el) return false;
                    if (el.tagName === 'SELECT') {
                        const target = norm(value);
                        const opt = Array.from(el.options).find(o =>
                            norm(o.text) === target ||
                            norm(o.value) === target ||
                            norm(o.text).includes(target)
                        );
                        if (opt) {
                            el.value = opt.value;
                            el.dispatchEvent(new Event('input',{bubbles:true}));
                            el.dispatchEvent(new Event('change',{bubbles:true}));
                            return true;
                        }
                    } else {
                        el.focus();
                        el.value = value;
                        el.dispatchEvent(new Event('input',{bubbles:true}));
                        el.dispatchEvent(new Event('change',{bubbles:true}));
                        return true;
                    }
                    return false;
                };

                let pokeOK = false;
                if (poke) {
                    poke.focus();
                    poke.value = pokemon;
                    poke.dispatchEvent(new Event('input',{bubbles:true}));
                    poke.dispatchEvent(new Event('change',{bubbles:true}));
                    pokeOK = true;
                }
                const isMaster = norm(league) === 'master';

                // Master League için PogoCoordinates tarafında PvP yerine Stats
                // seçilir ve minimum Level / IV değerleri doldurulur.
                // Kontrollerin HTML yapısı değişse bile label/alan metninden
                // bulmaya çalışıyoruz.
                const findSectionInputs = (sectionName) => {
                    const wanted = norm(sectionName);
                    const inputs = Array.from(document.querySelectorAll('input'))
                        .filter(e => ['text','number',''].includes((e.type || '').toLowerCase()));

                    // 1) input'un kendisi veya yakın üst elemanlarında bölüm adı.
                    for (const inp of inputs) {
                        let el = inp;
                        for (let depth = 0; depth < 6 && el; depth++, el = el.parentElement) {
                            const txt = norm((el.innerText || '').slice(0, 500));
                            if (txt.includes(wanted)) {
                                const group = Array.from(el.querySelectorAll('input'))
                                    .filter(x => ['text','number',''].includes((x.type || '').toLowerCase()));
                                if (group.length) return group;
                            }
                        }
                    }

                    // 2) Label metninden bölümü bul.
                    const labels = Array.from(document.querySelectorAll('label'));
                    const label = labels.find(l => norm(l.innerText || '').includes(wanted));
                    if (label) {
                        let el = label.parentElement;
                        for (let depth = 0; depth < 5 && el; depth++, el = el.parentElement) {
                            const group = Array.from(el.querySelectorAll('input'))
                                .filter(x => ['text','number',''].includes((x.type || '').toLowerCase()));
                            if (group.length) return group;
                        }
                    }
                    return [];
                };

                const setFirstMin = (sectionName, value) => {
                    const group = findSectionInputs(sectionName);
                    if (!group.length) return false;
                    // Ekrandaki düzen Min -> Max olduğu için ilk alan minimumdur.
                    return setValue(group[0], String(value));
                };

                if (isMaster) {
                    // Stats/PvP alanı bazı sürümlerde select değil buton/toggle olabilir.
                    let statsControl = Array.from(document.querySelectorAll('button,[role="button"],label'))
                        .find(e => norm(e.innerText || e.textContent || '').trim() === 'stats');
                    if (statsControl) {
                        try { statsControl.click(); } catch (e) {}
                    }
                    // Select olarak bulunmuşsa yine değerini ayarla.
                    const statResult = setValue(stat, 'Stats');
                    const levelResult = setFirstMin('level', 34);
                    const ivResult = setFirstMin('iv', 100);

                    return {
                        poke: pokeOK,
                        stat: !!(statResult || statsControl),
                        rank: true,
                        league: setValue(leagueCtl, league),
                        levelMin: levelResult,
                        ivMin: ivResult,
                        controls: controls.length
                    };
                }

                return {
                    poke: pokeOK,
                    stat: setValue(stat, 'PvP'),
                    rank: setValue(rank, '1'),
                    league: setValue(leagueCtl, league),
                    controls: controls.length
                };
            }
            """, {"pokemon": str(pokemon_name), "league": league_text})

            page.wait_for_timeout(1000)

            # Handle autocomplete: click an exact Pokémon suggestion if one appears.
            suggestion = page.locator("text=" + str(pokemon_name)).first
            try:
                if suggestion.is_visible(timeout=700):
                    suggestion.click(timeout=1200)
                    page.wait_for_timeout(300)
            except Exception:
                pass

            # Click a real Search/Ara button.
            buttons = page.locator("button, input[type=submit], input[type=button]")
            clicked = False
            for i in range(buttons.count()):
                b = buttons.nth(i)
                try:
                    label = ((b.inner_text(timeout=300) or "") + " " +
                             (b.get_attribute("value") or "") + " " +
                             (b.get_attribute("aria-label") or "")).strip().lower()
                    if "search" in label or "ara" in label:
                        b.click(timeout=1500)
                        clicked = True
                        break
                except Exception:
                    pass

            if not clicked:
                page.keyboard.press("Enter")

            page.wait_for_timeout(2500)
            return True

    except Exception as ex:
        try:
            msg = (
                "PogoCoordinates otomasyonu çalıştırılamadı.\n\n" + str(ex)
                + "\n\nGerekirse: pip install playwright"
            )
            root_obj = globals().get("_SINKA_ROOT") or getattr(globals().get("tk"), "_default_root", None)
            if root_obj is not None:
                root_obj.after(0, lambda m=msg: messagebox.showerror("SinKA PvP", m))
        except Exception:
            pass
        return False


def open_pogo_coordinates_search(pokemon_name, league_name):
    """PogoCoordinates otomasyonunu ayrı Python sürecünde çalıştırır ve eksikleri otomatik kurar."""
    name = str(pokemon_name or "").strip()
    league = str(league_name or "Great League").strip()
    # Sadece izin verilen lig adlarını kabul et; dışarıdan gelen belirsiz değer
    # Great'e sessizce düşmesin.
    league_aliases = {
        "great": "Great League", "great league": "Great League",
        "ultra": "Ultra League", "ultra league": "Ultra League",
        "master": "Master League", "master league": "Master League",
        "mega": "Mega", "mega league": "Mega", "gigamax": "Gigamax",
    }
    league = league_aliases.get(league.lower(), league)
    if league.lower() == "gigamax":
        _open_pogo_status_window(name, league)
        _pogo_status_message("ERROR|Gigamax için PogoCoordinates araması yapılamıyor.")
        root_obj = globals().get("_SINKA_ROOT")
        if root_obj is not None:
            root_obj.after(0, lambda: _show_pogo_error(
                "Gigamax için arama yapılamıyor.",
                "ERR-PG-GIGAMAX-NOT-SUPPORTED",
                "PogoCoordinates üzerinde Gigamax araması desteklenmiyor.",
                0
            ))
        return
    _open_pogo_status_window(name, league)
    _pogo_status_message("START|" + name + " için otomasyon başlatılıyor…")

    def launcher():
        try:
            import subprocess, tempfile, os, sys
            fd, path = tempfile.mkstemp(prefix="sinka_pogo_", suffix=".py")
            os.close(fd)
            Path(path).write_text(CHILD_CODE, encoding="utf-8")
            proc = subprocess.Popen(
                [sys.executable, path, json.dumps([_pogo_search_name(name)], ensure_ascii=False), league],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                close_fds=True,
            )
            for line in proc.stdout:
                line = line.strip()
                if line:
                    _pogo_status_message(line)
            rc = proc.wait()
            if rc:
                _pogo_status_message("ERROR|Otomasyon hata koduyla kapandı: " + str(rc))
            try:
                os.unlink(path)
            except Exception:
                pass
        except Exception as e:
            _pogo_status_message("ERROR|Alt işlem başlatılamadı: " + repr(e))

    import threading
    threading.Thread(target=launcher, daemon=True).start()


def open_pogo_coordinates_search_multiple(pokemon_names, league_name, silent=False, on_complete=None):
    """Çoklu PogoCoordinates araması. silent=True otomatik aramada popup açmaz;
    on_complete(payload) tamamlandığında ana thread üzerinde çağrılır."""
    names = list(dict.fromkeys(
        str(x).strip() for x in (pokemon_names or []) if str(x).strip()
    ))
    league = str(league_name or "Great League").strip()
    if not names:
        return

    if league.lower() == "gigamax":
        if not silent:
            _open_pogo_status_window("ÇOKLU POKÉMON ARAMA", league)
        _pogo_status_message("ERROR|Gigamax için PogoCoordinates araması yapılamıyor.")
        root_obj = globals().get("_SINKA_ROOT")
        if root_obj is not None and not silent:
            root_obj.after(0, lambda: _show_pogo_error(
                "Gigamax için arama yapılamıyor.",
                "ERR-PG-GIGAMAX-NOT-SUPPORTED",
                "PogoCoordinates üzerinde Gigamax araması desteklenmiyor.",
                0
            ))
        if on_complete:
            try:
                on_complete({"pokemon":"Çoklu Pokémon Arama", "league":league, "results":[], "errors":["Gigamax için PogoCoordinates araması yapılamıyor."], "noresult":False, "return_code":0})
            except Exception:
                pass
        return

    if not silent:
        _open_pogo_status_window("ÇOKLU POKÉMON ARAMA", league)
    _pogo_status_message(
        "START|" + str(len(names)) +
        " adet ☐ Pokémon tek aramada aranacak."
    )

    def launcher():
        import subprocess, tempfile, os, sys
        fd, path = tempfile.mkstemp(prefix="sinka_pogo_multi_", suffix=".py")
        os.close(fd)

        all_results = []
        errors = []
        noresult = False
        rc = 0

        try:
            child_code = globals().get("CHILD_CODE")
            if not child_code:
                code = "ERR-PG-CHILD-CODE-MISSING"
                detail = (
                    "CHILD_CODE tanımı bulunamadı. "
                    "Otomasyon alt süreci oluşturulamadı."
                )
                _pogo_status_message("ERROR|" + code + "|" + detail)
                root_obj = globals().get("_SINKA_ROOT")
                if root_obj is not None:
                    root_obj.after(
                        0,
                        lambda: _show_pogo_error(
                            "Çoklu Pokémon araması başlatılamadı.",
                            code, detail, -1
                        )
                    )
                return

            Path(path).write_text(child_code, encoding="utf-8")
            _pogo_status_message(
                "INFO|☐ Olan Pokémonlar: " + ", ".join(names)
            )

            proc = subprocess.Popen(
                [sys.executable, path, json.dumps([_pogo_search_name(n) for n in names], ensure_ascii=False), league],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                close_fds=True,
            )

            for raw in proc.stdout:
                line = raw.strip()
                if not line:
                    continue

                kind, _, detail = line.partition("|")

                if kind in ("ERROR", "PLAYWRIGHT_MISSING"):
                    errors.append(detail or "Bilinmeyen otomasyon hatası")
                    _pogo_status_message(
                        "ERROR|" + (detail or "Bilinmeyen otomasyon hatası")
                    )
                    continue

                if kind == "NORESULT":
                    # Child süreç arama tamamlanmadan NORESULT bildirebilir.
                    # Bu aşamada sonuç penceresini açma; sadece durumu kaydet.
                    # Son karar proc.wait() sonrasında, all_results kontrol edilerek verilir.
                    noresult = True
                    _pogo_status_message(
                        "INFO|⚠ Ara sonuç: henüz Pokémon sonucu gelmedi; işlem tamamlanması bekleniyor…"
                    )
                    continue

                if kind in ("RESULT", "RESULTS"):
                    try:
                        payload = json.loads(detail) if detail else {}
                        batch = payload.get("results", [])
                        if batch:
                            all_results.extend(batch)
                    except Exception as ex:
                        errors.append(
                            "ERR-PG-RESULT_JSON|" +
                            type(ex).__name__ + ": " + str(ex)
                        )
                    continue

                _pogo_status_message(kind + "|" + detail)

            rc = proc.wait()

            def finish():
                payload = {
                    "pokemon": "Çoklu Pokémon Arama",
                    "league": league,
                    "results": list(all_results),
                    "errors": list(errors),
                    "noresult": bool(noresult),
                    "return_code": rc,
                }
                for item in payload["results"]:
                    if isinstance(item, dict):
                        item.setdefault("league", league)

                if on_complete:
                    try:
                        on_complete(payload)
                    except Exception:
                        pass

                if silent:
                    return

                if all_results:
                    _show_pogo_results(payload)
                    return

                if errors or rc != 0:
                    code = "ERR-PG-UNKNOWN"
                    detail = errors[0] if errors else (
                        "Child işlem hata koduyla kapandı: " + str(rc)
                    )

                    if "|" in detail:
                        possible_code, possible_detail = detail.split("|", 1)
                        if possible_code.startswith("ERR-PG-"):
                            code = possible_code
                            detail = possible_detail

                    _show_pogo_error(
                        "Çoklu Pokémon araması hata verdi.",
                        code,
                        detail,
                        rc
                    )
                    return

                if noresult:
                    _show_pogo_noresult(
                        "Pokémon bulunamadı. Lütfen daha sonra tekrar deneyin."
                    )
                    return

                _show_pogo_error(
                    "Arama tamamlandı fakat sonuç durumu alınamadı.",
                    "ERR-PG-NO-STATUS",
                    "Child işleminden RESULT veya NORESULT mesajı gelmedi.",
                    rc
                )

            root_obj = globals().get("_SINKA_ROOT")
            if root_obj is not None:
                root_obj.after(0, finish)
            else:
                finish()

        except Exception as e:
            code = "ERR-PG-LAUNCH-" + type(e).__name__.upper()
            detail = str(e) or "(mesaj yok)"
            _pogo_status_message("ERROR|" + code + "|" + detail)

            root_obj = globals().get("_SINKA_ROOT")
            if root_obj is not None:
                root_obj.after(
                    0,
                    lambda c=code, d=detail: _show_pogo_error(
                        "Çoklu Pokémon araması başlatılamadı.",
                        c, d, -1
                    )
                )
        finally:
            try:
                os.unlink(path)
            except Exception:
                pass

    import threading
    threading.Thread(target=launcher, daemon=True).start()


CHILD_CODE = base64.b64decode('CmltcG9ydCBzeXMsIHRpbWUsIHRyYWNlYmFjaywgYXN5bmNpbywgc3VicHJvY2VzcywgaW1wb3J0bGliLnV0aWwsIG9zLCBqc29uLCByZQoKcmF3X25hbWVzID0gc3lzLmFyZ3ZbMV0KdHJ5OgogICAgcGFyc2VkX25hbWVzID0ganNvbi5sb2FkcyhyYXdfbmFtZXMpCiAgICBpZiBpc2luc3RhbmNlKHBhcnNlZF9uYW1lcywgbGlzdCk6CiAgICAgICAgbmFtZXMgPSBbc3RyKHgpLnN0cmlwKCkgZm9yIHggaW4gcGFyc2VkX25hbWVzIGlmIHN0cih4KS5zdHJpcCgpXQogICAgZWxzZToKICAgICAgICBuYW1lcyA9IFtzdHIocGFyc2VkX25hbWVzKS5zdHJpcCgpXQpleGNlcHQgRXhjZXB0aW9uOgogICAgbmFtZXMgPSBbc3RyKHJhd19uYW1lcykuc3RyaXAoKV0KaWYgbm90IG5hbWVzOgogICAgbmFtZXMgPSBbIiJdCm5hbWUgPSBuYW1lc1swXQpsZWFndWUgPSBzeXMuYXJndlsyXQp0cnk6CiAgICBzeXMuc3Rkb3V0LnJlY29uZmlndXJlKGVuY29kaW5nPSd1dGYtOCcsIGVycm9ycz0ncmVwbGFjZScpCiAgICBzeXMuc3RkZXJyLnJlY29uZmlndXJlKGVuY29kaW5nPSd1dGYtOCcsIGVycm9ycz0ncmVwbGFjZScpCmV4Y2VwdCBFeGNlcHRpb246CiAgICBwYXNzCgpkZWYgc2F5KGssIG0pOgogICAgcHJpbnQoayArICJ8IiArIG0sIGZsdXNoPVRydWUpCgpkZWYgZW5zdXJlX3BsYXl3cmlnaHQoKToKICAgIGlmIGltcG9ydGxpYi51dGlsLmZpbmRfc3BlYygicGxheXdyaWdodCIpIGlzIG5vdCBOb25lOgogICAgICAgIHJldHVybiBUcnVlCiAgICBzYXkoIklOU1RBTEwiLCAiUGxheXdyaWdodCBidWx1bmFtYWTEsS4gT3RvbWF0aWsga3VydWx1eW9y4oCmIikKICAgIHRyeToKICAgICAgICBwID0gc3VicHJvY2Vzcy5ydW4oCiAgICAgICAgICAgIFtzeXMuZXhlY3V0YWJsZSwgIi1tIiwgInBpcCIsICJpbnN0YWxsIiwgInBsYXl3cmlnaHQiXSwKICAgICAgICAgICAgY2FwdHVyZV9vdXRwdXQ9VHJ1ZSwgdGV4dD1UcnVlLCBlbmNvZGluZz0idXRmLTgiLCBlcnJvcnM9InJlcGxhY2UiLAogICAgICAgICAgICB0aW1lb3V0PTMwMAogICAgICAgICkKICAgICAgICBpZiBwLnJldHVybmNvZGUgIT0gMDoKICAgICAgICAgICAgc2F5KCJFUlJPUiIsICJQbGF5d3JpZ2h0IGt1cnVsYW1hZMSxOiAiICsgKHAuc3RkZXJyIG9yIHAuc3Rkb3V0KVstMTIwMDpdKQogICAgICAgICAgICByZXR1cm4gRmFsc2UKICAgICAgICBzYXkoIklOU1RBTEwiLCAiUGxheXdyaWdodCBQeXRob24gcGFrZXRpIGt1cnVsZHUuIikKICAgICAgICByZXR1cm4gVHJ1ZQogICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOgogICAgICAgIHNheSgiRVJST1IiLCAiUGxheXdyaWdodCBrdXJ1bHVtdSBiYcWfbGF0xLFsYW1hZMSxOiAiICsgcmVwcihlKSkKICAgICAgICByZXR1cm4gRmFsc2UKCmFzeW5jIGRlZiBtYWluKCk6CiAgICBpZiBub3QgZW5zdXJlX3BsYXl3cmlnaHQoKToKICAgICAgICByZXR1cm4gMTAKCiAgICB0cnk6CiAgICAgICAgZnJvbSBwbGF5d3JpZ2h0LmFzeW5jX2FwaSBpbXBvcnQgYXN5bmNfcGxheXdyaWdodAogICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOgogICAgICAgIHNheSgiRVJST1IiLCAiUGxheXdyaWdodCBrdXJ1bGR1IGFtYSB5w7xrbGVuZW1lZGk6ICIgKyByZXByKGUpKQogICAgICAgIHJldHVybiAxMAoKICAgIHNheSgiU1RBUlQiLCAiUGxheXdyaWdodCBoYXrEsXIuIFRhcmF5xLFjxLEgaGF6xLFybGFuxLF5b3LigKYiKQoKICAgIHRyeToKICAgICAgICBhc3luYyB3aXRoIGFzeW5jX3BsYXl3cmlnaHQoKSBhcyBwOgogICAgICAgICAgICBzYXkoIkJST1dTRVIiLCAiQ2hyb21pdW0ga29udHJvbCBlZGlsaXlvcuKApiIpCiAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgIGJyb3dzZXIgPSBhd2FpdCBwLmNocm9taXVtLmxhdW5jaChoZWFkbGVzcz1UcnVlKQogICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uIGFzIGU6CiAgICAgICAgICAgICAgICBtc2cgPSByZXByKGUpCiAgICAgICAgICAgICAgICBpZiAiRXhlY3V0YWJsZSBkb2Vzbid0IGV4aXN0IiBpbiBtc2cgb3IgImV4ZWN1dGFibGUiIGluIG1zZy5sb3dlcigpOgogICAgICAgICAgICAgICAgICAgIHNheSgiSU5TVEFMTCIsICJDaHJvbWl1bSBidWx1bmFtYWTEsS4gT3RvbWF0aWsga3VydWx1eW9y4oCmIikKICAgICAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgIHAyID0gc3VicHJvY2Vzcy5ydW4oCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBbc3lzLmV4ZWN1dGFibGUsICItbSIsICJwbGF5d3JpZ2h0IiwgImluc3RhbGwiLCAiY2hyb21pdW0iXSwKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGNhcHR1cmVfb3V0cHV0PVRydWUsIHRleHQ9VHJ1ZSwgZW5jb2Rpbmc9InV0Zi04IiwKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGVycm9ycz0icmVwbGFjZSIsIHRpbWVvdXQ9NjAwCiAgICAgICAgICAgICAgICAgICAgICAgICkKICAgICAgICAgICAgICAgICAgICAgICAgaWYgcDIucmV0dXJuY29kZSAhPSAwOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgc2F5KCJFUlJPUiIsICJDaHJvbWl1bSBrdXJ1bGFtYWTEsTogIiArIChwMi5zdGRlcnIgb3IgcDIuc3Rkb3V0KVstMTUwMDpdKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgcmV0dXJuIDEyCiAgICAgICAgICAgICAgICAgICAgICAgIHNheSgiSU5TVEFMTCIsICJDaHJvbWl1bSBrdXJ1bGR1LiBZZW5pZGVuIGJhxZ9sYXTEsWzEsXlvcuKApiIpCiAgICAgICAgICAgICAgICAgICAgICAgIGJyb3dzZXIgPSBhd2FpdCBwLmNocm9taXVtLmxhdW5jaChoZWFkbGVzcz1UcnVlKQogICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgaWU6CiAgICAgICAgICAgICAgICAgICAgICAgIHNheSgiRVJST1IiLCAiQ2hyb21pdW0gb3RvbWF0aWsga3VydWx1bXUgYmHFn2FyxLFzxLF6OiAiICsgcmVwcihpZSkpCiAgICAgICAgICAgICAgICAgICAgICAgIHJldHVybiAxMgogICAgICAgICAgICAgICAgZWxzZToKICAgICAgICAgICAgICAgICAgICBzYXkoIkVSUk9SIiwgIkNocm9taXVtIGJhxZ9sYXTEsWxhbWFkxLE6ICIgKyBtc2cpCiAgICAgICAgICAgICAgICAgICAgcmV0dXJuIDEyCgogICAgICAgICAgICBjb250ZXh0ID0gYXdhaXQgYnJvd3Nlci5uZXdfY29udGV4dCgKICAgICAgICAgICAgICAgIHZpZXdwb3J0PXsid2lkdGgiOiAxNDAwLCAiaGVpZ2h0IjogOTAwfQogICAgICAgICAgICApCiAgICAgICAgICAgIGF3YWl0IGNvbnRleHQuZ3JhbnRfcGVybWlzc2lvbnMoWyJjbGlwYm9hcmQtcmVhZCIsICJjbGlwYm9hcmQtd3JpdGUiXSkKICAgICAgICAgICAgYXdhaXQgY29udGV4dC5hZGRfaW5pdF9zY3JpcHQoIiIiCiAgICAgICAgICAgICgoKSA9PiB7CiAgICAgICAgICAgICAgICB0cnkgewogICAgICAgICAgICAgICAgICAgIGlmIChuYXZpZ2F0b3IuY2xpcGJvYXJkICYmIG5hdmlnYXRvci5jbGlwYm9hcmQud3JpdGVUZXh0KSB7CiAgICAgICAgICAgICAgICAgICAgICAgIGNvbnN0IG9yaWdpbmFsID0gbmF2aWdhdG9yLmNsaXBib2FyZC53cml0ZVRleHQuYmluZChuYXZpZ2F0b3IuY2xpcGJvYXJkKTsKICAgICAgICAgICAgICAgICAgICAgICAgbmF2aWdhdG9yLmNsaXBib2FyZC53cml0ZVRleHQgPSBhc3luYyAodGV4dCkgPT4gewogICAgICAgICAgICAgICAgICAgICAgICAgICAgd2luZG93Ll9fc2lua2FfbGFzdF9jb29yZHMgPSBTdHJpbmcodGV4dCB8fCAiIik7CiAgICAgICAgICAgICAgICAgICAgICAgICAgICB0cnkgeyByZXR1cm4gYXdhaXQgb3JpZ2luYWwodGV4dCk7IH0gY2F0Y2goZSkgeyByZXR1cm4gdW5kZWZpbmVkOyB9CiAgICAgICAgICAgICAgICAgICAgICAgIH07CiAgICAgICAgICAgICAgICAgICAgfQogICAgICAgICAgICAgICAgfSBjYXRjaChlKSB7fQogICAgICAgICAgICB9KSgpOwogICAgICAgICAgICAiIiIpCiAgICAgICAgICAgIHBhZ2UgPSBhd2FpdCBjb250ZXh0Lm5ld19wYWdlKCkKICAgICAgICAgICAgc2F5KCJQQUdFIiwgIlBvZ29Db29yZGluYXRlcyBhw6fEsWzEsXlvcuKApiIpCiAgICAgICAgICAgIGF3YWl0IHBhZ2UuZ290bygKICAgICAgICAgICAgICAgICJodHRwczovL3BvZ29jb29yZGluYXRlcy5jb20vc2VhcmNoIiwKICAgICAgICAgICAgICAgIHdhaXRfdW50aWw9ImRvbWNvbnRlbnRsb2FkZWQiLAogICAgICAgICAgICAgICAgdGltZW91dD02MDAwMAogICAgICAgICAgICApCiAgICAgICAgICAgIGF3YWl0IHBhZ2Uud2FpdF9mb3JfdGltZW91dCg2MDAwKQogICAgICAgICAgICBzYXkoIlBBR0UiLCAiU2F5ZmEgYcOnxLFsZMSxLiIpCgogICAgICAgICAgICBhc3luYyBkZWYgZmlyc3RfdmlzaWJsZShsb2NhdG9yKToKICAgICAgICAgICAgICAgIGZvciBpIGluIHJhbmdlKGF3YWl0IGxvY2F0b3IuY291bnQoKSk6CiAgICAgICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgICAgICB4ID0gbG9jYXRvci5udGgoaSkKICAgICAgICAgICAgICAgICAgICAgICAgaWYgYXdhaXQgeC5pc192aXNpYmxlKCk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICByZXR1cm4geAogICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICAgICAgICAgIHBhc3MKICAgICAgICAgICAgICAgIHJldHVybiBOb25lCgogICAgICAgICAgICBjb250cm9scyA9IHBhZ2UubG9jYXRvcigiaW5wdXQsIHNlbGVjdCwgdGV4dGFyZWEsIGJ1dHRvbiIpCiAgICAgICAgICAgIHZpc2libGUgPSAwCiAgICAgICAgICAgIGZvciBpIGluIHJhbmdlKG1pbihhd2FpdCBjb250cm9scy5jb3VudCgpLCA4MCkpOgogICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgIGMgPSBjb250cm9scy5udGgoaSkKICAgICAgICAgICAgICAgICAgICBpZiBhd2FpdCBjLmlzX3Zpc2libGUoKToKICAgICAgICAgICAgICAgICAgICAgICAgdmlzaWJsZSArPSAxCiAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgICAgICAgICAgICAgIHBhc3MKICAgICAgICAgICAgc2F5KCJDT05UUk9MUyIsICJHw7Zyw7xuZW4gZm9ybSBlbGVtYW7EsTogIiArIHN0cih2aXNpYmxlKSkKCiAgICAgICAgICAgICMgR0VSw4dFSyBQb2dvQ29vcmRpbmF0ZXMgZm9ybSBzxLFyYXPEsToKICAgICAgICAgICAgIyBBbnkgUG9rZW1vbiAtPiBhw6fEsWxhbiBBbHRhcmlhIMO2bmVyaXNpIC0+IFN0YXRzL1B2UCAtPiBUb3AgcmFua2luZyAxIC0+IExlYWd1ZSAtPiBTZWFyY2gKCiAgICAgICAgICAgICMgR2Vyw6dlayBzYXlmYWRha2kgaWxrIGFsYW4gY3VzdG9tIGJpciBjb21ib2JveCBvbGFiaWxpcjsKICAgICAgICAgICAgIyBidSBuZWRlbmxlIGlucHV0IHNlbGVjdG9yJ8O8bmUgYmHEn2zEsSBrYWxtxLF5b3J1ei4KICAgICAgICAgICAgIyBQb2dvQ29vcmRpbmF0ZXMgw6dva2x1IFBva8OpbW9uIGFsYW7EsToKICAgICAgICAgICAgIyBIZXIgc2XDp2ltZGVuIHNvbnJhIGlucHV0IERPTSd1IGRlxJ9pxZ9lYmlsZGnEn2kgacOnaW4gbG9jYXRvcifEsQogICAgICAgICAgICAjIGhlciBzZWZlcmluZGUgeWVuaWRlbiBidWx1eW9ydXouIELDtnlsZWNlIDIuLCAzLiAuLi4gUG9rw6ltb24KICAgICAgICAgICAgIyBla2xlbmlya2VuIGVza2kvc3RhbGUgZWxlbWVudCBrdWxsYW7EsWxtxLF5b3IuCiAgICAgICAgICAgIGFzeW5jIGRlZiBmaXJzdF92aXNpYmxlKGxvY2F0b3IpOgogICAgICAgICAgICAgICAgZm9yIGkgaW4gcmFuZ2UoYXdhaXQgbG9jYXRvci5jb3VudCgpKToKICAgICAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgIGVsID0gbG9jYXRvci5udGgoaSkKICAgICAgICAgICAgICAgICAgICAgICAgaWYgYXdhaXQgZWwuaXNfdmlzaWJsZSgpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgcmV0dXJuIGVsCiAgICAgICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbjoKICAgICAgICAgICAgICAgICAgICAgICAgcGFzcwogICAgICAgICAgICAgICAgcmV0dXJuIE5vbmUKCiAgICAgICAgICAgIGFzeW5jIGRlZiBnZXRfcG9rZW1vbl9jb250cm9sKCk6CiAgICAgICAgICAgICAgICAjIFY0NCd0ZSDDp2FsxLHFn2FuIGdlcsOnZWsgUG9nb0Nvb3JkaW5hdGVzIG1hbnTEscSfxLE6CiAgICAgICAgICAgICAgICAjIMOWbmNlIGlucHV0LCBidWx1bmFtYXpzYSBla3JhbmRha2kgIkFueSBQb2tlbW9uIiBhbGFuxLFuxLFuCiAgICAgICAgICAgICAgICAjIGfDtnLDvG5lbiBtZXRuaW5pIHTEsWtsYW5hYmlsaXIga29udHJvbCBvbGFyYWsga3VsbGFuLgogICAgICAgICAgICAgICAgY29udHJvbCA9IGF3YWl0IGZpcnN0X3Zpc2libGUocGFnZS5sb2NhdG9yKAogICAgICAgICAgICAgICAgICAgICdpbnB1dFtwbGFjZWhvbGRlcj0iQW55IFBva2Vtb24iXSwgJwogICAgICAgICAgICAgICAgICAgICdpbnB1dFtwbGFjZWhvbGRlcj0iQW55IFBva8OpbW9uIl0sICcKICAgICAgICAgICAgICAgICAgICAnaW5wdXRbcGxhY2Vob2xkZXIqPSJBbnkgUG9rZW1vbiIgaV0sICcKICAgICAgICAgICAgICAgICAgICAnaW5wdXRbYXJpYS1sYWJlbCo9IlBva2Vtb24iIGldLCAnCiAgICAgICAgICAgICAgICAgICAgJ2lucHV0W3JvbGU9ImNvbWJvYm94Il0sICcKICAgICAgICAgICAgICAgICAgICAnW3JvbGU9ImNvbWJvYm94Il0gaW5wdXQnCiAgICAgICAgICAgICAgICApKQogICAgICAgICAgICAgICAgaWYgY29udHJvbCBpcyBub3QgTm9uZToKICAgICAgICAgICAgICAgICAgICByZXR1cm4gY29udHJvbAoKICAgICAgICAgICAgICAgIGZvciBsYWJlbCBpbiAoIkFueSBQb2tlbW9uIiwgIkFueSBQb2vDqW1vbiIpOgogICAgICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICAgICAgdHh0ID0gcGFnZS5nZXRfYnlfdGV4dChsYWJlbCwgZXhhY3Q9VHJ1ZSkKICAgICAgICAgICAgICAgICAgICAgICAgY29udHJvbCA9IGF3YWl0IGZpcnN0X3Zpc2libGUodHh0KQogICAgICAgICAgICAgICAgICAgICAgICBpZiBjb250cm9sIGlzIG5vdCBOb25lOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgcmV0dXJuIGNvbnRyb2wKICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgICAgICAgICAgICAgICAgICBwYXNzCgogICAgICAgICAgICAgICAgIyBTb24gw6dhcmU6IGZvcm11biBpbGsgZ8O2csO8bmVuIHRleHQgaW5wdXQndS4KICAgICAgICAgICAgICAgIHJldHVybiBhd2FpdCBmaXJzdF92aXNpYmxlKHBhZ2UubG9jYXRvcigKICAgICAgICAgICAgICAgICAgICAnaW5wdXRbdHlwZT0idGV4dCJdLCBpbnB1dDpub3QoW3R5cGVdKScKICAgICAgICAgICAgICAgICkpCgogICAgICAgICAgICBwb2tlbW9uID0gYXdhaXQgZ2V0X3Bva2Vtb25fY29udHJvbCgpCiAgICAgICAgICAgIGlmIHBva2Vtb24gaXMgTm9uZToKICAgICAgICAgICAgICAgIHNheSgKICAgICAgICAgICAgICAgICAgICAiRVJST1IiLAogICAgICAgICAgICAgICAgICAgICJFUlItUEctRk9STS1QT0tFTU9OLUZJRUxEfFBvZ29Db29yZGluYXRlcyDDvHplcmluZGVraSAiCiAgICAgICAgICAgICAgICAgICAgIkFueSBQb2tlbW9uIGFsYW7EsSBidWx1bmFtYWTEsS4iCiAgICAgICAgICAgICAgICApCiAgICAgICAgICAgICAgICByZXR1cm4gMTMKCiAgICAgICAgICAgIHNheSgKICAgICAgICAgICAgICAgICJDT05UUk9MUyIsCiAgICAgICAgICAgICAgICBzdHIobGVuKG5hbWVzKSkgKyAiIFBva8OpbW9uIGF5bsSxIGFyYW1hIGFsYW7EsW5hIGVrbGVuZWNlay4iCiAgICAgICAgICAgICkKCiAgICAgICAgICAgIGFzeW5jIGRlZiBzZWxlY3RfcG9rZW1vbihwb2tlbW9uX25hbWUsIGluZGV4KToKICAgICAgICAgICAgICAgICMgSGVyIHNlw6dpbWRlbiBzb25yYSBQb2dvQ29vcmRpbmF0ZXMgRE9NJ3UgZGXEn2nFn2ViaWxpci4KICAgICAgICAgICAgICAgIGZpZWxkID0gYXdhaXQgZ2V0X3Bva2Vtb25fY29udHJvbCgpCiAgICAgICAgICAgICAgICBpZiBmaWVsZCBpcyBOb25lOgogICAgICAgICAgICAgICAgICAgIHNheSgKICAgICAgICAgICAgICAgICAgICAgICAgIkVSUk9SIiwKICAgICAgICAgICAgICAgICAgICAgICAgIkVSUi1QRy1GSUVMRC0iICsgc3RyKGluZGV4KSArCiAgICAgICAgICAgICAgICAgICAgICAgICJ8UG9rw6ltb24gYXJhbWEgYWxhbsSxIHllbmlkZW4gYnVsdW5hbWFkxLE6ICIgKwogICAgICAgICAgICAgICAgICAgICAgICBwb2tlbW9uX25hbWUKICAgICAgICAgICAgICAgICAgICApCiAgICAgICAgICAgICAgICAgICAgcmV0dXJuIEZhbHNlCgogICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgIGF3YWl0IGZpZWxkLmNsaWNrKCkKICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZXg6CiAgICAgICAgICAgICAgICAgICAgc2F5KAogICAgICAgICAgICAgICAgICAgICAgICAiRVJST1IiLAogICAgICAgICAgICAgICAgICAgICAgICAiRVJSLVBHLUZJRUxELUNMSUNLLSIgKyBzdHIoaW5kZXgpICsKICAgICAgICAgICAgICAgICAgICAgICAgInwiICsgdHlwZShleCkuX19uYW1lX18gKyAiOiAiICsgc3RyKGV4KQogICAgICAgICAgICAgICAgICAgICkKICAgICAgICAgICAgICAgICAgICByZXR1cm4gRmFsc2UKCiAgICAgICAgICAgICAgICAjIMOWTkVNTMSwOiBDdHJsK0EgLyBCYWNrc3BhY2UgeW9rLgogICAgICAgICAgICAgICAgIyDDh29rbHUgc2XDp2ltZGUgw7ZuY2VraSBjaGlwJ2xlcmUgZG9rdW5tYWRhbiB5YWxuxLF6Y2EKICAgICAgICAgICAgICAgICMga2xhdnllZGVuIHllbmkgUG9rw6ltb24gaXNtaW5pIHlhesSxeW9ydXouCgogICAgICAgICAgICAgICAgYXdhaXQgZmllbGQudHlwZShwb2tlbW9uX25hbWUsIGRlbGF5PTM1KQogICAgICAgICAgICAgICAgYXdhaXQgcGFnZS53YWl0X2Zvcl90aW1lb3V0KDgwMCkKCiAgICAgICAgICAgICAgICAjIERyb3Bkb3duIMO2bmVyaXNpbmkgaGVkZWZsZS4gw5ZuY2VsaWsgdGFtIGlzaW0uCiAgICAgICAgICAgICAgICBzdWdnZXN0aW9uID0gTm9uZQogICAgICAgICAgICAgICAgZXhhY3QgPSBwYWdlLmdldF9ieV90ZXh0KHBva2Vtb25fbmFtZSwgZXhhY3Q9VHJ1ZSkKICAgICAgICAgICAgICAgIGZvciBpIGluIHJhbmdlKGF3YWl0IGV4YWN0LmNvdW50KCkpOgogICAgICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICAgICAgZWwgPSBleGFjdC5udGgoaSkKICAgICAgICAgICAgICAgICAgICAgICAgaWYgbm90IGF3YWl0IGVsLmlzX3Zpc2libGUoKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGNvbnRpbnVlCiAgICAgICAgICAgICAgICAgICAgICAgIHRhZyA9IChhd2FpdCBlbC5ldmFsdWF0ZSgiKGUpPT5lLnRhZ05hbWUudG9Mb3dlckNhc2UoKSIpKS5sb3dlcigpCiAgICAgICAgICAgICAgICAgICAgICAgIGlmIHRhZyBpbiAoImlucHV0IiwgInRleHRhcmVhIiwgInNlbGVjdCIpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgY29udGludWUKICAgICAgICAgICAgICAgICAgICAgICAgIyBDaGlwIHllcmluZSBkcm9wZG93biBzZcOnZW5lxJ9pbmkgdGVyY2loIGV0LgogICAgICAgICAgICAgICAgICAgICAgICByb2xlID0gKGF3YWl0IGVsLmdldF9hdHRyaWJ1dGUoInJvbGUiKSBvciAiIikubG93ZXIoKQogICAgICAgICAgICAgICAgICAgICAgICBjbHMgPSAoYXdhaXQgZWwuZ2V0X2F0dHJpYnV0ZSgiY2xhc3MiKSBvciAiIikubG93ZXIoKQogICAgICAgICAgICAgICAgICAgICAgICBpZiByb2xlID09ICJvcHRpb24iIG9yICJvcHRpb24iIGluIGNscyBvciAiYXV0b2NvbXBsZXRlIiBpbiBjbHM6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBzdWdnZXN0aW9uID0gZWwKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGJyZWFrCiAgICAgICAgICAgICAgICAgICAgICAgIGlmIHN1Z2dlc3Rpb24gaXMgTm9uZToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHN1Z2dlc3Rpb24gPSBlbAogICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICAgICAgICAgIHBhc3MKCiAgICAgICAgICAgICAgICAjIEV4YWN0IGJ1bHVuYW1henNhIGfDtnLDvG5lbiBvcHRpb24vTEknbGFyZGEgYXJhLgogICAgICAgICAgICAgICAgaWYgc3VnZ2VzdGlvbiBpcyBOb25lOgogICAgICAgICAgICAgICAgICAgIGNhbmRpZGF0ZXMgPSBwYWdlLmxvY2F0b3IoCiAgICAgICAgICAgICAgICAgICAgICAgICdbcm9sZT0ib3B0aW9uIl0sIGxpLCBbZGF0YS12YWx1ZV0sIFtkYXRhLW9wdGlvbl0sICcKICAgICAgICAgICAgICAgICAgICAgICAgJ2J1dHRvbicKICAgICAgICAgICAgICAgICAgICApCiAgICAgICAgICAgICAgICAgICAgZm9yIGkgaW4gcmFuZ2UobWluKGF3YWl0IGNhbmRpZGF0ZXMuY291bnQoKSwgMzAwKSk6CiAgICAgICAgICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGVsID0gY2FuZGlkYXRlcy5udGgoaSkKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGlmIG5vdCBhd2FpdCBlbC5pc192aXNpYmxlKCk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgY29udGludWUKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHR4dCA9IChhd2FpdCBlbC5pbm5lcl90ZXh0KHRpbWVvdXQ9MzAwKSBvciAiIikuc3RyaXAoKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgdHh0Lmxvd2VyKCkgPT0gcG9rZW1vbl9uYW1lLmxvd2VyKCkgb3IgKAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIHBva2Vtb25fbmFtZS5sb3dlcigpIGluIHR4dC5sb3dlcigpIGFuZCBsZW4odHh0KSA8IDEwMAogICAgICAgICAgICAgICAgICAgICAgICAgICAgKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBzdWdnZXN0aW9uID0gZWwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBicmVhawogICAgICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgcGFzcwoKICAgICAgICAgICAgICAgIGlmIHN1Z2dlc3Rpb24gaXMgTm9uZToKICAgICAgICAgICAgICAgICAgICBzYXkoCiAgICAgICAgICAgICAgICAgICAgICAgICJFUlJPUiIsCiAgICAgICAgICAgICAgICAgICAgICAgIHBva2Vtb25fbmFtZSArICIgacOnaW4gYcOnxLFsYW4gc2XDp2ltIMO2bmVyaXNpIGJ1bHVuYW1hZMSxLiIKICAgICAgICAgICAgICAgICAgICApCiAgICAgICAgICAgICAgICAgICAgcmV0dXJuIEZhbHNlCgogICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgIGF3YWl0IHN1Z2dlc3Rpb24uY2xpY2soKQogICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbjoKICAgICAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IHN1Z2dlc3Rpb24ubG9jYXRvcigieHBhdGg9Li4iKS5jbGljaygpCiAgICAgICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbjoKICAgICAgICAgICAgICAgICAgICAgICAgIyBEcm9wZG93biBrbGF2eWUgaWxlIGRlIHNlw6dpbGViaWxpeW9yLgogICAgICAgICAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBhd2FpdCBmaWVsZC5wcmVzcygiQXJyb3dEb3duIikKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IGZpZWxkLnByZXNzKCJFbnRlciIpCiAgICAgICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZXg6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBzYXkoIkVSUk9SIiwgcG9rZW1vbl9uYW1lICsgIiBzZcOnaWxlbWVkaTogIiArIHJlcHIoZXgpKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgcmV0dXJuIEZhbHNlCgogICAgICAgICAgICAgICAgYXdhaXQgcGFnZS53YWl0X2Zvcl90aW1lb3V0KDYwMCkKCiAgICAgICAgICAgICAgICAjIFNlw6dpbSBzb25yYXPEsSBpbnB1dCB0ZWtyYXIgZ8O2csO8bsO8eW9yc2EgdmUgaMOibMOiIHlhesSxIHZhcnNhLAogICAgICAgICAgICAgICAgIyBFbnRlciBpbGUgY2hpcCdlIGTDtm7DvMWfdMO8cm1leWkgZGVuZS4gw5ZuY2VraSBjaGlwJ2xlcmUgZG9rdW5tYS4KICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICBjdXJyZW50X3ZhbHVlID0gKGF3YWl0IGZpZWxkLmlucHV0X3ZhbHVlKCkpLnN0cmlwKCkKICAgICAgICAgICAgICAgICAgICBpZiBjdXJyZW50X3ZhbHVlOgogICAgICAgICAgICAgICAgICAgICAgICBhd2FpdCBmaWVsZC5wcmVzcygiRW50ZXIiKQogICAgICAgICAgICAgICAgICAgICAgICBhd2FpdCBwYWdlLndhaXRfZm9yX3RpbWVvdXQoMzUwKQogICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbjoKICAgICAgICAgICAgICAgICAgICBwYXNzCiAgICAgICAgICAgICAgICBzYXkoCiAgICAgICAgICAgICAgICAgICAgIkNPTlRST0xTIiwKICAgICAgICAgICAgICAgICAgICBmIntpbmRleH0ve2xlbihuYW1lcyl9IOKAoiB7cG9rZW1vbl9uYW1lfSBzZcOnaWxkaS4iCiAgICAgICAgICAgICAgICApCiAgICAgICAgICAgICAgICByZXR1cm4gVHJ1ZQoKICAgICAgICAgICAgc2VsZWN0ZWRfbmFtZXMgPSBbXQogICAgICAgICAgICBmb3IgaW5kZXgsIHBva2Vtb25fbmFtZSBpbiBlbnVtZXJhdGUobmFtZXMsIDEpOgogICAgICAgICAgICAgICAgaWYgbm90IGF3YWl0IHNlbGVjdF9wb2tlbW9uKHBva2Vtb25fbmFtZSwgaW5kZXgpOgogICAgICAgICAgICAgICAgICAgIHJldHVybiAxOAogICAgICAgICAgICAgICAgc2VsZWN0ZWRfbmFtZXMuYXBwZW5kKHBva2Vtb25fbmFtZSkKCiAgICAgICAgICAgIHNheSgKICAgICAgICAgICAgICAgICJDT05UUk9MUyIsCiAgICAgICAgICAgICAgICAi4pyTICIgKyBzdHIobGVuKHNlbGVjdGVkX25hbWVzKSkgKwogICAgICAgICAgICAgICAgIiBQb2vDqW1vbiBheW7EsSBhcmFtYSBhbGFuxLFuYSBla2xlbmRpLiIKICAgICAgICAgICAgKQoKICAgICAgICAgICAgIyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KICAgICAgICAgICAgIyBMxLBHIEFZQVJJCiAgICAgICAgICAgICMgUG9nb0Nvb3JkaW5hdGVzJ2luIGN1c3RvbSBkcm9wZG93biB5YXDEsXPEsSBuZWRlbml5bGUgTWFzdGVyCiAgICAgICAgICAgICMgc2XDp2VuZcSfaSBoZXIgemFtYW4gRE9NJ2RhIDxzZWxlY3Q+IG9sYXJhayBidWx1bm1heWFiaWxpeW9yLgogICAgICAgICAgICAjIMOWbmNlIG5hdGl2ZSBzZWxlY3QnaSwgc29ucmEgZ8O2csO8bmVuIGRyb3Bkb3duIG1ldG5pbmkgZGVuZXJpei4KICAgICAgICAgICAgIyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KICAgICAgICAgICAgYXN5bmMgZGVmIHNldF9sZWFndWVfdWkod2FudGVkX2xlYWd1ZSk6CiAgICAgICAgICAgICAgICB3YW50ZWQgPSBzdHIod2FudGVkX2xlYWd1ZSBvciAiR3JlYXQgTGVhZ3VlIikuc3RyaXAoKQogICAgICAgICAgICAgICAgbG93ID0gd2FudGVkLmxvd2VyKCkKICAgICAgICAgICAgICAgIGlmICJncmVhdCIgaW4gbG93OgogICAgICAgICAgICAgICAgICAgIGFsaWFzZXMgPSBbIkdyZWF0IiwgIkdyZWF0IExlYWd1ZSJdCiAgICAgICAgICAgICAgICAgICAgdGFyZ2V0X3Nob3J0ID0gIkdyZWF0IgogICAgICAgICAgICAgICAgZWxpZiAidWx0cmEiIGluIGxvdzoKICAgICAgICAgICAgICAgICAgICBhbGlhc2VzID0gWyJVbHRyYSIsICJVbHRyYSBMZWFndWUiXQogICAgICAgICAgICAgICAgICAgIHRhcmdldF9zaG9ydCA9ICJVbHRyYSIKICAgICAgICAgICAgICAgIGVsc2U6CiAgICAgICAgICAgICAgICAgICAgYWxpYXNlcyA9IFt3YW50ZWRdCiAgICAgICAgICAgICAgICAgICAgdGFyZ2V0X3Nob3J0ID0gd2FudGVkCgogICAgICAgICAgICAgICAgc2F5KCJTVEVQIiwgIlsxLzZdIEdyZWF0L1VsdHJhIGFyYW1hc8SxIGJhxZ9sYXTEsWzEsXlvcjogIiArIHdhbnRlZCkKICAgICAgICAgICAgICAgIHNheSgiU1RFUCIsICJbMi82XSDDlm5jZSBQdlAgc2VrbWVzaSBhcmFuxLF5b3IgdmUgc2XDp2lsaXlvcuKApiIpCgogICAgICAgICAgICAgICAgIyBQb2dvQ29vcmRpbmF0ZXMgYcOnxLFsZMSxxJ/EsW5kYSB2YXJzYXnEsWxhbiBzZWttZSBQb2vDqW1vbidkdXIuCiAgICAgICAgICAgICAgICAjIEdyZWF0L1VsdHJhIHNlw6dpbWkgUHZQIHNla21lc2luaW4gacOnaW5kZWtpIGtvbnRyb2xkZSBidWx1bnVyLgogICAgICAgICAgICAgICAgcHZwX2NsaWNrZWQgPSBGYWxzZQogICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgIGNhbmRpZGF0ZXMgPSBwYWdlLmdldF9ieV90ZXh0KCJQdlAiLCBleGFjdD1UcnVlKQogICAgICAgICAgICAgICAgICAgIGZvciBpIGluIHJhbmdlKGF3YWl0IGNhbmRpZGF0ZXMuY291bnQoKSk6CiAgICAgICAgICAgICAgICAgICAgICAgIGVsID0gY2FuZGlkYXRlcy5udGgoaSkKICAgICAgICAgICAgICAgICAgICAgICAgaWYgYXdhaXQgZWwuaXNfdmlzaWJsZSgpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgYXdhaXQgZWwuY2xpY2sodGltZW91dD0xNTAwKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgYXdhaXQgcGFnZS53YWl0X2Zvcl90aW1lb3V0KDYwMCkKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHB2cF9jbGlja2VkID0gVHJ1ZQogICAgICAgICAgICAgICAgICAgICAgICAgICAgc2F5KCJDT05UUk9MUyIsICJQdlAgc2VrbWVzaSBzZcOnaWxkaS4iKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgYnJlYWsKICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZToKICAgICAgICAgICAgICAgICAgICBzYXkoIkRFQlVHIiwgIlB2UCBzZWttZXNpIHTEsWtsYW1hIGV4Y2VwdGlvbjogIiArIHJlcHIoZSkpCgogICAgICAgICAgICAgICAgaWYgbm90IHB2cF9jbGlja2VkOgogICAgICAgICAgICAgICAgICAgICMgQmF6xLEgc8O8csO8bWxlcmRlIGFyaWEtbGFiZWwvdGl0bGUgaWxlIGJ1bHVuYWJpbGlyLgogICAgICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICAgICAgbG9jID0gcGFnZS5sb2NhdG9yKCdidXR0b24sIFtyb2xlPSJidXR0b24iXScpLmZpbHRlcihoYXNfdGV4dD0iUHZQIikKICAgICAgICAgICAgICAgICAgICAgICAgZm9yIGkgaW4gcmFuZ2UoYXdhaXQgbG9jLmNvdW50KCkpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgZWwgPSBsb2MubnRoKGkpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiBhd2FpdCBlbC5pc192aXNpYmxlKCk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgYXdhaXQgZWwuY2xpY2sodGltZW91dD0xNTAwKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IHBhZ2Uud2FpdF9mb3JfdGltZW91dCg2MDApCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgcHZwX2NsaWNrZWQgPSBUcnVlCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgc2F5KCJDT05UUk9MUyIsICJQdlAgc2VrbWVzaSBzZcOnaWxkaSAoYnV0dG9uIGZhbGxiYWNrKS4iKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGJyZWFrCiAgICAgICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOgogICAgICAgICAgICAgICAgICAgICAgICBzYXkoIkRFQlVHIiwgIlB2UCBidXR0b24gZmFsbGJhY2sgZXhjZXB0aW9uOiAiICsgcmVwcihlKSkKCiAgICAgICAgICAgICAgICBpZiBub3QgcHZwX2NsaWNrZWQ6CiAgICAgICAgICAgICAgICAgICAgc2F5KCJFUlJPUiIsICJFUlItUEctUFZQLUNPTlRST0wtTk9ULUZPVU5EfFBvZ29Db29yZGluYXRlcyDDvHplcmluZGUgUHZQIHNla21lc2kgYnVsdW5hbWFkxLEuIikKICAgICAgICAgICAgICAgICAgICByZXR1cm4gRmFsc2UKCiAgICAgICAgICAgICAgICBzYXkoIlNURVAiLCAiWzMvNl0gUHZQIHNla21lc2kgYcOnxLFsZMSxOyBHcmVhdC9VbHRyYSBrb250cm9sw7wgYXJhbsSxeW9y4oCmIikKICAgICAgICAgICAgICAgICMgw5ZuY2UgZ2Vyw6dlayBzZWxlY3RsZXJpIGtvbnRyb2wgZXQuIEJ1cmFkYSB5YWxuxLF6Y2EgR3JlYXQgKyBVbHRyYQogICAgICAgICAgICAgICAgIyBzZcOnZW5la2xlcmkgb2xhbiBiaXIgc2VsZWN0IGthYnVsIGVkaWxpcjsgTWFzdGVyIGtlc2lubGlrbGUgYXJhbm1hei4KICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICByZXN1bHQgPSBhd2FpdCBwYWdlLmV2YWx1YXRlKCIiIgogICAgICAgICAgICAgICAgICAgIChwYXlsb2FkKSA9PiB7CiAgICAgICAgICAgICAgICAgICAgICAgIGNvbnN0IG5vcm0gPSBzID0+IChzIHx8ICcnKS50b1N0cmluZygpLnRyaW0oKS50b0xvd2VyQ2FzZSgpOwogICAgICAgICAgICAgICAgICAgICAgICBjb25zdCB3YW50ZWQgPSBub3JtKHBheWxvYWQudGFyZ2V0KTsKICAgICAgICAgICAgICAgICAgICAgICAgZm9yIChjb25zdCBzIG9mIEFycmF5LmZyb20oZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnc2VsZWN0JykpKSB7CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBjb25zdCBvcHRzID0gQXJyYXkuZnJvbShzLm9wdGlvbnMgfHwgW10pOwogICAgICAgICAgICAgICAgICAgICAgICAgICAgY29uc3QgdGV4dHMgPSBvcHRzLm1hcChvID0+IG5vcm0oby50ZXh0KSk7CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiAoISh0ZXh0cy5zb21lKHggPT4geCA9PT0gJ2dyZWF0JyB8fCB4LmluY2x1ZGVzKCdncmVhdCBsZWFndWUnKSkgJiYKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIHRleHRzLnNvbWUoeCA9PiB4ID09PSAndWx0cmEnIHx8IHguaW5jbHVkZXMoJ3VsdHJhIGxlYWd1ZScpKSkpIGNvbnRpbnVlOwogICAgICAgICAgICAgICAgICAgICAgICAgICAgY29uc3Qgb3B0ID0gb3B0cy5maW5kKG8gPT4gewogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGNvbnN0IHQ9bm9ybShvLnRleHQpLCB2PW5vcm0oby52YWx1ZSk7CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgcmV0dXJuIHQgPT09IHdhbnRlZCB8fCB2ID09PSB3YW50ZWQgfHwgdC5pbmNsdWRlcyh3YW50ZWQpIHx8IHYuaW5jbHVkZXMod2FudGVkKTsKICAgICAgICAgICAgICAgICAgICAgICAgICAgIH0pOwogICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgKCFvcHQpIHJldHVybiB7b2s6ZmFsc2UsIHJlYXNvbjonR3JlYXQvVWx0cmEgc2VsZWN0IGJ1bHVuZHUgZmFrYXQgaGVkZWYgYnVsdW5hbWFkxLEnLCBvcHRpb25zOm9wdHMubWFwKG89Pm8udGV4dCl9OwogICAgICAgICAgICAgICAgICAgICAgICAgICAgcy52YWx1ZT1vcHQudmFsdWU7CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBzLmRpc3BhdGNoRXZlbnQobmV3IEV2ZW50KCdpbnB1dCcse2J1YmJsZXM6dHJ1ZX0pKTsKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHMuZGlzcGF0Y2hFdmVudChuZXcgRXZlbnQoJ2NoYW5nZScse2J1YmJsZXM6dHJ1ZX0pKTsKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHJldHVybiB7b2s6dHJ1ZSx0ZXh0Om9wdC50ZXh0LHZhbHVlOm9wdC52YWx1ZX07CiAgICAgICAgICAgICAgICAgICAgICAgIH0KICAgICAgICAgICAgICAgICAgICAgICAgcmV0dXJuIHtvazpmYWxzZSwgcmVhc29uOidHcmVhdC9VbHRyYSBuYXRpdmUgc2VsZWN0IGJ1bHVuYW1hZMSxJ307CiAgICAgICAgICAgICAgICAgICAgfQogICAgICAgICAgICAgICAgICAgICIiIiwgeyJ0YXJnZXQiOiB0YXJnZXRfc2hvcnR9KQogICAgICAgICAgICAgICAgICAgIHNheSgiREVCVUciLCAiR3JlYXQvVWx0cmEgbmF0aXZlIHNlbGVjdCBzb251Y3U6ICIgKyBqc29uLmR1bXBzKHJlc3VsdCwgZW5zdXJlX2FzY2lpPUZhbHNlKSkKICAgICAgICAgICAgICAgICAgICBpZiByZXN1bHQgYW5kIHJlc3VsdC5nZXQoIm9rIik6CiAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IHBhZ2Uud2FpdF9mb3JfdGltZW91dCg3MDApCiAgICAgICAgICAgICAgICAgICAgICAgIHNheSgiQ09OVFJPTFMiLCAiTGVhZ3VlID0gIiArIHN0cihyZXN1bHQuZ2V0KCJ0ZXh0IikpICsgIiBzZcOnaWxkaS4iKQogICAgICAgICAgICAgICAgICAgICAgICByZXR1cm4gVHJ1ZQogICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOgogICAgICAgICAgICAgICAgICAgIHNheSgiREVCVUciLCAiR3JlYXQvVWx0cmEgbmF0aXZlIHNlbGVjdCBleGNlcHRpb246ICIgKyByZXByKGUpKQoKICAgICAgICAgICAgICAgIHNheSgiU1RFUCIsICJbNC82XSBBw6fEsWxhbiBQdlAga29udHJvbMO8bmRlICIgKyB0YXJnZXRfc2hvcnQgKyAiIGFyYW7EsXlvcuKApiIpCiAgICAgICAgICAgICAgICAjIE5hdGl2ZSBzZWxlY3QgeW9rc2EsIFB2UCBzZWttZXNpbmluIGHDp3TEscSfxLEgZ8O2csO8bsO8ciBidXRvbi9vcHRpb24KICAgICAgICAgICAgICAgICMgw7x6ZXJpbmRlIHNhZGVjZSBHcmVhdCB2ZXlhIFVsdHJhIGFyYW7EsXIuCiAgICAgICAgICAgICAgICBmb3IgdGFyZ2V0X25hbWUgaW4gYWxpYXNlczoKICAgICAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgIGxvYyA9IHBhZ2UuZ2V0X2J5X3RleHQodGFyZ2V0X25hbWUsIGV4YWN0PVRydWUpCiAgICAgICAgICAgICAgICAgICAgICAgIGNvdW50ID0gYXdhaXQgbG9jLmNvdW50KCkKICAgICAgICAgICAgICAgICAgICAgICAgc2F5KCJERUJVRyIsICJQdlAgaGVkZWZpICciICsgdGFyZ2V0X25hbWUgKyAiJyDihpIgIiArIHN0cihjb3VudCkgKyAiIGFkZXQiKQogICAgICAgICAgICAgICAgICAgICAgICBmb3IgaSBpbiByYW5nZShjb3VudCk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBlbCA9IGxvYy5udGgoaSkKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGlmIGF3YWl0IGVsLmlzX3Zpc2libGUoKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBhd2FpdCBlbC5jbGljayh0aW1lb3V0PTE1MDApCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgYXdhaXQgcGFnZS53YWl0X2Zvcl90aW1lb3V0KDcwMCkKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBzYXkoIkNPTlRST0xTIiwgIkxlYWd1ZSA9ICIgKyB0YXJnZXRfbmFtZSArICIgc2XDp2lsZGkuIikKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICByZXR1cm4gVHJ1ZQogICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZToKICAgICAgICAgICAgICAgICAgICAgICAgc2F5KCJERUJVRyIsICJQdlAgaGVkZWYgYXJhbWEvdMSxa2xhbWEgZXhjZXB0aW9uICciICsgdGFyZ2V0X25hbWUgKyAiJzogIiArIHJlcHIoZSkpCgogICAgICAgICAgICAgICAgc2F5KCJTVEVQIiwgIls1LzZdIEdyZWF0L1VsdHJhIGtvbnRyb2zDvCBidWx1bmFtYWTEsTsgZ8O2csO8bsO8ciBzZcOnZW5la2xlciB0ZcWfaGlzIGVkaWxpeW9y4oCmIikKICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICB2YWxzID0gYXdhaXQgcGFnZS5sb2NhdG9yKCdbcm9sZT0ib3B0aW9uIl0sIG9wdGlvbiwgbGksIGJ1dHRvbiwgW3JvbGU9ImJ1dHRvbiJdJykuYWxsX2lubmVyX3RleHRzKCkKICAgICAgICAgICAgICAgICAgICB2YWxzPVsoeCBvciAnJykuc3RyaXAoKS5yZXBsYWNlKCdcbicsJyAnKSBmb3IgeCBpbiB2YWxzIGlmICh4IG9yICcnKS5zdHJpcCgpXQogICAgICAgICAgICAgICAgICAgIHNheSgiREVCVUciLCAiR8O2csO8bsO8ciBQdlAgc2XDp2VuZWtsZXJpOiAiICsgIiB8ICIuam9pbih2YWxzWzo2MF0pKQogICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOgogICAgICAgICAgICAgICAgICAgIHNheSgiREVCVUciLCAiUHZQIHNlw6dlbmVrIHRlxZ9oaXNpIGJhxZ9hcsSxc8SxejogIiArIHJlcHIoZSkpCiAgICAgICAgICAgICAgICBzYXkoIkVSUk9SIiwgIkVSUi1QRy1HUkVBVC1VTFRSQS1DT05UUk9MLU5PVC1GT1VORHxQdlAgc2VrbWVzaW5kZW4gR3JlYXQvVWx0cmEgc2XDp2ltaSBidWx1bmFtYWTEsS4iKQogICAgICAgICAgICAgICAgcmV0dXJuIEZhbHNlCgogICAgICAgICAgICAjIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLQogICAgICAgICAgICAjIEzEsEcgREFWUkFOScWeSQogICAgICAgICAgICAjIE1hc3RlciBMZWFndWUgUG9nb0Nvb3JkaW5hdGVzJ3RlIEdyZWF0L1VsdHJhIGdpYmkgTGVhZ3VlCiAgICAgICAgICAgICMgZHJvcGRvd24nxLFuZGFuIHNlw6dpbG1lei4gS3VsbGFuxLFjxLEgaXN0ZcSfaW5lIGfDtnJlIE1hc3RlcidkYQogICAgICAgICAgICAjIGlsayBhxZ9hbWEgZG/En3J1ZGFuIFN0YXRzIGZpbHRyZXNpZGlyOyBhcmTEsW5kYW4gTGV2ZWwgTWluPTM0CiAgICAgICAgICAgICMgdmUgSVYgTWluPTEwMCB5YXrEsWzEsXIuIELDtnlsZWNlIE1hc3RlciBhcmFtYXPEsW5kYSBQcm9jZXNzIDE2CiAgICAgICAgICAgICMgw7xyZXRlbiBMZWFndWUgZHJvcGRvd24gZGVuZW1lc2kgdGFtYW1lbiBhdGxhbsSxci4KICAgICAgICAgICAgIyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KICAgICAgICAgICAgaXNfbWFzdGVyID0gKCJtYXN0ZXIiIGluIHN0cihsZWFndWUpLmxvd2VyKCkpIG9yICgibWVnYSIgaW4gc3RyKGxlYWd1ZSkubG93ZXIoKSkKICAgICAgICAgICAgaWYgaXNfbWFzdGVyOgogICAgICAgICAgICAgICAgc2F5KCJTVEVQIiwgIlsxLzVdICIgKyAoIk1lZ2EiIGlmICJtZWdhIiBpbiBzdHIobGVhZ3VlKS5sb3dlcigpIGVsc2UgIk1hc3RlciIpICsgIiBtb2R1OiBMZWFndWUgZHJvcGRvd24gQVlBUkxBTk1BWUFDQUsuIikKICAgICAgICAgICAgICAgIHNheSgiU1RFUCIsICJbMi81XSDDlm5jZSBTdGF0cyBzZcOnaW1pIGtvbnRyb2wgZWRpbGl5b3LigKYiKQogICAgICAgICAgICBlbHNlOgogICAgICAgICAgICAgICAgbGVhZ3VlX2RvbmUgPSBhd2FpdCBzZXRfbGVhZ3VlX3VpKGxlYWd1ZSkKICAgICAgICAgICAgICAgIGlmIG5vdCBsZWFndWVfZG9uZToKICAgICAgICAgICAgICAgICAgICBzYXkoIkVSUk9SIiwgIkxlYWd1ZSBheWFybGFuYW1hZMSxOiAiICsgbGVhZ3VlKQogICAgICAgICAgICAgICAgICAgIHJldHVybiAxNgogICAgICAgICAgICAgICAgc2F5KCJTVEVQIiwgIlsxLzVdICIgKyBsZWFndWUgKyAiIExlYWd1ZSBzZcOnaW1pIHRhbWFtbGFuZMSxLiIpCiAgICAgICAgICAgICAgICBzYXkoIlNURVAiLCAiWzIvNV0gUHZQIGZpbHRyZXNpIGhhesSxcmxhbsSxeW9y4oCmIikKCiAgICAgICAgICAgIGFzeW5jIGRlZiBjbGlja19leGFjdF92aXNpYmxlKHRleHQpOgogICAgICAgICAgICAgICAgbG9jID0gcGFnZS5nZXRfYnlfdGV4dCh0ZXh0LCBleGFjdD1UcnVlKQogICAgICAgICAgICAgICAgZm9yIGkgaW4gcmFuZ2UoYXdhaXQgbG9jLmNvdW50KCkpOgogICAgICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICAgICAgZWwgPSBsb2MubnRoKGkpCiAgICAgICAgICAgICAgICAgICAgICAgIGlmIGF3YWl0IGVsLmlzX3Zpc2libGUoKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IGVsLmNsaWNrKCkKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IHBhZ2Uud2FpdF9mb3JfdGltZW91dCgzMDApCiAgICAgICAgICAgICAgICAgICAgICAgICAgICByZXR1cm4gVHJ1ZQogICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICAgICAgICAgIHBhc3MKICAgICAgICAgICAgICAgIHJldHVybiBGYWxzZQoKICAgICAgICAgICAgYXN5bmMgZGVmIHNldF9taW5fZmllbGQobGFiZWxfbmFtZSwgdmFsdWUpOgogICAgICAgICAgICAgICAgIyBMYWJlbMSxbiBidWx1bmR1xJ91IGdydXAgacOnaW5kZWtpIGlsayBpbnB1dCA9IE1pbi4KICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICBsYWJlbHMgPSBwYWdlLmdldF9ieV90ZXh0KGxhYmVsX25hbWUsIGV4YWN0PVRydWUpCiAgICAgICAgICAgICAgICAgICAgZm9yIGkgaW4gcmFuZ2UoYXdhaXQgbGFiZWxzLmNvdW50KCkpOgogICAgICAgICAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBsYWIgPSBsYWJlbHMubnRoKGkpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiBub3QgYXdhaXQgbGFiLmlzX3Zpc2libGUoKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBjb250aW51ZQogICAgICAgICAgICAgICAgICAgICAgICAgICAgIyDDlm5jZSB5YWvEsW4gcGFyZW50IGdydWJ1bmRha2kgaW5wdXRsYXIuCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBwYXJlbnQgPSBsYWIKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGZvciBfIGluIHJhbmdlKDUpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIHBhcmVudCA9IHBhcmVudC5sb2NhdG9yKCJ4cGF0aD0uLiIpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgYXdhaXQgcGFyZW50LmNvdW50KCkgPT0gMDoKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgYnJlYWsKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBpbnB1dHMgPSBwYXJlbnQubG9jYXRvcignaW5wdXRbdHlwZT0ibnVtYmVyIl0sIGlucHV0W3R5cGU9InRleHQiXSwgaW5wdXQ6bm90KFt0eXBlXSknKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGlmIGF3YWl0IGlucHV0cy5jb3VudCgpID49IDE6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGZpcnN0ID0gaW5wdXRzLmZpcnN0CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGlmIGF3YWl0IGZpcnN0LmlzX3Zpc2libGUoKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IGZpcnN0LmNsaWNrKCkKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IGZpcnN0LmZpbGwoc3RyKHZhbHVlKSkKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IGZpcnN0LnByZXNzKCJUYWIiKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgcmV0dXJuIFRydWUKICAgICAgICAgICAgICAgICAgICAgICAgICAgICMgTGFiZWxkZW4gc29ucmFraSBpbGsgaW5wdXQuCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBjYW5kaWRhdGUgPSBsYWIubG9jYXRvcigieHBhdGg9Zm9sbG93aW5nOjppbnB1dFsxXSIpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiBhd2FpdCBjYW5kaWRhdGUuY291bnQoKSBhbmQgYXdhaXQgY2FuZGlkYXRlLmZpcnN0LmlzX3Zpc2libGUoKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBmaXJzdCA9IGNhbmRpZGF0ZS5maXJzdAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IGZpcnN0LmNsaWNrKCkKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBhd2FpdCBmaXJzdC5maWxsKHN0cih2YWx1ZSkpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgYXdhaXQgZmlyc3QucHJlc3MoIlRhYiIpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgcmV0dXJuIFRydWUKICAgICAgICAgICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbjoKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHBhc3MKICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICAgICAgcGFzcwogICAgICAgICAgICAgICAgcmV0dXJuIEZhbHNlCgogICAgICAgICAgICBpZiBpc19tYXN0ZXI6CiAgICAgICAgICAgICAgICAjIE1hc3RlcidkYSBMZWFndWUgc2XDp2ltaW5lIGhpw6cgZG9rdW5tYS4KICAgICAgICAgICAgICAgICMgU3RhdHMgYnV0b251IGJhesSxIHPDvHLDvG1sZXJkZSB6YXRlbiBrxLFybcSxesSxL2FrdGlmIGdlbGl5b3I7CiAgICAgICAgICAgICAgICAjIGFrdGlmc2UgdGVrcmFyIHTEsWtsYW1hayBQdlAneWUgZ2XDp2lyZWJpbGlyLiBCdSBuZWRlbmxlIMO2bmNlCiAgICAgICAgICAgICAgICAjIGdlcsOnZWsgVUkgZHVydW11bnUgdGVzcGl0IGVkaXlvciwgeWFsbsSxemNhIGFrdGlmIGRlxJ9pbHNlIHTEsWtsxLF5b3J1ei4KICAgICAgICAgICAgICAgIHN0YXRzX2RvbmUgPSBGYWxzZQogICAgICAgICAgICAgICAgc3RhdHMgPSBOb25lCiAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgbG9jID0gcGFnZS5nZXRfYnlfdGV4dCgiU3RhdHMiLCBleGFjdD1UcnVlKQogICAgICAgICAgICAgICAgICAgIGZvciBpIGluIHJhbmdlKGF3YWl0IGxvYy5jb3VudCgpKToKICAgICAgICAgICAgICAgICAgICAgICAgZWwgPSBsb2MubnRoKGkpCiAgICAgICAgICAgICAgICAgICAgICAgIGlmIGF3YWl0IGVsLmlzX3Zpc2libGUoKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHN0YXRzID0gZWwKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGJyZWFrCiAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uIGFzIGU6CiAgICAgICAgICAgICAgICAgICAgc2F5KCJERUJVRyIsICJTdGF0cyBrb250cm9sw7wgYXJhbmFtYWTEsTogIiArIHJlcHIoZSkpCgogICAgICAgICAgICAgICAgaWYgc3RhdHMgaXMgbm90IE5vbmU6CiAgICAgICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgICAgICBzdGF0ZSA9IGF3YWl0IHN0YXRzLmV2YWx1YXRlKCIiIgogICAgICAgICAgICAgICAgICAgICAgICBlID0+IHsKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGNvbnN0IGNoYWluPVtdOyBsZXQgeD1lOwogICAgICAgICAgICAgICAgICAgICAgICAgICAgZm9yKGxldCBpPTA7aTw1ICYmIHg7aSsrLHg9eC5wYXJlbnRFbGVtZW50KXsKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBjaGFpbi5wdXNoKHt0YWc6eC50YWdOYW1lLCBjbHM6eC5jbGFzc05hbWV8fCcnLCByb2xlOnguZ2V0QXR0cmlidXRlKCdyb2xlJyl8fCcnLAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBwcmVzc2VkOnguZ2V0QXR0cmlidXRlKCdhcmlhLXByZXNzZWQnKSwgc3RhdGU6eC5nZXRBdHRyaWJ1dGUoJ2RhdGEtc3RhdGUnKSwKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgc2VsZWN0ZWQ6eC5nZXRBdHRyaWJ1dGUoJ2FyaWEtc2VsZWN0ZWQnKX0pOwogICAgICAgICAgICAgICAgICAgICAgICAgICAgfQogICAgICAgICAgICAgICAgICAgICAgICAgICAgY29uc3QgYmc9Z2V0Q29tcHV0ZWRTdHlsZShlKS5iYWNrZ3JvdW5kQ29sb3I7CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBjb25zdCBwPWUucGFyZW50RWxlbWVudDsKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGNvbnN0IHNpYj1wID8gQXJyYXkuZnJvbShwLnF1ZXJ5U2VsZWN0b3JBbGwoJ2J1dHRvbixbcm9sZT1idXR0b25dJykpLm1hcCh4PT4oe3RleHQ6KHguaW5uZXJUZXh0fHwnJykudHJpbSgpLGJnOmdldENvbXB1dGVkU3R5bGUoeCkuYmFja2dyb3VuZENvbG9yLGNsczp4LmNsYXNzTmFtZXx8JycscHJlc3NlZDp4LmdldEF0dHJpYnV0ZSgnYXJpYS1wcmVzc2VkJyksc3RhdGU6eC5nZXRBdHRyaWJ1dGUoJ2RhdGEtc3RhdGUnKX0pKSA6IFtdOwogICAgICAgICAgICAgICAgICAgICAgICAgICAgcmV0dXJuIHtjaGFpbixiZyxzaWJsaW5nczpzaWJ9OwogICAgICAgICAgICAgICAgICAgICAgICB9CiAgICAgICAgICAgICAgICAgICAgICAgICIiIikKICAgICAgICAgICAgICAgICAgICAgICAgc2F5KCJERUJVRyIsICJTdGF0cyBVSSBkdXJ1bXU6ICIgKyBqc29uLmR1bXBzKHN0YXRlLCBlbnN1cmVfYXNjaWk9RmFsc2UpKQogICAgICAgICAgICAgICAgICAgICAgICBhY3RpdmUgPSBGYWxzZQogICAgICAgICAgICAgICAgICAgICAgICBmb3IgYyBpbiBzdGF0ZS5nZXQoImNoYWluIiwgW10pOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgc3RyKGMuZ2V0KCJwcmVzc2VkIiwgIiIpKS5sb3dlcigpID09ICJ0cnVlIiBvciBzdHIoYy5nZXQoInNlbGVjdGVkIiwgIiIpKS5sb3dlcigpID09ICJ0cnVlIiBvciBzdHIoYy5nZXQoInN0YXRlIiwgIiIpKS5sb3dlcigpIGluICgiYWN0aXZlIiwic2VsZWN0ZWQiLCJvbiIpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGFjdGl2ZSA9IFRydWUKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBicmVhawogICAgICAgICAgICAgICAgICAgICAgICAgICAgY2xzID0gc3RyKGMuZ2V0KCJjbHMiLCAiIikpLmxvd2VyKCkKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGlmIGFueShrIGluIGNscyBmb3IgayBpbiAoImFjdGl2ZSIsICJzZWxlY3RlZCIsICJpcy1hY3RpdmUiLCAiY2hlY2tlZCIsICJjdXJyZW50IikpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGFjdGl2ZSA9IFRydWUKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBicmVhawogICAgICAgICAgICAgICAgICAgICAgICAjIFBvZ29Db29yZGluYXRlcyBla3JhbsSxbmRhIFN0YXRzIGvEsXJtxLF6xLEsIFB2UCBrb3l1IGlzZQogICAgICAgICAgICAgICAgICAgICAgICAjIGV4cGxpY2l0IEFSSUEgeW9rc2EgZGEgU3RhdHMnxLEgYWt0aWYga2FidWwgZXQuCiAgICAgICAgICAgICAgICAgICAgICAgIGlmIG5vdCBhY3RpdmU6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBzaWJzPXN0YXRlLmdldCgic2libGluZ3MiLFtdKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgc3RhdHNfYmc9c3RhdGUuZ2V0KCJiZyIsIiIpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBwdnBfYmc9IiIKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGZvciBzIGluIHNpYnM6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgc3RyKHMuZ2V0KCJ0ZXh0IiwiIikpLnN0cmlwKCkubG93ZXIoKT09InB2cCI6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIHB2cF9iZz1zLmdldCgiYmciLCIiKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgc3RhdHNfYmcgYW5kIHB2cF9iZyBhbmQgc3RhdHNfYmcgIT0gcHZwX2JnOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGFjdGl2ZSA9IFRydWUKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBzYXkoIkRFQlVHIiwgIlN0YXRzL1B2UCBhcmthIHBsYW5sYXLEsSBmYXJrbMSxOyBTdGF0cyBha3RpZiBrYWJ1bCBlZGlsZGkuIikKCiAgICAgICAgICAgICAgICAgICAgICAgIGlmIGFjdGl2ZToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHN0YXRzX2RvbmUgPSBUcnVlCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBzYXkoIkNPTlRST0xTIiwgIlN0YXRzIHphdGVuIHNlw6dpbGk7IHRla3JhciB0xLFrbGFubWFkxLEuIikKICAgICAgICAgICAgICAgICAgICAgICAgZWxzZToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IHN0YXRzLmNsaWNrKHRpbWVvdXQ9MTUwMCkKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGF3YWl0IHBhZ2Uud2FpdF9mb3JfdGltZW91dCg0MDApCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBzdGF0c19kb25lID0gVHJ1ZQogICAgICAgICAgICAgICAgICAgICAgICAgICAgc2F5KCJDT05UUk9MUyIsICJTdGF0cyBzZcOnaWxkaS4iKQogICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZToKICAgICAgICAgICAgICAgICAgICAgICAgc2F5KCJERUJVRyIsICJTdGF0cyBkdXJ1bS9rbGlrIGhhdGFzxLE6ICIgKyByZXByKGUpKQoKICAgICAgICAgICAgICAgIGlmIG5vdCBzdGF0c19kb25lOgogICAgICAgICAgICAgICAgICAgIHNheSgiRVJST1IiLCAiTWFzdGVyL01lZ2EgacOnaW4gU3RhdHMga29udHJvbMO8IGJ1bHVuYW1hZMSxIHZleWEgc2XDp2lsZW1lZGkuIikKICAgICAgICAgICAgICAgICAgICByZXR1cm4gMTkKCiAgICAgICAgICAgICAgICBzYXkoIlNURVAiLCAiWzMvNV0gTGV2ZWwgTWluID0gMzQgeWF6xLFsxLF5b3LigKYiKQogICAgICAgICAgICAgICAgbGV2ZWxfZG9uZSA9IGF3YWl0IHNldF9taW5fZmllbGQoIkxldmVsIiwgMzQpCiAgICAgICAgICAgICAgICBzYXkoIkNPTlRST0xTIiwgIkxldmVsIE1pbj0zNCAiICsgKCJPSyIgaWYgbGV2ZWxfZG9uZSBlbHNlICJIQVRBIikpCgogICAgICAgICAgICAgICAgc2F5KCJTVEVQIiwgIls0LzVdIElWIE1pbiA9IDEwMCB5YXrEsWzEsXlvcuKApiIpCiAgICAgICAgICAgICAgICBpdl9kb25lID0gYXdhaXQgc2V0X21pbl9maWVsZCgiSVYiLCAxMDApCiAgICAgICAgICAgICAgICBzYXkoIkNPTlRST0xTIiwgIklWIE1pbj0xMDAgIiArICgiT0siIGlmIGl2X2RvbmUgZWxzZSAiSEFUQSIpKQoKICAgICAgICAgICAgICAgIGlmIG5vdCBsZXZlbF9kb25lIG9yIG5vdCBpdl9kb25lOgogICAgICAgICAgICAgICAgICAgIHNheSgiRVJST1IiLCAiTWFzdGVyL01lZ2EgTGVhZ3VlIGZpbHRyZSBhbGFubGFyxLFuZGFuIGJpcmkgYXlhcmxhbmFtYWTEsS4iKQogICAgICAgICAgICAgICAgICAgIHJldHVybiAxOQogICAgICAgICAgICAgICAgaWYgbm90IHN0YXRzX2RvbmUgb3Igbm90IGxldmVsX2RvbmUgb3Igbm90IGl2X2RvbmU6CiAgICAgICAgICAgICAgICAgICAgc2F5KCJFUlJPUiIsICJNYXN0ZXIvTWVnYSBMZWFndWUgZmlsdHJlIGFsYW5sYXLEsW5kYW4gYmlyaSBheWFybGFuYW1hZMSxLiIpCiAgICAgICAgICAgICAgICAgICAgcmV0dXJuIDE5CiAgICAgICAgICAgIGVsc2U6CiAgICAgICAgICAgICAgICAjIEdyZWF0IC8gVWx0cmEgbWV2Y3V0IGRhdnJhbsSxxZ86IFB2UCArIFRvcCByYW5raW5nIDEuCiAgICAgICAgICAgICAgICBwdnAgPSBOb25lCiAgICAgICAgICAgICAgICBwdnBfdGV4dCA9IHBhZ2UuZ2V0X2J5X3RleHQoIlB2UCIsIGV4YWN0PVRydWUpCiAgICAgICAgICAgICAgICBmb3IgaSBpbiByYW5nZShhd2FpdCBwdnBfdGV4dC5jb3VudCgpKToKICAgICAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgIGVsID0gcHZwX3RleHQubnRoKGkpCiAgICAgICAgICAgICAgICAgICAgICAgIGlmIGF3YWl0IGVsLmlzX3Zpc2libGUoKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIHB2cCA9IGVsCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBicmVhawogICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICAgICAgICAgIHBhc3MKCiAgICAgICAgICAgICAgICBpZiBwdnAgaXMgTm9uZToKICAgICAgICAgICAgICAgICAgICBzYXkoIkVSUk9SIiwgIlN0YXRzIGLDtmzDvG3DvG5kZWtpIFB2UCBzZcOnZW5lxJ9pIGJ1bHVuYW1hZMSxLiIpCiAgICAgICAgICAgICAgICAgICAgcmV0dXJuIDE0CgogICAgICAgICAgICAgICAgYXdhaXQgcHZwLmNsaWNrKCkKICAgICAgICAgICAgICAgIGF3YWl0IHBhZ2Uud2FpdF9mb3JfdGltZW91dCgzMDApCiAgICAgICAgICAgICAgICBzYXkoIkNPTlRST0xTIiwgIlN0YXRzID0gUHZQIHNlw6dpbGRpLiIpCgogICAgICAgICAgICAgICAgcmFuayA9IE5vbmUKICAgICAgICAgICAgICAgIGxhYmVsX3RleHQgPSBwYWdlLmdldF9ieV90ZXh0KCJUb3AgcmFua2luZyIsIGV4YWN0PVRydWUpCiAgICAgICAgICAgICAgICBmb3IgaSBpbiByYW5nZShhd2FpdCBsYWJlbF90ZXh0LmNvdW50KCkpOgogICAgICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICAgICAgbGFiZWwgPSBsYWJlbF90ZXh0Lm50aChpKQogICAgICAgICAgICAgICAgICAgICAgICBpZiBub3QgYXdhaXQgbGFiZWwuaXNfdmlzaWJsZSgpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgY29udGludWUKICAgICAgICAgICAgICAgICAgICAgICAgY2FuZGlkYXRlID0gbGFiZWwubG9jYXRvcigieHBhdGg9Zm9sbG93aW5nOjppbnB1dFsxXSIpCiAgICAgICAgICAgICAgICAgICAgICAgIGlmIGF3YWl0IGNhbmRpZGF0ZS5jb3VudCgpIGFuZCBhd2FpdCBjYW5kaWRhdGUuZmlyc3QuaXNfdmlzaWJsZSgpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgcmFuayA9IGNhbmRpZGF0ZS5maXJzdAogICAgICAgICAgICAgICAgICAgICAgICAgICAgYnJlYWsKICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgICAgICAgICAgICAgICAgICBwYXNzCgogICAgICAgICAgICAgICAgaWYgcmFuayBpcyBOb25lOgogICAgICAgICAgICAgICAgICAgIGlucHV0cyA9IHBhZ2UubG9jYXRvcignaW5wdXRbdHlwZT0ibnVtYmVyIl0sIGlucHV0W3R5cGU9InRleHQiXSwgaW5wdXQ6bm90KFt0eXBlXSknKQogICAgICAgICAgICAgICAgICAgIGZvciBpIGluIHJhbmdlKGF3YWl0IGlucHV0cy5jb3VudCgpKToKICAgICAgICAgICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgICAgICAgICAgZWwgPSBpbnB1dHMubnRoKGkpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiBhd2FpdCBlbC5pc192aXNpYmxlKCkgYW5kIGF3YWl0IGVsLmlucHV0X3ZhbHVlKCkgPT0gIjEiOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIHJhbmsgPSBlbAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGJyZWFrCiAgICAgICAgICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBwYXNzCgogICAgICAgICAgICAgICAgaWYgcmFuayBpcyBOb25lOgogICAgICAgICAgICAgICAgICAgIHNheSgiRVJST1IiLCAiVG9wIHJhbmtpbmcgYWxhbsSxIGJ1bHVuYW1hZMSxLiIpCiAgICAgICAgICAgICAgICAgICAgcmV0dXJuIDE1CgogICAgICAgICAgICAgICAgYXdhaXQgcmFuay5jbGljaygpCiAgICAgICAgICAgICAgICBhd2FpdCByYW5rLmZpbGwoIjEiKQogICAgICAgICAgICAgICAgYXdhaXQgcmFuay5wcmVzcygiVGFiIikKICAgICAgICAgICAgICAgIHNheSgiQ09OVFJPTFMiLCAiVG9wIHJhbmtpbmcgPSAxIHlhcMSxbGTEsS4iKQoKICAgICAgICAgICAgICAgICMgU8SxcmFsYW1hOiBJViAlIChoaWdoZXN0KQogICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgIHNvcnRfZG9uZSA9IEZhbHNlCiAgICAgICAgICAgICAgICAgICAgc29ydF9yZXN1bHQgPSBhd2FpdCBwYWdlLmV2YWx1YXRlKCIiIgogICAgICAgICAgICAgICAgICAgICgpID0+IHsKICAgICAgICAgICAgICAgICAgICAgICAgY29uc3Qgbm9ybSA9IHMgPT4gKHMgfHwgJycpLnRvU3RyaW5nKCkudHJpbSgpLnRvTG93ZXJDYXNlKCk7CiAgICAgICAgICAgICAgICAgICAgICAgIGZvciAoY29uc3QgcyBvZiBBcnJheS5mcm9tKGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJ3NlbGVjdCcpKSkgewogICAgICAgICAgICAgICAgICAgICAgICAgICAgY29uc3Qgb3B0cyA9IEFycmF5LmZyb20ocy5vcHRpb25zIHx8IFtdKTsKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGNvbnN0IG9wdCA9IG9wdHMuZmluZChvID0+IHsKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBjb25zdCB0ID0gbm9ybShvLnRleHQpOwogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGNvbnN0IHYgPSBub3JtKG8udmFsdWUpOwogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIHJldHVybiAodC5pbmNsdWRlcygnaXYgJScpICYmICh0LmluY2x1ZGVzKCdoaWdoZXN0JykgfHwgdC5pbmNsdWRlcygnaGlnaCcpKSkKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgfHwgKHYuaW5jbHVkZXMoJ2l2JykgJiYgKHYuaW5jbHVkZXMoJ2Rlc2MnKSkpOwogICAgICAgICAgICAgICAgICAgICAgICAgICAgfSk7CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpZiAob3B0KSB7CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgcy52YWx1ZSA9IG9wdC52YWx1ZTsKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBzLmRpc3BhdGNoRXZlbnQobmV3IEV2ZW50KCdpbnB1dCcsIHtidWJibGVzOnRydWV9KSk7CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgcy5kaXNwYXRjaEV2ZW50KG5ldyBFdmVudCgnY2hhbmdlJywge2J1YmJsZXM6dHJ1ZX0pKTsKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICByZXR1cm4gdHJ1ZTsKICAgICAgICAgICAgICAgICAgICAgICAgICAgIH0KICAgICAgICAgICAgICAgICAgICAgICAgfQogICAgICAgICAgICAgICAgICAgICAgICByZXR1cm4gZmFsc2U7CiAgICAgICAgICAgICAgICAgICAgfQogICAgICAgICAgICAgICAgICAgICIiIikKICAgICAgICAgICAgICAgICAgICBzb3J0X2RvbmUgPSBib29sKHNvcnRfcmVzdWx0KQogICAgICAgICAgICAgICAgICAgIGlmIG5vdCBzb3J0X2RvbmU6CiAgICAgICAgICAgICAgICAgICAgICAgIGZvciBsYWJlbCBpbiAoIklWICUgKGhpZ2hlc3QpIiwgIklWICUgKEhpZ2hlc3QpIiwgIklWICUgaGlnaGVzdCIpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgYXdhaXQgY2xpY2tfZXhhY3RfdmlzaWJsZShsYWJlbCk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgc29ydF9kb25lID0gVHJ1ZQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGJyZWFrCiAgICAgICAgICAgICAgICAgICAgc2F5KCJDT05UUk9MUyIsICJTxLFyYWxhbWEgPSBJViAlIChoaWdoZXN0KSIgaWYgc29ydF9kb25lIGVsc2UgIklWIHPEsXJhbGFtYXPEsSBidWx1bmFtYWTEsTsgYXJhbWEgZGV2YW0gZWRpeW9yLiIpCiAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uIGFzIGU6CiAgICAgICAgICAgICAgICAgICAgc2F5KCJJTkZPIiwgIklWIHPEsXJhbGFtYSBrb250cm9sw7wgYXRsYW5kxLE6ICIgKyByZXByKGUpKQoKICAgICAgICAgICAgaWYgaXNfbWFzdGVyOgogICAgICAgICAgICAgICAgc2F5KCJTVEVQIiwgIls1LzVdIE1hc3Rlci9NZWdhIGZpbHRyZWxlcmkgdGFtYW1sYW5kxLE7IFNlYXJjaCBidXRvbnUgYXJhbsSxeW9y4oCmIikKICAgICAgICAgICAgZWxzZToKICAgICAgICAgICAgICAgIHNheSgiU1RFUCIsICJbMy81XSBHcmVhdC9VbHRyYSBmaWx0cmVsZXJpIHRhbWFtbGFuZMSxOyBTZWFyY2ggYnV0b251IGFyYW7EsXlvcuKApiIpCgogICAgICAgICAgICAjIEVuIGFsdHRha2kga8Sxcm3EsXrEsSBTZWFyY2ggYnV0b251LgogICAgICAgICAgICBzZWFyY2ggPSBOb25lCiAgICAgICAgICAgIGJ1dHRvbnMgPSBwYWdlLmxvY2F0b3IoImJ1dHRvbiIpCiAgICAgICAgICAgIGZvciBpIGluIHJhbmdlKGF3YWl0IGJ1dHRvbnMuY291bnQoKSk6CiAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgZWwgPSBidXR0b25zLm50aChpKQogICAgICAgICAgICAgICAgICAgIGlmIGF3YWl0IGVsLmlzX3Zpc2libGUoKToKICAgICAgICAgICAgICAgICAgICAgICAgdCA9IChhd2FpdCBlbC5pbm5lcl90ZXh0KHRpbWVvdXQ9MzAwKSBvciAiIikuc3RyaXAoKQogICAgICAgICAgICAgICAgICAgICAgICBpZiB0Lmxvd2VyKCkgPT0gInNlYXJjaCI6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBzZWFyY2ggPSBlbAogICAgICAgICAgICAgICAgICAgICAgICAgICAgYnJlYWsKICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICAgICAgcGFzcwoKICAgICAgICAgICAgaWYgc2VhcmNoIGlzIE5vbmU6CiAgICAgICAgICAgICAgICBzYXkoIkVSUk9SIiwgIlNlYXJjaCBidXRvbnUgYnVsdW5hbWFkxLEuIikKICAgICAgICAgICAgICAgIHJldHVybiAxNwoKICAgICAgICAgICAgc2F5KCJTRUFSQ0giLCAiU2VhcmNoIGJ1dG9udW5hIGJhc8SxbMSxeW9y4oCmIikKICAgICAgICAgICAgYXdhaXQgc2VhcmNoLmNsaWNrKCkKICAgICAgICAgICAgYXdhaXQgcGFnZS53YWl0X2Zvcl90aW1lb3V0KDUwMDApCgogICAgICAgICAgICBib2R5X3RleHQgPSBhd2FpdCBwYWdlLmxvY2F0b3IoImJvZHkiKS5pbm5lcl90ZXh0KCkKICAgICAgICAgICAgbG93ID0gYm9keV90ZXh0Lmxvd2VyKCkKCiAgICAgICAgICAgICMgU29udcOnIHlva3NhIGHDp8SxayBiaXIgc29udcOnIG1lc2FqxLEgZ8O2bmRlci4KICAgICAgICAgICAgc3Ryb25nX3Jlc3VsdCA9IEZhbHNlCiAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgIHN0cm9uZ19yZXN1bHQgPSBhd2FpdCBwYWdlLmdldF9ieV9yb2xlKAogICAgICAgICAgICAgICAgICAgICJidXR0b24iLCBuYW1lPXJlLmNvbXBpbGUociJHZXQgY29vcmRzIiwgcmUuSSkKICAgICAgICAgICAgICAgICkuY291bnQoKSA+IDAKICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbjoKICAgICAgICAgICAgICAgIHBhc3MKICAgICAgICAgICAgaWYgbm90IHN0cm9uZ19yZXN1bHQ6CiAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgc3Ryb25nX3Jlc3VsdCA9IGF3YWl0IHBhZ2UuZ2V0X2J5X3RleHQoCiAgICAgICAgICAgICAgICAgICAgICAgICJHZXQgY29vcmRzIiwgZXhhY3Q9VHJ1ZQogICAgICAgICAgICAgICAgICAgICkuY291bnQoKSA+IDAKICAgICAgICAgICAgICAgIGV4Y2VwdCBFeGNlcHRpb246CiAgICAgICAgICAgICAgICAgICAgcGFzcwoKICAgICAgICAgICAgbm9fcmVzdWx0X21hcmtlcnMgPSAoCiAgICAgICAgICAgICAgICAic2hvd2luZyAwIHNwYXducyIsCiAgICAgICAgICAgICAgICAibm8gc3Bhd25zIiwKICAgICAgICAgICAgICAgICJubyBwb2tlbW9uIGZvdW5kIiwKICAgICAgICAgICAgICAgICJubyBwb2vDqW1vbiBmb3VuZCIsCiAgICAgICAgICAgICAgICAibm90aGluZyBmb3VuZCIKICAgICAgICAgICAgKQogICAgICAgICAgICBpZiAobm90IHN0cm9uZ19yZXN1bHQpIGFuZCBhbnkoeCBpbiBsb3cgZm9yIHggaW4gbm9fcmVzdWx0X21hcmtlcnMpOgogICAgICAgICAgICAgICAgc2F5KCJOT1JFU1VMVCIsICJQb2vDqW1vbiBidWx1bmFtYWTEsS4gTMO8dGZlbiBkYWhhIHNvbnJhIHRla3JhciBkZW5leWluLiIpCiAgICAgICAgICAgICAgICByZXR1cm4gMAoKICAgICAgICAgICAgIyAiR2V0IGNvb3JkcyIgYnV0b251bnUgYnVsLiDEsGxrIHNvbnXDpyBrYXJ0xLFuZGFraSBidXRvbnUgdGVyY2loIGV0LgogICAgICAgICAgICBnZXRfYnV0dG9ucyA9IHBhZ2UuZ2V0X2J5X3JvbGUoImJ1dHRvbiIsIG5hbWU9cmUuY29tcGlsZShyIkdldCBjb29yZHMiLCByZS5JKSkKICAgICAgICAgICAgaWYgYXdhaXQgZ2V0X2J1dHRvbnMuY291bnQoKSA9PSAwOgogICAgICAgICAgICAgICAgZ2V0X2J1dHRvbnMgPSBwYWdlLmdldF9ieV90ZXh0KCJHZXQgY29vcmRzIiwgZXhhY3Q9VHJ1ZSkKCiAgICAgICAgICAgIGlmIGF3YWl0IGdldF9idXR0b25zLmNvdW50KCkgPT0gMDoKICAgICAgICAgICAgICAgICMgU29udcOnIHNhecSxc8SxIGfDtnLDvG7DvHlvcnNhIGFtYSBidXRvbiBidWx1bm11eW9yc2EgeWluZSBzb251w6cKICAgICAgICAgICAgICAgICMgYmlsZ2lzaW5pIHZlcm1leWUgw6dhbMSxxZ87IGtvb3JkaW5hdCBvbG1hZGFuIGthcnTEsSBkw7ZuZMO8cm1lLgogICAgICAgICAgICAgICAgaWYgcmUuc2VhcmNoKHIic2hvd2luZ1xzK1xkK1xzK3NwYXducz8iLCBsb3cpOgogICAgICAgICAgICAgICAgICAgIHNheSgiTk9SRVNVTFQiLCAiUG9rw6ltb24gYnVsdW5kdSBhbmNhayBrb29yZGluYXQgZMO8xJ9tZXNpIGFsxLFuYW1hZMSxLiBMw7x0ZmVuIGRhaGEgc29ucmEgdGVrcmFyIGRlbmV5aW4uIikKICAgICAgICAgICAgICAgICAgICByZXR1cm4gMAogICAgICAgICAgICAgICAgc2F5KCJOT1JFU1VMVCIsICJQb2vDqW1vbiBidWx1bmFtYWTEsS4gTMO8dGZlbiBkYWhhIHNvbnJhIHRla3JhciBkZW5leWluLiIpCiAgICAgICAgICAgICAgICByZXR1cm4gMAoKICAgICAgICAgICAgIyDEsGxrIDQgc29udWN1IGFsLgogICAgICAgICAgICBidXR0b25fY291bnQgPSBtaW4oYXdhaXQgZ2V0X2J1dHRvbnMuY291bnQoKSwgNCkKICAgICAgICAgICAgcmVzdWx0cyA9IFtdCgogICAgICAgICAgICBhc3luYyBkZWYgZXh0cmFjdF9jYXJkX3RleHQoYnRuKToKICAgICAgICAgICAgICAgIGJlc3QgPSAiIgogICAgICAgICAgICAgICAgY2FyZCA9IGJ0bgogICAgICAgICAgICAgICAgZm9yIF8gaW4gcmFuZ2UoNyk6CiAgICAgICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgICAgICBjYXJkID0gY2FyZC5sb2NhdG9yKCJ4cGF0aD0uLiIpCiAgICAgICAgICAgICAgICAgICAgICAgIHR4dCA9IChhd2FpdCBjYXJkLmlubmVyX3RleHQodGltZW91dD0zMDApIG9yICIiKS5zdHJpcCgpCiAgICAgICAgICAgICAgICAgICAgICAgIGlmIGxlbih0eHQpID4gbGVuKGJlc3QpOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgYmVzdCA9IHR4dAogICAgICAgICAgICAgICAgICAgICAgICBpZiAoIkNQIiBpbiB0eHQgYW5kICJJViIgaW4gdHh0KSBvciAoIkRFU1BBV04iIGluIHR4dC51cHBlcigpKToKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGJyZWFrCiAgICAgICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbjoKICAgICAgICAgICAgICAgICAgICAgICAgYnJlYWsKICAgICAgICAgICAgICAgIHJldHVybiBiZXN0CgogICAgICAgICAgICBhc3luYyBkZWYgZ2V0X2Nvb3JkcyhidG4pOgogICAgICAgICAgICAgICAgY29vcmRzID0gIiIKICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICBhd2FpdCBwYWdlLmV2YWx1YXRlKCIoKSA9PiB7IHdpbmRvdy5fX3NpbmthX2xhc3RfY29vcmRzID0gJyc7IH0iKQogICAgICAgICAgICAgICAgICAgIGF3YWl0IGJ0bi5jbGljaygpCiAgICAgICAgICAgICAgICAgICAgYXdhaXQgcGFnZS53YWl0X2Zvcl90aW1lb3V0KDcwMCkKICAgICAgICAgICAgICAgICAgICBjb29yZHMgPSBhd2FpdCBwYWdlLmV2YWx1YXRlKCIoKSA9PiB3aW5kb3cuX19zaW5rYV9sYXN0X2Nvb3JkcyB8fCAnJyIpCiAgICAgICAgICAgICAgICAgICAgaWYgbm90IHJlLnNlYXJjaChyJy0/XGQrXC5cZCtccypbLCBdXHMqLT9cZCtcLlxkKycsIGNvb3JkcyBvciAiIik6CiAgICAgICAgICAgICAgICAgICAgICAgIGNvb3JkcyA9IGF3YWl0IHBhZ2UuZXZhbHVhdGUoIigpID0+IG5hdmlnYXRvci5jbGlwYm9hcmQucmVhZFRleHQoKS5jYXRjaCgoKT0+ICcnKSIpCiAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgICAgICAgICAgICAgIHBhc3MKCiAgICAgICAgICAgICAgICBtbSA9IHJlLnNlYXJjaChyJygtP1xkezEsM31cLlxkezMsfSlccypbLCBdXHMqKC0/XGR7MSwzfVwuXGR7Myx9KScsIGNvb3JkcyBvciAiIikKICAgICAgICAgICAgICAgIGlmIG1tOgogICAgICAgICAgICAgICAgICAgIHJldHVybiBtbS5ncm91cCgxKSArICIsICIgKyBtbS5ncm91cCgyKQoKICAgICAgICAgICAgICAgIGJvZHkgPSBhd2FpdCBwYWdlLmxvY2F0b3IoImJvZHkiKS5pbm5lcl90ZXh0KCkKICAgICAgICAgICAgICAgIG1tID0gcmUuc2VhcmNoKHInKC0/XGR7MSwzfVwuXGR7Myx9KVxzKlssIF1ccyooLT9cZHsxLDN9XC5cZHszLH0pJywgYm9keSkKICAgICAgICAgICAgICAgIHJldHVybiAobW0uZ3JvdXAoMSkgKyAiLCAiICsgbW0uZ3JvdXAoMikpIGlmIG1tIGVsc2UgIiIKCiAgICAgICAgICAgIGZvciBpZHggaW4gcmFuZ2UoYnV0dG9uX2NvdW50KToKICAgICAgICAgICAgICAgIHRyeToKICAgICAgICAgICAgICAgICAgICBidG4gPSBnZXRfYnV0dG9ucy5udGgoaWR4KQogICAgICAgICAgICAgICAgICAgIGlmIG5vdCBhd2FpdCBidG4uaXNfdmlzaWJsZSgpOgogICAgICAgICAgICAgICAgICAgICAgICBjb250aW51ZQoKICAgICAgICAgICAgICAgICAgICBzYXkoIlJFU1VMVCIsIGYie2lkeCsxfS4gUG9rw6ltb24gc29udWN1IG9rdW51eW9y4oCmIikKICAgICAgICAgICAgICAgICAgICBjYXJkX3RleHQgPSBhd2FpdCBleHRyYWN0X2NhcmRfdGV4dChidG4pCiAgICAgICAgICAgICAgICAgICAgY29vcmRzID0gYXdhaXQgZ2V0X2Nvb3JkcyhidG4pCiAgICAgICAgICAgICAgICAgICAgaWYgbm90IGNvb3JkczoKICAgICAgICAgICAgICAgICAgICAgICAgY29udGludWUKCiAgICAgICAgICAgICAgICAgICAgZGVmIGZpZWxkKHBhdHRlcm4pOgogICAgICAgICAgICAgICAgICAgICAgICBtbSA9IHJlLnNlYXJjaChwYXR0ZXJuLCBjYXJkX3RleHQsIHJlLkkgfCByZS5NKQogICAgICAgICAgICAgICAgICAgICAgICByZXR1cm4gbW0uZ3JvdXAoMSkuc3RyaXAoKSBpZiBtbSBlbHNlICIiCgogICAgICAgICAgICAgICAgICAgICMgU29udcOnIGthcnTEsW5kYWtpIEdFUsOHRUsgUG9rw6ltb24gYWTEsW7EsSDDp8Sxa2FyLgogICAgICAgICAgICAgICAgICAgICMgRGFoYSDDtm5jZSBidXJhZGEgYXJhbWEga3V0dXN1bmEgeWF6xLFsYW4gYG5hbWVgIGt1bGxhbsSxbMSxeW9yZHU7CiAgICAgICAgICAgICAgICAgICAgIyBidSB5w7x6ZGVuIMO2cm4uIFRpbmthdGluayBhcmFuxLFya2VuIGZhcmtsxLEgYmlyIHNvbnXDpyBkYQogICAgICAgICAgICAgICAgICAgICMgVGlua2F0aW5rIGRpeWUgZ8O2c3RlcmlsZWJpbGl5b3JkdS4KICAgICAgICAgICAgICAgICAgICBhY3R1YWxfbmFtZSA9ICIiCiAgICAgICAgICAgICAgICAgICAgdHJ5OgogICAgICAgICAgICAgICAgICAgICAgICAjIMOWbmNlbGlrOiBrYXJ0IGnDp2luZGVraSBQb2vDqWRleCBiYcSfbGFudMSxc8SxbsSxbiBtZXRuaS4KICAgICAgICAgICAgICAgICAgICAgICAgbGlua3MgPSBjYXJkLmxvY2F0b3IoImFbaHJlZio9Jy9wb2tlbW9uLWdvL3Bva2VkZXgvJ10iKQogICAgICAgICAgICAgICAgICAgICAgICBmb3IgbGkgaW4gcmFuZ2UoYXdhaXQgbGlua3MuY291bnQoKSk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICB0eHQgPSAoYXdhaXQgbGlua3MubnRoKGxpKS5pbm5lcl90ZXh0KHRpbWVvdXQ9MjUwKSBvciAiIikuc3RyaXAoKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgdHh0IGFuZCBsZW4odHh0KSA8IDgwOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGFjdHVhbF9uYW1lID0gdHh0CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgYnJlYWsKICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgICAgICAgICAgICAgICAgICBwYXNzCiAgICAgICAgICAgICAgICAgICAgaWYgbm90IGFjdHVhbF9uYW1lOgogICAgICAgICAgICAgICAgICAgICAgICB0cnk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICBpbWdzID0gY2FyZC5sb2NhdG9yKCJpbWciKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgZm9yIGlpIGluIHJhbmdlKGF3YWl0IGltZ3MuY291bnQoKSk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgaW1nX2VsID0gaW1ncy5udGgoaWkpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgYWx0ID0gKGF3YWl0IGltZ19lbC5nZXRfYXR0cmlidXRlKCJhbHQiKSBvciAiIikuc3RyaXAoKQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIHRpdGxlID0gKGF3YWl0IGltZ19lbC5nZXRfYXR0cmlidXRlKCJ0aXRsZSIpIG9yICIiKS5zdHJpcCgpCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgY2FuZGlkYXRlID0gYWx0IG9yIHRpdGxlCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgaWYgY2FuZGlkYXRlIGFuZCBsZW4oY2FuZGlkYXRlKSA8IDgwOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBhY3R1YWxfbmFtZSA9IGNhbmRpZGF0ZQogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICBicmVhawogICAgICAgICAgICAgICAgICAgICAgICBleGNlcHQgRXhjZXB0aW9uOgogICAgICAgICAgICAgICAgICAgICAgICAgICAgcGFzcwogICAgICAgICAgICAgICAgICAgIGlmIG5vdCBhY3R1YWxfbmFtZToKICAgICAgICAgICAgICAgICAgICAgICAgIyBLYXJ0IGJhxZ9sxLHEn8SxbmRhIGFyYW5hbiBpc2ltIGdlw6dpeW9yc2Egb251IGt1bGxhbjsKICAgICAgICAgICAgICAgICAgICAgICAgIyBha3NpIGhhbGRlIG1ldGluZGVuIGlsayBtYWt1bCBzYXTEsXLEsSBzZcOnLgogICAgICAgICAgICAgICAgICAgICAgICBsaW5lcyA9IFt4LnN0cmlwKCkgZm9yIHggaW4gY2FyZF90ZXh0LnNwbGl0bGluZXMoKSBpZiB4LnN0cmlwKCldCiAgICAgICAgICAgICAgICAgICAgICAgIGZvciBsbiBpbiBsaW5lczoKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGlmIHJlLnNlYXJjaChyJ1xiKENQfElWfExWTHxMT0NBVElPTnxERVNQQVdOfEdFVCBDT09SRFMpXGInLCBsbiwgcmUuSSk6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgY29udGludWUKICAgICAgICAgICAgICAgICAgICAgICAgICAgIGlmIGxlbihsbikgPD0gNjAgYW5kIG5vdCByZS5mdWxsbWF0Y2gocidbMC05Oi4sJSgpXC0gXSsnLCBsbik6CiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgYWN0dWFsX25hbWUgPSBsbgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIGJyZWFrCiAgICAgICAgICAgICAgICAgICAgYWN0dWFsX25hbWUgPSBhY3R1YWxfbmFtZSBvciBuYW1lCgogICAgICAgICAgICAgICAgICAgIHJlc3VsdHMuYXBwZW5kKHsKICAgICAgICAgICAgICAgICAgICAgICAgInBva2Vtb24iOiBhY3R1YWxfbmFtZSwKICAgICAgICAgICAgICAgICAgICAgICAgInNlYXJjaGVkX3Bva2Vtb24iOiBuYW1lLAogICAgICAgICAgICAgICAgICAgICAgICAiY3AiOiBmaWVsZChyJ1xiQ1BccysoWzAtOV0rKScpLAogICAgICAgICAgICAgICAgICAgICAgICAibGV2ZWwiOiBmaWVsZChyJ1xiTFZMXHMrKFswLTldKyg/OlwuWzAtOV0rKT8pJyksCiAgICAgICAgICAgICAgICAgICAgICAgICJpdiI6IGZpZWxkKHInXGJJVlxzKyhbMC05XStcJVxzKlwoW14pXStcKSknKSwKICAgICAgICAgICAgICAgICAgICAgICAgImxvY2F0aW9uIjogZmllbGQocidcYkxPQ0FUSU9OXHMrKC4rPykoPzpcbnwkKScpLAogICAgICAgICAgICAgICAgICAgICAgICAiZGVzcGF3biI6IGZpZWxkKHInXGJERVNQQVdOXHMrKFswLTk6XSspJyksCiAgICAgICAgICAgICAgICAgICAgICAgICJjb29yZHMiOiBjb29yZHMsCiAgICAgICAgICAgICAgICAgICAgfSkKICAgICAgICAgICAgICAgICAgICBzYXkoIlJFU1VMVCIsIGYie2xlbihyZXN1bHRzKX0gc29udcOnIGhhesSxcmxhbmTEsS4iKQogICAgICAgICAgICAgICAgICAgIGlmIGxlbihyZXN1bHRzKSA+PSA0OgogICAgICAgICAgICAgICAgICAgICAgICBicmVhawogICAgICAgICAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBleDoKICAgICAgICAgICAgICAgICAgICBzYXkoIklORk8iLCBmIntpZHgrMX0uIHNvbnXDpyBva3VuYW1hZMSxOiB7cmVwcihleCl9IikKCiAgICAgICAgICAgIGlmIG5vdCByZXN1bHRzOgogICAgICAgICAgICAgICAgc2F5KCJFUlJPUiIsICJQb2vDqW1vbiBidWx1bmR1IGZha2F0IHNvbnXDpyBrb29yZGluYXRsYXLEsSBhbMSxbmFtYWTEsS4iKQogICAgICAgICAgICAgICAgcmV0dXJuIDE5CgogICAgICAgICAgICBzYXkoIlJFU1VMVCIsIGpzb24uZHVtcHMoeyJyZXN1bHRzIjogcmVzdWx0c1s6NF19LCBlbnN1cmVfYXNjaWk9RmFsc2UpKQogICAgICAgICAgICByZXR1cm4gMAoKICAgIGV4Y2VwdCBFeGNlcHRpb24gYXMgZToKICAgICAgICBzYXkoIkVSUk9SIiwgdHlwZShlKS5fX25hbWVfXyArICI6ICIgKyBzdHIoZSkpCiAgICAgICAgdHJhY2ViYWNrLnByaW50X2V4YygpCiAgICAgICAgcmV0dXJuIDIwCgpzeXMuZXhpdChhc3luY2lvLnJ1bihtYWluKCkpKQo=').decode('utf-8')

def _open_pogo_status_window(pokemon_name, league_name):
    """Arama başlar başlamaz açılan temiz durum penceresi."""
    global _POGO_STATUS_WINDOW, _POGO_STATUS_TEXT, _POGO_STATUS_PROGRESS
    try:
        root_obj = globals().get("_SINKA_ROOT") or getattr(globals().get("tk"), "_default_root", None)
        if root_obj is None:
            return

        old = globals().get("_POGO_STATUS_WINDOW")
        if old is not None:
            try:
                old.destroy()
            except Exception:
                pass

        win = tk.Toplevel(root_obj)
        _POGO_STATUS_WINDOW = win
        win.title(f"SinKA PvP - {league_name} - Pokémon Aranıyor")
        win.geometry("720x620")
        win.minsize(680, 560)
        win.transient(root_obj)

        outer = tk.Frame(win, padx=20, pady=16)
        outer.pack(fill="both", expand=True)

        tk.Label(
            outer, text="🔎  POKÉMON ARANIYOR",
            font=("Arial", 19, "bold")
        ).pack(anchor="w")

        tk.Label(
            outer, text=f"{league_name}   •   {pokemon_name}",
            font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", pady=(4, 12))

        _POGO_STATUS_PROGRESS = ttk.Progressbar(
            outer, mode="indeterminate", length=650
        )
        _POGO_STATUS_PROGRESS.pack(fill="x", pady=(0, 12))
        try:
            _POGO_STATUS_PROGRESS.start(10)
        except Exception:
            pass

        frame = tk.LabelFrame(
            outer, text="  İşlem Durumu  ",
            font=("Segoe UI", 10, "bold"),
            padx=8, pady=8
        )
        frame.pack(fill="both", expand=True)

        _POGO_STATUS_TEXT = tk.Text(
            frame, height=20, width=78,
            font=("Consolas", 10),
            wrap="word", state="disabled",
            relief="flat", padx=8, pady=8
        )
        _POGO_STATUS_TEXT.pack(fill="both", expand=True)

        tk.Label(
            outer,
            text="Tüm işlemler arka planda yürütülüyor.",
            font=("Arial", 9)
        ).pack(anchor="w", pady=(8, 0))

        win.protocol("WM_DELETE_WINDOW", win.destroy)
        win.update_idletasks()
        win.lift()
        win.focus_force()
    except Exception:
        pass


def _pogo_status_add_line(kind, detail):
    try:
        box = globals().get("_POGO_STATUS_TEXT")
        if box is None:
            return

        icons = {
            "START": "▶", "INSTALL": "⚙", "BROWSER": "🌐",
            "PAGE": "🌐", "CONTROLS": "⚙", "SEARCH": "🔎",
            "RESULTS": "✓", "RESULT": "✓", "NORESULT": "⚠",
            "ERROR": "✖", "PLAYWRIGHT_MISSING": "✖", "INFO": "•"
        }
        prefix = icons.get(kind, "•")

        box.config(state="normal")
        box.insert("end", f"{prefix} {detail or kind}\n")
        box.see("end")
        box.config(state="disabled")
    except Exception:
        pass


def _pogo_result_image_loader(name):
    """Sonuç ekranı için PokémonDB/PokeAPI görselini mevcut SinKa önbelleğinden al."""
    try:
        root_obj = globals().get("_SINKA_ROOT") or getattr(globals().get("tk"), "_default_root", None)
        app = getattr(root_obj, "_sinka_app", None) if root_obj is not None else None
        if app is not None and hasattr(app, "download_image_bytes"):
            return app.download_image_bytes(str(name).strip())
    except Exception:
        pass
    return None

_POGO_RESULT_IMAGE_LOADER = _pogo_result_image_loader
_POGO_RESULT_PHOTOS = []

def _pogo_widget_exists(widget):
    """Tk widget hâlâ mevcut mu? Yoksa otomatik sonuç ekranı yeni pencere açabilsin."""
    try:
        return bool(widget is not None and widget.winfo_exists())
    except Exception:
        return False

def _modernize_global_popup(win):
    """Modern SinKA teması için global arama pencereleri."""
    bg, navy, card, fg, accent = "#F4F7FB", "#16233D", "#FFFFFF", "#182230", "#2F6BDE"
    try:
        win.configure(bg=bg)
    except Exception:
        pass
    def walk(w):
        try:
            cls=w.winfo_class()
            if cls=="Frame":
                w.configure(bg=bg)
            elif cls=="Label":
                w.configure(bg=card, fg=fg)
            elif cls=="Button":
                w.configure(bg=navy, fg="white", activebackground=accent,
                             activeforeground="white", relief="flat", bd=0,
                             font=("Segoe UI",10,"bold"), cursor="hand2")
            elif cls=="Entry":
                w.configure(bg="white", fg=fg, insertbackground=fg,
                             relief="solid", bd=1, font=("Segoe UI",10))
            elif cls=="Text":
                w.configure(bg="white", fg=fg, insertbackground=fg)
            elif cls=="Canvas":
                w.configure(bg=card)
            for c in w.winfo_children():
                walk(c)
        except Exception:
            pass
    walk(win)
    try:
        win.update_idletasks()
        win.lift()
    except Exception:
        pass

def _show_pogo_results(payload):
    """Durum penceresini büyük ve okunaklı sonuç kartlarına dönüştürür."""
    global _POGO_STATUS_WINDOW, _POGO_STATUS_PROGRESS

    try:
        win = globals().get("_POGO_STATUS_WINDOW")
        root_obj = globals().get("_SINKA_ROOT") or getattr(globals().get("tk"), "_default_root", None)
        if root_obj is None:
            return

        # Otomatik aramada status penceresi özellikle açılmıyor (silent=True).
        # Bu durumda eski status penceresine bağlı kalmak sonuç ekranının hiç açılmamasına
        # neden oluyordu. Pencere yoksa doğrudan yeni bir Toplevel oluştur.
        if win is None or not _pogo_widget_exists(win):
            win = tk.Toplevel(root_obj)
            globals()["_POGO_STATUS_WINDOW"] = win
        else:
            try:
                win.deiconify()
            except Exception:
                pass

        data = json.loads(payload) if isinstance(payload, str) else payload
        results = data.get("results", [])

        if not results:
            _show_pogo_noresult("Pokémon bulunamadı. Lütfen daha sonra tekrar deneyin.")
            return

        try:
            _POGO_STATUS_PROGRESS.stop()
            _POGO_STATUS_PROGRESS.pack_forget()
        except Exception:
            pass

        for widget in win.winfo_children():
            try:
                widget.destroy()
            except Exception:
                pass

        win.title("SinKA PvP - Pokémon Bulundu")
        win.geometry("1180x820")
        win.minsize(980, 680)

        outer = tk.Frame(win, padx=20, pady=16)
        outer.pack(fill="both", expand=True)

        tk.Label(
            outer, text="🎯  POKÉMON BULUNDU",
            font=("Segoe UI", 20, "bold")
        ).pack(anchor="w")
        active_league = str(data.get("league") or "Great League")
        league_cp = {
            "Great League": "1500 CP MAX",
            "Ultra League": "2500 CP MAX",
            "Master League": "SINIRSIZ CP MAX",
            "Mega": "SINIRSIZ CP MAX",
            "Gigamax": "SINIRSIZ CP MAX",
        }.get(active_league, "")

        tk.Label(
            outer,
            text=f"{active_league}  •  {league_cp}  •  {len(results)} sonuç bulundu  •  IV % (highest)",
            font=("Segoe UI", 11, "bold")
        ).pack(anchor="w", pady=(3, 12))

        if data.get("auto"):
            tk.Label(outer, text="ℹ️ Otomatik arama sonucu — aşağıdaki kartlarda koordinatlar gösterilir.", font=("Segoe UI", 10), anchor="w").pack(anchor="w", pady=(0, 8))

        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        content = tk.Frame(canvas)
        content.grid_columnconfigure(0, weight=1, uniform="resultcol")
        content.grid_columnconfigure(1, weight=1, uniform="resultcol")

        content.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        content_window = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        def _result_canvas_resize(event):
            try:
                canvas.itemconfigure(content_window, width=event.width)
            except Exception:
                pass

        canvas.bind("<Configure>", _result_canvas_resize)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for idx, item in enumerate(results, 1):
            card = tk.LabelFrame(
                content,
                text=f"  #{idx}  ",
                font=("Segoe UI", 12, "bold"),
                padx=14, pady=12
            )
            card.grid(
                row=(idx - 1) // 2,
                column=(idx - 1) % 2,
                sticky="nsew",
                padx=6,
                pady=6
            )

            pname = str(item.get("pokemon") or data.get("pokemon") or "Pokémon")

            # Sonuç kartında bulunan gerçek Pokémonun görselini göster.
            card_top = tk.Frame(card)
            card_top.pack(fill="x", anchor="w")
            image_label = tk.Label(card_top, width=100, height=100)
            image_label.pack(side="left", padx=(0, 14))
            name_box = tk.Frame(card_top)
            name_box.pack(side="left", fill="x", expand=True)

            tk.Label(
                name_box, text=pname,
                font=("Segoe UI", 17, "bold")
            ).pack(anchor="w")

            searched = str(item.get("searched_pokemon") or data.get("pokemon") or "").strip()
            if searched and searched.lower() != pname.lower():
                tk.Label(
                    name_box, text="Aranan: " + searched,
                    font=("Segoe UI", 10)
                ).pack(anchor="w", pady=(3, 0))

            try:
                image_data = None
                if callable(globals().get("_POGO_RESULT_IMAGE_LOADER")):
                    image_data = globals()["_POGO_RESULT_IMAGE_LOADER"](pname)
                if image_data and PIL_OK:
                    img = Image.open(BytesIO(image_data)).convert("RGBA")
                    img.thumbnail((100, 100), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    image_label.configure(image=photo, text="")
                    image_label.image = photo
                    if "_POGO_RESULT_PHOTOS" not in globals():
                        globals()["_POGO_RESULT_PHOTOS"] = []
                    globals()["_POGO_RESULT_PHOTOS"].append(photo)
            except Exception:
                pass

            info = []
            if item.get("cp"):
                info.append("CP " + str(item["cp"]))
            if item.get("level"):
                info.append("Level " + str(item["level"]))
            if item.get("iv"):
                info.append("IV " + str(item["iv"]))

            tk.Label(
                card,
                text="   •   ".join(info) if info else "Bilgi alınamadı",
                font=("Segoe UI", 11, "bold")
            ).pack(anchor="w", pady=(3, 4))

            if item.get("location"):
                tk.Label(
                    card,
                    text="📍  " + str(item["location"]),
                    font=("Segoe UI", 10),
                    wraplength=500,
                    justify="left"
                ).pack(anchor="w", pady=(0, 3))

            if item.get("despawn"):
                tk.Label(
                    card,
                    text="⏱  Despawn: " + str(item["despawn"]),
                    font=("Segoe UI", 10)
                ).pack(anchor="w", pady=(0, 7))

            tk.Label(
                card, text="KOORDİNAT",
                font=("Segoe UI", 9, "bold")
            ).pack(anchor="w")

            coord = str(item.get("coords") or "").strip()
            coord_var = tk.StringVar(
                value=coord if coord else "Koordinat alınamadı"
            )

            coord_entry = tk.Entry(
                card,
                textvariable=coord_var,
                font=("Consolas", 14, "bold"),
                justify="center",
                relief="solid",
                bd=1
            )
            coord_entry.pack(fill="x", pady=(3, 8), ipady=5)

            def copy_coord(c=coord, entry=coord_entry):
                if not c:
                    messagebox.showwarning(
                        "SinKA PvP",
                        "Bu sonuç için koordinat alınamadı."
                    )
                    return
                try:
                    root_obj.clipboard_clear()
                    root_obj.clipboard_append(c)
                    root_obj.update()
                    entry.selection_range(0, "end")
                    messagebox.showinfo(
                        "SinKA PvP",
                        "Koordinat panoya kopyalandı."
                    )
                except Exception as ex:
                    messagebox.showerror(
                        "SinKA PvP",
                        "Koordinat kopyalanamadı:\n" + str(ex)
                    )

            tk.Button(
                card,
                text="📋  KOORDİNATI KOPYALA",
                command=copy_coord,
                font=("Segoe UI", 11, "bold"),
                padx=12, pady=7
            ).pack(fill="x")

        app = getattr(root_obj, "_sinka_app", None)
        def add_found_pokemon():
            if app is None:
                messagebox.showwarning("SinKA PvP", "Ana uygulama bağlantısı bulunamadı.")
                return
            added = 0
            names = []
            league_name = str(data.get("league") or getattr(app, "loaded_league", None) or getattr(app, "league_var", tk.StringVar(value="Great League")).get())
            for item in results:
                if not isinstance(item, dict):
                    continue
                pname = str(item.get("pokemon") or item.get("searched_pokemon") or "").strip()
                if not pname:
                    continue
                # Gerçek bulunan Pokémonu Mevcut olarak işaretle.
                try:
                    app._set_owned(pname, True, league_name)
                    added += 1
                    names.append(pname)
                except Exception:
                    pass
            try:
                app.save_preferences()
                if getattr(app, "loaded_league", None) == league_name:
                    app.refresh()
            except Exception:
                pass
            messagebox.showinfo("SinKA PvP", f"{added} bulunan Pokémon Mevcut olarak eklendi.")

        # Sonuç işlem butonları her zaman listenin EN ALTINDA.
        action_bar = tk.Frame(outer, bg="#F4F7FB", height=58)
        action_bar.pack(fill="x", pady=(10, 0))
        action_bar.pack_propagate(False)

        tk.Button(
            action_bar,
            text="➕  BULUNAN POKÉMONLARI EKLE",
            command=add_found_pokemon,
            font=("Segoe UI", 10, "bold"),
            bg="#16233D", fg="white",
            activebackground="#263A62", activeforeground="white",
            relief="flat", bd=0,
            padx=18, pady=8
        ).pack(side="left", padx=(0, 8), pady=7)

        tk.Button(
            action_bar,
            text="KAPAT",
            command=win.destroy,
            font=("Segoe UI", 10, "bold"),
            bg="#E7ECF3", fg="#16233D",
            activebackground="#D9E0EA",
            relief="flat", bd=0,
            padx=22, pady=8
        ).pack(side="right", pady=7)

        _modernize_global_popup(win)
        win.update_idletasks()
        win.lift()

    except Exception as ex:
        messagebox.showerror(
            "SinKA PvP",
            "Sonuç ekranı oluşturulamadı:\n" + str(ex)
        )


def _show_pogo_error(title, code, detail, return_code=0):
    """Hataları ayrı hata ekranına taşımadan mevcut İşlem Durumu penceresine ekler."""
    try:
        win = globals().get("_POGO_STATUS_WINDOW")
        if win is None:
            # Durum penceresi yoksa son çare olarak normal hata kutusu kullanılır.
            messagebox.showerror(
                "SinKA PvP - Hata",
                title + "\n\nHATA KODU: " + str(code) +
                "\n\n" + str(detail) +
                "\n\nProcess kodu: " + str(return_code)
            )
            return

        try:
            win.title("SinKA PvP - Pokémon Aranıyor / Hata")
        except Exception:
            pass

        try:
            _POGO_STATUS_PROGRESS.stop()
        except Exception:
            pass

        # Hata, daha önce görülen tüm işlem satırlarının ALTINA yazılır.
        _pogo_status_add_line("ERROR", str(title))
        _pogo_status_add_line("ERROR", "HATA KODU: " + str(code))
        _pogo_status_add_line("ERROR", str(detail))
        _pogo_status_add_line("ERROR", "Process kodu: " + str(return_code))

        try:
            win.lift()
        except Exception:
            pass
    except Exception:
        pass


def _show_pogo_noresult(message):
    try:
        win = globals().get("_POGO_STATUS_WINDOW")
        if win is None:
            messagebox.showinfo("SinKA PvP", message)
            return

        try:
            _POGO_STATUS_PROGRESS.stop()
            _POGO_STATUS_PROGRESS.pack_forget()
        except Exception:
            pass

        for widget in win.winfo_children():
            try:
                widget.destroy()
            except Exception:
                pass

        win.title("SinKA PvP - Sonuç Yok")
        outer = tk.Frame(win, padx=28, pady=30)
        outer.pack(fill="both", expand=True)

        tk.Label(
            outer, text="⚠️  POKÉMON BULUNAMADI",
            font=("Arial", 19, "bold")
        ).pack(pady=(30, 12))

        tk.Label(
            outer,
            text="Lütfen daha sonra tekrar deneyin.",
            font=("Arial", 12)
        ).pack(pady=(0, 25))

        tk.Button(
            outer, text="Kapat",
            command=win.destroy,
            font=("Segoe UI", 11, "bold"),
            padx=20, pady=7
        ).pack()

        win.update_idletasks()
        win.lift()
    except Exception:
        pass


def _pogo_status_message(message):
    """Otomasyonun tüm durumlarını yalnızca durum/sonuç penceresinde gösterir."""
    try:
        kind, _, detail = str(message).partition("|")
        root_obj = globals().get("_SINKA_ROOT") or getattr(globals().get("tk"), "_default_root", None)

        def apply():
            try:
                if kind in ("RESULTS", "RESULT"):
                    try:
                        data = json.loads(detail) if detail else {}
                        if data.get("results"):
                            _pogo_status_add_line("RESULTS", "Sonuçlar bulundu. Sonuç ekranı hazırlanıyor…")
                            _show_pogo_results(detail)
                        else:
                            _show_pogo_noresult("Pokémon bulunamadı. Lütfen daha sonra tekrar deneyin.")
                    except Exception as ex:
                        _show_pogo_noresult("Sonuç okunamadı: " + str(ex))
                    return

                _pogo_status_add_line(kind, detail or str(message))

                if kind == "NORESULT":
                    _show_pogo_noresult(detail or "Pokémon bulunamadı. Lütfen daha sonra tekrar deneyin.")
                elif kind in ("ERROR", "PLAYWRIGHT_MISSING"):
                    win = globals().get("_POGO_STATUS_WINDOW")
                    if win is not None:
                        win.title("SinKA PvP - Otomasyon Hatası")
                        try:
                            _POGO_STATUS_PROGRESS.stop()
                        except Exception:
                            pass

            except Exception:
                pass

        if root_obj is not None:
            try:
                root_obj.after(0, apply)
            except Exception:
                apply()
    except Exception:
        pass


# ============================================================
# Otomatik PogoCoordinates araması
# Great -> 30 dk -> Ultra -> 30 dk -> Master -> 30 dk -> Great...
# ============================================================
_POGO_AUTO_INTERVAL_MS = 30 * 60 * 1000
_POGO_AUTO_LEAGUES = ["Great League", "Ultra League", "Master League"]
_POGO_AUTO_RUNNING = False
_POGO_AUTO_INDEX = 0
_POGO_AUTO_AFTER_ID = None
_POGO_AUTO_RESULTS = []

def _show_pogo_auto_notification(league, results):
    """Otomatik aramada sonuç geldiğinde sonuçları açıp eklemeyi sağlayan bildirim."""
    try:
        root_obj = globals().get("_SINKA_ROOT")
        if root_obj is None:
            return
        results = list(results or [])
        win = tk.Toplevel(root_obj)
        win.title("SinKA PvP - Pokémon Bulundu")
        win.resizable(False, False)
        win.attributes("-topmost", True)
        outer = tk.Frame(win, padx=22, pady=16)
        outer.pack(fill="both", expand=True)
        tk.Label(outer, text="🎯 Pokémon bulundu!", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        tk.Label(outer, text=f"{league}: {len(results)} sonuç bulundu.", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(5, 2))
        tk.Label(outer, text="Koordinatları görmek ve bulunan Pokémonları Mevcut olarak eklemek için sonuç ekranını açın.", font=("Segoe UI", 9), wraplength=500, justify="left").pack(anchor="w", pady=(0, 10))
        btns = tk.Frame(outer)
        btns.pack(fill="x")
        def show_results():
            try:
                win.destroy()
            except Exception:
                pass
            _show_pogo_results({"results": results, "pokemon": "Otomatik Arama", "league": league, "auto": True})
        tk.Button(btns, text="📍 Sonuçları ve Koordinatları Göster", command=show_results, font=("Segoe UI", 10, "bold"), padx=12, pady=7).pack(side="left", padx=(0, 7))
        tk.Button(btns, text="Kapat", command=win.destroy, padx=14, pady=7).pack(side="right")
        win.update_idletasks()
        try:
            x = root_obj.winfo_x() + root_obj.winfo_width() - win.winfo_width() - 25
            y = root_obj.winfo_y() + 55
            win.geometry(f"+{max(0,x)}+{max(0,y)}")
        except Exception:
            pass
        win.lift()
        win.focus_force()
        self._modernize_popup(win)
    except Exception:
        pass


class App:
    def _search_pokemon_row(self, pokemon_name):
        """Search exactly the Pokémon represented by this table row."""
        league = self._get_active_search_league()
        self.active_search_league = league
        open_pogo_coordinates_search(str(pokemon_name).strip(), league)


    def _open_multiple_pokemon_coordinates(self):
        """Sadece Mevcut sütununda ☐ olan Pokémonları aynı aramada ara."""
        names = []

        try:
            columns = list(self.tree["columns"])
        except Exception:
            columns = []

        for iid in self.tree.get_children(""):
            try:
                # Mevcut sütunundaki işaret esas alınır:
                # ☑ = mevcut, aranmaz
                # ☐ = mevcut değil, aranır
                current_mark = str(self.tree.set(iid, "current")).strip()
                if current_mark == "☑":
                    continue

                vals = self.tree.item(iid, "values")
                name = ""
                if "pokemon" in columns:
                    name = str(vals[columns.index("pokemon")]).strip()
                elif vals:
                    name = str(vals[0]).strip()

                if name:
                    names.append(name)
            except Exception:
                continue

        # Aynı Pokémon birden fazla görünüyorsa yalnızca bir kez ekle.
        names = list(dict.fromkeys(names))

        if not names:
            try:
                messagebox.showinfo(
                    "SinKA PvP",
                    "Mevcut (☑) olarak işaretlenmeyen Pokémon bulunamadı.\n\n"
                    "Çoklu arama yalnızca Mevcut sütununda ☐ olan Pokémonları arar."
                )
            except Exception:
                pass
            return

        league = self._get_active_search_league()
        self.active_search_league = league

        open_pogo_coordinates_search_multiple(names, league)

    def _auto_pogo_names_for_league(self, league, requested_count=None, show_shadows=None, include_legendary=None, include_mythical=None):
        """Çoklu aramanın kullandığı mantıkla ilgili ligdeki ☐ Pokémonları hazırlar.
        Aktif ligde mevcut görünür tabloyu, diğer liglerde güncel PvPoke sıralamasını kullanır.
        Ağır indirme gerekiyorsa çağıran kod bunu worker thread'de yapabilir; Tk değişkenleri
        önceden ana thread'den alınır."""
        if requested_count is None:
            requested_count = self.get_requested_count()
        if show_shadows is None:
            show_shadows = bool(self.shadow_var.get())
        if include_legendary is None:
            include_legendary = bool(self.legendary_var.get())
        if include_mythical is None:
            include_mythical = bool(self.mythical_var.get())

        owned = dict(self.owned_by_league.get(league, {}))

        if getattr(self, "loaded_league", None) == league and self.tree.get_children():
            names = []
            columns = list(self.tree["columns"])
            pidx = columns.index("pokemon") if "pokemon" in columns else 3
            for iid in self.tree.get_children(""):
                try:
                    if str(self.tree.set(iid, "current")).strip() == "☑":
                        continue
                    vals = self.tree.item(iid, "values")
                    name = str(vals[pidx]).strip() if vals and pidx < len(vals) else ""
                    if name and name not in names:
                        names.append(name)
                except Exception:
                    pass
            return names

        cp = LEAGUES.get(league)
        if cp is None:
            return []
        data = download_json(RANKING_URL.format(cp=cp)) or []
        candidates = []
        for p in data:
            if not isinstance(p, dict) or is_gigantamax(p):
                continue
            if not show_shadows and is_shadow(p):
                continue
            if include_legendary and not self.has_pokemon_tag(p, "legendary"):
                continue
            if include_mythical and not self.has_pokemon_tag(p, "mythical"):
                continue
            candidates.append(p)

        names = []
        for p in candidates:
            name = str(p.get("speciesName") or p.get("speciesId") or "").strip()
            if not name or bool(owned.get(name, False)):
                continue
            if name not in names:
                names.append(name)
            if len(names) >= requested_count:
                break
        return names

    def toggle_pogo_auto_search(self):
        """Otomatik Great -> Ultra -> Master aramasını kullanıcıdan alınan dakikayla başlat/durdur."""
        if getattr(self, "pogo_auto_running", False):
            self.pogo_auto_running = False
            try:
                if self.pogo_auto_after_id is not None:
                    self.root.after_cancel(self.pogo_auto_after_id)
            except Exception:
                pass
            self.pogo_auto_after_id = None
            try:
                self.pogo_auto_button.config(text="⏱ Otomatik Ara")
            except Exception:
                pass
            self.set_status("Otomatik Pokémon araması durduruldu.")
            return

        # Kullanıcı otomatik aramayı seçtiğinde aralık mutlaka sorulur.
        minutes = simpledialog.askinteger(
            "Otomatik Arama Aralığı",
            "Kaç dakikada bir arama yapayım?\n\nÖrnek: 30",
            parent=self.root,
            initialvalue=getattr(self, "pogo_auto_interval_minutes", 30),
            minvalue=1, maxvalue=1440
        )
        if minutes is None:
            self.set_status("Otomatik arama başlatılmadı.")
            return

        self.pogo_auto_interval_minutes = int(minutes)
        self.pogo_auto_interval_ms = self.pogo_auto_interval_minutes * 60 * 1000
        self.pogo_auto_running = True
        self.pogo_auto_index = 0
        self.pogo_auto_results = []
        try:
            self.pogo_auto_button.config(text="⏹ Otomatik Aramayı Durdur")
        except Exception:
            pass
        self.set_status(f"Otomatik arama başladı: Great League aranıyor. Aralık: {minutes} dakika.")
        self._pogo_auto_run_current()

    def _pogo_auto_run_current(self):
        if not getattr(self, "pogo_auto_running", False):
            return
        if self.pogo_auto_index >= len(_POGO_AUTO_LEAGUES):
            self.pogo_auto_index = 0

        league = _POGO_AUTO_LEAGUES[self.pogo_auto_index]
        self.set_status(f"⏱ Otomatik arama: {league} için çoklu Pokémon aranıyor...")

        # Tkinter değişkenlerini/Treeview'ı worker thread'de okumamak için değerleri şimdi al.
        try:
            requested_count = self.get_requested_count()
            show_shadows = bool(self.shadow_var.get())
            include_legendary = bool(self.legendary_var.get())
            include_mythical = bool(self.mythical_var.get())
            active_now = (getattr(self, "loaded_league", None) == league and bool(self.tree.get_children()))
        except Exception:
            requested_count = 36
            show_shadows = False
            include_legendary = False
            include_mythical = False
            active_now = False

        # Aktif ligde isimleri doğrudan ana thread'deki mevcut tablodan al.
        if active_now:
            try:
                names = self._auto_pogo_names_for_league(
                    league, requested_count, show_shadows, include_legendary, include_mythical
                )
            except Exception as ex:
                self.set_status("Otomatik arama hazırlama hatası: " + str(ex))
                self._pogo_auto_schedule_next()
                return
            self._pogo_auto_launch_names(league, names)
            return

        def prepare():
            try:
                names = self._auto_pogo_names_for_league(
                    league, requested_count, show_shadows, include_legendary, include_mythical
                )
            except Exception as ex:
                names = []
                self.root.after(0, lambda e=ex: self.set_status("Otomatik arama hazırlama hatası: " + str(e)))

            self.root.after(0, lambda n=names: self._pogo_auto_launch_names(league, n))

        threading.Thread(target=prepare, daemon=True).start()

    def _pogo_auto_launch_names(self, league, names):
        if not getattr(self, "pogo_auto_running", False):
            return
        if not names:
            self.set_status(f"{league}: aranacak ☐ Pokémon kalmadı. 30 dakika sonra sonraki lige geçilecek.")
            self._pogo_auto_schedule_next()
            return
        self.set_status(f"⏱ {league}: {len(names)} Pokémon PogoCoordinates'te aranıyor...")
        open_pogo_coordinates_search_multiple(
            names,
            league,
            silent=True,
            on_complete=self._pogo_auto_search_complete
        )

    def _pogo_auto_search_complete(self, payload):
        if not getattr(self, "pogo_auto_running", False):
            return
        league = str(payload.get("league") or _POGO_AUTO_LEAGUES[self.pogo_auto_index])
        results = payload.get("results") or []
        for item in results:
            if isinstance(item, dict):
                item.setdefault("league", league)
        if results:
            self.pogo_auto_results.extend(results)
            self.root.after(0, lambda l=league, r=list(results): _show_pogo_auto_notification(l, r))
            self.root.after(0, lambda l=league, n=len(results): self.set_status(f"🎯 {l}: {n} Pokémon bulundu. Otomatik arama devam edecek."))
        else:
            self.root.after(0, lambda l=league: self.set_status(f"{l}: sonuç bulunamadı. Sonraki otomatik arama 30 dakika sonra."))
        self.root.after(0, self._pogo_auto_schedule_next)

    def _pogo_auto_schedule_next(self):
        if not getattr(self, "pogo_auto_running", False):
            return
        self.pogo_auto_index = (self.pogo_auto_index + 1) % len(_POGO_AUTO_LEAGUES)
        try:
            if self.pogo_auto_after_id is not None:
                self.root.after_cancel(self.pogo_auto_after_id)
        except Exception:
            pass
        next_league = _POGO_AUTO_LEAGUES[self.pogo_auto_index]
        self.set_status(f"⏱ Otomatik arama aktif. Sonraki: {next_league} — {self.pogo_auto_interval_minutes} dakika sonra.")
        self.pogo_auto_after_id = self.root.after(getattr(self, "pogo_auto_interval_ms", _POGO_AUTO_INTERVAL_MS), self._pogo_auto_run_current)

    def _open_selected_pokemon_coordinates(self):
        """Seçili Pokémon varsa PogoCoordinates aramasını başlatır."""
        name = ""
        league = self._get_active_search_league()
        try:
            sel = self.tree.selection()
            if sel:
                vals = self.tree.item(sel[0], "values")
                cols = list(self.tree["columns"])
                if "pokemon" in cols:
                    idx = cols.index("pokemon")
                    if vals and idx < len(vals):
                        name = str(vals[idx]).strip()
        except Exception:
            pass

        # Satır seçilmiş olsa bile Pokémon seçilmemiş olabilir. Rank gibi değerleri
        # yanlışlıkla Pokémon adı kabul etme.
        if not name or name.isdigit() or name.lower() in {"rank", "pokemon", "pokémon"}:
            # Modal messagebox ana pencereye grab verdiği için Windows'ta bip sesi
            # sonrası uygulama kilitlenmiş gibi görünebiliyordu. Seçim yokken
            # yalnızca normal (grab almayan) bir bilgi penceresi açıyoruz.
            try:
                info = tk.Toplevel(self.root)
                info.title("SinKA PvP - Pokémon Seçimi")
                info.transient(self.root)
                info.resizable(False, False)
                info.protocol("WM_DELETE_WINDOW", info.destroy)
                frame = ttk.Frame(info, padding=18)
                frame.pack(fill="both", expand=True)
                ttk.Label(frame, text="Herhangi bir Pokémon seçili değil.",
                          font=("Segoe UI", 11, "bold")).pack(pady=(0, 10))
                ttk.Label(frame,
                          text="Lütfen listeden bir Pokémon seçin ve ardından tekrar\n"
                               "'Seçili Pokémon'u Ara' butonuna basın.",
                          justify="center").pack(pady=(0, 12))
                ttk.Button(frame, text="Tamam", command=info.destroy).pack()
                info.update_idletasks()
                try:
                    x = self.root.winfo_rootx() + (self.root.winfo_width() - info.winfo_width()) // 2
                    y = self.root.winfo_rooty() + (self.root.winfo_height() - info.winfo_height()) // 2
                    info.geometry(f"+{max(0, x)}+{max(0, y)}")
                except Exception:
                    pass
                info.lift()
                try:
                    info.focus_force()
                except Exception:
                    pass
            except Exception:
                pass
            return

        league = self._get_active_search_league()
        self.active_search_league = league
        try:
            _pogo_status_message(
                "START|" + name + " için " + league +
                " araması başlatılıyor…"
            )
        except Exception:
            pass
        open_pogo_coordinates_search(name, league)

    def __init__(self, root):
        self.root = root
        globals()["_SINKA_ROOT"] = root
        try:
            root._sinka_app = self
        except Exception:
            pass
        self.root.title("SinKA PvP - Türkçe - V134")
        self.root.geometry("1280x760")

        # Pokémon Coordinates - hızlı PvP Rank 1 araması
        try:
            self.top_action_row = ttk.Frame(self.root)
            self.top_action_row.pack(fill="x", padx=12, pady=(4, 2))

            self.pogo_button_box = ttk.LabelFrame(self.top_action_row, text="Pokémon Arama")
            self.pogo_button_box.pack(side="left", anchor="w", padx=(0, 8), pady=0)
            self.pogo_button_bar = tk.Frame(self.pogo_button_box)
            self.pogo_button_bar.pack(side="left", padx=5, pady=3)

            self.pogo_search_button = tk.Button(
                self.pogo_button_bar,
                text="🔎 Seçili Pokémon'u Ara",
                command=lambda: self._open_selected_pokemon_coordinates(),
                font=("Segoe UI", 11, "bold")
            )
            self.pogo_search_button.pack(side="left", padx=(0, 6))

            self.pogo_multi_search_button = tk.Button(
                self.pogo_button_bar,
                text="🔎 Çoklu Pokémon Ara",
                command=lambda: self._open_multiple_pokemon_coordinates(),
                font=("Segoe UI", 11, "bold")
            )
            self.pogo_multi_search_button.pack(side="left", padx=(0, 6))

            self.pogo_auto_button = tk.Button(
                self.pogo_button_bar,
                text="⏱ Otomatik Ara",
                command=lambda: self.toggle_pogo_auto_search(),
                font=("Segoe UI", 11, "bold")
            )
            self.pogo_auto_button.pack(side="left")

            # Pokémon Arama ile Araçlar arasında, Mevcut Pokémonların
            # Durum bilgilerinin tamamlanıp tamamlanmadığını gösterir.
            # Üst orta bölüm: Durum + Event bekleyen + Yeniden yakalanacak
            # bilgilerini tek kutu içinde gösterir.
            self.status_summary_frame = tk.Frame(
                self.top_action_row,
                relief="groove",
                bd=1,
                padx=8,
                pady=2
            )
            self.status_summary_frame.pack(side="left", padx=10, pady=2, fill="x", expand=True)

            self.status_summary_label = tk.Label(
                self.status_summary_frame,
                text="⚠ Durum belirtilmedi",
                font=("Segoe UI", 11, "bold"),
                anchor="center"
            )
            self.status_summary_label.pack(side="left", padx=5)

            self.event_waiting_label = tk.Label(
                self.status_summary_frame,
                text="",
                font=("Segoe UI", 10, "bold"),
                anchor="center"
            )
            self.event_waiting_label.pack(side="left", padx=5)

            self.recatch_label = tk.Label(
                self.status_summary_frame,
                text="",
                font=("Segoe UI", 10, "bold"),
                anchor="center"
            )
            self.recatch_label.pack(side="left", padx=5)

            # Arka planda event takibi için küçük durum alanı. Eşleşme bulunduğunda
            # aynı merkez kutu içinde büyük kırmızı uyarıya dönüşür.
            self.event_watch_label = tk.Label(
                self.status_summary_frame,
                text="🔎 Event takip ediliyor",
                font=("Segoe UI", 9, "bold"),
                anchor="center"
            )
            self.event_watch_label.pack(side="left", padx=(8, 5))

            self.event_alert_frame = tk.Frame(
                self.status_summary_frame,
                bd=0,
                relief="flat"
            )
            self.event_alert_frame.pack_forget()
            self.event_alert_image_label = None
            self.event_alert_text_label = None

            global _POGO_STATUS_LABEL
            _POGO_STATUS_LABEL = tk.Label(
                self.root,
                text="Hazır — Pokémon seçip 'Seçili Pokémon'u Ara' butonuna basın.",
                anchor="w",
                font=("Segoe UI", 10)
            )
            _POGO_STATUS_LABEL.pack(fill="x", padx=10, pady=(0, 4))
        except Exception:
            self.pogo_search_button = None
            self.pogo_multi_search_button = None
            self.pogo_auto_button = None
        self.root.minsize(1050, 650)

        self.ranking_data = []
        self.pokemon = {}
        self.translate_move = lambda x: x
        self.loaded_league = None
        self.active_search_league = "Great League"
        self.busy = False
        self.generation = 0
        self.image_cache = {}
        self.image_bytes = {}
        self.image_files = {}
        self.type_icon_bytes = {}
        self.type_icon_cache = {}
        self.type_icon_photos = {}
        # Hareket tipi çözümünü bir kez yap; her redraw sırasında Game Master taraması yapma.
        self._move_type_key_cache = {}
        self._main_move_types_cache = {}
        self._move_type_cache_loading = False
        self._move_type_cache_ready = False
        self.type_keys = {}
        self.evolution_names = {}
        self.evolution_photo_cache = {}
        self.evolution_photo_refs = []
        self.photo_refs = []
        # Otomatik Great/Ultra/Master PogoCoordinates araması durumu.
        self.pogo_auto_running = False
        self.pogo_auto_index = 0
        self.pogo_auto_after_id = None
        self.pogo_auto_results = []
        self.pogo_auto_interval_ms = 30 * 60 * 1000
        self.pogo_auto_interval_minutes = 30
        # Kullanıcı ayarları
        self.auto_save_data_var = tk.BooleanVar(value=True)
        self.team_source_var = tk.StringVar(value="Mevcut Pokémonlar")
        # Mevcut işaretleri lig bazında bağımsız tut.
        self.owned_by_league = {league: {} for league in LEAGUES}
        self.owned_status_by_league = {league: {} for league in LEAGUES}
        self.owned = {}  # Eski sürümlerle uyumluluk / aktif ligin görünümü
        self._owned_summary_buttons = {}
        self._owned_summary_photos = []
        # PvP'ye hazır Pokémonlar: Mevcut + Durum(+) olanları lig sıralamasını
        # bozmadan üst başlıkta gösterir.
        self._ready_pokemon_photos = {}
        self._ready_pokemon_photo_refs = []
        self._ready_pokemon_loading = set()
        self._ready_pokemon_names = []
        appdata = Path(os.environ.get('APPDATA', str(Path.home()))) / 'SinKA PvP'
        self.preferences_path = appdata / 'preferences.json'
        # Her lig kendi filtre ayarlarını bağımsız saklar.
        self._applying_league_filters = False
        self.league_filter_settings = {league: {
            'count': 36, 'shadow': False, 'legendary': False, 'mythical': False,
            'xl': False, 'special': False, 'original_rank': False
        } for league in LEAGUES}

        self.build_ui()
        # Filtre kayıt bağlamı: lig değişirken eski ligin durumunu yeni lige yazmayı önler.
        self._filter_context_league = self.league_var.get()
        self.load_preferences()
        self._filter_context_league = self.league_var.get()
        self._event_watch_started = False
        self._event_watch_busy = False
        self._event_watch_token = 0
        self._event_watch_image_photo = None
        self._event_watch_last_match_key = None
        self._event_watch_current_match = None
        self._event_alert_countdown_label = None
        self._manual_pokemon_matches = []
        self._manual_pokemon_selected = None
        self._manual_pokemon_active = None
        self._manual_pokemon_original_ranking_data = None
        self.root.after(1200, self._start_event_status_watcher)
        # Lig değişimi ayrı yönetilir; league_var trace ile save_preferences çağırmak
        # eski ligin filtrelerini yeni lige yazıyordu. Sadece gerçek filtre değişikliklerini izle.
        for preference_var in (self.count_var, self.shadow_var, self.legendary_var, self.mythical_var, self.xl_var, self.special_var, self.original_rank_var):
            preference_var.trace_add('write', lambda *_: (not getattr(self, '_applying_league_filters', False)) and self.save_preferences())
        threading.Thread(target=self.load_initial, daemon=True).start()

    def open_url(self, url):
        webbrowser.open(url)

    def load_preferences(self):
        try:
            with open(self.preferences_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            league = data.get('league')
            if league in LEAGUES:
                self.league_var.set(league)

            # Yeni format: her lig kendi Mevcut listesini tutar.
            saved_by_league = data.get('owned_by_league')
            if isinstance(saved_by_league, dict):
                for league_name in LEAGUES:
                    values = saved_by_league.get(league_name, {})
                    if isinstance(values, dict):
                        self.owned_by_league[league_name] = {
                            str(name): bool(value)
                            for name, value in values.items() if value
                        }

            # Durum bilgileri de lig bazında saklanır. Geçersiz/eski değerleri temizle.
            saved_status = data.get('owned_status_by_league')
            if isinstance(saved_status, dict):
                for league_name in LEAGUES:
                    values = saved_status.get(league_name, {})
                    if isinstance(values, dict):
                        self.owned_status_by_league[league_name] = {
                            str(name): ('X' if str(value) in ('✕', 'x', 'X') else ('−' if str(value) == '-' else str(value)))
                            for name, value in values.items()
                            if str(value) in ('✕', 'X', 'x', '+', '−', '-')
                        }

            # Eski sürümde tek bir 'owned' listesi vardı. Hangi lige ait olduğu
            # bilinmediği için, kaydedilmiş aktif lige bir kez aktarılır.
            legacy_owned = data.get('owned') or {}
            if not saved_by_league and isinstance(legacy_owned, dict):
                active_league = self.league_var.get()
                self.owned_by_league[active_league] = {
                    str(name): bool(value)
                    for name, value in legacy_owned.items() if value
                }

            self.owned = dict(self.owned_by_league.get(self.league_var.get(), {}))

            # Yeni format: filtreler de Mevcut listeleri gibi lig bazında bağımsızdır.
            saved_filters = data.get('league_filter_settings')
            if isinstance(saved_filters, dict):
                for league_name in LEAGUES:
                    vals = saved_filters.get(league_name, {})
                    if isinstance(vals, dict):
                        defaults = self.league_filter_settings[league_name]
                        try:
                            defaults['count'] = max(1, min(100, int(vals.get('count', defaults['count']))))
                        except Exception:
                            pass
                        for key in ('shadow', 'legendary', 'mythical', 'xl', 'special', 'original_rank'):
                            if key in vals:
                                defaults[key] = bool(vals.get(key))
            else:
                # Eski sürümlerdeki tekil filtreleri yalnızca kayıtlı aktif lige aktar.
                active = self.league_var.get()
                try:
                    self.league_filter_settings[active]['count'] = max(1, min(100, int(data.get('count', 36))))
                except Exception:
                    pass
                for key in ('shadow', 'legendary', 'mythical', 'xl', 'special', 'original_rank'):
                    self.league_filter_settings[active][key] = bool(data.get(key, False))

            self._apply_league_filter_settings(self.league_var.get())
            self.auto_save_data_var.set(bool(data.get('auto_save_data', True)))
            self.team_source_var.set(str(data.get('team_source', 'Mevcut Pokémonlar')))
            if self.team_source_var.get() not in ('Mevcut Pokémonlar', 'Güncel Meta'):
                self.team_source_var.set('Mevcut Pokémonlar')
            try:
                mins = int(data.get('auto_search_minutes', self.pogo_auto_interval_minutes))
                mins = max(1, min(1440, mins))
                self.pogo_auto_interval_minutes = mins
                self.pogo_auto_interval_ms = mins * 60 * 1000
            except Exception:
                pass
            self.toggle_original_rank()
            self._refresh_owned_summary()
        except Exception:
            pass

    def save_preferences(self):
        try:
            self._save_current_league_filter_settings()
            self.preferences_path.parent.mkdir(parents=True, exist_ok=True)
            try:
                saved_count = max(1, min(100, int(str(self.count_var.get()).strip())))
            except Exception:
                saved_count = 36
            active_league = self.league_var.get()
            self.owned = dict(self.owned_by_league.get(active_league, {}))
            data = {
                'owned': self.owned,  # eski sürümlerle uyumluluk
                'owned_by_league': self.owned_by_league,
                'owned_status_by_league': self.owned_status_by_league,
                'league': active_league,
                'count': saved_count,
                'shadow': self.shadow_var.get(),
                'legendary': self.legendary_var.get(),
                'mythical': self.mythical_var.get(),
                'xl': self.xl_var.get(),
                'special': self.special_var.get(),
                'original_rank': self.original_rank_var.get(),
                'league_filter_settings': self.league_filter_settings,
                'auto_save_data': self.auto_save_data_var.get(),
                'team_source': self.team_source_var.get(),
                'auto_search_minutes': int(getattr(self, 'pogo_auto_interval_minutes', 30)),
            }
            temporary = self.preferences_path.with_suffix('.tmp')
            with open(temporary, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            os.replace(temporary, self.preferences_path)
        except Exception:
            pass

    def _is_owned(self, name, league=None):
        """Pokémonun sadece belirtilen/aktif ligde Mevcut olup olmadığını döndürür."""
        league = league or self.league_var.get()
        return bool(self.owned_by_league.get(league, {}).get(str(name), False))

    def _set_owned(self, name, state, league=None):
        """Mevcut işaretini yalnızca belirtilen/aktif lige kaydeder."""
        league = league or self.league_var.get()
        bucket = self.owned_by_league.setdefault(league, {})
        key = str(name)
        if state:
            bucket[key] = True
        else:
            bucket.pop(key, None)
            self.owned_status_by_league.setdefault(league, {}).pop(key, None)
        self.owned = dict(bucket)
        try:
            self._refresh_owned_summary()
            self._refresh_status_summary()
            self._refresh_ready_pokemon_panel()
        except Exception:
            pass

    def _get_owned_status(self, name, league=None):
        """Mevcut Pokémon için lig bazlı Durum: X, + veya −."""
        league = league or self.league_var.get()
        if not self._is_owned(name, league):
            return ""
        value = str(self.owned_status_by_league.get(league, {}).get(str(name), "") or "")
        if value == "-":
            value = "−"
        return value if value in ("X", "+", "−") else ""

    def _set_owned_status(self, name, status, league=None):
        """Mevcut Pokémonun Durum seçimini lig bazında kaydeder."""
        league = league or self.league_var.get()
        key = str(name)
        if not self._is_owned(key, league):
            self.owned_status_by_league.setdefault(league, {}).pop(key, None)
            return
        value = str(status or "")
        bucket = self.owned_status_by_league.setdefault(league, {})
        if value in ("✕", "X", "x", "+", "−", "-"):
            bucket[key] = "X" if value in ("✕", "x") else ("−" if value == "-" else value)
        else:
            bucket.pop(key, None)
        self.save_preferences()
        self._refresh_status_cell(key, league)
        self._refresh_status_summary()
        self._refresh_ready_pokemon_panel()

    def _refresh_status_cell(self, name, league=None):
        try:
            league = league or self.league_var.get()
            iid = self.tree_items_by_name.get(str(name))
            if iid:
                self.tree.set(iid, "status", self._get_owned_status(name, league))
        except Exception:
            pass

    def _choose_owned_status(self, event):
        """Durum hücresine tıklanınca çalışan güvenilir seçim menüsü.

        Yalnızca Mevcut (☑) olan satırlarda çalışır. Kullanıcı hücreye
        tıkladığında üç açık seçenek gösterilir: X, + ve −.
        """
        try:
            col = self.tree.identify_column(event.x)
            item = self.tree.identify_row(event.y)
            if not item:
                return False

            display = list(self.tree["displaycolumns"])
            status_col = f"#{display.index('status') + 1}" if "status" in display else None
            if col != status_col:
                return False

            name = str(self.tree.set(item, "pokemon"))
            league = self.league_var.get()
            if not self._is_owned(name, league):
                return "break"

            current = self._get_owned_status(name, league)

            # Toplevel + FocusOut yerine Tk menüsü kullanıyoruz.
            # Böylece Windows'ta menünün anında kapanması veya seçim
            # yapılamaması problemi oluşmaz.
            menu = tk.Menu(self.root, tearoff=0, font=("Segoe UI", 10))

            def choose(value):
                self._set_owned_status(name, value, league)
                try:
                    self.tree.selection_set(item)
                    self.tree.focus(item)
                except Exception:
                    pass

            options = [
                ("X", "Güçlerinin değişmesi gerekiyor"),
                ("+", "Tavsiye edilen güçlere sahip"),
                ("−", "Evrimleştirilmemiş hali var"),
            ]
            for symbol, desc in options:
                label = f"{symbol}   {desc}"
                if symbol == current:
                    label = f"✓  {label}"
                menu.add_command(
                    label=label,
                    command=lambda v=symbol: choose(v)
                )

            # İsteğe bağlı olarak mevcut durumu kaldırabilmek için menünün
            # altına temizle seçeneği eklenir.
            menu.add_separator()
            menu.add_command(
                label="Durumu temizle",
                command=lambda: choose("")
            )

            try:
                menu.tk_popup(event.x_root, event.y_root + 3)
            finally:
                try:
                    menu.grab_release()
                except Exception:
                    pass
            return "break"
        except Exception:
            return "break"

    def _handle_tree_click(self, event):
        """Ana tablo tıklamalarını tek noktadan yönetir.

        Durum sütunu, Mevcut sütunundan önce değerlendirilir; böylece iki
        ayrı Button-1 binding'inin birbirini engellemesi önlenir.
        """
        try:
            col = self.tree.identify_column(event.x)
            item = self.tree.identify_row(event.y)
            if not item:
                return
            display = list(self.tree["displaycolumns"])

            if "status" in display and col == f"#{display.index('status') + 1}":
                return self._choose_owned_status(event)

            if "current" in display and col == f"#{display.index('current') + 1}":
                return self.toggle_current(event)
        except Exception:
            return

    @staticmethod
    def _event_watch_normalize_name(value):
        """Event ve uygulama isimlerini güvenli şekilde karşılaştırır."""
        s = str(value or "").strip().lower()
        s = re.sub(r"^(?:shadow|mega|g-max|gigantamax|dynamax)\s+", "", s)
        s = re.sub(r"[^a-z0-9ğüşıöçé\- ]+", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    @staticmethod
    def _event_watch_parse_datetime(value):
        """ISO event tarihini timezone-aware datetime'a çevirir."""
        if not value:
            return None
        try:
            s = str(value).strip()
            if s.endswith("Z"):
                s = s[:-1] + "+00:00"
            dt = datetime.fromisoformat(s)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt
        except Exception:
            return None

    def _event_watch_date_text(self, dt):
        try:
            local_dt = dt.astimezone()
            return (
                f"{local_dt.day} {TR_MONTHS.get(local_dt.month, local_dt.strftime('%B'))} "
                f"{local_dt.year} • {local_dt.strftime('%H:%M')}"
            )
        except Exception:
            return str(dt)

    def _start_event_status_watcher(self):
        """Event takibini sessizce arka planda başlatır."""
        if getattr(self, "_event_watch_started", False):
            return
        self._event_watch_started = True
        self._event_watch_schedule_next(0)

    def _event_watch_schedule_next(self, delay_ms=15 * 60 * 1000):
        try:
            self.root.after(max(1000, int(delay_ms)), self._event_watch_run)
        except Exception:
            pass

    def _event_watch_run(self):
        if getattr(self, "_event_watch_busy", False):
            self._event_watch_schedule_next()
            return

        self._event_watch_busy = True
        try:
            self.event_watch_label.config(text="🔎 Event takip ediliyor", fg="#555555")
        except Exception:
            pass

        threading.Thread(
            target=self._event_watch_worker,
            daemon=True
        ).start()

    def _event_watch_worker(self):
        try:
            events = self._fetch_leekduck_events_live()
            self._cached_leekduck_events = list(events or [])
            try:
                self.root.after(0, self._dashboard_refresh_upcoming_events)
            except Exception:
                pass
            now = datetime.now(timezone.utc)

            # ÖNEMLİ: Event takibi artık yalnızca seçili ligi değil,
            # bütün liglerde Mevcut + Durum kayıtlarını kontrol eder.
            # Durum = '-' (event bekleyen) VEYA '✕' (yeniden yakalanması gereken)
            # olan tüm Mevcut Pokémonlar event adayına dahil edilir.
            wanted = {}
            for league, owned_bucket in (self.owned_by_league or {}).items():
                status_bucket = self.owned_status_by_league.get(league, {}) or {}
                for name, state in (owned_bucket or {}).items():
                    if not state:
                        continue
                    status = str(status_bucket.get(str(name), "") or "")
                    if status not in ("−", "-", "X", "✕", "x"):
                        continue
                    key = self._event_watch_normalize_name(name)
                    if not key:
                        continue
                    wanted.setdefault(key, []).append({
                        "name": str(name),
                        "league": str(league),
                        "status": status,
                    })

            matches = []
            for event in events or []:
                if not isinstance(event, dict):
                    continue
                start = self._event_watch_parse_datetime(event.get("start_time"))
                if not start or start <= now:
                    continue

                event_names = _event_pokemon_names(event)
                matched_records = []
                seen_records = set()
                for event_name in event_names:
                    key = self._event_watch_normalize_name(event_name)
                    for record in wanted.get(key, []):
                        rkey = (record["name"], record["league"], record["status"])
                        if rkey not in seen_records:
                            matched_records.append(record)
                            seen_records.add(rkey)

                for record in matched_records:
                    matches.append((start, record, event))

            # En yakın gelecek event ilk sırada olsun.
            matches.sort(key=lambda x: x[0])
            match = matches[0] if matches else None

            self.root.after(
                0,
                lambda m=match: self._event_watch_apply_result(m)
            )
        except Exception as ex:
            # Ağ/event kaynağı geçici olarak erişilemezse uygulamayı rahatsız etme.
            self.root.after(
                0,
                lambda: self._event_watch_apply_error(str(ex))
            )

    def _event_watch_apply_error(self, error_text=""):
        self._event_watch_busy = False
        try:
            self.event_watch_label.config(
                text="🔎 Event takip ediliyor",
                fg="#555555"
            )
        except Exception:
            pass
        self._event_watch_schedule_next()

    def _event_watch_apply_result(self, match):
        self._event_watch_busy = False

        if not match:
            self._event_watch_last_match_key = None
            self._event_watch_current_match = None
            try:
                self.event_watch_label.config(
                    text="✓ Event takip ediliyor • Yakında gelecek event yok",
                    fg="#008000"
                )
            except Exception:
                pass
            self._event_watch_clear_alert()
            self._event_watch_schedule_next()
            return

        start, record, event = match
        pokemon_name = record.get("name", "Pokémon")
        league = record.get("league", "")
        status = record.get("status", "")
        title = _tr_event_title(event.get("title", "Etkinlik"))
        key = (
            self._event_watch_normalize_name(pokemon_name),
            league,
            status,
            str(event.get("start_time") or ""),
            str(event.get("url") or "")
        )
        self._event_watch_last_match_key = key
        self._event_watch_current_match = {
            "start": start,
            "pokemon_name": pokemon_name,
            "league": league,
            "status": status,
            "title": title,
            "event": event,
        }

        try:
            self.event_watch_label.config(
                text="⚠ Event takip ediliyor • Tüm ligler",
                fg="#c00000"
            )
        except Exception:
            pass

        self._event_watch_show_alert(
            pokemon_name,
            title,
            self._event_watch_date_text(start),
            league=league,
            status=status,
            start=start
        )

        # Görseli ağdan arka planda indir; Tk PhotoImage yalnızca ana thread'de oluşsun.
        threading.Thread(
            target=self._event_watch_load_image_worker,
            args=(pokemon_name, key),
            daemon=True
        ).start()

        self._event_watch_schedule_next()

    def _event_watch_countdown_text(self, start):
        """Event başlangıcına kalan süreyi Türkçe ve okunaklı biçimde döndürür."""
        try:
            now = datetime.now(timezone.utc)
            if start.tzinfo is None:
                start = start.replace(tzinfo=timezone.utc)
            total = int((start - now).total_seconds())
            if total <= 0:
                return "Şimdi başlıyor"
            days, rem = divmod(total, 86400)
            hours, rem = divmod(rem, 3600)
            minutes, seconds = divmod(rem, 60)
            parts = []
            if days:
                parts.append(f"{days} gün")
            if hours or days:
                parts.append(f"{hours} saat")
            parts.append(f"{minutes} dk")
            if not days and not hours:
                parts.append(f"{seconds} sn")
            return "Kalan: " + " ".join(parts)
        except Exception:
            return "Kalan süre hesaplanamadı"

    def _event_watch_update_countdown(self):
        """Gösterilen event için kalan süreyi saniyelik olarak günceller."""
        try:
            current = getattr(self, "_event_watch_current_match", None)
            label = getattr(self, "_event_alert_countdown_label", None)
            if not current or label is None or not label.winfo_exists():
                return
            start = current.get("start")
            text = self._event_watch_countdown_text(start)
            label.config(text=text)
            if text == "Şimdi başlıyor":
                # Event zamanı geldiyse yeni event listesini hemen kontrol et.
                self._event_watch_busy = False
                self.root.after(1000, self._event_watch_run)
                return
            self.root.after(1000, self._event_watch_update_countdown)
        except Exception:
            pass

    def _event_watch_clear_alert(self):
        try:
            if getattr(self, "event_alert_frame", None):
                self.event_alert_frame.pack_forget()
            self._event_watch_image_photo = None
            self._event_watch_current_match = None
            self._event_alert_countdown_label = None
        except Exception:
            pass

    def _event_watch_show_alert(self, pokemon_name, title, date_text, photo=None, league="", status="", start=None):
        try:
            frame = self.event_alert_frame
            if not frame.winfo_exists():
                return

            for child in frame.winfo_children():
                child.destroy()

            self._event_alert_countdown_label = None

            if photo is not None:
                image_label = tk.Label(frame, image=photo, bd=0)
                image_label.pack(side="left", padx=(8, 4))
                self._event_watch_image_photo = photo

            info = tk.Frame(frame, bd=0)
            info.pack(side="left", padx=4)

            tk.Label(
                info,
                text="⚠ YAKINDA EVENT GELECEKTİR!",
                font=("Segoe UI", 8, "bold"),
                wraplength=190,
                justify="left",
                fg="#c00000"
            ).pack(anchor="w")

            tk.Label(
                info,
                text=f"Pokémon: {pokemon_name}",
                font=("Segoe UI", 10, "bold")
            ).pack(anchor="w")

            tk.Label(
                info,
                text=f"Lig: {league}   •   Durum: {status}",
                font=("Segoe UI", 9, "bold")
            ).pack(anchor="w")

            self._event_alert_countdown_label = tk.Label(
                info,
                text=self._event_watch_countdown_text(start) if start else "Kalan süre: -",
                font=("Segoe UI", 10, "bold"),
                fg="#c00000"
            )
            self._event_alert_countdown_label.pack(anchor="w")

            tk.Label(
                info,
                text=f"📅 {date_text}",
                font=("Segoe UI", 9, "bold")
            ).pack(anchor="w")

            tk.Label(
                info,
                text=title,
                font=("Arial", 8),
                wraplength=260,
                justify="left"
            ).pack(anchor="w")

            frame.pack(fill="x", padx=10, pady=(2, 5), before=self.dashboard_upcoming)
            self.root.after(1000, self._event_watch_update_countdown)
        except Exception:
            pass

    def _event_watch_load_image_worker(self, pokemon_name, key):
        try:
            data = self.download_image_bytes(pokemon_name)
            if not data:
                return
            img = Image.open(BytesIO(data)).convert("RGBA")
            img.thumbnail((48, 48), Image.LANCZOS)

            def apply():
                try:
                    if key != getattr(self, "_event_watch_last_match_key", None):
                        return
                    photo = ImageTk.PhotoImage(img)
                    self._event_watch_show_current_alert_with_photo(photo)
                except Exception:
                    pass

            self.root.after(0, apply)
        except Exception:
            pass

    def _event_watch_show_current_alert_with_photo(self, photo):
        try:
            current = getattr(self, "_event_watch_current_match", None)
            if not current:
                return
            self._event_watch_image_photo = photo
            self._event_watch_show_alert(
                current.get("pokemon_name", "Pokémon"),
                current.get("title", "Etkinlik"),
                self._event_watch_date_text(current.get("start")),
                photo=photo,
                league=current.get("league", ""),
                status=current.get("status", ""),
                start=current.get("start")
            )
        except Exception:
            pass

    def _focus_manual_pokemon_search(self):
        """Sol menüdeki Web / Manuel > Manuel Pokemon Arama seçeneğini çalıştırır."""
        try:
            entry = getattr(self, "manual_pokemon_entry", None)
            if entry is not None and entry.winfo_exists():
                entry.focus_set()
                entry.select_range(0, tk.END)
                self.set_status("Manuel Pokémon arama hazır.")
            else:
                self.set_status("Manuel Pokémon arama alanı hazır değil.")
        except Exception:
            pass

    def _manual_pokemon_all_names(self):
        """Aktif ligdeki mevcut ranking verisinden benzersiz Pokémon adları."""
        names=[]
        for p in (getattr(self,"ranking_data",[]) or []):
            if not isinstance(p,dict):
                continue
            n=str(p.get("name") or p.get("speciesName") or p.get("pokemon") or "").strip()
            if n and n not in names:
                names.append(n)
        return names

    def _manual_pokemon_filter(self, event=None):
        try:
            q=self.manual_pokemon_var.get().strip().lower()
            lb=self.manual_pokemon_list
            lb.delete(0,tk.END)
            if not q:
                lb.place_forget()
                self._manual_pokemon_matches=[]
                self._manual_pokemon_selected=None
                return

            names=self._manual_pokemon_all_names()
            starts=[n for n in names if n.lower().startswith(q)]
            contains=[n for n in names if q in n.lower() and n not in starts]
            matches=starts+contains
            self._manual_pokemon_matches=matches

            for n in matches[:30]:
                lb.insert(tk.END,n)

            if matches:
                # Entry'nin hemen altında açılır liste.
                self.root.update_idletasks()
                x=self.manual_pokemon_entry.winfo_rootx()-self.root.winfo_rootx()
                y=self.manual_pokemon_entry.winfo_rooty()-self.root.winfo_rooty()+self.manual_pokemon_entry.winfo_height()
                lb.place(x=x,y=y)
                lb.lift()
                lb.selection_set(0)
                self._manual_pokemon_selected=matches[0]
            else:
                lb.place_forget()
                self._manual_pokemon_selected=None
        except Exception:
            pass

    def _manual_pokemon_list_select(self,event=None):
        try:
            sel=self.manual_pokemon_list.curselection()
            if not sel: return
            n=str(self.manual_pokemon_list.get(sel[0]))
            self._manual_pokemon_selected=n
            self.manual_pokemon_var.set(n)
        except Exception:
            pass

    def _manual_pokemon_show_selected(self,event=None):
        try:
            typed=str(self.manual_pokemon_var.get()).strip()
            name=self._manual_pokemon_selected or typed
            if not name:
                return "break"

            names=self._manual_pokemon_all_names()
            exact=next((n for n in names if n.lower()==name.lower()),None)
            if exact:
                name=exact
            else:
                matches=[n for n in names if n.lower().startswith(name.lower())]
                if not matches:
                    matches=[n for n in names if name.lower() in n.lower()]
                if not matches:
                    messagebox.showinfo("Manuel Pokémon Ara",f"'{name}' bulunamadı.")
                    return "break"
                name=matches[0]

            self._manual_pokemon_selected=name
            self.manual_pokemon_var.set(name)
            self.manual_pokemon_list.place_forget()
            self._manual_pokemon_active=name
            self._manual_pokemon_apply_filter(name)
        except Exception as ex:
            messagebox.showerror("Manuel Pokémon Ara",str(ex))
        return "break"

    def _manual_pokemon_apply_filter(self,name):
        """Seçilen Pokémon'u ana tabloya, mevcut satır üretim sistemiyle gösterir.

        Önemli: ranking_data kesinlikle daraltılmaz. Böylece manuel arama listesi
        ve lig verisi bozulmadan kalır; sadece görünür Treeview tek Pokémon olur.
        """
        target = str(name or "").strip().lower()
        if not target:
            return
        try:
            league = self.league_var.get()
            if league == "Gigamax":
                return

            source = getattr(self, "ranking_data", []) or []
            selected = None
            for p in source:
                if not isinstance(p, dict):
                    continue
                n = str(p.get("speciesName") or p.get("name") or p.get("pokemon") or "").strip()
                if n.lower() == target:
                    selected = p
                    break

            if selected is None:
                for p in source:
                    if not isinstance(p, dict):
                        continue
                    n = str(p.get("speciesName") or p.get("name") or p.get("pokemon") or "").strip()
                    if n.lower().startswith(target) or target in n.lower():
                        selected = p
                        break

            if selected is None:
                messagebox.showinfo("Manuel Pokémon Ara", f"'{name}' için veri bulunamadı.")
                return

            # Yeni generation: eski arka plan hesaplamaları bu satırı ezemez.
            self.generation += 1
            gen = self.generation

            for item in self.tree.get_children():
                self.tree.delete(item)
            self.evolution_names = {}

            self.info_var.set(
                f"{league} • Manuel Pokémon: {selected.get('speciesName') or name} • 1 Pokémon"
            )
            self.set_status("Seçilen Pokémon hesaplanıyor...")

            # calculate_rows mevcut ana tablo üretimini kullanır: Rank, tip, evrim,
            # Rank-1 IV, CP, XL, özel güç, hareketler, Mevcut ve Durum dahil.
            threading.Thread(
                target=self.calculate_rows,
                args=([selected], league, gen, str(selected.get("speciesName") or name)),
                daemon=True
            ).start()
        except Exception as ex:
            messagebox.showerror("Manuel Pokémon Ara", str(ex))

    def _manual_pokemon_clear_filter(self):
        original=self._manual_pokemon_original_ranking_data
        if isinstance(original,list):
            self.ranking_data=original
        self._manual_pokemon_original_ranking_data=None
        self._manual_pokemon_active=None

    def _refresh_status_summary(self):
        """Üst ortadaki durum kutusunu Mevcut Pokémonların durumlarına göre günceller."""
        try:
            status_label = getattr(self, "status_summary_label", None)
            event_label = getattr(self, "event_waiting_label", None)
            recatch_label = getattr(self, "recatch_label", None)
            if status_label is None:
                return

            league = self.league_var.get()
            owned = self.owned_by_league.get(league, {}) or {}
            owned_names = [str(name) for name, state in owned.items() if state]
            status_bucket = self.owned_status_by_league.get(league, {}) or {}

            statuses = [
                str(status_bucket.get(name, "") or "")
                for name in owned_names
            ]

            # Durum belirtilmediyse kırmızı.
            complete = bool(owned_names) and all(
                s in ("X", "✕", "x", "+", "−", "-") for s in statuses
            )
            if complete:
                status_label.config(
                    text="✓ Durum belirtildi",
                    fg="#008000"
                )
            else:
                status_label.config(
                    text="⚠ Durum belirtilmedi",
                    fg="#c00000"
                )

            # Eksi = event bekleyen Pokémon.
            event_count = sum(1 for s in statuses if s in ("−", "-"))
            # Çarpı = yeniden yakalanması gereken Pokémon.
            recatch_count = sum(1 for s in statuses if s in ("X", "✕", "x"))

            if event_count == 0:
                event_label.config(
                    text="• Event bekleyen Pokémon: 0",
                    fg="#008000"
                )
            else:
                event_label.config(
                    text=f"• Event bekleyen Pokémon: {event_count}",
                    fg="#c00000"
                )

            if recatch_count == 0:
                recatch_label.config(
                    text="• Yeniden yakalanması gereken Pokémon: 0",
                    fg="#008000"
                )
            else:
                recatch_label.config(
                    text=f"• Yeniden yakalanması gereken Pokémon: {recatch_count}",
                    fg="#c00000"
                )
        except Exception:
            pass

    def _ready_pokemon_ordered_names(self):
        """Aktif ligde Mevcut ve Durumu '+' olan Pokémonları ranking sırasıyla döndürür."""
        try:
            league = self.league_var.get()
            owned = self.owned_by_league.get(league, {}) or {}
            statuses = self.owned_status_by_league.get(league, {}) or {}
            allowed = {str(n) for n, state in owned.items() if state and str(statuses.get(str(n), "")) == "+"}
            if not allowed:
                return []

            # Öncelik aktif ligdeki ranking_data sırasıdır. Böylece PvP'ye hazır
            # kutusu ana sıralamayı hiçbir şekilde değiştirmez.
            result = []
            seen = set()
            for p in (self.ranking_data or []):
                if not isinstance(p, dict):
                    continue
                name = str(p.get("speciesName") or p.get("name") or "").strip()
                if name in allowed and name not in seen:
                    result.append(name)
                    seen.add(name)

            # Ranking verisinde bulunmayan eski Mevcut kayıtlarını da kaybetme;
            # bunlar en sonda gösterilir.
            for name in allowed:
                if name not in seen:
                    result.append(name)
                    seen.add(name)
            return result
        except Exception:
            return []

    def _bind_ready_mousewheel(self, widget=None):
        """Hazır Pokémon alanında fare tekerini tüm alt widgetlardan canvas'a yönlendirir."""
        canvas = getattr(self, "ready_pokemon_canvas", None)
        if canvas is None:
            return
        if widget is None:
            widget = getattr(self, "ready_pokemon_inner", None)
        if widget is None:
            return

        tag = "SinKAReadyScroll"
        try:
            # Bir kez sınıf etiketi oluştur; böylece kartların/etiketlerin
            # üzerinde fare varken de canvas kesin olarak kayar.
            if not getattr(self, "_ready_scroll_tag_bound", False):
                def wheel(event):
                    delta = getattr(event, "delta", 0)
                    if delta:
                        canvas.yview_scroll(-int(delta / 120), "units")
                        return "break"
                self.root.bind_class(tag, "<MouseWheel>", wheel)
                self.root.bind_class(tag, "<Button-4>", lambda e: (canvas.yview_scroll(-1, "units"), "break"))
                self.root.bind_class(tag, "<Button-5>", lambda e: (canvas.yview_scroll(1, "units"), "break"))
                self._ready_scroll_tag_bound = True

            tags = list(widget.bindtags())
            if tag not in tags:
                tags.insert(1, tag)
                widget.bindtags(tuple(tags))
            for child in widget.winfo_children():
                self._bind_ready_mousewheel(child)
        except Exception:
            pass

    def _refresh_ready_pokemon_panel(self):
        """PvP'ye hazır Pokémonlar panelini günceller; yalnızca UI thread'de çağrılır."""
        try:
            canvas = getattr(self, "ready_pokemon_canvas", None)
            inner = getattr(self, "ready_pokemon_inner", None)
            if canvas is None or inner is None:
                return

            names = self._ready_pokemon_ordered_names()
            self._ready_pokemon_names = names

            # Sayaç aktif ligdeki hazır Pokémon sayısını gösterir.
            count_label = getattr(self, "ready_pokemon_count_label", None)
            if count_label is not None:
                count_label.config(text=f"PvP HAZIR: {len(names)} Pokémon")

            for child in inner.winfo_children():
                child.destroy()

            if not names:
                tk.Label(
                    inner, text="Henüz hazır Pokémon yok",
                    font=("Segoe UI", 9), padx=8, pady=17
                ).pack(side="left")
            else:
                # Dar alanda 3 sütun kullan; kalan Pokémonlar aşağı satırlara iner.
                for idx, name in enumerate(names):
                    self._ready_pokemon_make_card(inner, name, idx)

            inner.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.yview_moveto(0)
        except Exception:
            pass

        self._bind_ready_mousewheel()

    def _ready_pokemon_make_card(self, parent, name, idx=0):
        card = tk.Frame(parent, bd=0, relief="flat", padx=2, pady=1)
        col = idx % 3
        row = idx // 3
        card.grid(row=row, column=col, padx=1, pady=1, sticky="nsew")
        for c in range(3):
            parent.grid_columnconfigure(c, weight=1, uniform="readycol")

        photo = self._ready_pokemon_photos.get(name)
        if photo is not None:
            img_label = tk.Label(card, image=photo, bd=0)
            img_label.image = photo
        else:
            img_label = tk.Label(card, text="◉", font=("Segoe UI", 20, "bold"), width=2, height=1, bd=0)
        img_label.pack(side="top")

        tk.Label(
            card, text=name, font=("Segoe UI", 8, "bold"),
            width=11, anchor="center"
        ).pack(side="top")

        if photo is None:
            self._ready_pokemon_load_image_async(name)

    def _ready_pokemon_load_image_async(self, name):
        if not PIL_OK or not name or name in self._ready_pokemon_loading:
            return
        if name in self._ready_pokemon_photos:
            return
        self._ready_pokemon_loading.add(name)

        def worker():
            data = None
            try:
                data = self.download_image_bytes(name)
            except Exception:
                data = None

            def apply():
                self._ready_pokemon_loading.discard(name)
                if not data or name not in self._ready_pokemon_names:
                    return
                try:
                    img = Image.open(BytesIO(data)).convert("RGBA")
                    img.thumbnail((42, 42), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    self._ready_pokemon_photos[name] = photo
                    self._ready_pokemon_photo_refs.append(photo)
                    self._refresh_ready_pokemon_panel()
                except Exception:
                    pass

            try:
                self.root.after(0, apply)
            except Exception:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _refresh_dashboard_owned_stats(self):
        """Dashboarddaki lig bazlı Mevcut / PvP Hazır sayaçlarını günceller."""
        try:
            labels = getattr(self, "dashboard_owned_stats", {})
            for league, refs in labels.items():
                owned = self.owned_by_league.get(league, {}) or {}
                statuses = self.owned_status_by_league.get(league, {}) or {}
                total = sum(1 for state in owned.values() if state)
                ready = sum(1 for name, state in owned.items() if state and str(statuses.get(str(name), "")) == "+")
                refs["total"].config(text=f"{total}")
                refs["ready"].config(text=f"{ready}")
        except Exception:
            pass

    def _refresh_owned_summary(self):
        """Üst sağdaki lig bazlı Mevcut Pokémon sayaçlarını günceller."""
        try:
            buttons = getattr(self, "_owned_summary_buttons", {})
            for league, button in buttons.items():
                count = len(self.owned_by_league.get(league, {}) or {})
                button.config(text=f"{league}: {count}")
        except Exception:
            pass
        self._refresh_status_summary()
        self._refresh_dashboard_owned_stats()

    def _load_owned_detail_image(self, name, label, win, size=120):
        """Mevcut Pokémon penceresi için SinKA'nın ana görsel motorunu kullanır.
        Ağ işlemi worker thread'de, PhotoImage oluşturma ana Tk thread'inde yapılır.
        """
        try:
            if not PIL_OK:
                return

            # Ana uygulamadaki çoklu kaynak + form + cache sistemini kullan.
            data = self.download_image_bytes(name)
            if not data:
                print(f"[Pokemon Resim] Görsel bulunamadı: {name}")
                return

            from io import BytesIO
            img = Image.open(BytesIO(data)).convert("RGBA")
            img.thumbnail((size, size), Image.LANCZOS)

            def apply():
                try:
                    if not win.winfo_exists() or not label.winfo_exists():
                        return
                    photo = ImageTk.PhotoImage(img)
                    label.configure(image=photo, text="")
                    # PhotoImage referansını label üzerinde tut.
                    label.image = photo
                except Exception as exc:
                    print(f"[Pokemon Resim UI] {name}: {exc}")

            win.after(0, apply)
        except Exception as exc:
            print(f"[Pokemon Resim] {name}: {exc}")

    def _show_owned_pokemon_window(self, league):
        """Mevcut Pokémonları gösteren sade, non-blocking pencere.
        Pencere açılırken ağ isteği yapılmaz; yalnızca bellekteki owned_by_league
        ve owned_status_by_league kayıtları kullanılır.
        """
        try:
            # Eski pencere varsa kapat.
            old = getattr(self, "_owned_detail_window", None)
            try:
                if old is not None and old.winfo_exists():
                    old.destroy()
            except Exception:
                pass

            win = tk.Toplevel(self.root)
            self._owned_detail_window = win
            win.title(f"Mevcut Pokémonlar — {league}")
            win.geometry("980x700")
            win.minsize(760, 520)
            win.configure(bg="#F4F7FB")

            # Pencereyi öne getir ama grab/focus_force kullanma.
            try:
                win.lift()
            except Exception:
                pass

            header = tk.Frame(win, bg="#16233D", height=74)
            header.pack(fill="x")
            header.pack_propagate(False)

            tk.Label(
                header, text=f"{league} — Mevcut Pokémonlar",
                bg="#16233D", fg="white",
                font=("Segoe UI", 18, "bold")
            ).pack(side="left", padx=24, pady=12)

            # İçerik alanı
            body = tk.Frame(win, bg="#F4F7FB")
            body.pack(fill="both", expand=True, padx=18, pady=(14, 8))

            # Sayaç
            count_var = tk.StringVar(value="")
            tk.Label(
                body, textvariable=count_var,
                bg="#F4F7FB", fg="#536174",
                font=("Segoe UI", 10, "bold")
            ).pack(anchor="w", pady=(0, 8))

            # Scrollable alan
            outer = tk.Frame(body, bg="#FFFFFF", bd=1, relief="solid")
            outer.pack(fill="both", expand=True)

            canvas = tk.Canvas(
                outer, bg="#FFFFFF", highlightthickness=0
            )
            scrollbar = ttk.Scrollbar(
                outer, orient="vertical", command=canvas.yview
            )
            cards = tk.Frame(canvas, bg="#FFFFFF")

            canvas_window = canvas.create_window(
                (0, 0), window=cards, anchor="nw"
            )
            canvas.configure(yscrollcommand=scrollbar.set)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            def _resize_inner(event=None):
                try:
                    canvas.configure(scrollregion=canvas.bbox("all"))
                    canvas.itemconfigure(canvas_window, width=canvas.winfo_width())
                except Exception:
                    pass

            cards.bind("<Configure>", _resize_inner)
            canvas.bind("<Configure>", _resize_inner)

            # Fare tekeri: Windows / macOS ve Linux için.
            def _wheel(event):
                try:
                    if getattr(event, "delta", 0):
                        step = -3 if event.delta > 0 else 3
                    else:
                        step = -3 if getattr(event, "num", 0) == 4 else 3
                    canvas.yview_scroll(step, "units")
                except Exception:
                    pass

            canvas.bind("<MouseWheel>", _wheel)
            canvas.bind("<Button-4>", _wheel)
            canvas.bind("<Button-5>", _wheel)

            # Tekerlek kartların üzerindeyken de çalışsın.
            def _bind_wheel_recursive(widget):
                try:
                    widget.bind("<MouseWheel>", _wheel)
                    widget.bind("<Button-4>", _wheel)
                    widget.bind("<Button-5>", _wheel)
                    for child in widget.winfo_children():
                        _bind_wheel_recursive(child)
                except Exception:
                    pass

            def _read_owned():
                """Owned kaydını güvenli biçimde listeye çevir."""
                data = getattr(self, "owned_by_league", {}) or {}
                league_data = data.get(league, {}) or {}
                result = []

                if isinstance(league_data, dict):
                    for name, value in league_data.items():
                        # Dict değeri farklı sürümlerde bool/dict olabilir.
                        if isinstance(value, dict):
                            is_owned = value.get("owned", value.get("current", True))
                        else:
                            is_owned = bool(value)
                        if is_owned:
                            result.append(str(name))
                elif isinstance(league_data, (list, tuple, set)):
                    result = [str(x) for x in league_data]

                # Eski/alternatif kayıt yapısı varsa fallback.
                if not result:
                    alt = getattr(self, "owned_status_by_league", {}) or {}
                    alt_data = alt.get(league, {}) or {}
                    if isinstance(alt_data, dict):
                        # Status kaydı bulunanlar Mevcut olarak kabul edilmez;
                        # yalnızca mevcut dict'te owned yoksa isim eşleşmesi için
                        # güvenli fallback.
                        for name in alt_data:
                            if name in league_data:
                                result.append(str(name))

                # Sıralamayı ana ranking_data ile mümkün olduğunca koru.
                ranking = getattr(self, "ranking_data", {}) or {}
                ordered = []
                seen = set()
                if isinstance(ranking, dict):
                    rows = ranking.get(league, [])
                else:
                    rows = []
                for row in rows if isinstance(rows, list) else []:
                    if isinstance(row, dict):
                        n = row.get("name") or row.get("pokemon") or row.get("speciesName")
                        if n and str(n) in result and str(n) not in seen:
                            ordered.append(str(n))
                            seen.add(str(n))
                for n in result:
                    if n not in seen:
                        ordered.append(n)
                        seen.add(n)
                return ordered

            def _status_for(name):
                statuses = getattr(self, "owned_status_by_league", {}) or {}
                d = statuses.get(league, {}) or {}
                value = d.get(name, "")
                if value in ("✕", "x", "X"):
                    return "X"
                if value in ("+", "−", "-"):
                    return "−" if value == "-" else value
                return "—"

            def _remove(name):
                """Mevcut kaydından kaldır ve durumunu temizle."""
                try:
                    owned = getattr(self, "owned_by_league", {}) or {}
                    ld = owned.get(league)
                    if isinstance(ld, dict):
                        ld.pop(name, None)
                    elif isinstance(ld, list):
                        while name in ld:
                            ld.remove(name)

                    statuses = getattr(self, "owned_status_by_league", {}) or {}
                    sd = statuses.get(league)
                    if isinstance(sd, dict):
                        sd.pop(name, None)

                    # Ana tabloyu güncelle; hata olursa pencere yine açık kalır.
                    try:
                        self._refresh_owned_summary()
                    except Exception:
                        pass
                    try:
                        self.refresh_table()
                    except Exception:
                        try:
                            self.change_league()
                        except Exception:
                            pass

                    _render()
                except Exception as exc:
                    # Sessiz messagebox yerine pencerenin içinde hata göster.
                    error_var.set(f"Kaldırılamadı: {exc}")

            error_var = tk.StringVar(value="")

            def _render():
                for child in cards.winfo_children():
                    child.destroy()

                names = _read_owned()
                count_var.set(f"{len(names)} Pokémon mevcut")

                if not names:
                    empty = tk.Frame(cards, bg="#FFFFFF", height=190)
                    empty.pack(fill="x", padx=20, pady=25)
                    empty.pack_propagate(False)
                    tk.Label(
                        empty,
                        text=f"Bu ligde Mevcut Pokémon bulunmuyor.",
                        bg="#FFFFFF", fg="#344054",
                        font=("Segoe UI", 14, "bold")
                    ).pack(pady=(55, 4))
                    tk.Label(
                        empty,
                        text="Ana listedeki 'Mevcut' kutusunu işaretlediğinde burada görünecek.",
                        bg="#FFFFFF", fg="#7A8699",
                        font=("Segoe UI", 10)
                    ).pack()
                    return

                for idx, name in enumerate(names):
                    card = tk.Frame(
                        cards, bg="#F8FAFD",
                        bd=1, relief="solid", height=94
                    )
                    card.pack(fill="x", padx=10, pady=7)
                    card.pack_propagate(False)

                    # Büyük ve tam görsel alanı. Label'a character-width vermiyoruz;
                    # böylece Pokémon resmi kesilmiyor.
                    image_frame = tk.Frame(
                        card, bg="#F8FAFD", width=142, height=132
                    )
                    image_frame.pack(side="left", padx=(10, 6), pady=3)
                    image_frame.pack_propagate(False)

                    img_label = tk.Label(
                        image_frame, text="Yükleniyor…",
                        bg="#F8FAFD", fg="#AAB4C3",
                        font=("Segoe UI", 8)
                    )
                    img_label.pack(fill="both", expand=True)

                    # Resim ağdan ayrı thread'de yüklenir; UI beklemez.
                    threading.Thread(
                        target=self._load_owned_detail_image,
                        args=(name, img_label, win, 120),
                        daemon=True
                    ).start()

                    mid = tk.Frame(card, bg="#F8FAFD")
                    mid.pack(side="left", fill="both", expand=True, padx=8)

                    tk.Label(
                        mid, text=name,
                        bg="#F8FAFD", fg="#182230",
                        font=("Segoe UI", 13, "bold"),
                        anchor="w"
                    ).pack(anchor="w", pady=(16, 2))

                    st = _status_for(name)
                    tk.Label(
                        mid, text=f"Durum: {st}",
                        bg="#F8FAFD", fg="#536174",
                        font=("Segoe UI", 10, "bold"),
                        anchor="w"
                    ).pack(anchor="w")

                    btn = tk.Button(
                        card, text="✕ KALDIR",
                        command=lambda n=name: _remove(n),
                        bg="#E8EDF4", fg="#26364D",
                        activebackground="#D9E0EA",
                        relief="flat", bd=0,
                        font=("Segoe UI", 9, "bold"),
                        padx=16, pady=8,
                        cursor="hand2"
                    )
                    btn.pack(side="right", padx=16)

                # Kartların üstündeyken tekerlek kaydırmayı etkinleştir.
                _bind_wheel_recursive(cards)

            # Alt durum + yenile
            bottom = tk.Frame(win, bg="#F4F7FB")
            bottom.pack(fill="x", padx=18, pady=(0, 14))

            tk.Label(
                bottom, textvariable=error_var,
                bg="#F4F7FB", fg="#B42318",
                font=("Segoe UI", 9, "bold")
            ).pack(side="left", padx=4)

            tk.Button(
                bottom, text="🔄 YENİLE",
                command=lambda: (_render(), error_var.set("")),
                bg="#16233D", fg="white",
                activebackground="#263A62",
                relief="flat", bd=0,
                font=("Segoe UI", 10, "bold"),
                padx=20, pady=9, cursor="hand2"
            ).pack(side="right")

            _render()

        except Exception as exc:
            # Ana uygulamayı durdurma; terminale yaz.
            print(f"[Mevcut Pokemon Penceresi] {type(exc).__name__}: {exc}")

    def _build_owned_summary_ui(self):
        """Varsayılan olarak üst satırda sayaç oluşturur (geriye dönük uyumluluk)."""
        return self._build_owned_summary_ui_in_parent(self.top_action_row)

    def _build_owned_summary_ui_in_parent(self, parent):
        """Lig sayaçlarını verilen ana çerçeveye yerleştirir."""
        try:
            box = ttk.LabelFrame(parent, text="Mevcut Pokémonlar")
            box.pack(side="right", anchor="ne", padx=(8, 0), pady=0)
            self.owned_summary_box = box
            for league in LEAGUES:
                button = ttk.Button(
                    box,
                    text=f"{league}: 0",
                    command=lambda lg=league: self._show_owned_pokemon_window(lg)
                )
                button.pack(side="left", padx=3, pady=5, ipadx=4, ipady=3)
                self._owned_summary_buttons[league] = button
            self._refresh_owned_summary()
        except Exception:
            pass

    def _dashboard_open_settings(self):
        """Kullanıcının uygulama davranışlarını yönetebileceği modern ayar paneli."""
        try:
            win = tk.Toplevel(self.root)
            win.title("SinKA PvP - Ayarlar")
            win.transient(self.root)
            win.resizable(False, False)
            win.configure(bg="#eef3f8")

            outer = tk.Frame(win, bg="#eef3f8")
            outer.pack(fill="both", expand=True, padx=18, pady=18)

            header = tk.Frame(outer, bg="#152b4d", height=76)
            header.pack(fill="x")
            header.pack_propagate(False)
            tk.Label(header, text="⚙  SinKA PvP Ayarları", bg="#152b4d", fg="white",
                     font=("Segoe UI", 17, "bold")).pack(anchor="w", padx=22, pady=(14, 0))
            tk.Label(header, text="Uygulamanın arama, kayıt ve takım oluşturma davranışlarını buradan belirleyin.",
                     bg="#152b4d", fg="#b9cbe3", font=("Segoe UI", 9)).pack(anchor="w", padx=22, pady=(2, 0))

            body = tk.Frame(outer, bg="white", highlightbackground="#d5deea", highlightthickness=1)
            body.pack(fill="both", expand=True, pady=(10, 10))

            def section(title, subtitle):
                f = tk.Frame(body, bg="white")
                f.pack(fill="x", padx=20, pady=(16, 4))
                tk.Label(f, text=title, bg="white", fg="#152b4d", font=("Segoe UI", 11, "bold")).pack(anchor="w")
                tk.Label(f, text=subtitle, bg="white", fg="#718096", font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))
                return f

            f1 = section("Kayıt ve veri", "Tercihleriniz ve Mevcut/Durum bilgileri için önerilen seçenekler.")
            tk.Checkbutton(f1, text="Verileri otomatik kaydet", variable=self.auto_save_data_var,
                           bg="white", activebackground="white", anchor="w", font=("Segoe UI", 10)).pack(fill="x", pady=(7, 2))
            tk.Label(f1, text="Önerilen: Açık — uygulama kapanmadan önce yaptığınız değişikliklerin korunmasını sağlar.",
                     bg="white", fg="#718096", font=("Segoe UI", 8)).pack(anchor="w", padx=(24, 0))

            f2 = section("Otomatik Pokémon arama", "Great / Ultra / Master aramalarının tekrar aralığını belirleyin.")
            r = tk.Frame(f2, bg="white")
            r.pack(fill="x", pady=(8, 2))
            tk.Label(r, text="Arama aralığı:", bg="white", fg="#26364d", font=("Segoe UI", 10)).pack(side="left")
            interval_var = tk.StringVar(value=str(getattr(self, 'pogo_auto_interval_minutes', 30)))
            interval_entry = ttk.Spinbox(r, from_=1, to=1440, textvariable=interval_var, width=8)
            interval_entry.pack(side="left", padx=10)
            tk.Label(r, text="dakika", bg="white", fg="#718096", font=("Segoe UI", 9)).pack(side="left")
            tk.Label(f2, text="Öneri: 30 dakika. Daha kısa aralık daha sık veri çeker; gereksiz yük oluşturabilir.",
                     bg="white", fg="#718096", font=("Segoe UI", 8)).pack(anchor="w", padx=(0, 0), pady=(2, 0))

            f3 = section("Takım oluşturma", "Takım motorunun ilk aramada hangi Pokémon havuzunu kullanacağını seçin.")
            team_combo = ttk.Combobox(f3, textvariable=self.team_source_var, state="readonly", width=30,
                                      values=("Mevcut Pokémonlar", "Güncel Meta"))
            team_combo.pack(anchor="w", pady=(8, 2))
            tk.Label(f3, text="Öneri: Mevcut Pokémonlar — elinizde gerçekten bulunan Pokémonlarla uygulanabilir takım üretir.",
                     bg="white", fg="#718096", font=("Segoe UI", 8), wraplength=620, justify="left").pack(anchor="w")
            tk.Label(f3, text="Güncel Meta seçilirse takım motoru mevcut listeniz yerine güncel PvPoke meta havuzunu başlangıç noktası alır.",
                     bg="white", fg="#718096", font=("Segoe UI", 8), wraplength=620, justify="left").pack(anchor="w", pady=(2, 0))

            footer = tk.Frame(outer, bg="#eef3f8")
            footer.pack(fill="x")
            def apply_settings():
                try:
                    mins = int(str(interval_var.get()).strip())
                    mins = max(1, min(1440, mins))
                except Exception:
                    mins = 30
                self.pogo_auto_interval_minutes = mins
                self.pogo_auto_interval_ms = mins * 60 * 1000
                self.save_preferences()
                self.set_status(f"Ayarlar kaydedildi • Otomatik arama: {mins} dakika • Takım kaynağı: {self.team_source_var.get()}")
                win.destroy()
            ttk.Button(footer, text="Kapat", command=win.destroy).pack(side="right", padx=(8, 0))
            ttk.Button(footer, text="Kaydet ve Uygula", command=apply_settings).pack(side="right")

            win.update_idletasks()
            win.geometry("700x650")
            x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - win.winfo_width()) // 2)
            y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - win.winfo_height()) // 2)
            win.geometry(f"700x650+{x}+{y}")
            win.lift(); win.focus_force()
        except Exception as ex:
            messagebox.showerror("SinKA PvP", f"Ayarlar açılamadı: {ex}")

    def _dashboard_show_about(self):
        """Hakkında menüsü için büyük, sade SinKA güncelleme ekranı."""
        try:
            win = tk.Toplevel(self.root)
            win.title("SinKA PvP - Hakkında")
            win.transient(self.root)
            win.configure(bg="#081a33")
            try:
                win.state("zoomed")
            except Exception:
                try:
                    win.attributes("-zoomed", True)
                except Exception:
                    win.geometry(f"{self.root.winfo_width()}x{self.root.winfo_height()}")

            center = tk.Frame(win, bg="#081a33")
            center.pack(fill="both", expand=True)
            tk.Label(center, text="SinKA PvP", bg="#081a33", fg="white",
                     font=("Segoe UI", 38, "bold")).pack(pady=(150, 8))
            tk.Label(center, text="Sistem ve veri güncellemeleri hazırlanıyor...", bg="#081a33", fg="#9fb5d3",
                     font=("Segoe UI", 16)).pack(pady=4)
            tk.Label(center, text="Güncellemeleri dua ile bekleyiniz :)", bg="#081a33", fg="#dce8f7",
                     font=("Segoe UI", 22, "bold")).pack(pady=(18, 18))
            tk.Label(center, text="PvPoke • Pokémon verileri • Event takibi • SinKA PvP", bg="#081a33", fg="#7189aa",
                     font=("Segoe UI", 10)).pack()
            ttk.Button(center, text="Kapat", command=win.destroy).pack(pady=32, ipadx=18, ipady=6)
        except Exception as ex:
            messagebox.showerror("SinKA PvP", f"Hakkında ekranı açılamadı: {ex}")

    def build_ui(self):
        """SinKA PvP modern dashboard arayüzü.

        Tasarım hedefi: solda sabit navigasyon, üstte lig kartları, üçlü
        bilgi paneli, sade kontrol/filtre alanı ve geniş veri tablosu.
        Mevcut hesaplama, event, owned/status, manuel arama ve kayıt
        işlevleri aynı callback'lerle korunur.
        """
        # Eski üst Pogo/status satırı __init__ içinde daha önce oluşturuldu.
        # Yeni dashboard bunu gizler; gerekli callback'ler için widget referansları
        # korunur. Yeni görünür kontroller aşağıda yeniden oluşturulur.
        try:
            self.top_action_row.pack_forget()
        except Exception:
            pass

        BG = "#f4f7fb"
        NAVY = "#0b2348"
        NAVY2 = "#102f61"
        BLUE = "#1d65da"
        BLUE2 = "#2b78e8"
        CYAN = "#20c2b4"
        TEXT = "#152b4d"
        MUTED = "#6d7e98"
        BORDER = "#d9e2ef"
        WHITE = "#ffffff"
        RED = "#e32635"
        GREEN = "#008f58"
        GOLD = "#d99b16"

        self.root.configure(bg=BG)
        self.root.geometry("1536x900")
        self.root.minsize(1180, 720)
        self.root.update_idletasks()
        # Windows'ta uygulama açılışta tam ekran boyutunda (maksimize) başlasın.
        try:
            self.root.state("zoomed")
        except Exception:
            try:
                self.root.attributes("-zoomed", True)
            except Exception:
                pass

        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Modern.TFrame", background=BG)
        style.configure("Card.TFrame", background=WHITE)
        style.configure("Card.TLabelframe", background=WHITE, bordercolor=BORDER)
        style.configure("Card.TLabelframe.Label", background=WHITE, foreground=TEXT, font=("Segoe UI", 10, "bold"))
        style.configure("Modern.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 9))
        style.configure("Card.TLabel", background=WHITE, foreground=TEXT, font=("Segoe UI", 9))
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 18, "bold"))
        style.configure("Modern.TButton", font=("Segoe UI", 9, "bold"), padding=(10, 6))
        style.configure("Blue.TButton", font=("Segoe UI", 9, "bold"), padding=(12, 7), foreground=WHITE, background=BLUE)
        style.map("Blue.TButton", background=[("active", BLUE2)])
        style.configure("League.TButton", font=("Segoe UI", 11, "bold"), padding=(12, 9))
        style.configure("Modern.Horizontal.TProgressbar", troughcolor="#e6edf6", background=BLUE, bordercolor="#e6edf6", lightcolor=BLUE, darkcolor=BLUE)

        # ---------- ana gövde ----------
        shell = tk.Frame(self.root, bg=BG)
        shell.pack(fill="both", expand=True)

        sidebar = tk.Frame(shell, bg=NAVY, width=238)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        self.dashboard_sidebar = sidebar

        content = tk.Frame(shell, bg=BG)
        content.pack(side="left", fill="both", expand=True)
        self.dashboard_content = content

        # ---------- sol menü ----------
        brand = tk.Frame(sidebar, bg=NAVY)
        brand.pack(fill="x", padx=18, pady=(18, 20))
        logo = tk.Canvas(brand, width=42, height=42, bg=NAVY, highlightthickness=0)
        logo.pack(side="left")
        logo.create_oval(3, 3, 39, 39, fill="#ef3340", outline="")
        logo.create_rectangle(3, 20, 39, 39, fill="white", outline="")
        logo.create_line(3, 21, 39, 21, fill=NAVY, width=3)
        logo.create_oval(15, 15, 27, 27, fill="white", outline=NAVY, width=2)
        tk.Label(brand, text="SinKA PvP", bg=NAVY, fg=WHITE, font=("Segoe UI", 19, "bold")).pack(anchor="w", padx=(10, 0))
        tk.Label(brand, text="Pokémon GO PvP Assistant", bg=NAVY, fg="#b9c8df", font=("Segoe UI", 8)).pack(anchor="w", padx=(10, 0))
        tk.Label(brand, text="v1.3", bg=NAVY, fg="#b9c8df", font=("Segoe UI", 8)).pack(anchor="e")

        def nav_button(text, command=None, active=False, badge=None):
            bg = BLUE if active else NAVY
            holder = tk.Frame(sidebar, bg=bg)
            holder.pack(fill="x", padx=10, pady=3)
            b = tk.Button(holder, text=text, command=command, bg=bg, fg=WHITE,
                          activebackground=BLUE2, activeforeground=WHITE,
                          relief="flat", bd=0, anchor="w", font=("Segoe UI", 11, "bold"),
                          padx=15, pady=9, cursor="hand2")
            b.pack(fill="x")
            if badge is not None:
                tk.Label(holder, text=str(badge), bg=RED, fg=WHITE, font=("Segoe UI", 9, "bold"), width=2).place(relx=0.90, rely=0.5, anchor="center")
            return b

        # Pokémon Arama ana bölümü ve üç ayrı arama seçeneği
        search_section = tk.Frame(sidebar, bg=NAVY)
        search_section.pack(fill="x", padx=10, pady=(0, 5))
        tk.Label(
            search_section, text="🔎  Pokémon Arama", bg=NAVY, fg=WHITE,
            font=("Segoe UI", 11, "bold"), anchor="w", padx=15, pady=8
        ).pack(fill="x")

        def sub_nav_button(text, command):
            b = tk.Button(
                search_section, text="   " + text, command=command,
                bg=NAVY, fg="#d7e3f5", activebackground="#173d73",
                activeforeground=WHITE, relief="flat", bd=0, anchor="w",
                font=("Segoe UI", 10, "bold"), padx=18, pady=7, cursor="hand2"
            )
            b.pack(fill="x", padx=(12, 0))
            return b

        self.sidebar_selected_search_button = sub_nav_button(
            "Seçili Pokémon'u Ara", self._open_selected_pokemon_coordinates
        )
        self.sidebar_multi_search_button = sub_nav_button(
            "Çoklu Pokémon Ara", self._open_multiple_pokemon_coordinates
        )
        self.sidebar_auto_search_button = sub_nav_button(
            "Otomatik Ara", self.toggle_pogo_auto_search
        )

        nav_button("♟  Takım Oluştur", self.build_team)
        nav_button("▣  Event Takibi", self.show_nearby_events, badge="!")
        nav_button("⚙  Ayarlar", self._dashboard_open_settings)
        nav_button("↻  Veri Güncelle", self.reload)

        # Web / Manuel kaynak ve arama bölümü
        web_manual_section = tk.Frame(sidebar, bg=NAVY)
        web_manual_section.pack(fill="x", padx=10, pady=(7, 5))
        tk.Label(
            web_manual_section, text="Web / Manuel", bg=NAVY, fg="#8fa9cc",
            font=("Segoe UI", 9, "bold"), anchor="w", padx=15, pady=5
        ).pack(fill="x")

        def web_manual_button(text, command):
            b = tk.Button(
                web_manual_section, text="   " + text, command=command,
                bg=NAVY, fg="#d7e3f5", activebackground="#173d73",
                activeforeground=WHITE, relief="flat", bd=0, anchor="w",
                font=("Segoe UI", 10, "bold"), padx=18, pady=6, cursor="hand2"
            )
            b.pack(fill="x", padx=(12, 0))
            return b

        self.sidebar_pvpoke_button = web_manual_button("PVPoke", lambda: self.open_url("https://pvpoke.com/rankings/"))
        self.sidebar_manual_button = web_manual_button("Manuel Pokemon Arama", lambda: self.open_url("https://pogocoordinates.com/search"))
        self.sidebar_event_button = web_manual_button("Event", lambda: self.open_url("https://leekduck.com/events/"))

        nav_button("ⓘ  Hakkında", self._dashboard_show_about)

        spacer = tk.Frame(sidebar, bg=NAVY)
        spacer.pack(fill="both", expand=True)
        mark = tk.Canvas(spacer, width=160, height=150, bg=NAVY, highlightthickness=0)
        mark.pack(anchor="center", pady=(20, 0))
        mark.create_oval(25, 15, 135, 125, outline="#284b7d", width=8)
        mark.create_line(25, 70, 135, 70, fill="#284b7d", width=8)
        mark.create_oval(62, 57, 98, 93, fill="#284b7d", outline="")
        tk.Label(spacer, text="Gotta Catch\nGreat Teams!", bg=NAVY, fg="#6c8ab5", font=("Segoe UI", 14, "italic", "bold"), justify="center").pack()
        tk.Label(sidebar, text="www.sinka.com", bg=NAVY, fg="#b9c8df", font=("Segoe UI", 8)).pack(anchor="w", padx=30)
        tk.Label(sidebar, text="SinKA PvP", bg=NAVY, fg="#b9c8df", font=("Segoe UI", 8)).pack(anchor="w", padx=30, pady=(3, 0))
        tk.Label(sidebar, text=datetime.now().strftime("%d %B %Y"), bg=NAVY, fg="#b9c8df", font=("Segoe UI", 8)).pack(anchor="w", padx=30, pady=(3, 18))

        # ---------- üst bar kaldırıldı ----------
        # Lig kartları artık doğrudan üstten başlar; gereksiz SinKA PvP / tarih /
        # Pokémon ara alanı kaldırılarak ana Pokémon listesine daha fazla alan bırakılır.

        # ---------- ince durum şeridi ----------
        # Durum bilgileri artık Yakındaki Eventler kartından tamamen ayrıdır.
        # Lig kartlarının hemen üstünde, ekranı yatay olarak dolduran ince bir şerit.
        status_strip = tk.Frame(content, bg=WHITE, highlightbackground=BORDER, highlightthickness=1, height=34)
        status_strip.pack(fill="x", padx=14, pady=(0, 6))
        status_strip.pack_propagate(False)
        self.dashboard_status_strip = status_strip

        self.status_summary_label = tk.Label(status_strip, text="⚠ Durum belirtilmedi", bg=WHITE, fg=RED,
                                              font=("Segoe UI", 9, "bold"), anchor="w")
        self.status_summary_label.pack(side="left", padx=(14, 28))
        self.event_waiting_label = tk.Label(status_strip, text="Event bekleyen: 0", bg=WHITE, fg=GREEN,
                                             font=("Segoe UI", 9, "bold"), anchor="w")
        self.event_waiting_label.pack(side="left", padx=(0, 28))
        self.recatch_label = tk.Label(status_strip, text="Yeniden yakalanacak: 0", bg=WHITE, fg=GREEN,
                                       font=("Segoe UI", 9, "bold"), anchor="w")
        self.recatch_label.pack(side="left")

        # ---------- lig kartları ----------
        league_cards = tk.Frame(content, bg=BG)
        league_cards.pack(fill="x", padx=14, pady=(0, 9))
        self.dashboard_league_buttons = {}
        league_specs = [
            ("Great League", "1500 CP MAX", "♢", BLUE),
            ("Ultra League", "2500 CP MAX", "♢", "#283f62"),
            ("Master League", "SINIRSIZ CP MAX", "♢", "#3e1b72"),
            ("Mega", "", "♧", "#17304e"),
            ("Gigamax", "", "♧", "#252247"),
        ]
        for i, (lg, sub, icon, color) in enumerate(league_specs):
            card = tk.Frame(league_cards, bg=color, height=76, cursor="hand2")
            card.pack(side="left", fill="x", expand=True, padx=4)
            card.pack_propagate(False)
            tk.Label(card, text=icon, bg=color, fg=WHITE, font=("Segoe UI", 21, "bold")).pack(side="left", padx=(12, 4))
            inner = tk.Frame(card, bg=color)
            inner.pack(side="left", fill="both", expand=True)
            tk.Label(inner, text=lg, bg=color, fg=WHITE, font=("Segoe UI", 14, "bold"), anchor="w").pack(fill="x", pady=(10, 0))
            if sub:
                tk.Label(inner, text=sub, bg=color, fg=WHITE, font=("Segoe UI", 9, "bold"), anchor="w").pack(fill="x")
            self.dashboard_league_buttons[lg] = card
            for w in (card, inner):
                w.bind("<Button-1>", lambda e, x=lg: self._dashboard_select_league(x))
            for child in inner.winfo_children():
                child.bind("<Button-1>", lambda e, x=lg: self._dashboard_select_league(x))
        self._update_dashboard_selected_league()

        # ---------- bilgi kartları ----------
        # V59: Üst dört bilgi kartı Y ekseninde biraz büyütüldü; içerikler aynı kalır.
        info_cards = tk.Frame(content, bg=BG)
        info_cards.pack(fill="x", padx=14, pady=(0, 9))
        # Üst bilgi alanları ekranın tamamını kullanır; sağda boşluk bırakmaz.
        info_cards.grid_columnconfigure(0, weight=2, uniform="info")
        info_cards.grid_columnconfigure(1, weight=6, uniform="info")
        info_cards.grid_columnconfigure(2, weight=3, uniform="info")
        info_cards.grid_columnconfigure(3, weight=3, uniform="info")
        info_cards.grid_rowconfigure(0, weight=1)

        # ---------- bilgi kartları: Hazır + Mevcut + Event ----------
        # PvP'ye Hazır kartı biraz daraltıldı; böylece lig bazlı Mevcut
        # Pokémon özeti için yeni ve okunaklı bir alan açıldı.
        ready_card = tk.Frame(info_cards, bg=WHITE, width=235, height=158, highlightbackground=BORDER, highlightthickness=1)
        ready_card.grid(row=0, column=0, sticky="nsew", padx=(0, 3))
        ready_card.pack_propagate(False)
        ready_head = tk.Frame(ready_card, bg=WHITE)
        ready_head.pack(fill="x", padx=10, pady=(7, 2))
        tk.Label(ready_head, text="✓  PvP HAZIR", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 10, "bold"), anchor="w").pack(side="left", fill="x", expand=True)
        self.ready_pokemon_count_label = tk.Label(ready_head, text="PvP HAZIR: 0", bg=WHITE, fg=BLUE,
                 font=("Segoe UI", 8, "bold"), anchor="e")
        self.ready_pokemon_count_label.pack(side="right")
        ready_body = tk.Frame(ready_card, bg=WHITE)
        ready_body.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.ready_pokemon_box = ready_body
        self.ready_pokemon_canvas = tk.Canvas(ready_body, height=107, bg=WHITE, highlightthickness=0, bd=0)
        self.ready_pokemon_canvas.pack(side="left", fill="both", expand=True)
        self.ready_pokemon_vscroll = ttk.Scrollbar(ready_body, orient="vertical", command=self.ready_pokemon_canvas.yview)
        self.ready_pokemon_vscroll.pack(side="right", fill="y")
        self.ready_pokemon_canvas.configure(yscrollcommand=self.ready_pokemon_vscroll.set)
        self.ready_pokemon_inner = tk.Frame(self.ready_pokemon_canvas, bg=WHITE)
        self.ready_pokemon_window = self.ready_pokemon_canvas.create_window((0, 0), window=self.ready_pokemon_inner, anchor="nw")
        self.ready_pokemon_inner.bind("<Configure>", lambda e: self.ready_pokemon_canvas.configure(scrollregion=self.ready_pokemon_canvas.bbox("all")))
        self.ready_pokemon_canvas.bind("<Configure>", lambda e: self.ready_pokemon_canvas.itemconfigure(self.ready_pokemon_window, width=max(1, e.width-2)))
        # Hazır Pokémonlar artık dikey kaydırılabilir. Fare tekeri hem canvas
        # hem de kartların üzerinde çalışır.
        self.ready_pokemon_canvas.bind("<MouseWheel>", lambda e: self.ready_pokemon_canvas.yview_scroll(-int(e.delta/120), "units"))
        self.ready_pokemon_canvas.bind("<Button-4>", lambda e: self.ready_pokemon_canvas.yview_scroll(-1, "units"))
        self.ready_pokemon_canvas.bind("<Button-5>", lambda e: self.ready_pokemon_canvas.yview_scroll(1, "units"))
        # Kartların/etiketlerin üzerinde tekerlek kullanıldığında da hazır listesi kayar.
        self._ready_mouse_over = False
        ready_card.bind("<Enter>", lambda e: setattr(self, "_ready_mouse_over", True))
        ready_card.bind("<Leave>", lambda e: setattr(self, "_ready_mouse_over", False))
        if not getattr(self, "_ready_global_wheel_bound", False):
            def _ready_global_wheel(event):
                if getattr(self, "_ready_mouse_over", False):
                    delta = getattr(event, "delta", 0)
                    if delta:
                        self.ready_pokemon_canvas.yview_scroll(-int(delta / 120), "units")
                    return "break"
            self.root.bind("<MouseWheel>", _ready_global_wheel, add="+")
            self.root.bind("<Button-4>", lambda e: self.ready_pokemon_canvas.yview_scroll(-1, "units") if getattr(self, "_ready_mouse_over", False) else None, add="+")
            self.root.bind("<Button-5>", lambda e: self.ready_pokemon_canvas.yview_scroll(1, "units") if getattr(self, "_ready_mouse_over", False) else None, add="+")
            self._ready_global_wheel_bound = True
        self._refresh_ready_pokemon_panel()

        # ---------- Raid Pokémonları ----------
        # Raid kartı bilinçli olarak daha büyük tutulur; aynı anda yaklaşık 4 raid görünür.
        raid_card = tk.Frame(info_cards, bg=WHITE, width=760, height=158, highlightbackground=BORDER, highlightthickness=1)
        raid_card.grid(row=0, column=1, sticky="nsew", padx=(0, 3))
        raid_card.pack_propagate(False)
        raid_head = tk.Frame(raid_card, bg=WHITE)
        raid_head.pack(fill="x", padx=10, pady=(7, 2))
        tk.Label(raid_head, text="⚔  RAID Pokémonları", bg=WHITE, fg=TEXT, font=("Segoe UI", 10, "bold"), anchor="w").pack(side="left")
        tk.Label(raid_head, text="5 Yıldız + Mega", bg=WHITE, fg=MUTED, font=("Segoe UI", 8, "bold")).pack(side="left", padx=(10,0))
        tk.Button(raid_head, text="Kaynak", command=lambda: self.open_url(RAID_SOURCE_URL), bg=WHITE, fg=BLUE, relief="flat", bd=0, font=("Segoe UI", 8, "bold"), cursor="hand2").pack(side="right")

        raid_scroll_frame = tk.Frame(raid_card, bg=WHITE)
        raid_scroll_frame.pack(fill="both", expand=True, padx=6, pady=(0, 1))
        self.dashboard_raid_canvas = tk.Canvas(raid_scroll_frame, bg=WHITE, highlightthickness=0, bd=0)
        self.dashboard_raid_scrollbar = ttk.Scrollbar(raid_scroll_frame, orient="vertical", command=self.dashboard_raid_canvas.yview)
        self.dashboard_raid_canvas.pack(side="left", fill="both", expand=True)
        self.dashboard_raid_scrollbar.pack(side="right", fill="y")
        self.dashboard_raid_canvas.configure(yscrollcommand=self.dashboard_raid_scrollbar.set)
        self.dashboard_raid_body = tk.Frame(self.dashboard_raid_canvas, bg=WHITE)
        self.dashboard_raid_window = self.dashboard_raid_canvas.create_window((0, 0), window=self.dashboard_raid_body, anchor="nw")
        self.dashboard_raid_body.bind("<Configure>", lambda e: self.dashboard_raid_canvas.configure(scrollregion=self.dashboard_raid_canvas.bbox("all")))
        self.dashboard_raid_canvas.bind("<Configure>", lambda e: self.dashboard_raid_canvas.itemconfigure(self.dashboard_raid_window, width=e.width))
        self.dashboard_raid_photo_refs = []
        self.dashboard_raid_canvas.bind("<MouseWheel>", lambda e: self.dashboard_raid_canvas.yview_scroll(-int(e.delta/120), "units"))
        self.dashboard_raid_canvas.bind("<Button-4>", lambda e: self.dashboard_raid_canvas.yview_scroll(-1, "units"))
        self.dashboard_raid_canvas.bind("<Button-5>", lambda e: self.dashboard_raid_canvas.yview_scroll(1, "units"))
        self._raid_mouse_over = False
        raid_card.bind("<Enter>", lambda e: setattr(self, "_raid_mouse_over", True))
        raid_card.bind("<Leave>", lambda e: setattr(self, "_raid_mouse_over", False))
        if not getattr(self, "_raid_global_wheel_bound", False):
            def _raid_global_wheel(event):
                if getattr(self, "_raid_mouse_over", False):
                    delta = getattr(event, "delta", 0)
                    if delta:
                        self.dashboard_raid_canvas.yview_scroll(-int(delta / 120), "units")
                    return "break"
            self.root.bind("<MouseWheel>", _raid_global_wheel, add="+")
            self.root.bind("<Button-4>", lambda e: self.dashboard_raid_canvas.yview_scroll(-1, "units") if getattr(self, "_raid_mouse_over", False) else None, add="+")
            self.root.bind("<Button-5>", lambda e: self.dashboard_raid_canvas.yview_scroll(1, "units") if getattr(self, "_raid_mouse_over", False) else None, add="+")
            self._raid_global_wheel_bound = True
        tk.Label(raid_card, text="LeekDuck • canlı veri", bg=WHITE, fg="#8a98a8", font=("Segoe UI", 7)).pack(anchor="e", padx=8, pady=(0,2))
        self._refresh_dashboard_raids()

        # Mevcut Pokémonlar: tüm ligleri tek bakışta gösterir.
        owned_card = tk.Frame(info_cards, bg=WHITE, width=300, height=158, highlightbackground=BORDER, highlightthickness=1)
        owned_card.grid(row=0, column=2, sticky="nsew", padx=(0, 3))
        owned_card.pack_propagate(False)
        tk.Label(owned_card, text="▣  Mevcut Pokémonlar", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 9, "bold"), anchor="w").pack(fill="x", padx=8, pady=(5, 2))
        stats_frame = tk.Frame(owned_card, bg=WHITE)
        stats_frame.pack(fill="both", expand=True, padx=6, pady=(0, 4))
        self.dashboard_owned_stats = {}
        owned_specs = [
            ("Great League", "Great"),
            ("Ultra League", "Ultra"),
            ("Master League", "Master"),
            ("Mega", "Mega"),
            ("Gigamax", "Gigamax"),
        ]
        for league, short in owned_specs:
            row = tk.Frame(stats_frame, bg=WHITE)
            row.pack(fill="x", pady=0)
            if league == "Gigamax":
                tk.Frame(stats_frame, bg="#d9e2ef", height=1).pack(fill="x", pady=(2, 1))
            # Lig adı artık doğrudan butondur. Basınca sadece o ligde
            # Mevcut olarak tiklenmiş Pokémonlar açılır.
            league_btn = tk.Button(
                row, text=short, bg=WHITE, fg=BLUE, relief="flat", bd=0,
                font=("Segoe UI", 7, "bold"), width=6, anchor="w", cursor="hand2",
                activebackground=WHITE, activeforeground=BLUE,
                command=lambda lg=league: self._show_owned_pokemon_window(lg)
            )
            league_btn.pack(side="left")
            total = tk.Label(row, text="0", bg="#eef4fb", fg=TEXT, font=("Segoe UI", 9, "bold"), width=4)
            total.pack(side="left", padx=(3, 5))
            tk.Label(row, text="Mevcut", bg=WHITE, fg=MUTED, font=("Segoe UI", 7)).pack(side="left")
            ready = tk.Label(row, text="0", bg="#edf8ef", fg=GREEN, font=("Segoe UI", 9, "bold"), width=4)
            ready.pack(side="left", padx=(8, 3))
            tk.Label(row, text="PvP hazır", bg=WHITE, fg=MUTED, font=("Segoe UI", 7)).pack(side="left")
            # Küçük bilgi etiketi: butonun üzerine gelince tıklanabilir olduğu belli olsun.
            league_btn.bind("<Enter>", lambda e, b=league_btn: b.config(bg="#f2f6fb"))
            league_btn.bind("<Leave>", lambda e, b=league_btn: b.config(bg=WHITE))
            self.dashboard_owned_stats[league] = {"total": total, "ready": ready, "button": league_btn}
        self._refresh_dashboard_owned_stats()

        upcoming = tk.Frame(info_cards, bg=WHITE, width=250, height=158, highlightbackground=BORDER, highlightthickness=1)
        upcoming.grid(row=0, column=3, sticky="nsew")
        upcoming.pack_propagate(False)
        head = tk.Frame(upcoming, bg=WHITE)
        head.pack(fill="x", padx=10, pady=(7, 3))
        tk.Label(head, text="▣  Yakındaki Eventler", bg=WHITE, fg=TEXT, font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Button(head, text="Tümünü Gör", command=self.show_nearby_events, bg=WHITE, fg=BLUE, relief="flat", bd=0, font=("Segoe UI", 8, "bold"), cursor="hand2").pack(side="right")
        self.event_watch_label = tk.Label(upcoming, text="✓ Event takip ediliyor • Tüm ligler", bg=WHITE, fg=GREEN,
                                           font=("Segoe UI", 8, "bold"), anchor="w")
        self.event_watch_label.pack(fill="x", padx=10, pady=(0, 2))
        self.event_alert_frame = tk.Frame(upcoming, bg=WHITE)
        self.event_alert_frame.pack_forget()
        self.event_alert_image_label = None
        self.event_alert_text_label = None
        self.dashboard_upcoming = tk.Frame(upcoming, bg=WHITE)
        self.dashboard_upcoming.pack(fill="both", expand=True, padx=10, pady=(2, 8))
        self._dashboard_refresh_upcoming_events()

        # ---------- kontrol paneli ----------
        # Lig seçimi yalnızca üstteki Great / Ultra / Master / Mega / Gigamax
        # kartlarından yapılır. Alt bölümde ayrıca lig seçici bulunmaz.
        controls = tk.Frame(content, bg=WHITE, highlightbackground=BORDER, highlightthickness=1)
        controls.pack(fill="x", padx=14, pady=(0, 8))
        self.dashboard_controls = controls

        if not hasattr(self, "league_var"):
            self.league_var = tk.StringVar(value="Great League")
        if not hasattr(self, "count_var"):
            self.count_var = tk.StringVar(value="20")

        # İnce kontrol satırı: Gösterilecek Pokémon ilk sırada.
        # Sıralama seçicisi tamamen kaldırıldı; liste her zaman PvPoke sıralamasını kullanır.
        filter_row = tk.Frame(controls, bg=WHITE)
        filter_row.pack(fill="x", padx=10, pady=(6, 7))

        tk.Label(filter_row, text="Gösterilecek Pokémon", bg=WHITE, fg=TEXT,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=(2, 5))
        self.count_entry = ttk.Spinbox(
            filter_row, from_=1, to=100, width=6, textvariable=self.count_var,
            command=self.refresh
        )
        self.count_entry.pack(side="left", padx=(0, 12))
        self.count_entry.bind("<Return>", lambda e: self.refresh())
        self.count_entry.bind("<FocusOut>", lambda e: self.refresh())

        self.shadow_var = tk.BooleanVar(value=False)
        self.legendary_var = tk.BooleanVar(value=False)
        self.mythical_var = tk.BooleanVar(value=False)
        self.xl_var = tk.BooleanVar(value=False)
        self.special_var = tk.BooleanVar(value=False)
        self.original_rank_var = tk.BooleanVar(value=False)

        ttk.Checkbutton(filter_row, text="Shadow'ları göster",
                        variable=self.shadow_var).pack(side="left", padx=5)
        ttk.Checkbutton(filter_row, text="Efsanevi dahil et",
                        variable=self.legendary_var).pack(side="left", padx=5)
        ttk.Checkbutton(filter_row, text="Mistik dahil et",
                        variable=self.mythical_var).pack(side="left", padx=5)
        ttk.Checkbutton(filter_row, text="Özel güç (+) göster",
                        variable=self.special_var).pack(side="left", padx=5)
        ttk.Checkbutton(filter_row, text="Orijinal Rank",
                        variable=self.original_rank_var,
                        command=self.toggle_original_rank).pack(side="left", padx=5)
        ttk.Checkbutton(filter_row, text="XL gerektirmeyeni göster",
                        variable=self.xl_var).pack(side="left", padx=5)

        # Filtre kutuları değiştiğinde alt açıklama anında güncellensin.
        for _filter_var in (self.count_var, self.shadow_var, self.legendary_var, self.mythical_var, self.xl_var, self.special_var, self.original_rank_var):
            try:
                _filter_var.trace_add("write", lambda *_args: self._refresh_selected_filters_label())
            except Exception:
                pass

        # Kayıt/güncelleme butonları da aynı ince satırda, sağ tarafta.
        tk.Button(filter_row, text="↻ Güncel", command=self.reload, bg=WHITE, fg=TEXT,
                  relief="solid", bd=1, cursor="hand2", padx=8, pady=3).pack(side="right", padx=3)
        tk.Button(filter_row, text="▣ Excel", command=self.save_csv, bg=WHITE, fg=TEXT,
                  relief="solid", bd=1, cursor="hand2", padx=8, pady=3).pack(side="right", padx=3)
        tk.Button(filter_row, text="▣ Resim Kaydet (PNG)", command=self.save_a4_png,
                  bg=BLUE, fg=WHITE, relief="flat", bd=0,
                  font=("Segoe UI", 9, "bold"), padx=9, pady=4, cursor="hand2").pack(side="right", padx=3)

        # Manuel Pokémon arama da aynı satırda, ayrı bir Filtreler düğmesi olmadan.
        tk.Label(filter_row, text="Manuel Pokémon ara:", bg=WHITE, fg=MUTED,
                 font=("Segoe UI", 8)).pack(side="left", padx=(14, 4))
        self.manual_pokemon_var = tk.StringVar()
        self.manual_pokemon_frame = tk.Frame(filter_row, bg=WHITE)
        self.manual_pokemon_frame.pack(side="left")
        self.manual_pokemon_entry = ttk.Entry(
            self.manual_pokemon_frame, textvariable=self.manual_pokemon_var, width=18
        )
        self.manual_pokemon_entry.pack(side="left")
        self.manual_pokemon_show_button = ttk.Button(
            self.manual_pokemon_frame, text="Göster",
            command=self._manual_pokemon_show_selected
        )
        self.manual_pokemon_show_button.pack(side="left", padx=4)

        self.manual_pokemon_list = tk.Listbox(
            self.root, height=6, width=28, exportselection=False
        )
        self.manual_pokemon_list.place_forget()
        self.manual_pokemon_entry.bind("<KeyRelease>", self._manual_pokemon_filter)
        self.manual_pokemon_list.bind("<<ListboxSelect>>", self._manual_pokemon_list_select)

        # ---------- durum / bilgi / progress ----------
        info_line = tk.Frame(content, bg=BG)
        info_line.pack(fill="x", padx=14, pady=(0, 3))
        self.info_var = tk.StringVar(value="Veriler yükleniyor...")
        tk.Label(info_line, textvariable=self.info_var, bg=BG, fg=MUTED, font=("Segoe UI", 7)).pack(side="left")
        self.pokemon_count_label = tk.Label(info_line, text="Sıralanan Pokémon: 0 | Mevcut: 0 | Eksik: 0", bg=BG, fg=TEXT, font=("Segoe UI", 9, "bold"))
        self.pokemon_count_label.pack(side="right")
        self.loading_var = tk.StringVar(value="Veriler yükleniyor...")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(content, orient="horizontal", mode="determinate", maximum=100, variable=self.progress_var, style="Modern.Horizontal.TProgressbar")
        self.progress.pack(fill="x", padx=14, pady=(0, 5))

        # ---------- tablo ----------
        table_wrap = tk.Frame(content, bg=BG)
        table_wrap.pack(fill="both", expand=True, padx=14, pady=(0, 0))
        cols = ("rank", "original_rank", "image", "pokemon", "type", "evolution", "score", "iv", "level", "cp", "xl", "special", "move1", "move2", "move3", "current", "status")
        self.style = style
        self.style.configure("Treeview", rowheight=50, font=("Segoe UI", 10), background=WHITE, fieldbackground=WHITE, foreground=TEXT)
        self.style.configure("Treeview.Heading", background=NAVY, foreground=WHITE, font=("Segoe UI", 9, "bold"), padding=(5, 8))
        self.style.map("Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", TEXT)])
        self.tree = ttk.Treeview(table_wrap, columns=cols, show="headings")
        headings = {"rank":"#", "original_rank":"Rank", "image":"Görsel", "pokemon":"Pokémon", "type":"Tip", "evolution":"Önceki Evrimler", "score":"PvPoke", "iv":"Rank 1 IV", "level":"Lvl", "cp":"CP", "xl":"XL", "special":"Özel Güç", "move1":"Güç 1", "move2":"Güç 2", "move3":"Güç 3", "current":"Mevcut", "status":"Durum"}
        widths = {"rank":45,"original_rank":55,"image":60,"pokemon":125,"type":150,"evolution":260,"score":65,"iv":90,"level":55,"cp":55,"xl":45,"special":70,"move1":145,"move2":145,"move3":145,"current":62,"status":55}
        for c in cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], minwidth=35, stretch=False, anchor="w" if c in ("pokemon","type","evolution","move1","move2","move3") else "center")
        self.tree["displaycolumns"] = tuple(c for c in cols if c != "original_rank")
        self.overlay_widgets = []
        self._overlay_job = None
        self.tree.tag_configure("red", foreground="#d61f2c")
        self.tree.tag_configure("green", foreground="#008f58")
        self.tree.tag_configure("odd", background="#ffffff")
        self.tree.tag_configure("even", background="#f2f6fb")
        self.tree.tag_configure("red_even", foreground="#d61f2c", background="#f2f6fb")
        self.tree.tag_configure("red_odd", foreground="#d61f2c", background="#ffffff")
        self.tree.tag_configure("green_even", foreground="#008f58", background="#f2f6fb")
        self.tree.tag_configure("green_odd", foreground="#008f58", background="#ffffff")
        vs = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        table_wrap.rowconfigure(0, weight=1)
        table_wrap.columnconfigure(0, weight=1)
        self.tree_items_by_name = {}
        self.tree.bind("<Button-1>", self._handle_tree_click)
        self.tree.bind("<Configure>", lambda e: self.schedule_overlay_redraw())
        self.tree.bind("<Expose>", lambda e: self.schedule_overlay_redraw())
        self.tree.bind("<MouseWheel>", lambda e: self.schedule_overlay_redraw())
        self.tree.bind("<Button-4>", lambda e: self.schedule_overlay_redraw())
        self.tree.bind("<Button-5>", lambda e: self.schedule_overlay_redraw())

        # ---------- tablo altı bilgi ----------
        # Durum açıklamaları ve AKTİF LİGE AİT filtreler aynı satırda,
        # tablonun hemen altında ve küçük puntoda gösterilir.
        footer = tk.Frame(content, bg=WHITE, highlightbackground=BORDER, highlightthickness=1, height=34)
        footer.pack(fill="x", padx=14, pady=(5, 7))
        footer.pack_propagate(False)

        legend = tk.Frame(footer, bg=WHITE)
        legend.pack(side="left", padx=(10, 4), pady=5)
        for sym, txt, color in (("+", "Evrimleştirilebilir veya tüm güçlere sahip", GREEN),
                                ("−", "Evrim için event beklenmeli", GOLD),
                                ("X", "Event dışında evrimleştirilmiş — yeniden yakala", RED)):
            tk.Label(legend, text=sym, bg=WHITE, fg=color, font=("Segoe UI", 8, "bold")).pack(side="left")
            tk.Label(legend, text=txt, bg=WHITE, fg=TEXT, font=("Segoe UI", 7)).pack(side="left", padx=(2, 8))

        sep = tk.Label(footer, text="|", bg=WHITE, fg=BORDER, font=("Segoe UI", 8))
        sep.pack(side="left", padx=2)

        self.selected_filters_label = tk.Label(
            footer, text="", bg=WHITE, fg=MUTED,
            font=("Segoe UI", 7), anchor="w", justify="left"
        )
        self.selected_filters_label.pack(side="left", padx=(6, 8), pady=5, fill="x", expand=True)

        self._refresh_selected_filters_label()

        self.status_var = tk.StringVar(value="Hazır.")
        # Eski ikinci alt durum çubuğu kaldırıldı; status_var yalnızca
        # dahili durum mesajları için tutuluyor. Böylece tablo altında tek
        # bilgi satırı kalıyor.

        self._refresh_status_summary()

    def _refresh_selected_filters_label(self):
        """Tablonun altındaki küçük satırda yalnızca aktif ligin filtrelerini gösterir."""
        try:
            league = str(self.league_var.get() or "Great League")
            vals = dict(getattr(self, "league_filter_settings", {}).get(league, {}))
            try:
                count = max(1, min(100, int(vals.get("count", self.count_var.get()))))
            except Exception:
                count = 36
            parts = [f"{league}: İlk {count}"]
            if vals.get("shadow"): parts.append("Shadow")
            if vals.get("legendary"): parts.append("Efsanevi")
            if vals.get("mythical"): parts.append("Mistik")
            if vals.get("xl"): parts.append("XL gerektirmeyen")
            if vals.get("special"): parts.append("Özel güç (+)")
            if vals.get("original_rank"): parts.append("Orijinal Rank")
            if len(parts) == 1:
                parts.append("Standart filtreler")
            if hasattr(self, "selected_filters_label") and self.selected_filters_label.winfo_exists():
                self.selected_filters_label.config(text="Seçili filtreler: " + " • ".join(parts))
        except Exception:
            pass

    def _dashboard_focus(self, widget):
        try:
            widget.lift()
            self.root.after(50, lambda: widget.focus_set())
        except Exception:
            pass

    def _update_dashboard_selected_league(self):
        """Üst lig kartında seçili ligi altı çizili gösterir."""
        try:
            active = str(self.league_var.get() or "Great League")
            for lg, card in getattr(self, "dashboard_league_buttons", {}).items():
                color = getattr(card, "_sinka_base_color", None) or card.cget("bg")
                # Kart üzerindeki tüm metinlerde seçili lig için underline.
                for child in card.winfo_children():
                    if isinstance(child, tk.Label):
                        base_font = child.cget("font")
                        try:
                            f = tkfont.Font(font=base_font)
                            child.configure(font=(f.actual("family"), f.actual("size"), "bold", "underline") if lg == active else (f.actual("family"), f.actual("size"), "bold"))
                        except Exception:
                            pass
                    elif isinstance(child, tk.Frame):
                        for sub in child.winfo_children():
                            if isinstance(sub, tk.Label):
                                base_font = sub.cget("font")
                                try:
                                    f = tkfont.Font(font=base_font)
                                    sub.configure(font=(f.actual("family"), f.actual("size"), "bold", "underline") if lg == active else (f.actual("family"), f.actual("size"), "bold"))
                                except Exception:
                                    pass
                card.configure(relief="solid" if lg == active else "flat", bd=2 if lg == active else 0)
        except Exception:
            pass

    def _dashboard_select_league(self, league):
        try:
            if league in LEAGUES:
                self.active_search_league = league
                self.league_var.set(league)
                self.change_league()
                self._update_dashboard_selected_league()
        except Exception:
            pass

    def _dashboard_toggle_filters(self):
        try:
            panel = self.dashboard_filter_panel
            if panel.winfo_viewable():
                panel.pack_forget()
            else:
                panel.pack(fill="x", padx=14, pady=(0, 7), before=self.dashboard_content.winfo_children()[-1])
        except Exception:
            pass

    def _raid_source_date_text(self, text):
        """GO Hub tarih aralığını kısa ve okunabilir hale getirir."""
        t = re.sub(r"\s+", " ", str(text or "")).strip()
        return t

    def _parse_raid_date_range(self, text, now):
        """Pokémon GO Hub tarih aralıklarını güvenilir biçimde parse eder.
        Örnekler: September 9 – 15, August 26 – September 8, September 30 – October 6.
        """
        t = self._raid_source_date_text(text)
        months = {
            'january':1,'february':2,'march':3,'april':4,'may':5,'june':6,
            'july':7,'august':8,'september':9,'october':10,'november':11,'december':12
        }
        # Tarih tablosundaki HTML metninden yıl varsa al, yoksa mevcut yıl.
        year_m = re.search(r'\b(20\d{2})\b', t)
        year = int(year_m.group(1)) if year_m else now.year
        m = re.search(
            r'([A-Za-z]+)\s+(\d{1,2})\s*[–-]\s*'
            r'(?:(?:([A-Za-z]+)\s+)?(\d{1,2}))', t
        )
        if not m:
            return None, None
        sm, sd = m.group(1).lower(), int(m.group(2))
        em = (m.group(3) or m.group(1)).lower()
        ed = int(m.group(4))
        if sm not in months or em not in months:
            return None, None
        start = datetime(year, months[sm], sd, 0, 0)
        end_year = year
        # Örn. December 30 – January 5.
        if months[em] < months[sm]:
            end_year += 1
        end = datetime(end_year, months[em], ed, 23, 59, 59)
        return start, end

    def _raid_status_text(self, date_text, now):
        start, end = self._parse_raid_date_range(date_text, now)
        if not start:
            return "Devam ediyor"
        if start.date() <= now.date() <= end.date():
            return "Devam ediyor"
        diff = (start.date() - now.date()).days
        if diff == 1:
            return "1 gün kaldı"
        if diff > 1:
            return f"{diff} gün kaldı"
        return "Sona erdi"

    def _raid_find_content_after_heading(self, heading, wanted):
        """Başlığın altındaki ilk uygun UL/TABLE'ı, bir sonraki heading'e geçmeden bulur."""
        wanted = set(wanted)
        for node in heading.find_all_next():
            if node is heading:
                continue
            if getattr(node, 'name', None) in ('h2', 'h3'):
                break
            if getattr(node, 'name', None) in wanted:
                return node
        return None

    def _raid_img_from_node(self, node):
        if not node:
            return None
        img = node.find('img') if hasattr(node, 'find') else None
        if not img:
            return None
        for attr in ('src', 'data-src', 'data-lazy-src', 'data-original'):
            src = img.get(attr)
            if src:
                return urllib.parse.urljoin(RAID_SOURCE_URL, src)
        return None

    def _raid_clean_name(self, text):
        text = re.sub(r'\s+', ' ', str(text or '')).strip()
        # CP bilgilerini ve gereksiz parantez içi raid bilgisini temizle.
        text = re.sub(r'\s+\d{3,6}\s*CP\b.*$', '', text, flags=re.I).strip()
        text = re.sub(r'\s*\([^)]*CP[^)]*\)', '', text, flags=re.I).strip()
        return text.strip('•-–— ')

    def _raid_add_item(self, items, seen, name, date_text, group, upcoming, image=None, now=None):
        name = self._raid_clean_name(name)
        if not name:
            return
        key = (name.lower(), self._raid_source_date_text(date_text), group.lower())
        if key in seen:
            return
        seen.add(key)
        status = self._raid_status_text(date_text, now) if upcoming else "Devam ediyor"
        items.append({
            'name': name,
            'date': self._raid_source_date_text(date_text),
            'type': 'Tier 5 Raids' if group == '5 Yıldız' else 'Mega Raids',
            'raid_group': group,
            'status': status,
            'image': image,
            'upcoming': bool(upcoming),
        })

    def _fetch_raid_data(self):
        """RAID verisini, çalışan Event sistemindeki ScrapedDuck/LeekDuck
        events.json yapısını kullanarak çıkarır.

        Mantık özellikle Event sistemiyle aynıdır: tüm kategoriler dolaşılır,
        event başlığı/raids alanı anahtar kelimelerle incelenir. Böylece HTML
        yapısına bağımlı kalınmaz.
        """
        try:
            # SinKA'daki çalışan Event çekme fonksiyonunu doğrudan kullan.
            events = self._fetch_leekduck_events_live()
            if not isinstance(events, list):
                return [], "LeekDuck Event verisi beklenen formatta değil"

            now = datetime.now().astimezone()
            items = []
            seen = set()

            def norm(v):
                return re.sub(r"\s+", " ", str(v or "")).strip()

            def clean_name(name):
                n = norm(name)
                # LeekDuck bazı kayıtlarda bölge/CP bilgisini isimle birlikte tutabilir.
                n = re.sub(r"\s+\d{3,6}\s*CP\b.*$", "", n, flags=re.I).strip()
                n = re.sub(r"\s*\([^)]*CP[^)]*\)", "", n, flags=re.I).strip()
                return n.strip("•,-–—:")

            def parse_dt(v):
                if not v:
                    return None
                txt = norm(v)
                try:
                    dt = datetime.fromisoformat(txt.replace("Z", "+00:00"))
                    if dt.tzinfo is None:
                        dt = dt.replace(tzinfo=now.tzinfo)
                    return dt.astimezone(now.tzinfo)
                except Exception:
                    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
                        try:
                            return datetime.strptime(txt, fmt).replace(tzinfo=now.tzinfo)
                        except Exception:
                            pass
                return None

            def name_list(values):
                out = []
                if isinstance(values, list):
                    for x in values:
                        if isinstance(x, dict):
                            val = x.get("name") or x.get("title") or x.get("pokemon")
                        else:
                            val = x
                        if val:
                            val = clean_name(val)
                            if val and val.lower() not in {a.lower() for a in out}:
                                out.append(val)
                elif values:
                    val = clean_name(values)
                    if val:
                        out.append(val)
                return out

            for ev in events:
                if not isinstance(ev, dict):
                    continue

                title = norm(ev.get("title"))
                category = norm(ev.get("category"))
                low = (title + " " + category).lower()
                raw = ev.get("raw") if isinstance(ev.get("raw"), dict) else {}
                details = raw.get("details") if isinstance(raw.get("details"), dict) else {}

                # Event sistemindeki gerçek raids alanını kullan.
                raid_names = name_list(ev.get("raids"))
                if not raid_names:
                    raid_names = name_list(details.get("raids"))

                # Başlıkta raid anahtar kelimesi yoksa bile category/raids alanı
                # varsa kabul et. Shadow kayıtlarını kesinlikle alma.
                if "shadow" in low:
                    continue
                is_5 = any(k in low for k in (
                    "5-star", "5 star", "tier 5", "5-star raid", "5 star raid"
                ))
                is_mega = "mega raid" in low or "mega raids" in low
                is_raid_event = (
                    "raid battle" in low or "raid battles" in low or
                    "raid" in category.lower() or bool(raid_names)
                )

                if not is_raid_event:
                    continue

                # Event başlığından grup belirlenemiyorsa category/raw metninde ara.
                if is_mega:
                    group = "Mega"
                elif is_5:
                    group = "5 Yıldız"
                else:
                    # Sadece istenen iki grup. 1/3 yıldız ve belirsiz kayıtlar yok.
                    continue

                start_dt = parse_dt(ev.get("start_time"))
                end_dt = parse_dt(ev.get("end_time"))

                if end_dt and end_dt < now:
                    continue

                if start_dt and start_dt > now:
                    upcoming = True
                    days = max(1, (start_dt.date() - now.date()).days)
                    status = "1 gün kaldı" if days == 1 else f"{days} gün kaldı"
                else:
                    upcoming = False
                    status = "Devam ediyor"

                if start_dt and end_dt:
                    date_text = f"{start_dt.strftime('%d.%m.%Y')} - {end_dt.strftime('%d.%m.%Y')}"
                elif start_dt:
                    date_text = start_dt.strftime('%d.%m.%Y')
                else:
                    date_text = ""

                article_url = str(ev.get("url") or raw.get("article_url") or "")
                # Bazı event kayıtlarında raids boş olabilir. Başlıktan güvenli
                # şekilde Pokémon adını çıkarmayı son çare olarak kullan.
                if not raid_names:
                    stripped = re.sub(
                        r"\s+in\s+(5[- ]?star|tier\s*5|mega)\s+raid\s*battles?\b.*$",
                        "", title, flags=re.I
                    )
                    if stripped and len(stripped) < 100:
                        raid_names = [clean_name(stripped)]

                for name in raid_names:
                    key = (name.lower(), group, date_text)
                    if key in seen:
                        continue
                    seen.add(key)
                    items.append({
                        "name": name,
                        "date": date_text,
                        "type": "Tier 5 Raids" if group == "5 Yıldız" else "Mega Raids",
                        "raid_group": group,
                        "status": status,
                        "image": None,
                        "upcoming": upcoming,
                        "url": article_url,
                        "start_dt": start_dt,
                        "end_dt": end_dt,
                    })

            # Aktifler üstte; sonra en yakın gelecek kayıtlar.
            items.sort(key=lambda x: (
                1 if x.get("upcoming") else 0,
                x.get("start_dt") or now,
                0 if x.get("raid_group") == "5 Yıldız" else 1,
                x.get("name", "")
            ))

            # Görsel: LeekDuck raid-bosses sayfası + isim eşleşmesi.
            items = self._load_raid_images(items)
            return items[:30], None

        except Exception as e:
            return [], f"Raid verisi alınamadı: {type(e).__name__}: {e}"

    def _load_raid_images(self, items):
        """LeekDuck raid-bosses sayfasını isimle tarar; bulunamazsa PokeAPI
        official-artwork URL'sini fallback olarak kullanır. Görseller ana
        thread'de oluşturulmaz; sadece URL hazırlanır.
        """
        if not items:
            return items
        try:
            if BS4_OK:
                req = urllib.request.Request(
                    RAID_IMAGE_SOURCE_URL,
                    headers={'User-Agent':'Mozilla/5.0 SinKA-PvP'}
                )
                with urllib.request.urlopen(req, timeout=15) as r:
                    html = r.read().decode('utf-8', errors='ignore')
                soup = BeautifulSoup(html, 'html.parser')
                imgs = []
                for img in soup.find_all('img'):
                    alt = re.sub(r'\s+', ' ', img.get('alt','')).strip().lower()
                    src = img.get('src') or img.get('data-src') or img.get('data-lazy-src') or img.get('data-original')
                    if src and alt:
                        imgs.append((alt, urllib.parse.urljoin(RAID_IMAGE_SOURCE_URL, src)))
                for item in items:
                    target = re.sub(r'^(mega|shadow)\s+', '', item['name'].lower()).strip()
                    for alt, src in imgs:
                        if target == alt or target in alt or alt in target:
                            item['image'] = src
                            break
        except Exception:
            pass

        # Sağlam fallback: isimden PokeAPI ID çözümü için bilinen isimleri kullan.
        fallback_ids = {
            'regirock': 377, 'regice': 378, 'registeel': 379,
            'zacian': 888, 'zamazenta': 889, 'xurkitree': 796,
            'buzzwole': 794, 'pheromosa': 795, 'xerneas': 716,
            'beedrill': 15, 'gyarados': 130, 'houndoom': 229,
            'venusaur': 3, 'malamar': 687, 'victreebel': 71,
        }
        for item in items:
            if item.get('image'):
                continue
            name = re.sub(r'^(mega|shadow)\s+', '', item.get('name','').lower()).strip()
            # Bölge eklerini ve parantezleri kaldır.
            name = re.sub(r'\s*\([^)]*\)', '', name).strip()
            key = name.replace(' ', '-').replace('.', '')
            pid = fallback_ids.get(key)
            if pid:
                item['image'] = f'https://raw.githubusercontent.com/PokeAPI/sprites/master/sprites/pokemon/other/official-artwork/{pid}.png'
        return items

    def _refresh_dashboard_raids(self):
        """Raid kartını arka planda yeniler; UI ana thread'de güncellenir."""
        if not hasattr(self, 'dashboard_raid_body'):
            return
        def worker():
            items, err = self._fetch_raid_data()
            if not err and items:
                self._preload_raid_images(items)
            try:
                self.root.after(0, lambda: self._render_dashboard_raids(items, err))
            except Exception:
                pass
        threading.Thread(target=worker, daemon=True).start()

    def _preload_raid_images(self, items):
        """Raid görsellerini UI çizilmeden önce arka planda topluca indirir.
        Böylece kartlar ekrana geldikten sonra kare/kademeli render oluşmaz."""
        if not PIL_OK or not items:
            return
        try:
            from concurrent.futures import ThreadPoolExecutor, as_completed
            cache = getattr(self, "_raid_image_bytes_cache", None)
            if cache is None:
                cache = {}
                self._raid_image_bytes_cache = cache

            def fetch(item):
                url = item.get("image")
                if not url:
                    return item, None
                if url in cache:
                    return item, cache[url]
                try:
                    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 SinKA-PvP"})
                    with urllib.request.urlopen(req, timeout=8) as r:
                        data = r.read()
                    if data:
                        cache[url] = data
                    return item, data
                except Exception:
                    return item, None

            with ThreadPoolExecutor(max_workers=min(8, max(1, len(items)))) as ex:
                futures = [ex.submit(fetch, item) for item in items]
                for fut in as_completed(futures):
                    item, data = fut.result()
                    if data:
                        item["image_bytes"] = data
        except Exception:
            pass

    def _render_dashboard_raids(self, items, err=None):
        if not hasattr(self, 'dashboard_raid_body'):
            return
        for w in self.dashboard_raid_body.winfo_children():
            w.destroy()
        self.dashboard_raid_photo_refs=[]
        if err:
            tk.Label(self.dashboard_raid_body, text="Raid verisi alınamadı", bg=WHITE, fg=MUTED, font=("Segoe UI",8)).pack(anchor='w',padx=6,pady=8)
            return
        if not items:
            tk.Label(self.dashboard_raid_body, text="Şu anda 5 yıldız / Mega raid verisi bulunamadı.", bg=WHITE, fg=MUTED, font=("Segoe UI",8)).pack(anchor='w',padx=6,pady=8)
            return
        # Dört sütunlu, dikey kaydırılabilir RAID listesi. Görseller render
        # öncesinde topluca hazırlandığı için kaydırma sırasında yeniden indirme yoktur.
        grid = tk.Frame(self.dashboard_raid_body, bg=WHITE)
        grid.pack(fill='both', expand=True, padx=2, pady=1)
        for idx, item in enumerate(items):
            col = idx % 4
            row_no = idx // 4
            card = tk.Frame(grid, bg=WHITE, highlightbackground="#e7edf5", highlightthickness=1, width=158, height=54)
            card.grid(row=row_no, column=col, padx=3, pady=3, sticky='nsew')
            card.grid_propagate(False)
            grid.grid_columnconfigure(col, weight=1, uniform='raidcol')
            img_label=tk.Label(card,bg=WHITE,width=34,height=34)
            img_label.pack(side='left',padx=(3,3),pady=3)
            mid=tk.Frame(card,bg=WHITE)
            mid.pack(side='left',fill='both',expand=True,pady=4)
            group=item.get('raid_group','Raid')
            tk.Label(mid,text=group,bg=WHITE,fg=BLUE,font=("Segoe UI",6,"bold"),anchor='w').pack(fill='x')
            tk.Label(mid,text=item.get('name',''),bg=WHITE,fg=TEXT,font=("Segoe UI",8,"bold"),anchor='w').pack(fill='x')
            date_txt=item.get('date','')
            if date_txt:
                tk.Label(mid,text=date_txt,bg=WHITE,fg=MUTED,font=("Segoe UI",5),anchor='w').pack(fill='x')
            tk.Label(mid,text=item.get('status','Devam ediyor'),bg=WHITE,fg=GREEN if not item.get('upcoming') else BLUE,font=("Segoe UI",6,"bold"),anchor='w').pack(fill='x')
            data = item.get("image_bytes")
            if data and PIL_OK:
                try:
                    im = Image.open(BytesIO(data)).convert("RGBA")
                    im.thumbnail((34,34), Image.LANCZOS)
                    photo = ImageTk.PhotoImage(im)
                    img_label.config(image=photo)
                    img_label.image = photo
                    self.dashboard_raid_photo_refs.append(photo)
                except Exception:
                    img_label.config(text='⚔', fg=BLUE, font=("Segoe UI",18,"bold"))
            else:
                img_label.config(text='⚔',fg=BLUE,font=("Segoe UI",18,"bold"))

    def _dashboard_refresh_upcoming_events(self):
        """Sağdaki Yakındaki Eventler kartında event listesini yeniler.

        Event takip özeti ve eşleşme uyarısı aynı kartın üst kısmındadır;
        bu fonksiyon yalnızca normal yaklaşan event listesini yeniden çizer.
        """
        try:
            frame = getattr(self, "dashboard_upcoming", None)
            if frame is None:
                return
            for child in frame.winfo_children():
                child.destroy()
            events = getattr(self, "_cached_leekduck_events", None) or []
            shown = 0
            now = datetime.now(timezone.utc)
            for ev in events:
                if not isinstance(ev, dict):
                    continue
                start = self._event_watch_parse_datetime(ev.get("start_time"))
                if start and start <= now:
                    continue
                title = _tr_event_title(ev.get("title", "Etkinlik"))
                names = _event_pokemon_names(ev)
                nm = names[0] if names else "Event"
                dt = self._event_watch_date_text(start) if start else str(ev.get("date_info", [""])[0] if ev.get("date_info") else "")
                row = tk.Frame(frame, bg=WHITE)
                row.pack(fill="x", pady=1)
                top = tk.Frame(row, bg=WHITE)
                top.pack(fill="x")
                tk.Label(top, text=nm[:28], bg=WHITE, fg=TEXT, font=("Segoe UI", 8, "bold"), anchor="w").pack(side="left", fill="x", expand=True)
                tk.Label(top, text=dt[:30], bg=WHITE, fg=MUTED, font=("Segoe UI", 7), anchor="e").pack(side="right")
                tk.Label(row, text=title[:90], bg=WHITE, fg=MUTED, font=("Segoe UI", 6), anchor="w", justify="left", wraplength=285).pack(fill="x", padx=(0,2))
                shown += 1
                if shown >= 4:
                    break
            if not shown:
                tk.Label(frame, text="Yaklaşan event bilgisi bekleniyor…", bg=WHITE, fg=MUTED, font=("Segoe UI", 8)).pack(anchor="w", padx=5, pady=10)
        except Exception:
            pass

    def _format_leekduck_event_detail(self, event):
        """Yakındaki Eventler penceresinde event içeriğini okunabilir özetler."""
        if not isinstance(event, dict):
            return str(event)
        parts = []
        if event.get("date_info"):
            parts.append(" • ".join(event["date_info"][:3]))
        if event.get("pokemon"):
            # Gereksiz genel görselleri azalt; Pokémon adlarını koru.
            names = []
            for p in event["pokemon"]:
                if p not in names:
                    names.append(p)
            if names:
                parts.append("Pokémon: " + ", ".join(names[:18]))
        return "\n".join(parts)


    def _fetch_single_leekduck_event_detail(self, event_or_url):
        """Seçilen event için JSON'daki tüm yapılandırılmış bilgileri döndürür."""
        if isinstance(event_or_url, dict):
            return event_or_url
        return {
            "url": str(event_or_url),
            "pokemon": [],
            "features": [],
            "spawns": [],
            "raids": [],
            "eggs": [],
            "research": [],
            "shiny": [],
            "moves": [],
            "bonuses": [],
            "date_info": [],
            "headings": [],
            "text": "",
        }

    def _modernize_popup(self, win):
        """SinKA modern teması: açılan yardımcı pencereleri ana arayüzle uyumlu yapar."""
        bg = "#F4F7FB"
        navy = "#16233D"
        text_color = "#182230"
        muted = "#667085"
        card = "#FFFFFF"
        accent = "#2F6BDE"

        try:
            win.configure(bg=bg)
        except Exception:
            pass

        def walk(widget):
            try:
                cls = widget.winfo_class()
                if cls == "Frame":
                    # Başlık/özel koyu frame'leri bozma; diğerlerini açık tema yap.
                    try:
                        current = str(widget.cget("bg"))
                    except Exception:
                        current = ""
                    if current not in {navy, "#17233D", "#16233D"}:
                        widget.configure(bg=bg)
                elif cls == "Label":
                    try:
                        current = str(widget.cget("bg"))
                        if current not in {navy, "#17233D", "#16233D"}:
                            widget.configure(bg=card)
                    except Exception:
                        pass
                    try:
                        widget.configure(fg=text_color)
                    except Exception:
                        pass
                elif cls == "Button":
                    try:
                        widget.configure(
                            bg=navy, fg="white",
                            activebackground=accent,
                            activeforeground="white",
                            relief="flat", bd=0,
                            font=("Segoe UI", 10, "bold"),
                            cursor="hand2"
                        )
                    except Exception:
                        pass
                elif cls == "Entry":
                    try:
                        widget.configure(
                            bg="white", fg=text_color,
                            insertbackground=text_color,
                            relief="solid", bd=1,
                            font=("Segoe UI", 10)
                        )
                    except Exception:
                        pass
                elif cls == "Text":
                    try:
                        widget.configure(
                            bg="white", fg=text_color,
                            insertbackground=text_color,
                            selectbackground="#DCE7FA"
                        )
                    except Exception:
                        pass
                elif cls == "Canvas":
                    try:
                        widget.configure(bg=card)
                    except Exception:
                        pass
                elif cls == "Labelframe":
                    try:
                        widget.configure(bg=card, fg=text_color)
                    except Exception:
                        pass
                for child in widget.winfo_children():
                    walk(child)
            except Exception:
                pass

        walk(win)

        # Ortak yardımcı pencere görünümü.
        try:
            win.update_idletasks()
            win.lift()
        except Exception:
            pass

    def show_nearby_events(self):
        """Hızlı event listesi; detay sadece seçilen event için yüklenir."""
        import threading
        import webbrowser

        existing = getattr(self, "_nearby_event_window", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.deiconify()
                    existing.lift()
                    existing.focus_force()
                    return
            except Exception:
                self._nearby_event_window = None

        win = tk.Toplevel(self.root)
        self._nearby_event_window = win
        win.title("🔎 Yakındaki Eventler")
        win.geometry("1050x720")
        win.minsize(850, 580)
        win.transient(self.root)

        header = tk.Frame(win)
        header.pack(fill="x", padx=12, pady=(10, 6))

        tk.Label(
            header, text="🔎 Yakındaki Eventler",
            font=("Segoe UI", 16, "bold")
        ).pack(side="left")

        status = tk.Label(
            header, text="🔄 Event verileri alınıyor…",
            anchor="e",
            font=("Segoe UI", 12, "bold")
        )
        status.pack(side="right")

        # Event verileri ağdan alınırken kullanıcı beklediğini anlasın.
        loading_bar = ttk.Progressbar(header, mode="indeterminate", length=260)
        loading_bar.pack(side="right", padx=(12, 8))
        loading_bar.start(10)

        def stop_event_loading(final_text=None):
            # Eventler tamamen geldikten sonra animasyonu kesin olarak durdur.
            try:
                loading_bar.stop()
                loading_bar.pack_forget()
            except Exception:
                pass
            if final_text is not None:
                try:
                    status.config(text=final_text)
                except Exception:
                    pass

        main = tk.Frame(win)
        main.pack(fill="both", expand=True, padx=10, pady=5)

        left = tk.Frame(main, bd=1, relief="solid")
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))

        right = tk.Frame(main, bd=1, relief="solid", width=390)
        right.pack(side="right", fill="both", padx=(6, 0))
        right.pack_propagate(False)

        list_canvas = tk.Canvas(left, highlightthickness=0)
        list_scroll = tk.Scrollbar(left, orient="vertical", command=list_canvas.yview)
        list_content = tk.Frame(list_canvas)
        list_window = list_canvas.create_window((0, 0), window=list_content, anchor="nw")
        list_canvas.configure(yscrollcommand=list_scroll.set)

        list_content.bind(
            "<Configure>",
            lambda e: list_canvas.configure(scrollregion=list_canvas.bbox("all"))
        )
        list_canvas.bind(
            "<Configure>",
            lambda e: list_canvas.itemconfigure(list_window, width=e.width)
        )
        list_canvas.pack(side="left", fill="both", expand=True)
        list_scroll.pack(side="right", fill="y")

        detail_title = tk.Label(
            right, text="Bir etkinlik seçin",
            font=("Segoe UI", 14, "bold"),
            wraplength=350, justify="left", anchor="w"
        )
        detail_title.pack(fill="x", padx=12, pady=(14, 8))

        # Event başlangıç/bitiş geri sayımı: devam eden yeşil, başlamamış kırmızı.
        detail_countdown = tk.Label(
            right, text="", font=("Segoe UI", 20, "bold"),
            justify="center", anchor="center", padx=10, pady=10,
            relief="solid", bd=2
        )
        detail_countdown.pack(fill="x", padx=12, pady=(0, 8))

        detail_status = tk.Label(
            right, text="Detayları görmek için soldaki etkinliğe tıklayın.",
            justify="left", anchor="nw", wraplength=350
        )
        detail_status.pack(fill="x", padx=12, pady=5)

        # Özel vuruşu sağ panelin en üstünde büyük ve belirgin göster.
        special_move_frame = tk.Frame(right, bd=2, relief="solid", padx=10, pady=8)
        special_move_label = tk.Label(
            special_move_frame, text="⚡ ÖZEL VURUŞ",
            font=("Segoe UI", 12, "bold"), anchor="w"
        )
        special_move_label.pack(fill="x")
        special_move_value = tk.Label(
            special_move_frame, text="",
            font=("Segoe UI", 18, "bold"), anchor="w", justify="left", wraplength=350
        )
        special_move_value.pack(fill="x", pady=(3, 0))

        special_pokemon_frame = ttk.Frame(right)
        special_pokemon_title = ttk.Label(
            special_pokemon_frame, text="🎯 BU HAREKETİ ALABİLEN POKÉMONLAR",
            font=("Segoe UI", 10, "bold")
        )
        special_pokemon_title.pack(fill="x", pady=(0, 3))
        special_pokemon_grid = ttk.Frame(special_pokemon_frame)
        special_pokemon_grid.pack(fill="x")
        special_pokemon_refs = []

        detail_text = tk.Text(
            right, wrap="word", state="disabled",
            font=("Segoe UI", 10)
        )
        # Sağ tarafta event Pokémon görselleri. Ana SinKa görsel/cache
        # altyapısı kullanılır; burada yeni ağ isteği yapılmaz.
        detail_images_frame = ttk.Frame(right)
        detail_images_frame.pack(fill="x", padx=12, pady=(0, 4))
        detail_image_refs = []

        detail_text.pack(fill="both", expand=True, padx=12, pady=8)

        def _event_detail_names(d):
            names = []
            groups = (d.get("pokemon", []), d.get("features", []), d.get("spawns", []),
                      d.get("raids", []), d.get("research", []),
                      d.get("eggs", []), d.get("shiny", []), d.get("moves", []))
            for group in groups:
                for value in (group or []):
                    text = str(value).strip()
                    # JSON move/feature alanları bazen sözlük olabilir.
                    if isinstance(value, dict):
                        text = str(value.get("name") or value.get("pokemon") or "").strip()
                    # Basit metinlerden bilinen Pokémon isimlerini ana veri/cache
                    # üzerinden yakala. Böylece isim yazımı ne olursa olsun mevcut
                    # SinKa görseli kullanılabilir.
                    for key in (getattr(self, "image_cache", {}) or {}).keys():
                        if key and re.search(r"(?<![A-Za-z])" + re.escape(str(key)) + r"(?![A-Za-z])", text, re.I):
                            if key not in names:
                                names.append(key)
                    for key in (getattr(self, "image_bytes", {}) or {}).keys():
                        if key and re.search(r"(?<![A-Za-z])" + re.escape(str(key)) + r"(?![A-Za-z])", text, re.I):
                            if key not in names:
                                names.append(key)
            # Ana PvP tablosunda bulunmayan event Pokémonları da göster.
            # Bu isimler yalnızca POGO_EVENT_IMAGE_IDS içindeki ortak event görsel
            # önbelleğini kullanır; aynı Pokémon için ikinci bir ağ isteği yapılmaz.
            combined_text = " ".join(str(d.get(k, "")) for k in ("title", "text", "featured_attack_text"))
            combined_text += " " + " ".join(str(x) for k in ("pokemon", "features", "spawns", "raids", "shiny") for x in (d.get(k, []) or []))
            for key in POGO_EVENT_IMAGE_IDS:
                if key not in names and re.search(r"(?<![A-Za-z])" + re.escape(key) + r"(?![A-Za-z])", combined_text, re.I):
                    names.append(key)
            # Mega/Gigantamax adlarında temel tür adını 추출.
            for n in list(names):
                base = re.sub(r"^(?:Mega|G-Max|Gigantamax)\s+", "", str(n), flags=re.I).strip()
                if base in POGO_EVENT_IMAGE_IDS and base not in names:
                    names.append(base)
            return names[:12]

        def _show_event_detail_images(d):
            detail_image_refs.clear()
            for child in detail_images_frame.winfo_children():
                child.destroy()

            names = _event_detail_names(d)
            if not names:
                return

            # Aynı Pokémon için tekrar kart oluşturma.
            shown = set()
            for name in names:
                if name in shown:
                    continue
                shown.add(name)
                photo = None
                clean_name = re.sub(r"^(?:Mega|G-Max|Gigantamax)\s+", "", str(name), flags=re.I).strip()
                candidates = [str(name), str(name).lower(), clean_name, clean_name.lower()]

                # 1) Ana tablodaki PhotoImage cache'i.
                for key in candidates:
                    obj = self.image_cache.get(key) if isinstance(self.image_cache, dict) else None
                    if obj is not None:
                        photo = obj
                        break

                # 2) Daha önce indirilmiş dosya.
                if photo is None and PIL_OK and isinstance(self.image_files, dict):
                    for key in candidates:
                        fp = self.image_files.get(key)
                        if fp and os.path.exists(str(fp)):
                            try:
                                img = Image.open(str(fp)).convert("RGBA")
                                img.thumbnail((70, 70), Image.LANCZOS)
                                photo = ImageTk.PhotoImage(img)
                                detail_image_refs.append(photo)
                                break
                            except Exception:
                                pass

                # 3) Daha önce indirilmiş raw bytes. Yine ağ isteği yok.
                if photo is None and PIL_OK and isinstance(self.image_bytes, dict):
                    raw = None
                    for key in candidates:
                        raw = self.image_bytes.get(key)
                        if raw:
                            break
                    if raw:
                        try:
                            img = Image.open(BytesIO(raw)).convert("RGBA")
                            img.thumbnail((70, 70), Image.LANCZOS)
                            photo = ImageTk.PhotoImage(img)
                            detail_image_refs.append(photo)
                        except Exception:
                            pass

                # 4) Event görsel önbelleği: Gible/Staraptor gibi ana PvP tablosunda
                # o anda olmayan Pokémonlar için tek seferlik ortak cache kullan.
                if photo is None and clean_name in POGO_EVENT_IMAGE_IDS:
                    try:
                        photo = _event_load_image(clean_name, size=(82, 82))
                    except Exception:
                        photo = None

                if photo is not None:
                    # PhotoImage'in yaşamını detay penceresi açık kaldığı sürece koru.
                    detail_image_refs.append(photo)
                    item = ttk.Frame(detail_images_frame)
                    item.pack(side="left", padx=(0, 10))
                    lbl = ttk.Label(item, image=photo)
                    lbl.pack()
                    ttk.Label(item, text=str(name), font=("Segoe UI", 9, "bold")).pack()

        def set_detail_text(value):
            detail_text.config(state="normal")
            detail_text.delete("1.0", "end")
            detail_text.insert("1.0", value)
            detail_text.config(state="disabled")

        def _special_move_pokemon_names(d):
            # Yalnızca özel hareketi gerçekten öğrenen/alan Pokémonları çıkar.
            names = []
            raw = " ".join([
                str(d.get("featured_attack_text", "")),
                str(d.get("text", "")),
                " ".join(str(x) for x in (d.get("moves", []) or [])),
                " ".join(str(x) for x in (d.get("pokemon", []) or [])),
            ])
            low = raw.lower()

            # Mega etkinliklerde hareket doğrudan Mega Pokémon adına bağlıdır.
            for x in (d.get("moves", []) or []):
                xname = x.get("name") if isinstance(x, dict) else str(x)
                if xname and str(xname).lower() not in low:
                    continue
                if str(xname) and str(xname) not in names:
                    names.append(str(xname))

            # Metindeki tipik 'get a Garchomp that knows...' /
            # 'evolve ... to get a Garchomp' kalıpları.
            patterns = [
                r"get (?:a|an) ([A-Z][A-Za-z'’\- ]{2,30}) (?:that )?(?:knows|learns)",
                r"get (?:a|an) ([A-Z][A-Za-z'’\- ]{2,30}) (?:with|knows)",
                r"to get (?:a|an) ([A-Z][A-Za-z'’\- ]{2,30}) (?:that )?(?:knows|learns)",
                r"([A-Z][A-Za-z'’\- ]{2,30}) (?:that )?(?:knows|learns) (?:the )?(?:Charged Attack|Fast Attack)",
            ]
            for pat in patterns:
                for m in re.finditer(pat, raw, re.I):
                    n = re.sub(r"\s+", " ", m.group(1)).strip(" .,;:")
                    if n and n not in names and n.lower() not in {"the", "charged attack", "fast attack"}:
                        names.append(n)

            # 'Mega Evolved' özel hareketi olan eventlerde hedef genellikle başlıktaki Pokémon'dur.
            if "mega evolved" in low or "mega evolve" in low:
                title = str(d.get("title", ""))
                for candidate in POGO_EVENT_IMAGE_IDS:
                    if re.search(r"(?<![A-Za-z])" + re.escape(candidate) + r"(?![A-Za-z])", title, re.I):
                        if candidate not in names:
                            names.append(candidate)

            # Event veri alanındaki evrim zincirinden son evrimi seç; örn. Gible → Gabite → Garchomp.
            if not names:
                for value in (d.get("pokemon", []) or []):
                    text = value.get("name") if isinstance(value, dict) else str(value)
                    parts = [x.strip() for x in re.split(r"→|->", text or "") if x.strip()]
                    if len(parts) > 1:
                        names.append(parts[-1])

            # Görsel sisteminin tanıdığı isimlere normalize et.
            out=[]
            for n in names:
                base = re.sub(r"^(?:Mega|G-Max|Gigantamax)\s+", "", str(n), flags=re.I).strip()
                for cand in (str(n).strip(), base):
                    if cand and cand in POGO_EVENT_IMAGE_IDS and cand not in out:
                        out.append(cand)
            return out[:20]

        def _show_special_pokemon(d):
            special_pokemon_refs.clear()
            for child in special_pokemon_grid.winfo_children():
                child.destroy()
            names = _special_move_pokemon_names(d)
            if not names:
                special_pokemon_frame.pack_forget()
                return
            special_pokemon_frame.pack(fill="x", padx=12, pady=(0, 7), before=detail_images_frame)
            for i, name in enumerate(names):
                photo = None
                clean = re.sub(r"^(?:Mega|G-Max|Gigantamax)\s+", "", str(name), flags=re.I).strip()
                candidates = [str(name), str(name).lower(), clean, clean.lower()]
                for key in candidates:
                    obj = self.image_cache.get(key) if isinstance(self.image_cache, dict) else None
                    if obj is not None:
                        photo = obj; break
                if photo is None and PIL_OK and isinstance(self.image_files, dict):
                    for key in candidates:
                        fp = self.image_files.get(key)
                        if fp and os.path.exists(str(fp)):
                            try:
                                img=Image.open(str(fp)).convert("RGBA"); img.thumbnail((72,72), Image.LANCZOS)
                                photo=ImageTk.PhotoImage(img); special_pokemon_refs.append(photo); break
                            except Exception: pass
                if photo is None and PIL_OK and isinstance(self.image_bytes, dict):
                    for key in candidates:
                        rawb=self.image_bytes.get(key)
                        if rawb:
                            try:
                                img=Image.open(BytesIO(rawb)).convert("RGBA"); img.thumbnail((72,72), Image.LANCZOS)
                                photo=ImageTk.PhotoImage(img); special_pokemon_refs.append(photo); break
                            except Exception: pass
                if photo is None and clean in POGO_EVENT_IMAGE_IDS:
                    try:
                        photo=_event_load_image(clean, size=(82,82))
                    except Exception: photo=None
                if photo is not None:
                    special_pokemon_refs.append(photo)
                    card=ttk.Frame(special_pokemon_grid, padding=2)
                    card.grid(row=i//4, column=i%4, padx=4, pady=2, sticky="w")
                    ttk.Label(card, image=photo).pack()
                    ttk.Label(card, text=clean, font=("Segoe UI",9,"bold")).pack()

        def _parse_event_dt(value):
            """ScrapedDuck ISO tarihini timezone-aware datetime'a çevirir."""
            if not value:
                return None
            try:
                text = str(value).strip()
                if text.endswith("Z"):
                    text = text[:-1] + "+00:00"
                dt = datetime.fromisoformat(text)
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=timezone.utc)
                return dt
            except Exception:
                return None

        def _event_countdown(event):
            """Event durumunu ve kullanıcıya gösterilecek büyük geri sayımı döndürür."""
            start = _parse_event_dt(event.get("start_time"))
            end = _parse_event_dt(event.get("end_time"))
            # JSON'dan gelmeyen durumlarda date_info içindeki iki ISO tarihi dene.
            if start is None or end is None:
                di = event.get("date_info", []) or []
                parsed = [_parse_event_dt(x) for x in di]
                parsed = [x for x in parsed if x is not None]
                if start is None and parsed:
                    start = parsed[0]
                if end is None and len(parsed) > 1:
                    end = parsed[1]
            now = datetime.now(timezone.utc)

            def fmt(delta):
                sec = max(0, int(delta.total_seconds()))
                days, sec = divmod(sec, 86400)
                hours, sec = divmod(sec, 3600)
                minutes, seconds = divmod(sec, 60)
                parts = []
                if days:
                    parts.append(f"{days} gün")
                if hours or days:
                    parts.append(f"{hours} saat")
                if minutes or hours or days:
                    parts.append(f"{minutes} dk")
                parts.append(f"{seconds} sn")
                return " ".join(parts)

            if start and now < start:
                return ("not_started", "🔴 BAŞLAMADI", f"Başlamasına {fmt(start-now)} kaldı")
            if end and now < end:
                return ("active", "🟢 DEVAM EDİYOR", f"Bitmesine {fmt(end-now)} kaldı")
            if end and now >= end:
                return ("ended", "⚪ SÜRESİ DOLDU", "Bu etkinlik sona erdi")
            return ("unknown", "⚫ ZAMAN BİLGİSİ YOK", "Etkinlik zamanı alınamadı")

        def _update_countdown_label(label, event):
            try:
                if not label.winfo_exists():
                    return
                state, title_txt, remaining = _event_countdown(event)
                if state == "active":
                    bg, fg = "#dff6e7", "#008000"
                elif state == "not_started":
                    bg, fg = "#ffe4e4", "#c00000"
                elif state == "ended":
                    bg, fg = "#eeeeee", "#666666"
                else:
                    bg, fg = "#eeeeee", "#333333"
                label.config(text=f"{title_txt}\n{remaining}", bg=bg, fg=fg)
            except Exception:
                pass

        countdown_labels = []
        countdown_job = [None]

        def _refresh_countdowns():
            if not win.winfo_exists():
                return
            for label, event in list(countdown_labels):
                _update_countdown_label(label, event)
            countdown_job[0] = win.after(1000, _refresh_countdowns)

        def _clear_countdowns():
            try:
                if countdown_job[0]:
                    win.after_cancel(countdown_job[0])
            except Exception:
                pass
            countdown_job[0] = None
            countdown_labels.clear()

        def show_detail(event):
            if not isinstance(event, dict):
                return
            title = event.get("title", "Etkinlik")
            url = event.get("url", "")

            detail_title.config(text=_tr_event_title(title))
            _update_countdown_label(detail_countdown, event)
            detail_status.config(text="⏳ Etkinlik ayrıntıları alınıyor...")
            special_move_value.config(text="Yükleniyor…")
            special_move_frame.pack(fill="x", padx=12, pady=(0, 6), before=detail_images_frame)
            set_detail_text("Lütfen bekleyin...")

            def worker():
                try:
                    d = self._fetch_single_leekduck_event_detail(event)
                    lines = []

                    if d.get("date_info"):
                        lines.append("📅 TARİH / SAAT\n" + "\n".join(_tr_date(x) for x in d["date_info"][:10]))

                    if d.get("category"):
                        lines.append("🏷️ KATEGORİ\n" + _tr_category(d["category"]))

                    section_map = [
                        ("features", "⭐ ÖNE ÇIKAN POKÉMONLAR"),
                        ("spawns", "🌿 ETKİNLİKTE ÇIKAN POKÉMONLAR"),
                        ("raids", "⚔️ BASKIN POKÉMONLARI"),
                        ("research", "🔬 ARAŞTIRMA KARŞILAŞMALARI"),
                        ("eggs", "🥚 YUMURTALAR"),
                        ("moves", "⚡ ÖZEL HAREKET"),
                        ("shiny", "✨ PARLAK (SHINY)"),
                    ]

                    for key, label in section_map:
                        vals = list(dict.fromkeys(d.get(key, []) or []))
                        if vals:
                            lines.append(label + "\n" + ", ".join(_tr_common(v) for v in vals[:40]))

                    if d.get("bonuses"):
                        bonuses = list(dict.fromkeys(d["bonuses"]))
                        lines.append("🎁 BONUSLAR / GÜÇLER\n" + "\n".join("• " + _tr_common(b) for b in bonuses[:20]))

                    if d.get("featured_attack_text"):
                        lines.append(
                            "⚡ ÖZEL HAREKET DETAYI\n" +
                            _tr_common(str(d["featured_attack_text"]))
                        )
                    elif d.get("text"):
                        body = str(d["text"])
                        low = body.lower()
                        attack_terms = (
                            "featured attack", "special move",
                            "charged attack", "fast attack",
                            "exclusive attack"
                        )
                        pos = [low.find(t) for t in attack_terms if low.find(t) >= 0]
                        if pos:
                            p0 = min(pos)
                            snippet = body[p0:p0 + 900]
                            lines.append(
                                "⚡ ÖZEL HAREKET DETAYI\n" +
                                "Etkinlikte belirtilen Pokémon evrimleştirilerek "
                                "özel hareketi öğrenebilir.\n\n" +
                                _tr_common(snippet)
                            )
                        else:
                            lines.append("ℹ️ AÇIKLAMA\n" + _tr_common(body[:1200]))

                    if not lines:
                        lines.append("LeekDuck sayfası açılabilir; bu event için okunabilir detay bulunamadı.")

                    result = "\n\n".join(lines)

                    def apply():
                        detail_status.config(text="✅ Ayrıntılar hazır")
                        _show_special_pokemon(d)
                        _show_event_detail_images(d)
                        actual_move = str(d.get("actual_move") or "").strip()
                        if not actual_move:
                            # Geri dönüş: metindeki bilinen hareket adlarını ara.
                            body = str(d.get("featured_attack_text") or d.get("text") or "")
                            known_moves = (
                                "Earth Power", "Fell Stinger+", "Dark Pulse+",
                                "Hydro Cannon", "Blast Burn", "Frenzy Plant",
                                "Meteor Mash", "Psychic", "Rock Slide"
                            )
                            for mv in known_moves:
                                if mv.lower() in body.lower():
                                    actual_move = mv
                                    break
                        if not actual_move:
                            moves0 = list(dict.fromkeys(d.get("moves", []) or []))
                            actual_move = " / ".join(moves0[:10]) if moves0 else "Özel hareket ayrıntısı bulundu"
                        special_move_value.config(text=MOVE_TR.get(actual_move, _tr_common(actual_move)))
                        special_move_frame.pack(fill="x", padx=12, pady=(0, 6), before=detail_images_frame)
                        set_detail_text(result)

                    self.root.after(0, apply)

                except Exception as exc:
                    self.root.after(
                        0,
                        lambda: detail_status.config(
                            text=f"❌ Ayrıntılar alınamadı: {type(exc).__name__}"
                        )
                    )

            threading.Thread(target=worker, daemon=True).start()

            # Kullanıcı isterse kaynağı doğrudan tarayıcıda açabilir.
            if url:
                open_btn = getattr(show_detail, "_open_btn", None)
                if open_btn is None:
                    pass

        def render(events):
            if not win.winfo_exists():
                return

            for w in list_content.winfo_children():
                w.destroy()

            # SADECE özel hareket içeren etkinlikler.
            # Bazı LeekDuck kayıtlarında hareket "moves" alanında değil,
            # "Featured Attack" açıklaması içinde bulunuyor.
            move_events = []
            attack_words = (
                "featured attack", "special move", "charged attack",
                "fast attack", "exclusive attack", "knows the "
            )
            for event in events or []:
                if not isinstance(event, dict):
                    continue
                moves = event.get("moves", []) or []
                raw_text = str(event.get("text", "") or "")
                raw = str(event.get("raw", "") or "")
                featured = str(event.get("featured_attack_text", "") or "")
                searchable = (raw_text + " " + raw + " " + featured).lower()
                has_attack_text = any(w in searchable for w in attack_words)
                if moves or has_attack_text:
                    move_events.append(event)

            _clear_countdowns()
            if not move_events:
                tk.Label(
                    list_content,
                    text="Şu anda özel hareket içeren yaklaşan etkinlik bulunamadı.",
                    font=("Segoe UI", 12)
                ).pack(pady=30)
                stop_event_loading("Özel hareketli event yok")
                return

            status.config(
                text=f"{len(move_events)} özel hareketli etkinlik • ScrapedDuck / LeekDuck"
            )

            for event in move_events:
                title = _tr_event_title(event.get("title", "Etkinlik"))
                url = event.get("url", "")
                moves = list(dict.fromkeys(event.get("moves", []) or []))
                pokemon = list(dict.fromkeys(event.get("pokemon", []) or []))
                date_info = list(dict.fromkeys(event.get("date_info", []) or []))
                move_display = ", ".join(moves)
                if not move_display:
                    raw_text = str(event.get("featured_attack_text") or event.get("text", "") or "")
                    low = raw_text.lower()
                    positions = [p for p in (
                        low.find("featured attack"),
                        low.find("special move"),
                        low.find("charged attack"),
                        low.find("fast attack"),
                    ) if p >= 0]
                    if positions:
                        p0 = min(positions)
                        move_display = raw_text[p0:p0 + 500].replace("\n", " ")

                row = tk.Frame(list_content, bd=1, relief="solid")
                row.pack(fill="x", padx=5, pady=5)

                # Büyük ve canlı geri sayım: event devam ediyorsa yeşil, başlamadıysa kırmızı.
                countdown_label = tk.Label(
                    row, text="", font=("Segoe UI", 15, "bold"),
                    justify="center", anchor="center", padx=8, pady=6, bd=1, relief="solid"
                )
                countdown_label.pack(fill="x", padx=5, pady=(5, 4))
                countdown_labels.append((countdown_label, event))
                _update_countdown_label(countdown_label, event)

                # Event adı
                tk.Button(
                    row,
                    text="📅 " + title,
                    anchor="w",
                    justify="left",
                    font=("Segoe UI", 10, "bold"),
                    command=lambda e=event: show_detail(e)
                ).pack(fill="x", padx=5, pady=(5, 2))

                # Kullanıcının asıl istediği bilgi: Pokémon + özel hareket.
                tk.Label(
                    row,
                    text="⚡ ÖZEL HAREKET: " + move_display,
                    anchor="w",
                    justify="left",
                    font=("Segoe UI", 10, "bold")
                ).pack(fill="x", padx=10, pady=2)

                if pokemon:
                    tk.Label(
                        row,
                        text="🐾 Pokémon: " + ", ".join(_tr_common(p) for p in pokemon[:20]),
                        anchor="w",
                        justify="left"
                    ).pack(fill="x", padx=10, pady=2)

                if date_info:
                    tk.Label(
                        row,
                        text="📅 " + " → ".join(date_info[:2]),
                        anchor="w",
                        justify="left"
                    ).pack(fill="x", padx=10, pady=2)

                if url:
                    tk.Button(
                        row,
                        text="🌐 LeekDuck'ta aç",
                        command=lambda u=url: webbrowser.open(u)
                    ).pack(anchor="e", padx=5, pady=(2, 5))

            list_canvas.yview_moveto(0)
            if countdown_job[0] is None:
                countdown_job[0] = win.after(1000, _refresh_countdowns)
            stop_event_loading(
                f"{len(move_events)} özel hareketli etkinlik • ScrapedDuck / LeekDuck"
            )

        def worker():
            try:
                self.root.after(0, lambda: status.config(text="🔄 Event verileri alınıyor…"))
                events = self._fetch_leekduck_events_live()
                self.root.after(0, lambda: render(events))
            except Exception as exc:
                self.root.after(
                    0,
                    lambda: (loading_bar.stop(), loading_bar.pack_forget(), status.config(
                        text=f"❌ Event verileri alınamadı: {type(exc).__name__}"
                    ))
                )

        threading.Thread(target=worker, daemon=True).start()

        def close():
            _clear_countdowns()
            try:
                win.destroy()
            finally:
                self._nearby_event_window = None

        win.protocol("WM_DELETE_WINDOW", close)

    def _fetch_leekduck_events_live(self):
        """ScrapedDuck events.json'u doğru şemasıyla okuyup Pokémon,
        raid, spawn, research, shiny, bonus ve move bilgilerini çıkarır.
        """
        import urllib.request
        import json

        url = "https://raw.githubusercontent.com/zhenga8533/leak-duck/data/events.json"
        req = urllib.request.Request(
            url, headers={"User-Agent": "SinKa-PvP-100/19"}
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            data = json.loads(r.read().decode("utf-8", errors="ignore"))

        events = []
        seen = set()

        def names(items):
            out = []
            if isinstance(items, list):
                for x in items:
                    if isinstance(x, dict):
                        n = x.get("name")
                        if n and n not in out:
                            out.append(str(n))
                    elif isinstance(x, str) and x not in out:
                        out.append(x)
            return out

        def shiny_names(items):
            out = []
            if isinstance(items, list):
                for x in items:
                    if isinstance(x, dict):
                        n = x.get("name")
                        if n and n not in out:
                            out.append(str(n))
            return out

        def move_names(items):
            # ScrapedDuck'un moves alanı eventte öne çıkan Pokémonları belirtir.
            return names(items)

        # Gerçek JSON şeması: Event, Max Mondays, Season, Raid Battles vb.
        if not isinstance(data, dict):
            return []

        for category, rows in data.items():
            if not isinstance(rows, list):
                continue

            for item in rows:
                if not isinstance(item, dict):
                    continue

                title = str(item.get("title") or "Pokémon GO Etkinliği")
                article_url = str(item.get("article_url") or "")
                details = item.get("details") or {}
                if not isinstance(details, dict):
                    details = {}

                key = article_url.lower() or title.lower()
                if key in seen:
                    continue
                seen.add(key)

                features = names(details.get("features", []))
                spawns = names(details.get("spawns", []))
                raids = names(details.get("raids", []))
                eggs = names(details.get("eggs", []))
                research = names(details.get("research", []))
                shiny = shiny_names(details.get("shiny", []))
                moves = move_names(details.get("moves", []))
                bonuses = details.get("bonuses", [])
                if not isinstance(bonuses, list):
                    bonuses = [str(bonuses)]

                # Max Monday gibi bazı kayıtlarda details boş, Pokémon açıklamada.
                description = str(item.get("description") or "")
                pokemon = []
                for group in (features, spawns, raids, eggs, research, moves):
                    for n in group:
                        if n not in pokemon:
                            pokemon.append(n)

                date_info = []
                if item.get("start_time"):
                    date_info.append(str(item["start_time"]))
                if item.get("end_time"):
                    date_info.append(str(item["end_time"]))

                events.append({
                    "title": title,
                    "url": article_url,
                    "category": str(item.get("category") or category),
                    "start_time": item.get("start_time"),
                    "end_time": item.get("end_time"),
                    "pokemon": pokemon,
                    "features": features,
                    "spawns": spawns,
                    "raids": raids,
                    "eggs": eggs,
                    "research": research,
                    "shiny": shiny,
                    "moves": moves,
                    "bonuses": [str(x) for x in bonuses],
                    "headings": [str(item.get("category") or category)],
                    "date_info": date_info,
                    "text": description,
                    "raw": item,
                })

        # ScrapedDuck JSON bazı etkinliklerde "Featured Attack" bilgisini
        # henüz ayrı bir moves alanına koymayabiliyor (ör. Community Day Classic).
        # Bu nedenle yalnızca özel hareket ihtimali yüksek olan etkinlik sayfalarını
        # PARALEL şekilde kontrol ediyoruz. Böylece tüm eventleri tek tek bekletmiyoruz.
        import concurrent.futures
        import urllib.request
        from html.parser import HTMLParser
        import re

        candidates = []
        for ev in events:
            title = str(ev.get("title", "")).lower()
            category = str(ev.get("category", "")).lower()
            likely = (
                "community day" in title or
                "community day" in category or
                "raid day" in title or
                "raid day" in category or
                "evolution" in title or
                "evolve" in title or
                "featured attack" in str(ev.get("text", "")).lower()
            )
            if likely and ev.get("url"):
                candidates.append(ev)

        class AttackParser(HTMLParser):
            def __init__(self):
                super().__init__()
                self.parts=[]
            def handle_data(self,data):
                s=" ".join(data.split())
                if s:
                    self.parts.append(s)

        def enrich(ev):
            try:
                req=urllib.request.Request(
                    ev["url"],
                    headers={"User-Agent":"Mozilla/5.0 (SinKA PvP)"}
                )
                with urllib.request.urlopen(req, timeout=7) as r:
                    html=r.read().decode("utf-8", errors="ignore")
                p=AttackParser()
                p.feed(html)
                body=" ".join(p.parts)
                low=body.lower()

                terms=("featured attack","special move","charged attack",
                       "fast attack","exclusive attack")
                positions=[low.find(t) for t in terms if low.find(t)>=0]
                if not positions:
                    return ev

                p0=min(positions)
                snippet=body[p0:p0+1200]
                ev["featured_attack_text"]=snippet

                # Sayfada gerçek hareket adı varsa ayrıca çıkar.
                move_patterns = [
                    r"(?:featured attack|special move|charged attack|fast attack|exclusive attack)[^.!:]{0,120}(?:will know|learn|with|is)[: ]+([A-Z][A-Za-z'’+\- ]{2,40}?)(?=\.|,|\s+Trainer Battles|\s+Antrenör)",
                    r"(?:charged attack|fast attack)[: ]+([A-Z][A-Za-z'’+\- ]{2,40}?)(?=\.|,|\s+Trainer Battles)"
                ]
                for pat in move_patterns:
                    mm=re.search(pat, snippet, re.I)
                    if mm:
                        mv=mm.group(1).strip()
                        # Pokémon adını hareket sanma.
                        if mv.lower() not in {str(x).lower() for x in ev.get("pokemon", [])}:
                            ev["actual_move"]=mv
                            break

                # Harekete dair başlık sonrasında gelen Pokémon adı ve hareketi
                # genellikle ilk 2-3 cümlede bulunur. UI'ya metni de taşıyoruz.
                if not ev.get("text"):
                    ev["text"]=snippet
                return ev
            except Exception:
                return ev

        if candidates:
            with concurrent.futures.ThreadPoolExecutor(max_workers=min(5, len(candidates))) as pool:
                enriched=list(pool.map(enrich, candidates))
            by_url={e.get("url"):e for e in enriched}
            for ev in events:
                if ev.get("url") in by_url:
                    extra=by_url[ev["url"]]
                    if extra.get("featured_attack_text"):
                        ev["featured_attack_text"]=extra["featured_attack_text"]
                        ev["text"]=extra.get("text", ev.get("text",""))

        return events

    def _refresh_event_window(self, win, cards_frame, canvas, status):
        for w in cards_frame.winfo_children():
            w.destroy()
        self._event_window_refs = []

        today = datetime.now().date()

        for idx, event in enumerate(FALLBACK_EVENTS, 1):
            card = ttk.LabelFrame(
                cards_frame,
                text=f"  {idx}. {event['title']}  ",
                padding=12
            )
            card.pack(fill="x", padx=5, pady=7)

            names = _event_pokemon_names(event)
            if names:
                image_row = ttk.Frame(card)
                image_row.pack(anchor="w", pady=(0, 5))

                for name in names:
                    item = ttk.Frame(image_row, padding=4)
                    item.pack(side="left", padx=8)
                    photo = _event_load_image(name)
                    if photo:
                        self._event_window_refs.append(photo)
                        ttk.Label(item, image=photo).pack()
                    else:
                        ttk.Label(item, text="🖼️", font=("Segoe UI", 30)).pack()
                    ttk.Label(
                        item, text=name, font=("Segoe UI", 10, "bold")
                    ).pack()

            ttk.Label(
                card,
                text=f"📅 {event['date']}    ⏰ {event['time']}",
                font=("Segoe UI", 11, "bold")
            ).pack(anchor="w")

            try:
                d = datetime.strptime(event["date"], "%d.%m.%Y").date()
                days = (d - today).days
                if days < 0:
                    status_text = "GEÇMİŞ"
                elif days == 0:
                    status_text = "🔴 BUGÜN"
                elif days <= 2:
                    status_text = f"🔴 {days} GÜN KALDI"
                elif days <= 7:
                    status_text = f"🟠 {days} GÜN KALDI"
                else:
                    status_text = f"🟢 {days} GÜN KALDI"
            except Exception:
                status_text = "🟢 TAKVİMDE"

            ttk.Label(
                card, text=f"⏳ Durum: {status_text}",
                font=("Segoe UI", 10, "bold")
            ).pack(anchor="w", pady=(3, 0))
            ttk.Label(
                card, text=f"🔄 Evrim: {event['evolution']}"
            ).pack(anchor="w")
            ttk.Label(
                card,
                text=f"⚡ Özel hareket: {event['attack']}",
                font=("Segoe UI", 10, "bold")
            ).pack(anchor="w")
            ttk.Label(
                card, text=f"⏱ Evrim için son zaman: {event['deadline']}"
            ).pack(anchor="w")
            ttk.Label(
                card,
                text=f"📝 {event['detail']}",
                wraplength=820,
                justify="left"
            ).pack(anchor="w", pady=(3, 0))

        status.set(f"{len(FALLBACK_EVENTS)} yakın event gösteriliyor.")
    def _team_pokemon_image(self, pokemon_name, size=(58, 58)):
        """Takım ekranı için görseli SinKa'nın mevcut verisinden/cache'inden bulur."""
        if not PIL_OK:
            return None

        try:
            name = str(pokemon_name).strip()

            # 1) SinKa'da zaten bulunan görsel yükleme fonksiyonlarını kullanmayı dene.
            # Fonksiyon imzaları farklı olabileceği için kontrollü çağrılır.
            candidates = [
                "load_pokemon_image",
                "get_pokemon_image",
                "load_image",
                "get_image",
                "pokemon_image",
            ]
            for method_name in candidates:
                method = getattr(self, method_name, None)
                if not callable(method) or method_name == "_team_pokemon_image":
                    continue
                for args in ((name, size), (name,),):
                    try:
                        result = method(*args)
                        if result is not None:
                            if isinstance(result, ImageTk.PhotoImage):
                                return result
                            if hasattr(result, "getbbox"):
                                img = result.copy()
                                img.thumbnail(size, Image.LANCZOS)
                                return ImageTk.PhotoImage(img)
                    except Exception:
                        pass

            # 2) SinKa'nın veri/cache yapısından ID bulmayı dene.
            pid = None
            clean = re.sub(r"\s*\(.*?\)", "", name).strip()

            # Tree/data içinde ID tutuluyorsa onu kullan.
            for attr in ("pokemon_data", "pokemon_map", "pokemon_cache", "data"):
                obj = getattr(self, attr, None)
                if isinstance(obj, dict):
                    for key in (name, clean, name.lower(), clean.lower()):
                        item = obj.get(key)
                        if isinstance(item, dict):
                            pid = item.get("id") or item.get("dex") or item.get("pokemonId")
                            if pid:
                                break
                    if pid:
                        break

            # 3) İsim → Pokédex ID için yaygın temel eşleşmeler.
            fallback_ids = {
                "bulbasaur": 1, "ivysaur": 2, "venusaur": 3,
                "charmander": 4, "charmeleon": 5, "charizard": 6,
                "squirtle": 7, "wartortle": 8, "blastoise": 9,
                "pikachu": 25, "raichu": 26,
                "machop": 66, "machoke": 67, "machamp": 68,
                "gastly": 92, "haunter": 93, "gengar": 94,
                "lapras": 131, "eevee": 133, "vaporeon": 134,
                "jolteon": 135, "flareon": 136,
                "gible": 443, "gabite": 444, "garchomp": 445,
                "phantump": 708, "trevenant": 709,
                "staraptor": 398,
            }

            key = clean.lower().replace(" ", "-")
            pid = pid or fallback_ids.get(key)

            if not pid:
                # 4) PokeAPI'den ID çöz; görsel indirme arka thread'de yapılır.
                url = f"https://pokeapi.co/api/v2/pokemon/{key}"
                with urllib.request.urlopen(url, timeout=3) as r:
                    data = json.loads(r.read().decode("utf-8"))
                pid = data.get("id")

            if not pid:
                return None

            cache_dir = os.path.join(os.path.expanduser("~"), "PokemonGO_Event_Images")
            os.makedirs(cache_dir, exist_ok=True)
            cache = os.path.join(cache_dir, f"team_{pid}.png")

            if not os.path.exists(cache):
                img_url = (
                    "https://raw.githubusercontent.com/PokeAPI/sprites/master/"
                    f"sprites/pokemon/other/official-artwork/{pid}.png"
                )
                with urllib.request.urlopen(img_url, timeout=5) as r:
                    raw = r.read()
                with open(cache, "wb") as f:
                    f.write(raw)

            img = Image.open(cache).convert("RGBA")
            img.thumbnail(size, Image.LANCZOS)
            return ImageTk.PhotoImage(img)

        except Exception:
            return None

    def _team_move_type_key_uncached(self, move_id):
        """Bir hareket ID/name için PvPoke Game Master tipini bulur."""
        raw = str(move_id or "").strip()
        if not raw or raw == "-":
            return ""
        # Zaten sözlük verilmişse doğrudan kullan.
        if isinstance(move_id, dict):
            val = move_id.get("type") or move_id.get("moveType") or move_id.get("typeId")
            return str(val or "").lower().replace("type_", "")
        cache = getattr(self, "_sinka_pvpoke_moves_cache", {})
        if isinstance(cache, dict):
            m = cache.get(raw)
            if isinstance(m, dict):
                val = m.get("type") or m.get("moveType") or m.get("typeId")
                if val:
                    return str(val).lower().replace("type_", "")
            # Bazı veri sürümlerinde proto/id üzerinden eşleşme gerekebilir.
            ru = raw.upper()
            for key, m in cache.items():
                if not isinstance(m, dict):
                    continue
                if str(key).upper() == ru or str(m.get("moveId") or "").upper() == ru or str(m.get("name") or "").upper() == ru:
                    val = m.get("type") or m.get("moveType") or m.get("typeId")
                    if val:
                        return str(val).lower().replace("type_", "")
        return ""

    def _team_move_type_key(self, move_id):
        """Hareket tipini cache'leyerek hızlı döndürür. UI redraw sırasında ağ/GM taraması yapmaz."""
        if isinstance(move_id, dict):
            raw = str(move_id.get("moveId") or move_id.get("id") or move_id.get("name") or "").strip()
        else:
            raw = str(move_id or "").strip()
        if not raw or raw == "-":
            return ""
        cached = self._move_type_key_cache.get(raw)
        if cached is not None:
            return cached
        key = self._team_move_type_key_uncached(move_id)
        self._move_type_key_cache[raw] = key or ""
        return key or ""

    def _main_move_type_keys(self, name):
        """Ana lig tablosundaki Güç 1/2/3 hareketlerinin gerçek Pokémon GO tip ikon anahtarlarını döndürür."""
        cache_key = str(name or "").strip()
        cached = getattr(self, "_main_move_types_cache", {}).get(cache_key)
        if cached is not None:
            return list(cached)
        out = []
        moves = []

        # ÖNEMLİ: Ana lig tablosunda gösterilen hareketler PvPoke ranking
        # verisindeki "moveset" alanıdır. Game Master pokemon.json içinde
        # hareketlerin tipi bulunmadığı için yalnızca pokemon.json'a bakmak
        # çoğu zaman boş sonuç veriyordu. Önce görünür ranking kaydından
        # hareket ID'lerini alıyoruz.
        try:
            target_norm = self._pokemon_name_norm(str(name)) if hasattr(self, "_pokemon_name_norm") else re.sub(r"[^a-z0-9]+", "", str(name).lower())
            for rp in (self.ranking_data or []):
                if not isinstance(rp, dict):
                    continue
                rn = str(rp.get("speciesName") or rp.get("name") or rp.get("speciesId") or "")
                rn_norm = self._pokemon_name_norm(rn) if hasattr(self, "_pokemon_name_norm") else re.sub(r"[^a-z0-9]+", "", rn.lower())
                if rn_norm == target_norm or rn.lower() == str(name).lower():
                    raw_rank_moves = rp.get("moveset") or []
                    if isinstance(raw_rank_moves, (list, tuple)):
                        moves = list(raw_rank_moves)[:3]
                    break
        except Exception:
            pass

        # Ranking kaydı bulunamazsa Game Master pokemon kaydına geri dön.
        if not moves:
            try:
                sid, poke = self._find_pokemon_data_by_name(str(name))
            except Exception:
                sid, poke = None, None
            if isinstance(poke, dict):
                raw = poke.get("moveset") or poke.get("moves")
                if not raw:
                    fast = poke.get("fastMoves") or poke.get("fast_moves") or []
                    charged = poke.get("chargedMoves") or poke.get("charged_moves") or []
                    raw = list(fast) + list(charged)
                if isinstance(raw, dict):
                    raw = list(raw.values())
                if isinstance(raw, (list, tuple)):
                    moves = list(raw)[:3]
                elif raw:
                    moves = [raw]

        # Öncelik: Game Master hareket ID'si/tipi.
        for mv in moves:
            key = self._team_move_type_key(mv)
            if not key and isinstance(mv, dict):
                key = str(mv.get("type") or mv.get("moveType") or mv.get("typeId") or "").lower().replace("type_", "")
            out.append(key)
        while len(out) < 3:
            out.append("")
        result = out[:3]
        if not hasattr(self, "_main_move_types_cache"):
            self._main_move_types_cache = {}
        self._main_move_types_cache[cache_key] = list(result)
        return result

    def _move_type_key_from_display_name(self, move_name):
        """Çevrilmiş hareket adından Game Master tipini bulmak için güvenli geri dönüş."""
        raw = str(move_name or "").strip().rstrip("*")
        if not raw or raw == "-":
            return ""
        key = self._team_move_type_key(raw)
        if key:
            return key
        cache = getattr(self, "_sinka_pvpoke_moves_cache", {})
        if isinstance(cache, dict):
            for _, mv in cache.items():
                if not isinstance(mv, dict):
                    continue
                en = str(mv.get("name") or mv.get("moveId") or "").strip()
                try:
                    tr = str(self.translate_move(en)).strip().rstrip("*")
                except Exception:
                    tr = en
                if raw.lower() in {en.lower(), tr.lower()}:
                    val = mv.get("type") or mv.get("moveType") or mv.get("typeId")
                    if val:
                        return str(val).lower().replace("type_", "")
        return ""

    def _team_type_keys(self, name):
        """Game Master verisinden bir Pokémonun tiplerini cache'li şekilde döndürür."""
        cache = getattr(self, "_sinka_type_cache", None)
        if cache is None:
            cache = {}
            self._sinka_type_cache = cache
        clean = re.sub(r"\s*\((?:shadow|purified)\)\s*$", "", str(name), flags=re.I).strip()
        target = self._pokemon_name_norm(clean) if hasattr(self, "_pokemon_name_norm") else re.sub(r"[^a-z0-9]+", "", clean.lower())
        if target in cache:
            return cache[target]
        for sid, poke in (self.pokemon or {}).items():
            vals = [sid, poke.get("speciesId"), poke.get("name"), poke.get("speciesName")]
            norms = [re.sub(r"[^a-z0-9]+", "", str(v).lower()) for v in vals if v]
            if target in norms or any(target and (target == n or target in n or n in target) for n in norms):
                raw = poke.get("types") or poke.get("type") or []
                out = []
                for t in raw if isinstance(raw, list) else [raw]:
                    if isinstance(t, dict):
                        t = t.get("name") or t.get("type")
                    if t:
                        out.append(str(t).lower().replace("type_", ""))
                result = tuple(dict.fromkeys(out))
                cache[target] = result
                return result
        cache[target] = tuple()
        return tuple()

    def _fresh_team_ranking(self, league):
        """Takım oluşturma için butona basıldığı anda güncel PvPoke Overall verisini çeker."""
        if league == "Mega":
            return download_json(MEGA_RANKING_URL)
        if league not in LEAGUES or LEAGUES.get(league) is None:
            return list(self.ranking_data or [])
        return download_json(RANKING_URL.format(cp=LEAGUES[league]))

    def _training_team_bonus(self, candidate_names, league):
        """PvPoke Training Analysis'ten küçük kullanım sinyali; erişilemezse 0."""
        try:
            if league not in ("Great League", "Ultra League", "Master League"):
                return 0.0
            cache = getattr(self, "_sinka_training_html_cache", {})
            if league not in cache:
                url = "https://pvpoke.com/might-and-mastery/train/analysis/"
                req = urllib.request.Request(url, headers={"User-Agent": "SinKA-PvP-TeamBuilder/3.0"})
                with urllib.request.urlopen(req, timeout=15) as r:
                    html = r.read().decode("utf-8", errors="ignore")
                low = re.sub(r"<[^>]+>", " ", html)
                cache[league] = re.sub(r"\s+", " ", low).lower()
                self._sinka_training_html_cache = cache
            low = cache[league]
            names = [re.sub(r"\s*\(shadow\)\s*$", "", str(x), flags=re.I).lower() for x in candidate_names]
            hit = sum(1 for n in names if n and n in low)
            return 4.0 if hit == 3 else (1.5 if hit == 2 else 0.0)
        except Exception:
            return 0.0

    def _generate_meta_teams(self, league, ranking_data, show_shadows=False, progress_cb=None, owned_only=False):
        """SinKA Elite 90+ Team Engine.

        Katmanlar:
          1) Güncel Top-200 meta havuzu
          2) Kullanıcının geçerli Pokémon havuzu
          3) Hızlı matchup / coverage / core ön elemesi
          4) 6 rol dizilimi
          5) Gerçek PvPoke savaş motorunun tam portu olmayan, fakat
             PvPoke matchup + rol verilerini shield/energy/bait/timing
             senaryolarında yeniden değerlendiren "battle-scenario engine"
          6) Alignment / recovery / RPS / consistency / bulk / robustness
          7) Diversity + sonuç açıklaması

        Önemli: Bu fonksiyon tam turn-by-turn PvPoke JS battle engine değildir.
        Move/energy/CMP verisi ranking JSON'ında bulunmadığında sahte kesinlik
        üretmez; senaryo katmanı mevcut PvPoke ratinglerini kontrollü şekilde
        stres testine sokar.
        """
        from itertools import combinations, permutations

        def progress(percent, stage, detail=""):
            if progress_cb:
                try:
                    progress_cb(percent, stage, detail)
                except Exception:
                    pass

        def norm_name(value):
            try:
                return self._pokemon_name_norm(value)
            except Exception:
                return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())

        def sid_of(p):
            return str(p.get("speciesId") or p.get("speciesName") or p.get("name") or "")

        def display_name(p):
            return str(p.get("speciesName") or p.get("name") or p.get("speciesId") or "Bilinmeyen")

        def clamp(v, lo=0.0, hi=100.0):
            return max(lo, min(hi, float(v)))

        def pct_from_rating(v):
            return clamp(float(v) / 10.0)

        progress(5, "1/7 Güncel meta hazırlanıyor", "Tüm uygun PvPoke sıralaması taranıyor...")
        raw_all = [x for x in (ranking_data or []) if isinstance(x, dict) and not is_gigantamax(x)]
        if not show_shadows:
            raw_all = [x for x in raw_all if not is_shadow(x)]
        if len(raw_all) < 3:
            progress(100, "✗ Yeterli veri yok", "Filtre sonrası 3'ten az Pokémon kaldı.")
            return []

        # Aynı species/form kaydının tekrarını engelle.
        opponent_pool, seen_opp = [], set()
        for x in raw_all:
            sid = sid_of(x)
            if not sid or sid in seen_opp:
                continue
            seen_opp.add(sid)
            opponent_pool.append(x)
            if len(opponent_pool) >= 200:
                break
        if len(opponent_pool) < 3:
            return []

        # Kullanıcı havuzu.
        owned_names = set()
        if owned_only:
            owned = dict(getattr(self, "owned_by_league", {}).get(league, {}))
            for name, state in owned.items():
                if state:
                    owned_names.add(norm_name(name))

        if owned_only:
            candidate_source = []
            for x in raw_all:
                fields = (x.get("speciesId"), x.get("speciesName"), x.get("name"))
                if any(norm_name(v) in owned_names for v in fields if v):
                    candidate_source.append(x)
        else:
            candidate_source = list(raw_all)

        if len(candidate_source) < 3:
            progress(100, "✗ Yeterli Pokémon bulunamadı", f"Uygun aday sayısı: {len(candidate_source)}")
            return []

        candidate_source.sort(key=lambda x: float(x.get("score", 0) or 0), reverse=True)
        # İlk aramada TÜM mevcut Pokémonlar aday havuzuna alınır. Daha sonra
        # hepsi hızlı bireysel filtrede değerlendirilir ve yalnızca güçlü adaylar
        # kombinasyon aşamasına geçirilir. Böylece kullanıcı havuzu kaybolmaz,
        # fakat gereksiz üçlü kombinasyonları hesaplamayız.
        # Tüm uygun meta havuzunu tarıyoruz. Kombinasyon aşaması öncesinde
        # hızlı kalite filtresiyle daraltacağız; böylece düşük sıralı fakat takım
        # sinerjisi yüksek Pokémonların daha baştan kaybolmasını engelliyoruz.
        candidate_pool = list(candidate_source)

        progress(14, "2/7 Aday havuzu hazırlandı",
                 f"Top-200 rakip: {len(opponent_pool)} • Aday: {len(candidate_pool)}" +
                 (" • Sadece Mevcut Pokémonlar" if owned_only else ""))

        # ------------------------- temel veriler -------------------------
        type_map, bulk_map, stat_map = {}, {}, {}
        for p in candidate_pool:
            sid = sid_of(p)
            try:
                type_map[sid] = set(self._team_type_keys(display_name(p)))
            except Exception:
                type_map[sid] = set()
            stats = p.get("stats") or {}
            try:
                hp = float(stats.get("hp", stats.get("stamina", 0)) or 0)
                defense = float(stats.get("defense", stats.get("def", 0)) or 0)
                attack = float(stats.get("attack", stats.get("atk", 0)) or 0)
            except Exception:
                hp = defense = attack = 0.0
            bulk_map[sid] = {"bulk": hp * defense, "hp": hp, "defense": defense, "attack": attack}
            stat_map[sid] = {"attack": attack, "defense": defense, "hp": hp}
        max_bulk = max([v["bulk"] for v in bulk_map.values()] or [1.0])

        def matchup_value(p, opponent_sid):
            for key in ("matchups", "counters"):
                for m in p.get(key, []) or []:
                    if not isinstance(m, dict):
                        continue
                    if str(m.get("opponent")) == opponent_sid:
                        try:
                            return float(m.get("rating", 500) or 500)
                        except Exception:
                            return 500.0
            return 500.0

        opponent_ids = [sid_of(x) for x in opponent_pool]
        opponent_weights = {}
        opponent_rank = {}
        for rank_idx, o in enumerate(opponent_pool, 1):
            sid = sid_of(o)
            pscore = clamp(float(o.get("score", 50) or 50), 0, 100)
            rank_factor = 1.0 / (1.0 + 0.0105 * (rank_idx - 1))
            score_factor = 0.70 + 0.30 * (pscore / 100.0)
            opponent_weights[sid] = max(0.15, rank_factor * score_factor)
            opponent_rank[sid] = rank_idx
        total_weight = max(1.0, sum(opponent_weights.values()))

        role_map, individual_score, matchup_matrix = {}, {}, {}
        for p in candidate_pool:
            sid = sid_of(p)
            matchup_matrix[sid] = {oid: matchup_value(p, oid) for oid in opponent_ids if oid != sid}
            base = float(p.get("score", 0) or 0)
            individual_score[sid] = base
            scores = list(p.get("scores") or [])
            while len(scores) < 6:
                scores.append(base)
            vals = []
            for v in scores[:6]:
                try:
                    vals.append(clamp(float(v)))
                except Exception:
                    vals.append(clamp(base))
            role_map[sid] = {
                "lead": vals[0], "closer": vals[1], "switch": vals[2],
                "charger": vals[3], "attacker": vals[4], "consistency": vals[5]
            }

        # Tüm mevcutlar yukarıda değerlendirildi; kombinasyon sayısını azaltmak için
        # hızlı bireysel ön eleme yapıyoruz. Bu aşama yalnızca ranking + weighted
        # Top-50 matchup + roller kullanır; ağır battle simülasyonu değildir.
        fast_meta_opponents = opponent_pool[:50]
        fast_meta_weight = max(1.0, sum(opponent_weights.get(sid_of(o), 0.15) for o in fast_meta_opponents))
        fast_candidates = []
        for p in candidate_pool:
            pid = sid_of(p)
            weighted_match = 0.0
            for o in fast_meta_opponents:
                oid = sid_of(o)
                if oid == pid:
                    continue
                weighted_match += matchup_matrix[pid].get(oid, 500.0) * opponent_weights.get(oid, 0.15)
            weighted_match /= fast_meta_weight
            rm = role_map[pid]
            fast_score = (0.48 * individual_score[pid] + 0.22 * weighted_match / 10.0 +
                          0.10 * rm["lead"] + 0.10 * rm["switch"] + 0.10 * rm["closer"])
            fast_candidates.append((fast_score, p))
        fast_candidates.sort(key=lambda z: z[0], reverse=True)
        # Ön eleme tüm havuzu görür; yalnızca kombinasyon patlamasını önlemek için
        # en güçlü 60 aday üçlü üretim aşamasına geçirilir. Mevcut Pokémon filtresi
        # aktifse tüm sahip olunan adaylar yine önce değerlendirilir.
        fast_cap = 60 if not owned_only else min(60, max(36, len(fast_candidates)))
        candidate_pool = [p for _, p in fast_candidates[:min(fast_cap, len(fast_candidates))]]

        progress(25, "3/7 Matchup ve rol matrisi hazırlanıyor",
                 f"Top-200 • {len(candidate_source)} aday hızlı tarandı → {len(candidate_pool)} güçlü aday kombinasyona geçti")

        # -------------------- battle-scenario engine ----------------------
        # Turn-by-turn motor için güncel PvPoke Game Master move verisini bir kez
        # yükle. Ağ erişimi yoksa mevcut proxy motoru güvenli biçimde devralır.
        real_battle_engine = None
        real_battle_status = "proxy"
        try:
            move_cache = getattr(self, "_sinka_pvpoke_moves_cache", None)
            if not isinstance(move_cache, dict) or len(move_cache) < 100:
                move_cache = download_json(PVPoke_MOVES_URL)
                if isinstance(move_cache, list):
                    move_cache = {str(m.get("moveId")): m for m in move_cache if isinstance(m, dict) and m.get("moveId")}
                else:
                    move_cache = {}
                self._sinka_pvpoke_moves_cache = move_cache
            if move_cache:
                real_battle_engine = SinkaBattleEngine(move_cache, self.pokemon)
                real_battle_status = "turn-by-turn"
        except Exception:
            real_battle_engine = None
            real_battle_status = "proxy"

        # Proxy katmanı, turn-by-turn motorun erişemediği uzun kuyruktaki
        # matchup'ları hızlı değerlendirmek için korunur. Ağır gerçek simülasyon
        # yalnızca final aday havuzunda çalıştırılır.
        scenario_defs = (
            ("2-2", 2, 2, 1.00),
            ("1-1", 1, 1, 1.05),
            ("0-0", 0, 0, 0.95),
            ("2-1", 2, 1, 0.85),
            ("1-2", 1, 2, 0.85),
            ("0-1", 0, 1, 0.70),
            ("1-0", 1, 0, 0.70),
        )
        energy_steps = (-6, -3, 0, 3, 6)

        def scenario_rating(pid, oid, own_shields, opp_shields, energy_delta=0, bait=True):
            base = matchup_matrix.get(pid, {}).get(oid, 500.0)
            base100 = pct_from_rating(base)
            rm = role_map.get(pid, {})
            consistency = rm.get("consistency", 50.0)
            charger = rm.get("charger", 50.0)
            attacker = rm.get("attacker", 50.0)

            # Shield leverage: daha fazla shield alan tarafın baskı kurma
            # ihtimali artar; 0-0 closers ise consistency/bulk tarafına daha çok yaslanır.
            shield_delta = own_shields - opp_shields
            shield_adj = shield_delta * (2.0 + 0.035 * (charger - 50.0))
            if own_shields == 0 and opp_shields == 0:
                shield_adj += 0.03 * (rm.get("closer", 50.0) - 50.0)

            # Energy advantage: +energy özellikle charger/attacker profillerinde daha değerlidir.
            energy_adj = (energy_delta / 3.0) * (0.55 + 0.009 * (charger - 50.0) + 0.006 * (attacker - 50.0))

            # Bait/no-bait stres testi. Başarı ihtimali consistency ile sınırlanır.
            if bait and own_shields > 0:
                bait_prob = 0.30 + 0.0045 * consistency
                bait_swing = 2.0 + 0.055 * max(0.0, 100.0 - consistency)
                bait_adj = (bait_prob * bait_swing) - ((1.0 - bait_prob) * 0.55)
            else:
                bait_adj = 0.0

            # CMP/timing proxy: ranking verisinde gerçek IV/CMP turn bilgisi yoksa
            # yalnızca mevcut saldırı statı + attacker rolü ile küçük robustness sinyali.
            # Gerçek IV/CMP verisi daha sonra gelirse bu katman değiştirilebilir.
            timing_adj = 0.0
            if stat_map.get(pid, {}).get("attack", 0) > 0:
                timing_adj = 0.35 * (attacker - 50.0) / 10.0

            return clamp(base100 + shield_adj + energy_adj + bait_adj + timing_adj)

        # Senaryo katmanı Top-200'ün tamamını tekrar tekrar taramak yerine en güçlü
        # 40 meta rakip üzerinde hesaplanır. Top-200 coverage hesabı ise aşağıda
        # tam olarak korunur. Bu ayrım büyük hız kazandırır.
        scenario_opponents = opponent_pool[:40]
        scenario_cache = {}
        scenario_weight_total = max(1.0, sum(opponent_weights.get(sid_of(o), 0.15) for o in scenario_opponents))
        for p in candidate_pool:
            pid = sid_of(p)
            for o in scenario_opponents:
                oid = sid_of(o)
                if oid == pid:
                    continue
                for _, sh, oh, sw in scenario_defs:
                    for energy in energy_steps:
                        for bait in (False, True):
                            scenario_cache[(pid, oid, sh, oh, energy, bait)] = scenario_rating(pid, oid, sh, oh, energy, bait)

        def cached_scenario(pid, oid, sh, oh, energy=0, bait=True):
            key = (pid, oid, sh, oh, energy, bait)
            value = scenario_cache.get(key)
            if value is not None:
                return value
            value = scenario_rating(pid, oid, sh, oh, energy, bait)
            scenario_cache[key] = value
            return value

        candidate_role_match = {}
        role_scenario_cache = {}
        for p in candidate_pool:
            pid = sid_of(p)
            vals = []
            for sname, sh, oh, sw in scenario_defs:
                total = 0.0
                for o in scenario_opponents:
                    oid = sid_of(o)
                    if oid == pid:
                        continue
                    ow = opponent_weights.get(oid, 0.15)
                    total += cached_scenario(pid, oid, sh, oh, 0, bait=(sh > 0)) * (ow / scenario_weight_total)
                vals.append((sname, total * sw))
            # Rolün kendi PvPoke puanı + senaryo ortalaması.
            scenario_avg = sum(v for _, v in vals) / max(1, sum(sw for _,_,_,sw in scenario_defs))
            candidate_role_match[pid] = clamp(0.62 * scenario_avg + 0.38 * individual_score[pid])
            role_scenario_cache[pid] = vals

        # Pair core cache: ortak zayıflık + birbirini kurtarma + type diversity.
        pair_cache = {}
        for ai, a in enumerate(candidate_pool):
            aid = sid_of(a); ar = matchup_matrix[aid]
            for b in candidate_pool[ai + 1:]:
                bid = sid_of(b); br = matchup_matrix[bid]
                cover = shared_bad = hard_shared = recovery = 0.0
                for o in opponent_pool[:80]:
                    oid = sid_of(o); ow0 = opponent_weights.get(oid, 0.15)
                    if oid in (aid, bid):
                        continue
                    w = ow0 / total_weight
                    va, vb = ar.get(oid, 500.0), br.get(oid, 500.0)
                    if va < 470 and vb >= 530:
                        cover += 1.0 * w
                    elif vb < 470 and va >= 530:
                        cover += 1.0 * w
                    if va < 450 and vb < 450:
                        shared_bad += w
                    if va < 400 and vb < 400:
                        hard_shared += w
                    if va < 480 and vb >= 560:
                        recovery += w
                    elif vb < 480 and va >= 560:
                        recovery += w
                type_a, type_b = type_map.get(aid, set()), type_map.get(bid, set())
                type_bonus = 2.0 if type_a.isdisjoint(type_b) else 0.0
                pair_cache[frozenset((aid, bid))] = {
                    "core": clamp(50.0 + 90.0 * cover + 35.0 * recovery + type_bonus - 120.0 * shared_bad - 180.0 * hard_shared),
                    "coverage": clamp(100.0 * cover),
                    "shared_bad": 100.0 * shared_bad,
                    "recovery": clamp(100.0 * recovery),
                }

        def pair_data(a, b):
            return pair_cache.get(frozenset((sid_of(a), sid_of(b))), {"core": 50.0, "coverage": 0.0, "shared_bad": 0.0, "recovery": 50.0})

        # ------------------ takım common özellikleri ---------------------
        common_cache = {}
        opponent_rank_map = {}
        for _o in opponent_pool:
            try:
                opponent_rank_map[sid_of(_o)] = int(float(_o.get("rank", 999) or 999))
            except Exception:
                opponent_rank_map[sid_of(_o)] = 999

        def team_common(a, b, c):
            ids = [sid_of(a), sid_of(b), sid_of(c)]
            union_types = set().union(*(type_map.get(x, set()) for x in ids))
            pair_ab, pair_ac, pair_bc = pair_data(a,b), pair_data(a,c), pair_data(b,c)

            weighted_best = weighted_second = weighted_worst = 0.0
            good = strong = double_bad = triple_bad = alignment = hard_loss = 0.0
            threat_rows = []
            for oid, ow0 in opponent_weights.items():
                if oid in ids:
                    continue
                w = ow0 / total_weight
                vals = sorted((matchup_matrix[pid].get(oid, 500.0) for pid in ids), reverse=True)
                best, second, worst = vals
                weighted_best += w * best
                weighted_second += w * second
                weighted_worst += w * worst
                if best >= 550: good += w
                if best >= 600: strong += w
                if best < 450: double_bad += 0 if second >= 450 else w
                if best < 450 and second < 450 and worst < 450: triple_bad += w
                if best < 450 and second < 450: alignment += w
                if best < 400: hard_loss += w
                threat_rows.append((w, best, second, oid))

            coverage = clamp(0.56 * pct_from_rating(weighted_best) + 0.29 * pct_from_rating(weighted_second) + 0.15 * pct_from_rating(weighted_worst))
            role_consistency = sum(role_map[x]["consistency"] for x in ids) / 3.0
            consistency = clamp(0.72 * role_consistency + 0.28 * (100.0 - alignment * 100.0))
            pair_core = clamp((pair_ab["core"] + pair_ac["core"] + pair_bc["core"]) / 3.0)
            pair_coverage = clamp((pair_ab["coverage"] + pair_ac["coverage"] + pair_bc["coverage"]) / 3.0)
            pair_recovery = clamp((pair_ab["recovery"] + pair_ac["recovery"] + pair_bc["recovery"]) / 3.0)

            bulk = sum((bulk_map.get(x, {}).get("bulk", 0.0) / max_bulk) * 100.0 for x in ids) / 3.0
            type_count = len(union_types)
            same_type_pairs = sum(
                1 for i in range(3) for j in range(i+1,3)
                if type_map.get(ids[i], set()) and type_map.get(ids[j], set()) and type_map.get(ids[i], set()) == type_map.get(ids[j], set())
            )
            # Aynı tip sayısını tek başına ceza sebebi yapma. Asıl önemli olan,
            # güçlü Top-200 tehditlerinin takımın 2 veya 3 üyesini aynı anda
            # kötü eşleşmeye sokup sokmadığıdır.
            # Ortak tehdit analizi: sadece typing değil, gerçek Top-200 matchup
            # skorları kullanılır. 500 altı dezavantaj, 450 altı ciddi dezavantaj,
            # 400 altı ağır kayıp kabul edilir. Meta sırası yükseldikçe tehdidin
            # ağırlığı artar. Böylece 2/3 ortak açık gerçek bir takım cezasına dönüşür.
            weighted_shared_2 = 0.0
            weighted_shared_3 = 0.0
            weighted_single_answer = 0.0
            weighted_severe_shared_2 = 0.0
            weighted_critical_triple = 0.0
            critical_threat_ids = []
            severe_shared_threat_ids = []
            shared_threat_rows = []
            for oid, ow0 in opponent_weights.items():
                if oid in ids:
                    continue
                w = ow0 / max(1e-9, total_weight)
                vals = [matchup_matrix[pid].get(oid, 500.0) for pid in ids]
                bad_count = sum(1 for v in vals if v < 500.0)
                severe_count = sum(1 for v in vals if v < 450.0)
                very_bad_count = sum(1 for v in vals if v < 400.0)
                strong_count = sum(1 for v in vals if v >= 550.0)
                rank = opponent_rank_map.get(oid, 999)
                # Rank 1-50 tehditleri özellikle ciddiye al.
                meta_factor = 1.0 + (0.35 if rank <= 25 else 0.18 if rank <= 50 else 0.0)
                if bad_count >= 2:
                    weighted_shared_2 += w * meta_factor
                if bad_count == 3:
                    weighted_shared_3 += w * meta_factor
                if strong_count == 1 and bad_count >= 2:
                    weighted_single_answer += w * meta_factor
                if severe_count >= 2:
                    weighted_severe_shared_2 += w * meta_factor
                    severe_shared_threat_ids.append(oid)
                if very_bad_count == 3 and rank <= 50:
                    weighted_critical_triple += w * meta_factor
                    critical_threat_ids.append(oid)
                if bad_count >= 2:
                    severity = (0.75 + 0.55 * bad_count + 0.65 * severe_count + 0.85 * very_bad_count) * meta_factor
                    shared_threat_rows.append((w * severity, oid, bad_count, strong_count, min(vals), rank))
            shared_threat_rows.sort(reverse=True)
            shared_threat_risk = clamp(weighted_shared_2 * 100.0)
            triple_threat_risk = clamp(weighted_shared_3 * 100.0)
            single_answer_dependency = clamp(weighted_single_answer * 100.0)
            severe_shared_risk = clamp(weighted_severe_shared_2 * 100.0)
            critical_triple_risk = clamp(weighted_critical_triple * 100.0)

            # Aynı tip otomatik ceza değildir. Sadece ortak tehdit riski yüksekse
            # çeşitlilik düşer. İki Water, örneğin üçüncü üye açıkları kapatıyorsa
            # korunabilir.
            same_type_penalty = min(3.0, same_type_pairs * 0.75)
            diversity = clamp(60.0 + type_count * 5.0 - same_type_penalty
                              - shared_threat_risk * 0.42
                              - severe_shared_risk * 0.28
                              - triple_threat_risk * 0.85
                              - critical_triple_risk * 1.15)
            role_balance = clamp(0.45 * sum(role_map[x]["lead"] for x in ids) / 3.0 + 0.35 * sum(role_map[x]["switch"] for x in ids) / 3.0 + 0.20 * sum(role_map[x]["closer"] for x in ids) / 3.0)
            avg_individual = sum(individual_score[x] for x in ids) / 3.0

            # Top-200 threat list: en kötü rakipleri açıklama üretiminde kullan.
            threat_rows.sort(key=lambda z: (z[1], -z[0]))
            worst_threats = [oid for _, best, _, oid in threat_rows[:5]]

            # Senaryo robustness: 40 güçlü meta rakip × 7 shield senaryosu × 5 enerji.
            # Önceden hesaplanmış cache kullanılır; aynı matchup tekrar simüle edilmez.
            scenario_scores = []
            for o in scenario_opponents:
                oid = sid_of(o)
                if oid in ids:
                    continue
                w = opponent_weights.get(oid, 0.15) / scenario_weight_total
                scenario_vals = []
                for _, sh, oh, sw in scenario_defs:
                    for energy in energy_steps:
                        vals = [cached_scenario(pid, oid, sh, oh, energy, bait=(sh > 0)) for pid in ids]
                        scenario_vals.append(max(vals) * sw)
                scenario_avg = sum(scenario_vals) / max(1.0, sum(sw for _,_,_,sw in scenario_defs) * len(energy_steps))
                scenario_scores.append((w, scenario_avg))
            battle_robustness = clamp(sum(w * v for w, v in scenario_scores)) if scenario_scores else 50.0

            # Energy / shield / bait dependency: takım senaryo dağılımının tabanı.
            no_bait = []
            bait = []
            for o in scenario_opponents:
                oid = sid_of(o)
                if oid in ids:
                    continue
                w = opponent_weights.get(oid, 0.15) / scenario_weight_total
                vals_no = [cached_scenario(pid, oid, 1, 1, 0, bait=False) for pid in ids]
                vals_bait = [cached_scenario(pid, oid, 1, 1, 0, bait=True) for pid in ids]
                no_bait.append((w, max(vals_no)))
                bait.append((w, max(vals_bait)))
            no_bait_score = sum(w*v for w,v in no_bait) if no_bait else 50.0
            bait_score = sum(w*v for w,v in bait) if bait else 50.0
            bait_dependency = clamp(bait_score - no_bait_score + (100.0 - consistency) * 0.15)
            shield_dependency = clamp(abs((battle_robustness - coverage)) * 0.65)
            energy_dependency = clamp(sum(abs(energy) for energy in energy_steps) / len(energy_steps) * 0.35)

            return {
                "ids": ids,
                "coverage": coverage,
                "meta_score": clamp(0.60 * battle_robustness + 0.25 * (good*100.0) + 0.15 * avg_individual),
                "consistency": consistency,
                "pair_syn": pair_core - 50.0,
                "core": pair_core,
                "diversity": diversity,
                "hard_loss": hard_loss * 100.0,
                "alignment_risk": alignment * 100.0,
                "good_coverage": good * 100.0,
                "strong_coverage": strong * 100.0,
                "double_bad": double_bad * 100.0,
                "triple_bad": triple_bad * 100.0,
                "worst": min((r[1] for r in threat_rows), default=500.0),
                "best": max((r[1] for r in threat_rows), default=500.0),
                "avg_individual": avg_individual,
                "bulk": bulk,
                "charger": sum(role_map[x]["charger"] for x in ids) / 3.0,
                "attacker": sum(role_map[x]["attacker"] for x in ids) / 3.0,
                "rps_dependency": clamp(alignment * 100.0 + double_bad * 35.0),
                "pair_coverage": pair_coverage,
                "pair_recovery": pair_recovery,
                "battle_robustness": battle_robustness,
                "bait_dependency": bait_dependency,
                "shield_dependency": shield_dependency,
                "energy_dependency": energy_dependency,
                "worst_threats": worst_threats,
                "shared_threat_risk": shared_threat_risk,
                "triple_threat_risk": triple_threat_risk,
                "single_answer_dependency": single_answer_dependency,
                "severe_shared_risk": severe_shared_risk,
                "critical_triple_risk": critical_triple_risk,
                "critical_threat_count": len(critical_threat_ids),
                "critical_threat_ids": critical_threat_ids[:12],
                "severe_shared_threat_count": len(severe_shared_threat_ids),
                "role_balance": role_balance,
                "same_type_pairs": same_type_pairs,
                "shared_threat_rows": shared_threat_rows[:8],
            }

        def common_for(a,b,c):
            key = tuple(sorted((sid_of(a), sid_of(b), sid_of(c))))
            if key not in common_cache:
                common_cache[key] = team_common(a,b,c)
            return common_cache[key]

        # Lead -> switch güvenliği ve recovery cache. Lead'in kötü olduğu
        # rakiplere safe switch'in gerçek senaryo proxy'si uygulanır.
        switch_cache = {}
        for lead in candidate_pool:
            lid = sid_of(lead)
            for switch in candidate_pool:
                sid = sid_of(switch)
                if sid == lid:
                    continue
                total = 0.0; wsum = 0.0; rec = 0.0; stable = 0.0
                switch_weight_total = max(1.0, sum(opponent_weights.get(sid_of(o),0.15) for o in scenario_opponents))
                for o in scenario_opponents:
                    oid = sid_of(o)
                    if oid in (lid, sid):
                        continue
                    w = opponent_weights.get(oid,0.15) / switch_weight_total
                    lead_v = matchup_matrix[lid].get(oid, 500.0)
                    if lead_v < 500:
                        pressure = min(2.8, max(0.5, (500.0 - lead_v) / 90.0))
                    else:
                        pressure = 0.10
                    sv = cached_scenario(sid, oid, 1, 1, 3, bait=True)
                    total += sv * pressure * w
                    wsum += pressure * w
                    if lead_v < 480:
                        if sv >= 60: rec += w
                        elif sv >= 52: stable += w
                safety = total / max(1e-9, wsum)
                safety_score = clamp(0.58 * safety + 0.42 * role_map[sid]["switch"])
                recovery = clamp(0.70 * rec * 100.0 + 0.30 * stable * 100.0)
                switch_cache[(lid,sid)] = (safety_score, recovery)

        def switch_safety(pid, lead_id):
            return switch_cache.get((lead_id,pid), (candidate_role_match.get(pid,50.0),50.0))[0]
        def recovery_score(pid, lead_id):
            return switch_cache.get((lead_id,pid), (50.0,50.0))[1]

        # Hız için takım oluşturma sırasında Training Analysis web isteği yapılmaz.
        # Bu veri yalnızca yardımcı sinyal olduğundan sonuç kalitesine göre maliyeti
        # gereksizdir; mevcut cache varsa bile ağır hesaplamaya dahil edilmez.
        training_bonus_value = 0.0

        progress(40, "4/7 Battle-scenario engine hazırlanıyor", "Shield • Energy • Bait • Timing • Recovery...")

        # Ağır battle-scenario aşaması için hızlı ön eleme: önce bireysel role +
        # meta + pair sinyaliyle en umut verici üçlüleri çıkaracağız.
        n = len(candidate_pool)
        triple_count = n * (n-1) * (n-2) // 6
        total_orders = triple_count * 6
        progress(46, "5/7 Takımlar senaryo motorundan geçiriliyor", f"{triple_count:,} üçlü × 6 dizilim = {total_orders:,} değerlendirme")

        teams, done, last_pct = [], 0, -1
        for a,b,c in combinations(candidate_pool,3):
            common = common_for(a,b,c)
            # Bariz zayıf çekirdekleri battle engine'e göndermeden eliyoruz.
            # Aynı Top-50 tehdidine üçlü ağır kayıp varsa bu takım yapısal olarak
            # güvenilmezdir; gereksiz 6 rol simülasyonunu da yapma.
            # Top-50 bir tehdide karşı 3/3 çok kötü eşleşme varsa takım ELENİR.
            # Bu, yüzde puanlamasından önce uygulanan gerçek bir Elite Gate'tir.
            if common.get("critical_threat_count", 0) >= 1:
                done += 6
                continue
            if common.get("critical_triple_risk", 0.0) >= 2.0:
                done += 6
                continue
            if common.get("triple_threat_risk", 0.0) > 28.0 and common["strong_coverage"] < 22.0:
                done += 6
                continue
            if common["triple_bad"] > 22.0 and common["strong_coverage"] < 18.0 and common["battle_robustness"] < 48.0:
                done += 6
                continue

            for lead,switch,closer in permutations((a,b,c),3):
                lid,sid,cid = sid_of(lead),sid_of(switch),sid_of(closer)
                lead_role = role_map[lid]["lead"]
                switch_role = role_map[sid]["switch"]
                closer_role = role_map[cid]["closer"]
                lead_mu = candidate_role_match[lid]
                switch_mu = switch_safety(sid,lid)
                closer_mu = candidate_role_match[cid]
                recovery = recovery_score(sid,lid)

                role_score = clamp(0.42*lead_role + 0.34*switch_role + 0.24*closer_role)
                role_match_score = clamp(0.40*lead_mu + 0.34*switch_mu + 0.26*closer_mu)
                battle_performance = clamp(0.52*common["battle_robustness"] + 0.28*role_match_score + 0.20*recovery)
                safety = switch_mu
                meta_strength = common["meta_score"]
                core = common["core"]
                consistency_score = common["consistency"]

                # Alignment + role fit + scenario robustness. Coverage ve meta
                # birbirinden özellikle ayrı tutulur.
                score = (
                    0.26 * battle_performance +
                    0.14 * role_score +
                    0.14 * common["coverage"] +
                    0.10 * safety +
                    0.08 * consistency_score +
                    0.08 * recovery +
                    0.07 * core +
                    0.05 * common["bulk"] +
                    0.08 * meta_strength
                )
                score -= min(8.0, common["hard_loss"] * 0.20)
                score -= min(7.0, common["alignment_risk"] * 0.10)
                score -= min(5.0, common["rps_dependency"] * 0.055)
                score -= min(4.0, common["bait_dependency"] * 0.045)
                score -= min(3.0, common["energy_dependency"] * 0.025)
                score -= min(3.0, common["double_bad"] * 0.035)
                score -= min(7.0, common["shared_threat_risk"] * 0.16)
                score -= min(9.0, common["triple_threat_risk"] * 0.42)
                score -= min(4.0, common["single_answer_dependency"] * 0.10)
                score -= min(7.0, common.get("severe_shared_risk", 0.0) * 0.18)
                score -= min(10.0, common.get("critical_triple_risk", 0.0) * 0.42)
                # Aynı güçlü tehdide 3/3 kötü eşleşme: rank 1-50 ise takımın
                # tek başına bu nedenle finale kalmasını engelle.
                if common.get("critical_triple_risk", 0.0) >= 2.0:
                    score -= 8.0
                if common.get("triple_threat_risk", 0.0) >= 10.0:
                    score -= 6.0
                if common.get("shared_threat_risk", 0.0) >= 22.0:
                    score -= 5.0
                score -= min(2.5, max(0.0, 50.0 - common["worst"]/10.0) * 0.04)
                score += max(-1.0, min(1.0, (common["diversity"] - 50.0) * 0.015))

                # Küçük training sinyali.
                if league in ("Great League","Ultra League","Master League"):
                    try:
                        cache_text = getattr(self, "_sinka_training_html_cache", {}).get(league, "")
                        hits = sum(1 for p in (lead,switch,closer) if norm_name(display_name(p)) and norm_name(display_name(p)) in cache_text)
                        if hits == 3: score += 0.45
                        elif hits == 2: score += 0.18
                    except Exception:
                        pass

                # Rol diziliminin kendi robustness'ı.
                ordering_stability = clamp(0.45*lead_role + 0.35*switch_role + 0.20*closer_role)
                score = 0.96*score + 0.04*ordering_stability

                # Açıklanabilir sonuç alanları.
                threat_names = []
                for oid in common["worst_threats"][:3]:
                    try:
                        op = next((q for q in opponent_pool if sid_of(q)==oid), None)
                        threat_names.append(display_name(op) if op else oid)
                    except Exception:
                        threat_names.append(oid)
                if common["alignment_risk"] <= 12 and common["good_coverage"] >= 75:
                    archetype = "ABC • Dengeli"
                elif common["pair_coverage"] >= 48 and common["alignment_risk"] > 18:
                    archetype = "ABB/ABA • Baskı"
                elif common["battle_robustness"] >= 62 and common["rps_dependency"] < 20:
                    archetype = "Core + Güvenli Değişim"
                else:
                    archetype = "Hibrit"

                reasons = []
                if common["good_coverage"] >= 70: reasons.append("Top-200'ün büyük bölümüne en az bir güçlü cevap veriyor")
                if common["strong_coverage"] >= 45: reasons.append("Çok güçlü matchup kapsaması yüksek")
                if recovery >= 65: reasons.append("Lead kaybettiğinde Safe Switch ile toparlanma güçlü")
                if common["rps_dependency"] < 18: reasons.append("RPS / alignment bağımlılığı düşük")
                if common["bait_dependency"] < 12: reasons.append("Bait'e düşük bağımlılık")
                if common["battle_robustness"] >= 60: reasons.append("Shield + enerji senaryolarında sağlam")
                if common["core"] >= 65: reasons.append("İkili/üçlü çekirdek sinerjisi güçlü")
                if common["shared_threat_risk"] < 12: reasons.append("Ortak zayıflık riski düşük; cevaplar birbirini tamamlıyor")
                elif common["shared_threat_risk"] < 22: reasons.append("Ortak tehditlerin çoğuna alternatif cevap bulunuyor")
                if not reasons: reasons.append("Meta, rol ve matchup sinyalleri dengeli")
                warnings = []
                if threat_names: warnings.append("Dikkat: " + ", ".join(threat_names))
                if common["triple_threat_risk"] > 6: warnings.append("🚨 Üçlü ortak tehdit riski yüksek")
                elif common["shared_threat_risk"] > 22: warnings.append("⚠ İki Pokémonu aynı anda zorlayan tehditler var")
                if common["single_answer_dependency"] > 18: warnings.append("Tek cevaba bağımlılık yüksek")
                if common["alignment_risk"] > 22: warnings.append("Alignment riski yüksek")
                if common["bait_dependency"] > 18: warnings.append("Bait'e bağımlılık artıyor")

                teams.append({
                    "lead": lead, "safe_switch": switch, "closer": closer,
                    "score": float(score),
                    "avg_score": round(common["avg_individual"],1),
                    "meta_score": round(common["meta_score"],1),
                    "synergy": round(common["core"],1),
                    "coverage": round(common["coverage"],1),
                    "consistency": round(common["consistency"],1),
                    "hard_loss": round(common["hard_loss"],1),
                    "alignment_risk": round(common["alignment_risk"],1),
                    "rps_dependency": round(common["rps_dependency"],1),
                    "good_coverage": round(common["good_coverage"],1),
                    "strong_coverage": round(common["strong_coverage"],1),
                    "double_bad": round(common["double_bad"],1),
                    "triple_bad": round(common["triple_bad"],1),
                    "recovery_score": round(recovery,1),
                    "bulk_score": round(common["bulk"],1),
                    "role_balance": round(common.get("role_balance", ordering_stability),1),
                    "worst_matchup": round(common["worst"],0),
                    "best_matchup": round(common["best"],0),
                    "lead_score": round(lead_mu,1),
                    "switch_score": round(switch_mu,1),
                    "closer_score": round(closer_mu,1),
                    "charger_score": round(role_map[sid]["charger"],1),
                    "attacker_score": round(role_map[cid]["attacker"],1),
                    "type_count": len(set().union(*(type_map.get(x,set()) for x in common["ids"]))),
                    "top200_size": len(opponent_pool),
                    "battle_robustness": round(common["battle_robustness"],1),
                    "bait_dependency": round(common["bait_dependency"],1),
                    "shield_dependency": round(common["shield_dependency"],1),
                    "energy_dependency": round(common["energy_dependency"],1),
                    "shared_threat_risk": round(common["shared_threat_risk"],1),
                    "triple_threat_risk": round(common["triple_threat_risk"],1),
                    "single_answer_dependency": round(common["single_answer_dependency"],1),
                    "severe_shared_risk": round(common.get("severe_shared_risk",0.0),1),
                    "critical_triple_risk": round(common.get("critical_triple_risk",0.0),1),
                    "critical_threat_count": int(common.get("critical_threat_count",0) or 0),
                    "same_type_pairs": int(common["same_type_pairs"]),
                    "role_balance": round(common["role_balance"],1),
                    "archetype": archetype,
                    "why": reasons[:4],
                    "warnings": warnings[:3],
                    "scenario_engine": "PvPoke matchup + shield/energy/bait/timing proxy",
                })

            done += 6
            pct = int(46 + (done / max(1,total_orders)) * 47)
            if pct != last_pct:
                last_pct = pct
                progress(pct, "5/7 Takımlar senaryo motorundan geçiriliyor", f"{done:,} / {total_orders:,} dizilim ({done/max(1,total_orders)*100:.1f}%)")

        # ================================================================
        # 5B — GERÇEK TURN-BY-TURN DOĞRULAMA
        # ================================================================
        # Tüm milyonlarca kombinasyonu ağır simüle etmek yerine önce hızlı
        # skorla daraltılmış adayların en iyilerini gerçek battle engine'e sokuyoruz.
        # Bu, performans ile doğruluk arasında kontrollü bir denge sağlar.
        if real_battle_engine is not None and teams:
            progress(84, "5/7 Gerçek 3v3 doğrulama", "En güçlü adaylar artık üçlü takım state-machine motorunda test ediliyor...")
            # 3v3 rakip havuzu: Top-200'ün üst kısmından farklı çekirdekler.
            # Her adayı yüzlerce rakip üçlüyle boğmak yerine gerçekçi ve çeşitli
            # 8 meta çekirdeği kullanıyoruz; ağır turn-by-turn yalnızca quick_top'ta çalışır.
            opp_base = opponent_pool[:9]
            opp_patterns = [(0,1,2),(0,2,4),(0,3,6),(1,4,7),(2,5,8),(0,5,8),(1,3,8),(2,4,7),(3,5,6),(0,4,8),(1,5,7),(2,3,8)]
            opponent_teams = []
            seen_opp_teams = set()
            for pat in opp_patterns:
                if max(pat) >= len(opp_base):
                    continue
                trio = tuple(opp_base[i] for i in pat)
                key = tuple(sorted(sid_of(x) for x in trio))
                if key not in seen_opp_teams:
                    seen_opp_teams.add(key); opponent_teams.append(trio)
                if len(opponent_teams) >= 6:
                    break
            # Ağır 1v1 doğrulama: yalnızca hızlı elemenin ilk 60 adayı.
            quick_top = sorted(teams, key=lambda x: x.get("score", 0), reverse=True)[:60]
            real_opponents = opponent_pool[:10]

            def team_rating_lookup(attacker_sid, defender_sid):
                row = None
                for q in opponent_pool:
                    if sid_of(q) == attacker_sid:
                        row = q; break
                if row is None:
                    for q in quick_top[0:1]:
                        for pos in ("lead","safe_switch","closer"):
                            if sid_of(q.get(pos, {})) == attacker_sid:
                                row = q.get(pos); break
                if row is None:
                    return None
                for m in row.get("matchups", []) or []:
                    if isinstance(m, dict) and str(m.get("opponent")) == defender_sid:
                        try: return float(m.get("rating",500) or 500)
                        except Exception: return 500.0
                return 500.0

            def real_3v3_eval(team):
                team_list = [team.get("lead"), team.get("safe_switch"), team.get("closer")]
                if any(not isinstance(x, dict) for x in team_list) or not opponent_teams:
                    return None
                lookup_rows = list(opponent_pool) + team_list
                def local_rating(attacker_sid, defender_sid):
                    row = next((q for q in lookup_rows if sid_of(q) == attacker_sid), None)
                    if row is None:
                        return 500.0
                    for m in row.get("matchups", []) or []:
                        if isinstance(m, dict) and str(m.get("opponent")) == defender_sid:
                            try: return float(m.get("rating",500) or 500)
                            except Exception: return 500.0
                    return 500.0
                vals=[]; weighted=0.0; wsum=0.0; cmp_total=bait_a=bait_s=switches=farm=0; turns_total=0; samples=0
                for opp_team in opponent_teams:
                    # Rakibin üç olası lead'i; bizim dizilimimiz rol sırasını korur.
                    opp_weight = sum(opponent_weights.get(sid_of(x),0.15) for x in opp_team)/3.0
                    for opp_lead in range(3):
                        for sh in (2,1):
                            res = real_battle_engine.simulate_3v3(
                                team_list, list(opp_team), lead_a=0, lead_b=opp_lead, shields=sh,
                                rating_lookup=local_rating, bait_mode=True
                            )
                            if not res: continue
                            vals.append(res["score"]); weighted += res["score"]*opp_weight; wsum += opp_weight
                            cmp_total += int(res.get("cmp",0)); bait_a += int(res.get("bait_attempts",0)); bait_s += int(res.get("bait_successes",0))
                            switches += int(res.get("switches",0)); farm += int(res.get("farm_events",0)); turns_total += int(res.get("turns",0)); samples += 1
                if not vals: return None
                return {"score":weighted/max(1e-9,wsum),"cmp_rate":cmp_total/max(1,samples),
                        "bait_rate":bait_s/max(1,bait_a),"switches":switches/max(1,samples),
                        "farm":farm/max(1,samples),"avg_turns":turns_total/max(1,samples),"samples":samples}

            def real_role_eval(actor, role):
                if not isinstance(actor, dict):
                    return None
                if role == "lead":
                    shield_pairs = ((2,2), (1,1), (0,0))
                    energies = (0, 3, -3)
                elif role == "switch":
                    shield_pairs = ((1,1), (1,2), (2,1))
                    energies = (3, 0, -3)
                else:
                    shield_pairs = ((0,0), (0,1), (1,0))
                    energies = (0, 3, -3)

                vals = []
                weighted = 0.0
                wsum = 0.0
                cmp_total = 0
                bait_attempts = 0
                bait_successes = 0
                turns_total = 0
                count = 0
                for opp in real_opponents:
                    if sid_of(opp) == sid_of(actor):
                        continue
                    ow = opponent_weights.get(sid_of(opp), 0.15)
                    for sh_a, sh_b in shield_pairs:
                        for energy in energies:
                            for bait in (False, True):
                                result = real_battle_engine.scenario(
                                    actor, opp,
                                    shields=(sh_a, sh_b),
                                    energy_delta=energy,
                                    bait=bait
                                )
                                if not result:
                                    continue
                                vals.append(result["score"])
                                weighted += result["score"] * ow
                                wsum += ow
                                cmp_total += int(result.get("cmp", 0) or 0)
                                bait_attempts += int(result.get("bait_attempts", 0) or 0)
                                bait_successes += int(result.get("bait_successes", 0) or 0)
                                turns_total += int(result.get("turns", 0) or 0)
                                count += 1
                if not vals:
                    return None
                return {
                    "score": weighted / max(1e-9, wsum),
                    "cmp_rate": (cmp_total / max(1, count)),
                    "bait_rate": bait_successes / max(1, bait_attempts),
                    "avg_turns": turns_total / max(1, count),
                    "samples": count,
                }

            for idx, team in enumerate(quick_top, 1):
                lead_r = real_role_eval(team.get("lead"), "lead")
                switch_r = real_role_eval(team.get("safe_switch"), "switch")
                closer_r = real_role_eval(team.get("closer"), "closer")
                if not all((lead_r, switch_r, closer_r)):
                    continue
                real_score = clamp(
                    0.40 * lead_r["score"] +
                    0.34 * switch_r["score"] +
                    0.26 * closer_r["score"]
                )
                # Gerçek motor artık final puanın anlamlı bir parçası; hızlı
                # Top-200/coverage skoru ise geniş meta görünümünü koruyor.
                old_score = float(team.get("score", 0) or 0)
                team["proxy_score"] = round(old_score, 1)
                team["real_battle_score"] = round(real_score, 1)
                team["real_lead_score"] = round(lead_r["score"], 1)
                team["real_switch_score"] = round(switch_r["score"], 1)
                team["real_closer_score"] = round(closer_r["score"], 1)
                team["cmp_rate"] = round((lead_r["cmp_rate"] + switch_r["cmp_rate"] + closer_r["cmp_rate"]) / 3.0, 2)
                team["bait_success_rate"] = round((lead_r["bait_rate"] + switch_r["bait_rate"] + closer_r["bait_rate"]) / 3.0 * 100.0, 1)
                team["battle_turns_avg"] = round((lead_r["avg_turns"] + switch_r["avg_turns"] + closer_r["avg_turns"]) / 3.0, 1)
                team["battle_samples"] = int(lead_r["samples"] + switch_r["samples"] + closer_r["samples"])
                team["score"] = round(0.65 * old_score + 0.35 * real_score, 3)
                team["scenario_engine"] = "SinKA turn-by-turn • HP/Energy/Shield/CMP/Bait"
                if real_score >= old_score + 3:
                    team.setdefault("why", []).append("Gerçek turn-by-turn simülasyonda proxy skorundan daha güçlü")
                elif real_score + 3 < old_score:
                    team.setdefault("warnings", []).append("Turn-by-turn simülasyonda proxy skorundan zayıf")
                if idx % 25 == 0:
                    progress(86 + int(idx / max(1, len(quick_top)) * 5), "5/7 1v1 turn-by-turn + rol doğrulama", f"{idx}/{len(quick_top)} güçlü aday işlendi")

            # 3v3 state-machine en pahalı katman. 90+ arayışında ilk 60 güçlü
            # aday gerçek 3v3'e alınır; böylece proxy puanı biraz düşük olan ama
            # gerçek savaşta çok güçlü farklı core'lar kaybolmaz.
            three_v3_top = sorted([t for t in quick_top if t.get("real_battle_score", 0) > 0],
                                  key=lambda x: x.get("score",0), reverse=True)[:60]
            for idx, team in enumerate(three_v3_top, 1):
                res3 = real_3v3_eval(team)
                if not res3:
                    continue
                team["real_3v3_score"] = round(res3["score"],1)
                team["three_v_three_samples"] = int(res3["samples"])
                team["three_v_three_turns"] = round(res3["avg_turns"],1)
                team["three_v_three_switches"] = round(res3["switches"],2)
                team["three_v_three_cmp"] = round(res3["cmp_rate"],2)
                team["three_v_three_bait"] = round(res3["bait_rate"]*100.0,1)
                team["three_v_three_farm"] = round(res3["farm"],2)
                before = float(team.get("score",0) or 0)
                team["score"] = round(0.72*before + 0.28*res3["score"],3)
                team["scenario_engine"] = "SinKA 3v3 state-machine • HP/Energy/Shield/CMP/Switch/Farm/Bait"
                if res3["score"] >= before + 3:
                    team.setdefault("why",[]).append("3v3 gerçek state-machine simülasyonunda güçlü")
                elif res3["score"] + 3 < before:
                    team.setdefault("warnings",[]).append("3v3 state-machine sonucunda proxy skordan zayıf")
                if idx % 25 == 0:
                    progress(91 + int(idx / max(1, len(three_v3_top)) * 3), "5/7 Gerçek 3v3 doğrulama", f"{idx}/{len(three_v3_top)} takım 3v3 simüle edildi")

        # ================================================================
        # 5C — YENİ 100 PUAN ELITE TEAM SCORER
        # ================================================================
        # Puan artık mevcut proxy skorun üstüne eklenmiyor. Her takım aynı 100
        # puanlık ölçeğe yeniden normalize edilir. Böylece 90+ gerçekten Elite
        # seviyesini temsil eder; puan yapay olarak şişirilmez.
        def elite_team_score(team):
            three = float(team.get("real_3v3_score", team.get("battle_robustness", 50.0)) or 50.0)
            one = float(team.get("real_battle_score", team.get("battle_robustness", 50.0)) or 50.0)
            coverage = clamp(float(team.get("coverage", 50.0) or 50.0))
            shared = float(team.get("shared_threat_risk", 0.0) or 0.0)
            severe = float(team.get("severe_shared_risk", 0.0) or 0.0)
            triple = float(team.get("triple_threat_risk", 0.0) or 0.0)
            critical = float(team.get("critical_triple_risk", 0.0) or 0.0)
            threat_safety = clamp(100.0 - 0.55*shared - 0.85*severe - 1.10*triple - 1.35*critical)
            role = clamp(0.40*float(team.get("lead_score", 50.0) or 50.0) +
                          0.35*float(team.get("switch_score", 50.0) or 50.0) +
                          0.25*float(team.get("closer_score", 50.0) or 50.0))
            bait_dep = float(team.get("bait_dependency", 0.0) or 0.0)
            shield_dep = float(team.get("shield_dependency", 0.0) or 0.0)
            energy_dep = float(team.get("energy_dependency", 0.0) or 0.0)
            shield_energy = clamp(100.0 - 0.45*bait_dep - 0.35*shield_dep - 0.25*energy_dep)
            synergy = clamp(float(team.get("synergy", 50.0) or 50.0))
            bulk = clamp(float(team.get("bulk_score", 50.0) or 50.0))
            meta = clamp(float(team.get("meta_score", 50.0) or 50.0))
            score = (
                0.25*three +
                0.20*coverage +
                0.15*threat_safety +
                0.10*role +
                0.08*shield_energy +
                0.07*one +
                0.07*synergy +
                0.04*bulk +
                0.04*meta
            )
            # Elite için zorunlu güvenlik kapıları.
            if int(team.get("critical_threat_count", 0) or 0) > 0:
                return 0.0, {"three_v3":three,"coverage":coverage,"threat_safety":0.0,"role":role,"shield_energy":shield_energy,"one_v_one":one,"synergy":synergy,"bulk":bulk,"meta":meta}
            if shared >= 55.0 or severe >= 30.0 or triple >= 18.0:
                score -= min(18.0, 0.22*shared + 0.28*severe + 0.35*triple)
            if float(team.get("single_answer_dependency", 0.0) or 0.0) >= 25.0:
                score -= min(8.0, (float(team.get("single_answer_dependency",0.0))-25.0)*0.30)
            if three < 70.0:
                score -= (70.0-three)*0.20
            return clamp(score), {"three_v3":three,"coverage":coverage,"threat_safety":threat_safety,"role":role,"shield_energy":shield_energy,"one_v_one":one,"synergy":synergy,"bulk":bulk,"meta":meta}

        # Ön doğrulama olmayan adaylarda da aynı final ölçeği kullanılır; gerçek
        # 3v3 sonucu varsa o sonucu merkez alır.
        for team in teams:
            final_score, components = elite_team_score(team)
            team["score_components"] = components
            team["elite_score"] = round(final_score, 3)
            team["score"] = round(final_score, 3)
            if final_score >= 90.0:
                team.setdefault("why", []).append("90+ ELITE: 3v3 + coverage + threat safety + rol dengesi birlikte yüksek")
            elif final_score >= 85.0:
                team.setdefault("why", []).append("85+ VERY STRONG: Elite seviyesine yakın dengeli takım")

        # Gerçek 3v3 sonuçları geldiyse final puanı bir kez daha güncelle.
        if real_battle_engine is not None:
            for team in teams:
                final_score, components = elite_team_score(team)
                team["score_components"] = components
                team["elite_score"] = round(final_score, 3)
                team["score"] = round(final_score, 3)

        if real_battle_engine is not None:
            # Gerçek motor erişilebilirken 3v3 doğrulaması yapılmamış takımın
            # sıfır/boş simülasyonla finale sızmasına izin verme.
            verified = [t for t in teams if int(t.get("three_v_three_samples", 0) or 0) >= 18]
            if verified:
                teams = verified
            else:
                # Motor veri eksikliği yaşarsa kullanıcıyı boş bırakmamak için
                # yalnızca gerçek battle sonucu bulunan adayları kullan.
                teams = [t for t in teams if float(t.get("real_battle_score", 0) or 0) > 0]

        progress(95, "6/7 En iyi takımlar seçiliyor", "90+ Elite Gate + takım çeşitliliği + core tekrar kontrolü...")
        # Aynı üç Pokémonun farklı rol sıralamalarını tek takım say.
        best_by_triple = {}
        for t in teams:
            key = tuple(sorted(sid_of(t[x]) for x in ("lead","safe_switch","closer")))
            if key not in best_by_triple or t["score"] > best_by_triple[key]["score"]:
                best_by_triple[key] = t
        ranked = sorted(best_by_triple.values(), key=lambda x: x.get("score",0), reverse=True)

        # 90+ takımlar önce gelir. 85+ ikinci kalite katmanı, sonra en iyi kalanlar.
        elite = [t for t in ranked if t.get("score",0) >= 90.0]
        very_strong = [t for t in ranked if 85.0 <= t.get("score",0) < 90.0]
        strong = [t for t in ranked if t.get("score",0) < 85.0]

        def diversity_penalty(cand, selected):
            ids = {sid_of(cand[x]) for x in ("lead","safe_switch","closer")}
            penalty = 0.0
            for old in selected:
                old_ids = {sid_of(old[x]) for x in ("lead","safe_switch","closer")}
                shared = len(ids & old_ids)
                if shared >= 2:
                    penalty += 5.0 + 2.5*(shared-2)
                # Aynı core davranışına da küçük ceza.
                if cand.get("archetype") and cand.get("archetype") == old.get("archetype") and shared >= 1:
                    penalty += 1.5
            return penalty

        def select_diverse(pool, target=5):
            selected = []
            remaining = list(pool)
            while remaining and len(selected) < target:
                best = max(remaining, key=lambda t: float(t.get("score",0) or 0) - diversity_penalty(t, selected))
                selected.append(best)
                remaining.remove(best)
            return selected

        result = select_diverse(elite, 5)
        if len(result) < 5:
            result.extend(select_diverse([t for t in very_strong if t not in result], 5-len(result)))
        if len(result) < 5:
            result.extend(select_diverse([t for t in strong if t not in result], 5-len(result)))
        if len(result) < 5:
            result.extend(select_diverse([t for t in ranked if t not in result], 5-len(result)))

        # Sonuç kategorileri artık doğrudan gerçek final puanını yansıtır.
        for idx, t in enumerate(result[:5], 1):
            sc = float(t.get("score",0) or 0)
            if sc >= 90.0:
                t["category"] = "ELITE 90+"
            elif sc >= 85.0:
                t["category"] = "VERY STRONG"
            elif sc >= 80.0:
                t["category"] = "STRONG"
            else:
                t["category"] = "ALTERNATİF"
            t["selection_rank"] = idx

        progress(100, "✓ Tamamlandı", f"Elite 90+ scorer + Top-200 threat + 1v1 + 3v3 state-machine + diversity tamamlandı. {len(result[:5])} takım hazır.")
        return result[:5]

    def _open_team_progress(self, league, include_shadows=False, owned_only=False):
        win = tk.Toplevel(self.root)
        self._team_progress_window = win
        win.title("⚔ SinKA META TAKIM MOTORU")
        win.geometry("650x430")
        win.resizable(True, True)
        win.transient(self.root)
        win.protocol("WM_DELETE_WINDOW", lambda: win.withdraw())

        ttk.Label(win, text="⚔ SinKA META TAKIM MOTORU", font=("Segoe UI", 16, "bold")).pack(pady=(14,4))
        ttk.Label(win, text=f"Lig: {league}  •  Güncel PvPoke verisi + Top-200 + 3v3 Battle Engine", font=("Segoe UI", 10)).pack(pady=(0,4))
        ttk.Label(win, text=f"Shadow: {'dahil' if include_shadows else 'hariç'}  •  Aday havuzu: {'sadece Mevcut Pokémonlar' if owned_only else 'tüm meta'}", foreground="#555555").pack(pady=(0,8))
        self._team_progress_stage = tk.StringVar(value="İşlem başlatılıyor...")
        ttk.Label(win, textvariable=self._team_progress_stage, font=("Segoe UI", 11, "bold")).pack(pady=4)
        self._team_progress_bar = ttk.Progressbar(win, mode="determinate", maximum=100)
        self._team_progress_bar.pack(fill="x", padx=25, pady=(5,10))
        frame = ttk.Frame(win); frame.pack(fill="both", expand=True, padx=20, pady=5)
        self._team_progress_text = tk.Text(frame, height=14, wrap="word", state="disabled", font=("Consolas",9))
        scroll = ttk.Scrollbar(frame, command=self._team_progress_text.yview); self._team_progress_text.configure(yscrollcommand=scroll.set)
        self._team_progress_text.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        self._team_progress_detail = tk.StringVar(value="İşlem devam ediyor...")
        ttk.Label(win, textvariable=self._team_progress_detail).pack(pady=5)
        self._team_progress_close = ttk.Button(win, text="Arka Plana Al", command=win.withdraw)
        self._team_progress_close.pack(pady=(0,12))
        self._append_team_progress("▶ Takım motoru başlatıldı.")
        self._append_team_progress(f"▶ {league} için güncel veri isteniyor...")
        win.update_idletasks()
        return win

    def _append_team_progress(self, text):
        try:
            box = self._team_progress_text
            box.configure(state="normal")
            box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n")
            box.see("end")
            box.configure(state="disabled")
        except Exception:
            pass

    def _update_team_progress(self, percent, stage, detail=""):
        try:
            win = self._team_progress_window
            if not win or not win.winfo_exists(): return
            old = getattr(self, "_team_progress_percent", -1)
            self._team_progress_percent = percent
            self._team_progress_bar["value"] = percent
            self._team_progress_stage.set(stage)
            self._team_progress_detail.set(detail)
            if percent == 100 or int(percent//2) != int(old//2):
                self._append_team_progress(f"{stage}: {detail}" if detail else stage)
            win.update_idletasks()
        except Exception:
            pass

    def build_team(self):
        """Takım motorunu anında ilerleme penceresiyle başlatır.
        İlk arama yalnızca kullanıcının Mevcut Pokémon havuzundan yapılır;
        sonuç penceresinde bu filtre kapatılarak tüm meta havuzu seçilebilir.
        """
        if getattr(self, "_team_building", False):
            try:
                self._team_progress_window.deiconify(); self._team_progress_window.lift(); self._team_progress_window.focus_force()
            except Exception:
                pass
            return
        try:
            old = getattr(self, "_team_window", None)
            if old is not None and old.winfo_exists():
                old.lift(); old.focus_force(); return
        except Exception:
            self._team_window = None

        league = self.league_var.get()
        # Ayarlardaki takım kaynağına göre ilk arama havuzunu seç.
        # Varsayılan ve önerilen seçenek Mevcut Pokémonlar'dır.
        owned_only = (self.team_source_var.get() != 'Güncel Meta')
        self._start_team_build(league, include_shadows=False, owned_only=owned_only)

    def _start_team_build(self, league, include_shadows=False, owned_only=False):
        """Takım hesaplamasını worker thread'de başlatır ve ilerlemeyi ana thread'e aktarır."""
        if getattr(self, "_team_building", False):
            try:
                self._team_progress_window.deiconify(); self._team_progress_window.lift()
            except Exception:
                pass
            return

        self._team_building = True
        self._team_progress_percent = -1
        self.set_status(f"{league}: takım motoru çalışıyor...")
        self._open_team_progress(league, include_shadows, owned_only)

        def progress(percent, stage, detail=""):
            try:
                self.root.after(0, lambda p=percent,s=stage,d=detail: self._update_team_progress(p,s,d))
            except Exception:
                pass

        def worker():
            try:
                progress(3, "1/7 Güncel PvPoke sıralaması indiriliyor", "Butona basıldığı anda yeni veri çekiliyor...")
                ranking = self._fresh_team_ranking(league)
                if not ranking:
                    raise RuntimeError("PvPoke sıralama verisi boş geldi.")

                progress(10, "2/7 Takım filtresi hazırlanıyor", f"Shadow: {'dahil' if include_shadows else 'hariç'} • Mevcut: {'sadece mevcut' if owned_only else 'tüm meta'}")
                teams = self._generate_meta_teams(
                    league,
                    ranking,
                    show_shadows=include_shadows,
                    progress_cb=progress,
                    owned_only=owned_only
                )

                def done():
                    self._team_building = False
                    self._update_team_progress(100, "✓ Tamamlandı", f"{len(teams)} takım bulundu.")
                    self.set_status(f"{league}: güncel meta ile {len(teams)} takım oluşturuldu.")
                    try:
                        self._team_progress_window.destroy()
                    except Exception:
                        pass
                    self._show_generated_teams(teams, league, include_shadows, owned_only)

                self.root.after(0, done)
            except Exception as e:
                def failed(err=str(e)):
                    self._team_building = False
                    self._update_team_progress(100, "✗ İşlem başarısız", err)
                    self._append_team_progress("✗ HATA: " + err)
                    self.set_status("Takım oluşturma hatası.")
                self.root.after(0, failed)

        threading.Thread(target=worker, daemon=True).start()

    def _create_team_a4_image(self, teams, league, shadow_var=None, owned_var=None):
        """Takım sonuçlarını A4 yatay, tek PNG paftaya dönüştürür."""
        if not PIL_OK:
            messagebox.showerror("SinKA PvP", "A4 görsel oluşturmak için Pillow gerekli. CMD'de: py -m pip install pillow")
            return
        if not teams:
            messagebox.showinfo("SinKA PvP", "Görsel oluşturulacak takım bulunamadı.")
            return

        path = filedialog.asksaveasfilename(
            title="Takım sonuçlarını A4 yatay görsel olarak kaydet",
            defaultextension=".png",
            filetypes=[("PNG görseli", "*.png"), ("JPEG görseli", "*.jpg;*.jpeg"), ("Tüm dosyalar", "*.*")],
            initialfile=f"SinKA_{league.replace(' ', '_')}_Takimlar_A4.png"
        )
        if not path:
            return

        include_shadows = bool(shadow_var.get()) if shadow_var is not None else False
        owned_only = bool(owned_var.get()) if owned_var is not None else False

        # Worker'ın Tkinter nesnelerine erişmemesi için tüm metin ve görsel verisini
        # ana thread'de mümkün olduğunca snapshot olarak hazırla.
        team_snapshot = []
        needed_names = set()
        for idx, team in enumerate(teams[:5], 1):
            roles = []
            for role_key, role_title in (("lead", "LİDER / LEAD"), ("safe_switch", "GÜVENLİ DEĞİŞİM / SAFE SWITCH"), ("closer", "KAPATICI / CLOSER")):
                pdata = team.get(role_key, {}) if isinstance(team, dict) else {}
                if not isinstance(pdata, dict):
                    pdata = {}
                name = str(pdata.get("speciesName") or pdata.get("name") or pdata.get("speciesId") or "Bilinmeyen")
                types = self._team_type_keys(name)
                moves = pdata.get("moveset") or pdata.get("moves") or []
                if isinstance(moves, dict):
                    moves = list(moves.values())
                if not isinstance(moves, (list, tuple)):
                    moves = [moves]
                move_names = []
                move_types = []
                for mv in list(moves)[:3]:
                    try:
                        move_names.append(self.translate_move(str(mv)))
                    except Exception:
                        move_names.append(str(mv))
                    move_types.append(self._team_move_type_key(mv))
                while len(move_names) < 3:
                    move_names.append("-")
                    move_types.append("")
                roles.append({
                    "role": role_title,
                    "name": name,
                    "types": list(types)[:2],
                    "moves": move_names[:3],
                    "move_types": move_types[:3],
                })
                if name:
                    needed_names.add(name)

            team_snapshot.append({
                "number": idx,
                "score": float(team.get("score", 0) or 0),
                "meta": float(team.get("meta_score", 0) or 0),
                "synergy": float(team.get("synergy", 0) or 0),
                "coverage": float(team.get("coverage", 0) or 0),
                "consistency": float(team.get("consistency", 0) or 0),
                "lead_score": float(team.get("lead_score", 0) or 0),
                "switch_score": float(team.get("switch_score", 0) or 0),
                "closer_score": float(team.get("closer_score", 0) or 0),
                "good_coverage": float(team.get("good_coverage", 0) or 0),
                "alignment_risk": float(team.get("alignment_risk", 0) or 0),
                "rps_dependency": float(team.get("rps_dependency", 0) or 0),
                "recovery_score": float(team.get("recovery_score", 0) or 0),
                "bulk_score": float(team.get("bulk_score", 0) or 0),
                "strong_coverage": float(team.get("strong_coverage", 0) or 0),
                "shared_threat_risk": float(team.get("shared_threat_risk", 0) or 0),
                "triple_threat_risk": float(team.get("triple_threat_risk", 0) or 0),
                "single_answer_dependency": float(team.get("single_answer_dependency", 0) or 0),
                "severe_shared_risk": float(team.get("severe_shared_risk", 0) or 0),
                "critical_triple_risk": float(team.get("critical_triple_risk", 0) or 0),
                "same_type_pairs": int(team.get("same_type_pairs", 0) or 0),
                "real_battle_score": float(team.get("real_battle_score", 0) or 0),
                "cmp_rate": float(team.get("cmp_rate", 0) or 0),
                "bait_success_rate": float(team.get("bait_success_rate", 0) or 0),
                "battle_turns_avg": float(team.get("battle_turns_avg", 0) or 0),
                "real_3v3_score": float(team.get("real_3v3_score", 0) or 0),
                "three_v_three_samples": int(team.get("three_v_three_samples", 0) or 0),
                "three_v_three_turns": float(team.get("three_v_three_turns", 0) or 0),
                "three_v_three_switches": float(team.get("three_v_three_switches", 0) or 0),
                "three_v_three_cmp": float(team.get("three_v_three_cmp", 0) or 0),
                "three_v_three_bait": float(team.get("three_v_three_bait", 0) or 0),
                "three_v_three_farm": float(team.get("three_v_three_farm", 0) or 0),
                "scenario_engine": str(team.get("scenario_engine", "")),
                "roles": roles,
            })

        image_snapshot = {}
        for name in needed_names:
            try:
                data = self.image_bytes.get(name)
                if data:
                    image_snapshot[name] = data
            except Exception:
                pass

        type_snapshot = dict(getattr(self, "type_icon_bytes", {}) or {})
        snapshot = {
            "league": league,
            "cp_max": LEAGUES.get(league, 1500),
            "include_shadows": include_shadows,
            "owned_only": owned_only,
            "teams": team_snapshot,
            "image_bytes": image_snapshot,
            "type_icon_bytes": type_snapshot,
            "type_names": dict(getattr(self, "TYPE_TR", {}) or {}),
            "needed_names": list(needed_names),
        }
        self.set_status("Takım A4 görseli oluşturuluyor... Program çalışmaya devam edecek.")
        threading.Thread(target=self._save_team_a4_png_worker, args=(path, snapshot), daemon=True).start()

    def _save_team_a4_png_worker(self, path, snapshot):
        """Takım sonuç paftasını Tkinter'a dokunmadan arka planda üretir."""
        try:
            from PIL import ImageDraw, ImageFont
            from datetime import datetime

            W, H = 3508, 2480  # A4 yatay, 300 DPI
            img = Image.new("RGBA", (W, H), (247, 249, 252, 255))
            draw = ImageDraw.Draw(img)

            NAVY = (10, 35, 82)
            NAVY2 = (20, 55, 112)
            BLUE = (42, 111, 220)
            CYAN = (30, 194, 180)
            GOLD = (239, 174, 44)
            TEXT = (24, 38, 62)
            MUTED = (86, 98, 118)
            WHITE = (255, 255, 255)
            GRID = (214, 222, 232)
            GREEN = (0, 145, 76)
            RED = (205, 46, 54)

            def load_font(size, bold=False):
                candidates = [
                    r"C:\\Windows\\Fonts\\segoeuib.ttf" if bold else r"C:\\Windows\\Fonts\\segoeui.ttf",
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                    "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
                ]
                for f in candidates:
                    try:
                        return ImageFont.truetype(f, size)
                    except Exception:
                        pass
                return ImageFont.load_default()

            f_brand = load_font(54, True)
            f_title = load_font(76, True)
            f_subtitle = load_font(30, True)
            f_team = load_font(34, True)
            f_role = load_font(23, True)
            f_name = load_font(37, True)
            f_text = load_font(24, False)
            f_small = load_font(21, False)
            f_small_bold = load_font(21, True)
            f_score = load_font(45, True)
            f_footer = load_font(22, False)

            # Üst başlık alanı
            margin = 52
            header_h = 245
            draw.rectangle((margin, margin, W-margin, margin+header_h), fill=WHITE, outline=NAVY, width=4)
            draw.rectangle((margin, margin+header_h-12, W-margin, margin+header_h), fill=NAVY)
            draw.text((margin+35, margin+36), "SINKA", fill=NAVY, font=f_brand)
            draw.line((margin+35, margin+104, margin+260, margin+104), fill=CYAN, width=6)
            league = snapshot.get("league", "Great League")
            cp_max = snapshot.get("cp_max", 1500)
            draw.text((W//2, margin+68), league.upper(), fill=NAVY, font=f_title, anchor="mm")
            team_cp_text = "SINIRSIZ CP MAX" if league == "Master League" else f"{cp_max} CP MAX"
            draw.text((W//2, margin+145), f"{team_cp_text}  •  EN İYİ 5 TAKIM", fill=NAVY2, font=f_subtitle, anchor="mm")
            date_text = datetime.now().strftime("%d.%m.%Y")
            draw.text((W-margin-35, margin+48), date_text, fill=MUTED, font=f_small, anchor="ra")
            filter_text = "Shadow dahil" if snapshot.get("include_shadows") else "Shadow hariç"
            if snapshot.get("owned_only"):
                filter_text += "  •  Sadece mevcut Pokémonlar"
            else:
                filter_text += "  •  Meta havuzu"
            draw.text((W-margin-35, margin+90), filter_text, fill=GREEN if not snapshot.get("include_shadows") else RED, font=f_small_bold, anchor="ra")

            # Takımları 2 sütun + alt tek geniş kart şeklinde yerleştir.
            teams = snapshot.get("teams", [])[:5]
            card_gap = 28
            left = margin
            right = W-margin
            grid_top = margin + header_h + 28
            card_w = (right-left-card_gap)//2
            card_h = 600
            positions = []
            for i in range(min(4, len(teams))):
                row = i//2
                col = i%2
                x = left + col*(card_w+card_gap)
                y = grid_top + row*(card_h+card_gap)
                positions.append((x,y,card_w,card_h))
            if len(teams) > 4:
                y = grid_top + 2*(card_h+card_gap)
                positions.append((left, y, right-left, 600))

            image_cache = dict(snapshot.get("image_bytes", {}) or {})
            type_icons = dict(snapshot.get("type_icon_bytes", {}) or {})
            type_names = dict(snapshot.get("type_names", {}) or {})

            def get_image(name):
                data = image_cache.get(name)
                if data:
                    return data
                try:
                    # Worker thread olduğu için ağ erişimi güvenlidir.
                    data = self.download_image_bytes(name)
                    if data:
                        image_cache[name] = data
                        return data
                except Exception:
                    pass
                return None

            for idx, team in enumerate(teams):
                if idx >= len(positions):
                    break
                x, y, cw, ch = positions[idx]
                draw.rounded_rectangle((x, y, x+cw, y+ch), radius=18, fill=WHITE, outline=GRID, width=3)
                # takım başlığı
                draw.rectangle((x, y, x+cw, y+72), fill=NAVY2)
                draw.text((x+24, y+36), f"TAKIM {team['number']}", fill=WHITE, font=f_team, anchor="lm")
                draw.text((x+cw-24, y+36), f"{team['score']:.1f} / 100", fill=WHITE, font=f_score, anchor="rm")

                metrics = f"Meta {team['meta']:.1f}   •   Sinerji {team['synergy']:.1f}   •   Coverage {team['coverage']:.1f}   •   Consistency {team['consistency']:.1f}   •   Recovery {team.get('recovery_score',0):.1f}   •   1v1 {team.get('real_battle_score',0):.1f}   •   3v3 {team.get('real_3v3_score',0):.1f}"
                draw.text((x+24, y+99), metrics, fill=MUTED, font=f_small_bold)
                sub = f"Lead {team['lead_score']:.1f}  |  Switch {team['switch_score']:.1f}  |  Closer {team['closer_score']:.1f}  |  Top-200 avantaj: {team['good_coverage']:.1f}%  |  3v3 örnek: {team.get('three_v_three_samples',0)}  |  Switch: {team.get('three_v_three_switches',0):.2f}  |  CMP: {team.get('three_v_three_cmp',0):.2f}"
                draw.text((x+24, y+132), sub, fill=MUTED, font=f_small)
                threat_line = f"Ortak tehdit: {team.get('shared_threat_risk',0):.1f}%  •  Ciddi ortak açık: {team.get('severe_shared_risk',0):.1f}%  •  Üçlü tehdit: {team.get('triple_threat_risk',0):.1f}%  •  Kritik üçlü: {team.get('critical_triple_risk',0):.1f}%  •  Tek cevaba bağımlılık: {team.get('single_answer_dependency',0):.1f}%"
                draw.text((x+24, y+157), threat_line, fill=RED if team.get('triple_threat_risk',0) > 6 else MUTED, font=f_small_bold)

                # Rol satırı: isim/tip orta alanda, büyük Pokémon görseli sağda.
                role_y = y+185
                role_h = 116
                for r_i, role in enumerate(team["roles"]):
                    ry = role_y + r_i*role_h
                    if r_i % 2 == 1:
                        draw.rounded_rectangle((x+15, ry-5, x+cw-15, ry+role_h-8), radius=10, fill=(247,249,252))
                    draw.text((x+27, ry+18), role["role"], fill=NAVY, font=f_role)
                    name = role["name"]
                    # Sağda büyük görsel alanı
                    data = get_image(name)
                    image_box_x = x+cw-235
                    image_box_y = ry+5
                    if data:
                        try:
                            pimg = Image.open(BytesIO(data)).convert("RGBA")
                            pimg.thumbnail((104,104), Image.LANCZOS)
                            px = image_box_x + (104-pimg.width)//2
                            py = image_box_y + (104-pimg.height)//2
                            img.alpha_composite(pimg, (px, py))
                        except Exception:
                            pass
                    # İsim/tip orta alanda; büyük görsel sağda. Hareketlerin
                    # element ikonları isimlerinin hemen solunda gösterilir.
                    draw.text((x+27, ry+57), name, fill=TEXT, font=f_name, anchor="lm")
                    type_text = " / ".join(type_names.get(str(k).lower(), str(k).title()) for k in role.get("types", [])) or "-"
                    draw.text((x+27, ry+91), type_text, fill=MUTED, font=f_small)
                    moves = role.get("moves", ["-","-","-"])
                    move_types = role.get("move_types", ["","",""])
                    move_x = image_box_x - 28
                    labels = ("Hızlı", "Güç 1", "Güç 2")
                    for mi, (lab, mv) in enumerate(zip(labels, moves)):
                        yy = ry + 43 + mi*25
                        txt = f"{lab}: {mv}"
                        try:
                            tw = draw.textbbox((0,0), txt, font=f_small_bold)[2]
                        except Exception:
                            tw = len(txt)*11
                        icon_key = str(move_types[mi] or "").lower() if mi < len(move_types) else ""
                        icon_data = type_icons.get(icon_key)
                        if not icon_data and icon_key:
                            try:
                                tid = self.TYPE_ICON_ID.get(icon_key)
                                if tid:
                                    url = f"{self.TYPE_ICON_BASE}{tid}.png"
                                    req = urllib.request.Request(url, headers={"User-Agent":"SinKA-PvP-105"})
                                    with urllib.request.urlopen(req, timeout=5) as rr:
                                        icon_data = rr.read()
                                    type_icons[icon_key] = icon_data
                            except Exception:
                                icon_data = None
                        if icon_data:
                            try:
                                ic = Image.open(BytesIO(icon_data)).convert("RGBA")
                                ic.thumbnail((20,20), Image.LANCZOS)
                                ix = move_x - tw - ic.width - 7
                                img.alpha_composite(ic, (ix, yy-10))
                                draw.text((ix+ic.width+6, yy), txt, fill=TEXT, font=f_small_bold, anchor="lm")
                                continue
                            except Exception:
                                pass
                        draw.text((move_x, yy), txt, fill=TEXT, font=f_small, anchor="rm")


            footer_y = H - 72
            draw.line((margin+20, footer_y-20, W-margin-20, footer_y-20), fill=CYAN, width=4)
            draw.text((margin+25, footer_y+5), "SinKA PvP", fill=NAVY, font=f_footer)
            draw.text((W//2, footer_y+5), "Güncel Top-200 • 1v1 Turn-by-Turn • 3v3 State • Rol • Coverage • Recovery • Core", fill=MUTED, font=f_footer, anchor="mm")
            draw.text((W-margin-25, footer_y+5), "www.sinka.com.tr", fill=NAVY, font=f_footer, anchor="ra")

            ext = Path(path).suffix.lower()
            if ext in (".jpg", ".jpeg"):
                img.convert("RGB").save(path, "JPEG", quality=95, dpi=(300,300), optimize=True)
            else:
                img.save(path, "PNG", dpi=(300,300), optimize=True)

            self.root.after(0, lambda p=path: (self.set_status(f"Takım A4 görseli kaydedildi: {p}"), self._show_a4_result_dialog("Takım Görseli Kaydedildi", f"A4 yatay takım görseli başarıyla kaydedildi.\\n\\n{p}")))
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
            self.root.after(0, lambda err=err: self._show_a4_result_dialog("Takım Görseli Hatası", f"A4 takım görseli oluşturulamadı.\\n\\n{err}"))

    def _show_generated_teams(self, teams, league, include_shadows=False, owned_only=False):
        """Takım sonuçlarını seçenekler + Pokémon görselleri + roller ile gösterir."""
        self._team_image_refs = []
        win = tk.Toplevel(self.root)
        self._team_window = win
        win.title(f"Önerilen PvP Takımları - {league}")
        win.transient(self.root)
        win.protocol("WM_DELETE_WINDOW", lambda: (win.destroy(), setattr(self, "_team_window", None)))
        win.geometry("1050x760")
        win.minsize(900, 600)

        # Üst başlık + sağ üstte takım sonuçlarını A4 yatay paftaya dönüştüren buton.
        header_bar = ttk.Frame(win)
        header_bar.pack(fill="x", padx=14, pady=(8, 5))
        ttk.Label(
            header_bar,
            text=f"⚔ {league} İçin En İdeal 3\'lü Takımlar",
            font=("Segoe UI", 15, "bold")
        ).pack(side="left", padx=(6, 10))
        ttk.Button(
            header_bar,
            text="🖼 Resim Oluştur",
            command=lambda: self._create_team_a4_image(teams, league, shadow_var, owned_var)
        ).pack(side="right", padx=(6, 0))

        # Kullanıcı özellikle bu iki seçeneğin sonuç penceresinin üstünde olmasını istedi.
        options = ttk.LabelFrame(win, text="Takım Oluşturma Seçenekleri")
        options.pack(fill="x", padx=14, pady=(0, 8))

        shadow_var = tk.BooleanVar(value=not include_shadows)
        owned_var = tk.BooleanVar(value=owned_only)
        ttk.Checkbutton(
            options,
            text="🌑 Shadowları dahil etme",
            variable=shadow_var
        ).pack(side="left", padx=(12, 8), pady=8)
        ttk.Checkbutton(
            options,
            text="🎒 Sadece Mevcut Pokémonları kullan",
            variable=owned_var
        ).pack(side="left", padx=8, pady=8)

        option_status = tk.StringVar(value=(
            f"Shadow {'hariç' if shadow_var.get() else 'dahil'} • "
            f"{'Sadece mevcut Pokémonlar' if owned_var.get() else 'Genel meta'}"
        ))
        ttk.Label(options, textvariable=option_status, foreground="#555555").pack(side="left", padx=12)
        ttk.Label(options, text="(İşareti kaldırırsanız tüm meta havuzu kullanılır)", foreground="#777777").pack(side="left", padx=4)

        def refresh_options():
            # Eski sonuç penceresini gizlemek yerine kapat; böylece her yenilemede
            # arkada görünmez Toplevel birikmez. Yeni hesaplama ilerleme penceresiyle
            # hemen başlar.
            option_status.set("Yeni seçeneklerle takım hesaplanıyor...")
            try:
                win.destroy()
                if getattr(self, "_team_window", None) is win:
                    self._team_window = None
            except Exception:
                pass
            self._start_team_build(
                league,
                include_shadows=not shadow_var.get(),
                owned_only=owned_var.get()
            )

        ttk.Button(options, text="🔄 Seçenekleri Uygula / Takımı Yenile", command=refresh_options).pack(side="right", padx=12, pady=6)

        # Takımlar bulunamadığında bile nedenini açıkça göster.
        if not teams:
            msg = (
                "Bu filtrelerle uygun 3 Pokémon bulunamadı.\n\n"
                "• İlk arama yalnızca Mevcut Pokémonlarla yapılır; en az 3 uygun Pokémon işaretleyin.\n"
                "• Tüm meta ile denemek için 'Sadece Mevcut Pokémonları kullan' işaretini kaldırıp yenileyin.\n"
                "• Shadowları kullanmak istiyorsanız 'Shadowları dahil etme' seçeneğinin işaretini kaldırın."
            )
            ttk.Label(win, text=msg, justify="left", font=("Segoe UI", 10)).pack(fill="x", padx=25, pady=25)
            return

        scroll_host = ttk.Frame(win)
        scroll_host.pack(fill="both", expand=True, padx=8, pady=4)
        canvas = tk.Canvas(scroll_host, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_host, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        image_jobs = []

        def display_data(p_data):
            name = str(p_data.get("speciesName") or p_data.get("name") or p_data.get("speciesId") or "Bilinmeyen")
            types = self._team_type_keys(name)
            type_text = " / ".join(t.title() for t in types) if types else "-"
            moveset = p_data.get("moveset") or p_data.get("moves") or []
            if isinstance(moveset, dict):
                moveset = list(moveset.values())
            if not isinstance(moveset, (list, tuple)):
                moveset = [moveset]
            move_names = []
            move_types = []
            for mv in list(moveset)[:3]:
                try:
                    move_names.append(self.translate_move(str(mv)))
                except Exception:
                    move_names.append(str(mv))
                move_types.append(self._team_move_type_key(mv))
            while len(move_names) < 3:
                move_names.append("-")
                move_types.append("")
            return name, type_text, move_names, move_types

        for t_idx, team in enumerate(teams[:5], 1):
            category = str(team.get("category", "ALTERNATİF"))
            archetype = str(team.get("archetype", "Hibrit"))
            lf = ttk.LabelFrame(
                scroll_frame,
                text=(
                    f" Takım {t_idx} • {category} • {archetype} • TOPLAM: {team.get('score', 0):.1f}/100 "
                    f"• Meta: {team.get('meta_score', 0)} "
                    f"• Sinerji: {team.get('synergy', 0)} "
                    f"• Coverage: {team.get('coverage', 0)} "
                )
            )
            lf.pack(fill="x", expand=True, padx=10, pady=7)

            metrics = (
                f"Lead {team.get('lead_score', 0)}  |  "
                f"Safe Switch {team.get('switch_score', 0)}  |  "
                f"Closer {team.get('closer_score', 0)}  |  "
                f"Consistency {team.get('consistency', 0)}  |  "
                f"Top-200 avantaj {team.get('good_coverage', 0)}%  |  "
                f"Aynı anda kayıp {team.get('alignment_risk', 0)}%  |  RPS bağımlılığı {team.get('rps_dependency', 0)}%"
            )
            ttk.Label(lf, text=metrics, foreground="#555555", font=("Segoe UI", 8)).pack(
                fill="x", padx=10, pady=(4, 2)
            )

            engine_metrics = (
                f"Turn-by-turn {team.get('real_battle_score', '—')}  |  "
                f"Proxy {team.get('proxy_score', team.get('score', 0))}  |  "
                f"Recovery {team.get('recovery_score', 0)}  |  "
                f"CMP {team.get('cmp_rate', 0)}  |  "
                f"Bait başarı {team.get('bait_success_rate', 0)}%  |  "
                f"Ort. tur {team.get('battle_turns_avg', '—')}"
            )
            ttk.Label(lf, text=(
                f"3v3 State: {team.get('real_3v3_score', '—')}  |  "
                f"3v3 örnek {team.get('three_v_three_samples', 0)}  |  "
                f"3v3 ort. tur {team.get('three_v_three_turns', '—')}  |  "
                f"Switch {team.get('three_v_three_switches', 0)}  |  "
                f"CMP {team.get('three_v_three_cmp', 0)}  |  "
                f"Bait {team.get('three_v_three_bait', 0)}%  |  "
                f"Farm {team.get('three_v_three_farm', 0)}"
            ), foreground="#174a8b", font=("Segoe UI", 8, "bold")).pack(fill="x", padx=10, pady=(0, 3))
            ttk.Label(lf, text=engine_metrics, foreground="#444444", font=("Segoe UI", 8, "bold")).pack(
                fill="x", padx=10, pady=(0, 3)
            )
            threat_metrics = (
                f"Ortak tehdit {team.get('shared_threat_risk', 0)}%  |  "
                f"Ciddi ortak açık {team.get('severe_shared_risk', 0)}%  |  "
                f"Üçlü tehdit {team.get('triple_threat_risk', 0)}%  |  "
                f"Kritik üçlü {team.get('critical_triple_risk', 0)}%  |  "
                f"Tek cevaba bağımlılık {team.get('single_answer_dependency', 0)}%  |  "
                f"Aynı tip çifti {team.get('same_type_pairs', 0)}"
            )
            ttk.Label(lf, text=threat_metrics, foreground="#174a8b", font=("Segoe UI", 8, "bold")).pack(
                fill="x", padx=10, pady=(0, 3)
            )

            why = team.get("why", []) or []
            warnings = team.get("warnings", []) or []
            why_text = " • ".join(str(x) for x in why[:3])
            if why_text:
                ttk.Label(lf, text=f"🎯 Neden bu takım? {why_text}", foreground="#174a8b", font=("Segoe UI", 8)).pack(
                    fill="x", padx=10, pady=(1, 2)
                )
            if warnings:
                ttk.Label(lf, text="⚠ " + " • ".join(str(x) for x in warnings[:2]), foreground="#9b3d00", font=("Segoe UI", 8)).pack(
                    fill="x", padx=10, pady=(0, 3)
                )
            roles = [
                ("Lider (Lead)", team.get("lead", {})),
                ("Güvenli Değişim (Safe Switch)", team.get("safe_switch", {})),
                ("Kapatıcı (Closer)", team.get("closer", {})),
            ]
            for role_title, p_data in roles:
                if not isinstance(p_data, dict):
                    continue
                name, type_text, move_names, move_types = display_data(p_data)
                r_frame = ttk.Frame(lf)
                r_frame.pack(fill="x", padx=10, pady=5)

                ttk.Label(r_frame, text=f"• {role_title}:", font=("Segoe UI", 9, "bold"), width=27).pack(side="left")
                img_label = ttk.Label(r_frame, text="⏳", width=8, anchor="center")
                img_label.pack(side="left", padx=(0, 8))
                info_frame = ttk.Frame(r_frame)
                info_frame.pack(side="left", fill="x", expand=True)
                ttk.Label(info_frame, text=name, font=("Segoe UI", 11, "bold"), foreground="#0055aa").pack(anchor="w")
                ttk.Label(info_frame, text=f"[{type_text}]", foreground="#555555").pack(anchor="w")
                moves_frame = ttk.Frame(r_frame)
                moves_frame.pack(side="right", padx=(8, 2))
                for mi, (lab, mv) in enumerate(zip(("Hızlı", "Güç 1", "Güç 2"), move_names)):
                    mf = ttk.Frame(moves_frame)
                    mf.pack(anchor="e", pady=1)
                    key = move_types[mi] if mi < len(move_types) else ""
                    icon_label = ttk.Label(mf, text="", width=2)
                    icon_label.pack(side="left")
                    ttk.Label(mf, text=f"{lab}: {mv}", font=("Segoe UI", 7)).pack(side="left")
                    if key:
                        # Ağ erişimi worker'da; PhotoImage oluşturma yalnızca Tk ana thread'inde yapılır.
                        def load_move_icon(label=icon_label, type_key=key):
                            try:
                                data = self.download_type_icon_bytes(type_key)
                                if not data:
                                    return
                                def apply_icon():
                                    try:
                                        if not label.winfo_exists():
                                            return
                                        img = Image.open(BytesIO(data)).convert("RGBA")
                                        img.thumbnail((22, 22), Image.LANCZOS)
                                        ph = ImageTk.PhotoImage(img)
                                        label.configure(image=ph, text="")
                                        label.image = ph
                                        self._team_image_refs.append(ph)
                                    except Exception:
                                        pass
                                self.root.after(0, apply_icon)
                            except Exception:
                                pass
                        threading.Thread(target=load_move_icon, daemon=True).start()
                image_jobs.append((img_label, name))

        # Görselleri UI'yi kilitlemeden arka planda getir.
        def load_one_image(label, name):
            if not PIL_OK:
                return
            try:
                # Önce mevcut PhotoImage cache'ini ana thread'de kontrol et.
                photo = None
                for key in (name, name.lower()):
                    obj = self.image_cache.get(key) if isinstance(self.image_cache, dict) else None
                    if isinstance(obj, ImageTk.PhotoImage):
                        photo = obj; break
                if photo is not None:
                    label.configure(image=photo, text="")
                    label.image = photo
                    self._team_image_refs.append(photo)
                    return

                def worker_image():
                    raw = None
                    try:
                        raw = self.image_bytes.get(name) if isinstance(self.image_bytes, dict) else None
                    except Exception:
                        raw = None
                    if not raw:
                        try:
                            raw = self.download_image_bytes(name)
                        except Exception:
                            raw = None
                    if raw:
                        def apply():
                            try:
                                if not win.winfo_exists(): return
                                img = Image.open(BytesIO(raw)).convert("RGBA")
                                img.thumbnail((64, 64), Image.LANCZOS)
                                ph = ImageTk.PhotoImage(img)
                                self._team_image_refs.append(ph)
                                label.configure(image=ph, text="")
                                label.image = ph
                            except Exception:
                                pass
                        self.root.after(0, apply)
                    else:
                        try: self.root.after(0, lambda: label.configure(text="🖼"))
                        except Exception: pass
                threading.Thread(target=worker_image, daemon=True).start()
            except Exception:
                pass

        for label, name in image_jobs:
            load_one_image(label, name)

        # Fare tekerleğiyle kaydırma.
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        win.update_idletasks()
        try:
            screen_w, screen_h = win.winfo_screenwidth(), win.winfo_screenheight()
            x = max((screen_w - min(1100, screen_w - 60)) // 2, 0)
            y = max((screen_h - min(780, screen_h - 80)) // 2, 0)
            w = min(1100, screen_w - 60); h = min(780, screen_h - 80)
            win.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass


    def _fetch_dittobase_max_attackers(self):
        """Dittobase Max Battles / Overall listesini isim + Maksimum Hareket ile alır."""
        try:
            req = urllib.request.Request(
                DITTOBASE_MAX_ATTACKERS_URL,
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/151 Safari/537.36"}
            )
            with urllib.request.urlopen(req, timeout=20) as r:
                html = r.read().decode("utf-8", errors="ignore")

            from html.parser import HTMLParser

            class PokeLinkParser(HTMLParser):
                def __init__(self):
                    super().__init__(convert_charrefs=True)
                    self.active = False
                    self.buf = []
                    self.records = []
                    self.current_href = ""

                def handle_starttag(self, tag, attrs):
                    if tag.lower() != "a":
                        return
                    d = dict(attrs)
                    href = str(d.get("href", ""))
                    if "/pokemon-go/pokedex/" in href and "cp-iv-chart" not in href:
                        self.active = True
                        self.current_href = href
                        self.buf = []

                def handle_data(self, data):
                    if self.active:
                        self.buf.append(data)

                def handle_endtag(self, tag):
                    if tag.lower() == "a" and self.active:
                        name = re.sub(r"\s+", " ", " ".join(self.buf)).strip()
                        if name:
                            self.records.append((name, self.current_href))
                        self.active = False
                        self.current_href = ""
                        self.buf = []

            parser = PokeLinkParser()
            parser.feed(html)

            # Aynı Pokémon sayfasına verilen yardımcı bağlantıları ayıklayıp ilk 100
            # sıralı kayıt için isimleri koruyoruz.
            records = []
            seen = set()
            for name, href in parser.records:
                key = (name.strip().lower(), href.strip().lower())
                if key in seen:
                    continue
                seen.add(key)
                records.append({"name": name.strip(), "move": "—", "href": href})
                if len(records) >= 100:
                    break

            # Dittobase tablosundaki Max Move alanı, Pokémon bağlantısının hemen
            # sonrasında HTML içinde metin olarak bulunuyor. Her kayıt için bir sonraki
            # Pokémon bağlantısına kadar olan parçadan uygun hareket adını çıkar.
            anchors = list(re.finditer(r'<a[^>]+href=["\']([^"\']*/pokemon-go/pokedex/[^"\']*)["\'][^>]*>(.*?)</a>', html, flags=re.I | re.S))
            anchor_data = []
            for m in anchors:
                label = re.sub(r"<[^>]+>", " ", m.group(2))
                label = re.sub(r"\s+", " ", label).strip()
                if not label or "cp-iv-chart" in m.group(1):
                    continue
                anchor_data.append((m.start(), m.end(), label))

            # Sayfada Pokémon adı dışında linkler de olabildiği için benzersiz isim
            # sırasını temel alıyoruz.
            unique_anchor_data = []
            seen_names = set()
            for pos0, pos1, label in anchor_data:
                key = label.lower()
                if key in seen_names:
                    continue
                seen_names.add(key)
                unique_anchor_data.append((pos0, pos1, label))
                if len(unique_anchor_data) >= 100:
                    break

            # Hareket adlarını çıkarırken bilinen Max Move biçimlerini önceliklendir.
            # Örn. G-Max Terror, G-Max Hydrosnipe, Dynamax Cannon, Ground, Fire vb.
            type_words = {
                "normal", "fire", "water", "electric", "grass", "ice", "fighting",
                "poison", "ground", "flying", "psychic", "bug", "rock", "ghost",
                "dragon", "dark", "steel", "fairy"
            }
            for idx, rec in enumerate(records):
                if idx >= len(unique_anchor_data):
                    break
                start = unique_anchor_data[idx][1]
                end = unique_anchor_data[idx + 1][0] if idx + 1 < len(unique_anchor_data) else len(html)
                segment = html[start:end]
                text = re.sub(r"<script.*?</script>|<style.*?</style>", " ", segment, flags=re.I | re.S)
                text = re.sub(r"<[^>]+>", " ", text)
                text = re.sub(r"&(?:nbsp|amp|quot|#39);", " ", text, flags=re.I)
                text = re.sub(r"\s+", " ", text).strip()

                candidates = []
                # En güçlü aday: G-Max / Dynamax isimleri.
                candidates += re.findall(r"(?:G-Max|Dynamax)\s+[A-Za-z][A-Za-z' -]{1,60}", text, flags=re.I)
                # Normal Max Move'lar için sayfadaki hücre sırasından kısa metinleri dene.
                for token in re.split(r"\s{2,}| \| ", text):
                    token = token.strip(" -—")
                    if not token or token.lower() in type_words:
                        continue
                    if len(token) <= 60 and not re.fullmatch(r"[0-9.%,]+", token):
                        if token.lower() not in {rec["name"].lower(), "attack", "max damage", "% of best"}:
                            candidates.append(token)

                if candidates:
                    # Önce gerçekten hareket gibi görünen ilk adayı al.
                    move = candidates[0].strip()
                    move = re.sub(r"\s+", " ", move)
                    rec["move"] = move

            return records
        except Exception:
            return []

    def _max_attacker_clean_name(self, name):
        return str(name).replace("G-Max ", "Gigantamax ").strip()

    def _load_max_attacker_images_async(self, rows, refs, section_gen):
        if not PIL_OK or not rows:
            return
        from concurrent.futures import ThreadPoolExecutor, as_completed
        from io import BytesIO
        def fetch(item):
            name = item[1]
            return name, self.download_image_bytes(name)
        with ThreadPoolExecutor(max_workers=min(10, len(rows))) as pool:
            futures=[pool.submit(fetch, r) for r in rows]
            for f in as_completed(futures):
                try: name,data=f.result()
                except Exception: continue
                if not data or section_gen != getattr(self, '_max_section_gen', 0):
                    continue
                def finish(n=name,d=data):
                    if section_gen != getattr(self, '_max_section_gen', 0): return
                    try:
                        img=Image.open(BytesIO(d)).convert('RGBA')
                        img.thumbnail((52,52), Image.LANCZOS)
                        photo=ImageTk.PhotoImage(img)
                        self._max_image_cache[n]=photo
                        self._max_section_refs.append(photo)
                        iid=self._max_items.get(n)
                        if iid:
                            self._max_tree.item(iid, image=photo)
                    except Exception: pass
                self.root.after(0, finish)

    def _build_max_attacker_section(self):
        """Ana ekranın altında Dittobase Max Battles listesini doğrudan gösterir."""
        if not hasattr(self, '_max_section'):
            return
        for w in self._max_section.winfo_children():
            w.destroy()
        self._max_image_cache={}
        self._max_section_refs=[]
        self._max_items={}
        section = self._max_section
        ttk.Separator(section).pack(fill='x', pady=(12,8))
        ttk.Label(section, text='⚡ DYNAMAX & GIGANTAMAX — MAX BATTLE EN İYİ SALDIRICILAR', font=('Segoe UI',15,'bold')).pack(pady=(2,2))
        ttk.Label(section, text='Kaynak: Dittobase • Max Battles / Overall • Güncel liste doğrudan kaynaktan alınır', font=('Segoe UI',9)).pack(pady=(0,8))
        status=ttk.Label(section, text='Dittobase verileri yükleniyor…')
        status.pack(pady=(0,5))
        body=ttk.Frame(section)
        body.pack(fill='x', expand=True)
        cols=('rank','pokemon','move','damage','mmw','pct')
        tree=ttk.Treeview(body, columns=cols, show='tree headings', height=12)
        tree.heading('#0', text='Görsel')
        tree.column('#0', width=75, anchor='center', stretch=False)
        for c,h,w in [('rank','Rank',60),('pokemon','Pokémon',220),('move','Max Move',220),('damage','Max Hasar',100),('mmw','MMW',90),('pct','En İyinin %',110)]:
            tree.heading(c,text=h); tree.column(c,width=w,anchor='center')
        sb=ttk.Scrollbar(body,orient='vertical',command=tree.yview); tree.configure(yscrollcommand=sb.set)
        tree.pack(side='left',fill='x',expand=True); sb.pack(side='right',fill='y')
        self._max_tree=tree
        self._max_status=status
        try:
            rows=self._fetch_dittobase_max_attackers()
        except Exception:
            rows=[]
        if not rows:
            status.config(text='Dittobase verileri alınamadı.')
            return
        display=[]
        # Kaynaktaki ilk 100 için bilinen ilk 50 metrikleri HTML'den ayrıştırmaya çalış.
        # İsimler eksikse yine sıralama ve görseller gösterilir.
        for i,rec in enumerate(rows[:100],1):
            name=rec.get("name", "") if isinstance(rec, dict) else str(rec)
            move=rec.get("move", "—") if isinstance(rec, dict) else "—"
            clean=self._max_attacker_clean_name(name)
            iid=tree.insert('', 'end', text='', values=(i,clean,move,'—','—','—'))
            self._max_items[name]=iid
            display.append((i,name))
        status.config(text=f'{len(display)} Max Battle saldırıcısı yüklendi • 1–{len(display)}')
        self._max_section_gen=getattr(self,'_max_section_gen',0)+1
        gen=self._max_section_gen
        threading.Thread(target=self._load_max_attacker_images_async,args=(display,self._max_section_refs,gen),daemon=True).start()

    def _start_max_inline_load(self):
        if getattr(self,'_max_loading',False): return
        self._max_loading=True
        self._max_section_gen=getattr(self,'_max_section_gen',0)+1
        gen=self._max_section_gen
        if hasattr(self,'_max_status'):
            self._max_status.config(text='Dittobase Max Battles verileri yükleniyor…')
        def worker():
            try:
                rows=self._fetch_dittobase_max_attackers()
                self.root.after(0, lambda r=rows,g=gen: self._render_max_inline_rows(r,g))
            except Exception as e:
                self.root.after(0, lambda: self._max_status.config(text='Dittobase verileri alınamadı.'))
            finally:
                self.root.after(0, lambda: setattr(self,'_max_loading',False))
        threading.Thread(target=worker,daemon=True).start()

    def _render_max_inline_rows(self, rows, gen):
        if gen != getattr(self,'_max_section_gen',0) or not hasattr(self,'_max_tree'): return
        tree=self._max_tree
        for iid in tree.get_children(): tree.delete(iid)
        self._max_items={}; self._max_section_refs=[]; self._max_image_cache={}
        for i,rec in enumerate(rows[:100],1):
            name=rec.get("name", "") if isinstance(rec, dict) else str(rec)
            move=rec.get("move", "—") if isinstance(rec, dict) else "—"
            clean=self._max_attacker_clean_name(name)
            iid=tree.insert('', 'end', text='', values=(i,clean,move,'—','—','—'))
            self._max_items[name]=iid
        self._max_status.config(text=f'{min(100,len(rows))} Max Battle saldırıcısı yüklendi • Kaynak: Dittobase')
        display=[(i,(r.get("name", "") if isinstance(r, dict) else str(r))) for i,r in enumerate(rows[:100],1)]
        threading.Thread(target=self._load_max_attacker_images_async,args=(display,self._max_section_refs,gen),daemon=True).start()

        self._max_section = ttk.Frame(self.root)
        self._max_section.pack(fill="x", padx=16, pady=(0, 10))
        self._max_loading = False


    def _load_gigamax_main_thread(self):
        """Gigamax seçiliyken Dittobase Max Battles listesini ana tabloya yükler."""
        try:
            rows = self._fetch_dittobase_max_attackers()
            self.root.after(0, lambda r=rows: self._show_gigamax_main_rows(r))
        except Exception as e:
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "SinKA PvP",
                    "Dittobase Gigamax verileri alınamadı:\n\n" + str(e)
                )
            )
        finally:
            self.root.after(0, self.finish_busy)

    def _gigamax_base_name(self, name):
        """Max/G-Max/Gigantamax adından gerçek temel species adını çıkarır."""
        s = str(name or "").strip()
        # Dittobase/PvPoke farklı yazımlar kullanabiliyor:
        # G-Max Charizard, Gigantamax Charizard, Dynamax Charizard,
        # Charizard (Gigantamax) gibi varyantların tamamını normalize et.
        s = re.sub(r"^(?:Gigantamax|G-Max|Dynamax)\s+", "", s, flags=re.I).strip()
        s = re.sub(r"\s*\((?:Gigantamax|G-Max|Dynamax)\)\s*$", "", s, flags=re.I).strip()
        return s

    def _pokemon_name_norm(self, value):
        """Pokémon isimlerini form/işaret farklarından arındırarak karşılaştırır."""
        s = str(value or "").strip().lower()
        s = re.sub(r"^(?:gigantamax|g-max|dynamax)\s+", "", s)
        s = re.sub(r"\s*\((?:gigantamax|g-max|dynamax)\)\s*$", "", s)
        # Shadow vb. Game Master varyantlarını temel species ile eşleştir.
        s = re.sub(r"(?:_shadow|_purified)$", "", s)
        return re.sub(r"[^a-z0-9]+", "", s)

    def _find_pokemon_data_by_name(self, name):
        """PvPoke Game Master içinden Max/G-Max adıyla gerçek species kaydını bulur."""
        target = self._pokemon_name_norm(self._gigamax_base_name(name))
        if not target:
            return None, None

        # 1) Tam normalize eşleşme. speciesId, name ve speciesName birlikte kontrol edilir.
        for sid, poke in self.pokemon.items():
            if not isinstance(poke, dict):
                continue
            candidates = (
                sid,
                poke.get("speciesId"),
                poke.get("name"),
                poke.get("speciesName"),
            )
            if any(self._pokemon_name_norm(c) == target for c in candidates if c):
                return sid, poke

        # 2) Bazı Game Master kayıtlarında isim sonuna form eki gelebilir.
        # Temel species ile başlayan ilk uygun kaydı seç.
        for sid, poke in self.pokemon.items():
            if not isinstance(poke, dict):
                continue
            for cand in (poke.get("name"), poke.get("speciesName"), sid):
                norm = self._pokemon_name_norm(cand)
                if norm and (norm.startswith(target) or target.startswith(norm)):
                    # Çok kısa isimlerde yanlış eşleşmeyi önle.
                    if min(len(norm), len(target)) >= 4:
                        return sid, poke
        return None, None

    def _gigamax_evolution_data(self, sid, poke):
        """GigaMax satırı için normal species'ın önceki evrimlerini güvenilir biçimde üretir."""
        if not poke:
            return "Evrim yok", []

        # Önce normal evolution_stage mantığını kullan.
        evolution = self.evolution_stage(sid, poke)
        current_name = poke.get("name") or poke.get("speciesName") or sid
        evo_list = list(self.evolution_names.get(current_name, []))

        # evolution_stage hiçbir şey bulamadıysa parent zincirini doğrudan takip et.
        if not evo_list:
            family = poke.get("family") or {}
            parent = family.get("parent")
            seen = set()
            previous = []
            while parent and parent not in seen:
                seen.add(parent)
                parent_poke = self.pokemon.get(parent)
                if not parent_poke and str(parent).endswith("_shadow"):
                    parent_poke = self.pokemon.get(str(parent)[:-7])
                if not parent_poke:
                    # Normalize edilmiş speciesId ile tekrar ara.
                    _, parent_poke = self._find_pokemon_data_by_name(parent)
                if not parent_poke:
                    break
                parent_name = (parent_poke.get("name") or parent_poke.get("speciesName") or parent)
                previous.append(self.translate_pokemon_name(parent_name))
                parent_family = parent_poke.get("family") or {}
                parent = parent_family.get("parent")
            evo_list = list(reversed(previous))

        # Görsel overlay'in satır adıyla çalışması için her iki anahtarı da doldur.
        self.evolution_names[str(current_name)] = list(evo_list)
        return (" → ".join(evo_list) if evo_list else "Evrim yok"), evo_list

    def _show_gigamax_main_rows(self, names):
        """Dittobase listesini mevcut ana tabloya tam Pokémon bilgileriyle yükler."""
        self.generation += 1
        gen = self.generation

        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.tree_items_by_name = {}
        self.evolution_names = {}
        self.type_keys = {}
        # Ana tablo görsel yükleyicisi, Gigantamax/Dynamax görünen adlarını
        # gerçek species adına yönlendirecek.
        self._main_image_aliases = {}

        requested = self.get_requested_count()
        if isinstance(names, dict):
            names = list(names.values())
        normalized_records = []
        seen_names = set()
        for x in (names or []):
            if isinstance(x, dict):
                raw = str(x.get("name", "")).strip()
                move = str(x.get("move", "—")).strip() or "—"
            else:
                raw = str(x).strip()
                move = "—"
            if raw and raw.lower() not in seen_names:
                seen_names.add(raw.lower())
                normalized_records.append({"name": raw, "move": move})
        normalized_records = normalized_records[:requested]

        rows = []
        for rank, rec in enumerate(normalized_records, 1):
            raw_name = rec["name"]
            max_move = rec.get("move", "—") or "—"
            display_name = self._max_attacker_clean_name(raw_name)
            base_name = self._gigamax_base_name(display_name)
            sid, poke = self._find_pokemon_data_by_name(base_name)

            # Dittobase'deki form adı ile gerçek görsel/species adı arasında bağ kur.
            self._main_image_aliases[display_name] = (
                "Gigantamax " + base_name if self.loaded_league == "Gigamax" else base_name
            )

            if poke:
                actual_name = (
                    poke.get("name") or poke.get("speciesName") or base_name
                )
                # Pokémon tipi: normal PvP tablosundakiyle aynı TYPE_TR + ikon sistemi.
                type_text = self.pokemon_type_text(poke, display_name)

                # GigaMax için normal species'ın evrim zincirini kullan.
                evolution, evo_list = self._gigamax_evolution_data(sid or base_name, poke)
                self.evolution_names[display_name] = list(evo_list)
                self.evolution_names[str(actual_name)] = list(evo_list)

                owned = bool(
                    self._is_owned(display_name)
                    or self._is_owned(base_name)
                    or self._is_owned(actual_name)
                )
                tag = "green"
            else:
                type_text = "-"
                evolution = "Evrim yok"
                self.evolution_names[display_name] = []
                owned = self._is_owned(display_name)
                tag = "red"

            rows.append((
                rank,                       # rank
                rank,                       # original_rank
                "",                        # image
                display_name,              # pokemon
                type_text,                 # gerçek Pokémon tipi
                evolution,                 # önceki evrimler
                "Max Battle",              # score/source indicator
                "—",                       # iv
                "—",                       # level
                "—",                       # cp
                "—",                       # xl
                "—",                       # special
                max_move,                  # move1 = Dittobase Max Move
                "—",                       # move2
                "—",                       # move3
                "☑" if owned else "☐",  # current
                tag
            ))

        # Normal ana tablo altyapısını aynen kullanıyoruz:
        # tip ikonları + evrim ikonları + Pokémon görselleri burada devreye girer.
        self.show_rows(rows, gen, "Gigamax")
        self.info_var.set(
            f"Gigamax • Dittobase Max Battles / Overall • {len(rows)} Pokémon"
        )
        self.set_status(
            f"Gigamax: {len(rows)} Pokémon hazır • Kaynak: Dittobase Max Battles / Overall"
        )
        self._update_pokemon_count()

    def _activate_gigamax(self):
        if self.busy:
            return
        self.busy = True
        self.set_status("Gigamax / Dittobase verileri indiriliyor...")
        threading.Thread(
            target=self._load_gigamax_main_thread,
            daemon=True
        ).start()

    def load_initial(self):
        try:
            self.set_status("Pokémon ve hareket verileri indiriliyor...")
            pokemon_list = download_json(POKEMON_URL)
            tr_moves = download_json(TR_MOVES_URL)
            pogo_moves = download_json(POGO_MOVES_URL)

            self.pokemon = {
                p.get("speciesId"): p for p in pokemon_list
                if p.get("speciesId")
            }
            self.translate_move = build_move_translator(pogo_moves, tr_moves)

            self.set_status("Great League sıralaması indiriliyor...")
            self.load_league_data("Great League")
        except Exception as e:
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "SinKA PvP",
                    "Veriler indirilemedi:\n\n" + str(e)
                )
            )
            self.set_status("Veri yükleme hatası.")

    def load_league_data(self, league):
        if league == "Mega":
            data = download_json(MEGA_RANKING_URL)
        else:
            cp = LEAGUES[league]
            data = download_json(RANKING_URL.format(cp=cp))
        self.ranking_data = data
        self.loaded_league = league
        self.root.after(0, self.refresh)

    def _get_active_search_league(self):
        """Arama için ekranda seçili olan üst lig kartını kesin kaynak kabul eder."""
        # 1) Görünür lig seçicisi: en yüksek öncelik.
        try:
            league = str(self.league_var.get()).strip()
            if league in LEAGUES:
                self.active_search_league = league
                return league
        except Exception:
            pass

        # 2) Son seçilen aktif lig.
        try:
            league = str(getattr(self, "active_search_league", "") or "").strip()
            if league in LEAGUES:
                return league
        except Exception:
            pass

        # 3) Yüklenmiş lig.
        try:
            league = str(getattr(self, "loaded_league", "") or "").strip()
            if league in LEAGUES:
                self.active_search_league = league
                return league
        except Exception:
            pass

        self.active_search_league = "Great League"
        return "Great League"

    def _save_current_league_filter_settings(self, league=None):
        # league parametresi, Combobox yeni değere geçmiş olsa bile eski ligi
        # güvenli şekilde kaydetmemizi sağlar. Böylece ligler birbirine karışmaz.
        league = league or getattr(self, '_filter_context_league', None) or self.league_var.get()
        if league not in LEAGUES or not hasattr(self, 'league_filter_settings'):
            return
        try:
            count = max(1, min(100, int(str(self.count_var.get()).strip())))
        except Exception:
            count = 36
        self.league_filter_settings[league] = {
            'count': count,
            'shadow': bool(self.shadow_var.get()),
            'legendary': bool(self.legendary_var.get()),
            'mythical': bool(self.mythical_var.get()),
            'xl': bool(self.xl_var.get()),
            'special': bool(self.special_var.get()),
            'original_rank': bool(self.original_rank_var.get()),
        }

    def _apply_league_filter_settings(self, league):
        vals = self.league_filter_settings.setdefault(league, {
            'count': 36, 'shadow': False, 'legendary': False, 'mythical': False,
            'xl': False, 'special': False, 'original_rank': False
        })
        try:
            count = max(1, min(100, int(vals.get('count', 36))))
        except Exception:
            count = 36
        self._applying_league_filters = True
        try:
            self.count_var.set(str(count))
            self.shadow_var.set(bool(vals.get('shadow', False)))
            self.legendary_var.set(bool(vals.get('legendary', False)))
            self.mythical_var.set(bool(vals.get('mythical', False)))
            self.xl_var.set(bool(vals.get('xl', False)))
            self.special_var.set(bool(vals.get('special', False)))
            self.original_rank_var.set(bool(vals.get('original_rank', False)))
            self.toggle_original_rank()
            self._refresh_selected_filters_label()
        finally:
            self._applying_league_filters = False

    def change_league(self):
        if self.busy:
            return

        # Önce ekranda gösterilen eski ligin filtrelerini eski lige kaydet.
        # league_var artık yeni ligi gösteriyor olabileceği için doğrudan league_var
        # kullanmak hatalıydı; _filter_context_league eski ligi tutar.
        previous_league = getattr(self, '_filter_context_league', None)
        league = self.league_var.get()
        if previous_league in LEAGUES and previous_league != league:
            self._save_current_league_filter_settings(previous_league)

        # Yeni lig açıldıktan sonra yalnızca o lige ait kayıtlı filtreleri yükle.
        self._applying_league_filters = True
        try:
            self._apply_league_filter_settings(league)
        finally:
            self._applying_league_filters = False
        self._filter_context_league = league
        self.active_search_league = league
        self.owned = dict(self.owned_by_league.get(league, {}))
        self._refresh_owned_summary()
        self._refresh_status_summary()
        self._refresh_ready_pokemon_panel()
        self.save_preferences()

        if league == "Gigamax":
            self.loaded_league = "Gigamax"
            self._activate_gigamax()
            return

        if league == self.loaded_league:
            self.refresh()
            return

        self.busy = True
        self.set_status(league + " verisi indiriliyor...")
        threading.Thread(
            target=self.load_league_thread,
            args=(league,),
            daemon=True
        ).start()

    def load_league_thread(self, league):
        try:
            if league == "Mega":
                data = download_json(MEGA_RANKING_URL)
            else:
                cp = LEAGUES[league]
                data = download_json(RANKING_URL.format(cp=cp))
            self.ranking_data = data
            self.loaded_league = league
            self.root.after(0, self.refresh)
        except Exception as e:
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "SinKA PvP",
                    "Lig verisi indirilemedi:\n\n" + str(e)
                )
            )
        finally:
            self.root.after(0, self.finish_busy)

    def finish_busy(self):
        self.busy = False

    def reload(self):
        if self.busy:
            return
        self.busy = True
        self.set_status("Güncel veriler indiriliyor...")
        threading.Thread(target=self.reload_thread, daemon=True).start()

    def reload_thread(self):
        try:
            pokemon_list = download_json(POKEMON_URL)
            tr_moves = download_json(TR_MOVES_URL)
            pogo_moves = download_json(POGO_MOVES_URL)
            league = self.league_var.get()
            if league == "Gigamax":
                self.pokemon = {
                    p.get("speciesId"): p for p in pokemon_list
                    if p.get("speciesId")
                }
                self.translate_move = build_move_translator(pogo_moves, tr_moves)
                self.loaded_league = "Gigamax"
                self.root.after(0, self._activate_gigamax)
                return

            if league == "Mega":
                data = download_json(MEGA_RANKING_URL)
            else:
                data = download_json(RANKING_URL.format(cp=LEAGUES[league]))

            self.pokemon = {
                p.get("speciesId"): p for p in pokemon_list
                if p.get("speciesId")
            }
            self.translate_move = build_move_translator(pogo_moves, tr_moves)
            self.ranking_data = data
            self.loaded_league = league
            self.root.after(0, self.refresh)
        except Exception as e:
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    "SinKA PvP", "Yenileme başarısız:\n\n" + str(e)
                )
            )
        finally:
            self.root.after(0, self.finish_busy)

    def get_requested_count(self):
        try:
            value = int(str(self.count_var.get()).strip())
        except Exception:
            value = 36
        value = max(1, min(100, value))
        self.count_var.set(str(value))
        return value

    def has_pokemon_tag(self, item, tag):
        try:
            sid = str(item.get("speciesId", ""))
            poke = self.pokemon.get(sid)
            if not poke and sid.endswith("_shadow"):
                poke = self.pokemon.get(sid[:-7])
            if not poke:
                return False
            return tag in (poke.get("tags") or [])
        except Exception:
            return False

    def refresh(self):
        league = self.league_var.get()

        if league == "Gigamax":
            self._activate_gigamax()
            return

        if not self.ranking_data or not self.pokemon:
            return

        self.generation += 1
        gen = self.generation

        show_shadows = self.shadow_var.get()
        include_legendary = self.legendary_var.get()
        include_mythical = self.mythical_var.get()
        xl_only_no = self.xl_var.get()
        special_only_no = self.special_var.get()
        requested_count = self.get_requested_count()
        self.save_preferences()

        self.evolution_names = {}
        
        candidates = [
            p for p in self.ranking_data
            if not is_gigantamax(p)
            and (show_shadows or not is_shadow(p))
            and (not include_legendary or not self.has_pokemon_tag(p, "legendary"))
            and (not include_mythical or not self.has_pokemon_tag(p, "mythical"))
        ]

        if xl_only_no or special_only_no:
            self.info_var.set(
                f"{league} • filtre uygulanıyor..."
            )
        else:
            candidates = candidates[:requested_count]
            rarity_info = (
                f"Efsanevi {'hariç' if include_legendary else 'dahil'} • "
                f"Mistik {'hariç' if include_mythical else 'dahil'}"
            )
            self.info_var.set(
                f"{league} • {'Shadow dahil' if show_shadows else 'Shadow hariç'} • "
                f"{rarity_info} • {len(candidates)} Pokémon"
            )
        filtered = candidates

        self.set_status("Rank-1 IV hesaplanıyor... 0/" + str(len(filtered)))

        for item in self.tree.get_children():
            self.tree.delete(item)

        threading.Thread(
            target=self.calculate_rows,
            args=(filtered, league, gen),
            daemon=True
        ).start()

    def translate_pokemon_name(self, name):
        translations = {
            "Lickitung": "Lickitung",
            "Lickilicky": "Lickilicky",
            "Bulbasaur": "Bulbasaur",
            "Ivysaur": "Ivysaur",
            "Venusaur": "Venusaur",
        }
        return translations.get(name, name)

    def evolution_stage(self, sid, poke):
        if not poke:
            return "Evrim yok"

        family = poke.get("family") or {}
        parent = family.get("parent")

        if not parent:
            return "Evrim yok"

        previous = []
        current = parent
        seen = set()

        while current and current not in seen:
            seen.add(current)
            parent_poke = self.pokemon.get(current)

            if not parent_poke and current.endswith("_shadow"):
                parent_poke = self.pokemon.get(current[:-7])

            if not parent_poke:
                break

            parent_name = (
                parent_poke.get("name")
                or parent_poke.get("speciesName")
                or current
            )
            previous.append(self.translate_pokemon_name(parent_name))

            pf = parent_poke.get("family") or {}
            current = pf.get("parent")

        result = list(reversed(previous))
        current_name = (poke.get("name") or poke.get("speciesName") or sid)
        self.evolution_names[current_name] = result
        return " → ".join(result) if result else "Evrim yok"

    def calculate_rows(self, filtered, league, gen, manual_name=None):
        cp_limit = LEAGUES[league]
        rows = []
        source_index = 0
        target_count = 0
        need_xl_filter = self.xl_var.get()
        need_no_special_filter = self.special_var.get()
        requested_count = 1 if manual_name else self.get_requested_count()

        for p in filtered:
            if gen != self.generation:
                return

            source_index += 1
            original_rank = p.get("rank")
            try:
                original_rank = int(original_rank)
            except Exception:
                try:
                    original_rank = self.ranking_data.index(p) + 1
                except Exception:
                    original_rank = source_index

            sid = p.get("speciesId", "")
            name = p.get("speciesName", sid)
            score = f"{float(p.get('score', 0)):.1f}"

            poke = self.pokemon.get(sid)
            if not poke and sid.endswith("_shadow"):
                poke = self.pokemon.get(sid[:-7])

            if not poke or "baseStats" not in poke:
                if target_count < requested_count:
                    target_count += 1
                    rows.append(
                        (
                            target_count,
                            original_rank,
                            "",
                            name,
                            "-",
                            "-",
                            score,
                            "-",
                            "-",
                            "-",
                            "?",
                            "?",
                            "?",
                            "✖",
                            "✖",
                            "red"
                        )
                    )
            else:
                ivs, level, cp = rank1_iv(poke["baseStats"], cp_limit)

                moveset = p.get("moveset") or []
                elite = set(poke.get("eliteMoves", []))
                legacy = set(poke.get("legacyMoves", []))
                special_moves = elite | legacy

                move_names = []
                has_special = False
                for move in moveset[:3]:
                    name_tr = self.translate_move(move)
                    if move in special_moves:
                        name_tr += "*"
                        has_special = True
                    move_names.append(name_tr)

                while len(move_names) < 3:
                    move_names.append("-")

                xl = level > 40.0
                evolution = self.evolution_stage(sid, poke)
                self.evolution_names[str(name)] = list(self.evolution_names.get(
                    (poke.get("name") or poke.get("speciesName") or sid), []
                ))

                if manual_name:
                    # Manuel aramada seçilen Pokémon, üstteki XL/Özel güç
                    # filtrelerinden bağımsız olarak mutlaka gösterilir.
                    accepted = True
                elif need_xl_filter and xl:
                    accepted = False
                elif need_no_special_filter and has_special:
                    accepted = False
                else:
                    accepted = True

                if accepted and target_count < requested_count:
                    target_count += 1
                    rows.append((
                        target_count,
                        original_rank,
                        "",
                        name,
                        self.pokemon_type_text(poke, name),
                        evolution,
                        score,
                        f"{ivs[0]}/{ivs[1]}/{ivs[2]}",
                        f"{level:g}",
                        cp,
                        "✖" if xl else "✓",
                        "✖" if has_special else "✓",
                        move_names[0],
                        move_names[1],
                        move_names[2],
                        "☑" if self._is_owned(name) else "☐",
                        "red" if (xl or has_special) else "green"
                    ))

            if source_index % 5 == 0 or source_index == len(filtered):
                percent = (source_index / max(1, len(filtered))) * 100
                self.root.after(
                    0,
                    lambda i=source_index, total=len(filtered), pc=percent:
                    self.update_progress(pc, f"Veriler işleniyor... {i}/{total}")
                )

            if target_count >= requested_count and (need_xl_filter or need_no_special_filter):
                break

        self.root.after(0, lambda: self.show_rows(rows, gen, league))

    def update_progress(self, percent, text):
        if hasattr(self, "progress_var"):
            self.progress_var.set(percent)
        if hasattr(self, "loading_var"):
            self.loading_var.set(text)
        self.root.update_idletasks()

    TYPE_TR = {
        "normal": "Normal", "fire": "Ateş", "water": "Su", "electric": "Elektrik",
        "grass": "Çim", "ice": "Buz", "fighting": "Dövüş", "poison": "Zehir",
        "ground": "Yer", "flying": "Uçan", "psychic": "Psişik", "bug": "Böcek",
        "rock": "Kaya", "ghost": "Hayalet", "dragon": "Ejderha", "dark": "Karanlık",
        "steel": "Çelik", "fairy": "Peri"
    }
    TYPE_ICON_ID = {
        "normal": 1, "fighting": 2, "flying": 3, "poison": 4,
        "ground": 5, "rock": 6, "bug": 7, "ghost": 8, "steel": 9,
        "fire": 10, "water": 11, "grass": 12, "electric": 13,
        "psychic": 14, "ice": 15, "dragon": 16, "dark": 17, "fairy": 18
    }

    TYPE_ICON_BASE = "https://raw.githubusercontent.com/WatWowMap/wwm-uicons/main/type/"
    def format_types(self, types):
        if not types:
            return "-"
        vals = []
        for t in types:
            if isinstance(t, dict):
                t = t.get("name") or t.get("type", {}).get("name")
            t = str(t).lower()
            vals.append(self.TYPE_TR.get(t, t.title()))
        return " / ".join(vals)

    def pokemon_type_text(self, poke, name):
        if isinstance(poke, dict):
            candidates = [
                poke.get("types"),
                poke.get("type"),
                poke.get("typing"),
            ]
            if not any(candidates):
                t1 = poke.get("type1")
                t2 = poke.get("type2")
                if t1 or t2:
                    candidates.append([x for x in (t1, t2) if x])
            for value in candidates:
                if value:
                    if isinstance(value, str):
                        value = [x.strip() for x in re.split(r"[/,]", value) if x.strip()]
                    keys = []
                    for t in value:
                        if isinstance(t, dict):
                            t = t.get("name") or t.get("type", {}).get("name")
                        t = str(t).lower()
                        if t in self.TYPE_TR:
                            keys.append(t)
                    self.type_keys[str(name)] = keys
                    result = self.format_types(keys)
                    if result != "-":
                        return result
        return "-"

    def download_type_icon_bytes(self, type_key):
        if type_key in self.type_icon_bytes:
            return self.type_icon_bytes[type_key]
        type_id = self.TYPE_ICON_ID.get(type_key)
        if not type_id:
            return None
        url = f"{self.TYPE_ICON_BASE}{type_id}.png"
        try:
            req = urllib.request.Request(
                url,
                headers={"User-Agent": "SinKa-PvP-100/14.0"}
            )
            with urllib.request.urlopen(req, timeout=8) as r:
                data = r.read()
            self.type_icon_bytes[type_key] = data
            return data
        except Exception:
            return None

    def load_type_icon(self, type_key):
        if not PIL_OK or not type_key:
            return None
        if type_key in self.type_icon_cache:
            return self.type_icon_cache[type_key]
        try:
            from io import BytesIO
            data = self.download_type_icon_bytes(type_key)
            if not data:
                return None
            img = Image.open(BytesIO(data)).convert("RGBA")
            img.thumbnail((28, 28), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.type_icon_cache[type_key] = photo
            return photo
        except Exception:
            return None

    def gigamax_image_url(self, name):
        """Gigantamax formunun gerçek Home görselini döndürür."""
        raw = str(name).strip()
        raw = re.sub(r"^(?:Gigantamax|G-Max|Dynamax)\s+", "", raw, flags=re.I).strip()
        slug = raw.lower()
        slug = slug.replace("♀", "-f").replace("♂", "-m")
        slug = re.sub(r"[^a-z0-9-]+", "-", slug).strip("-")
        aliases = {
            "farfetch-d": "farfetchd",
            "sirfetch-d": "sirfetchd",
            "mr-mime": "mr-mime",
            "mime-jr": "mime-jr",
            "ho-oh": "ho-oh",
        }
        slug = aliases.get(slug, slug)
        return f"https://img.pokemondb.net/sprites/home/normal/{slug}-gigantamax.png"

    def mega_image_url(self, name):
        """Mega formunun PokémonDB Home görselini döndürür."""
        raw = str(name).strip()
        raw = re.sub(r"^Mega\s+", "", raw, flags=re.I).strip()
        slug = raw.lower().replace("♀", "-f").replace("♂", "-m")
        slug = re.sub(r"[^a-z0-9-]+", "-", slug).strip("-")
        # PvPoke -> PokémonDB form adlandırmaları.
        aliases = {
            "charizard-x": "charizard-mega-x", "charizard-y": "charizard-mega-y",
            "mewtwo-x": "mewtwo-mega-x", "mewtwo-y": "mewtwo-mega-y",
        }
        slug = aliases.get(slug, slug + "-mega" if not slug.endswith("-mega") else slug)
        return f"https://img.pokemondb.net/sprites/home/normal/{slug}.png"

    def pokemon_image_url(self, name):
        raw = str(name).strip()
        base = re.sub(r"\s*\(shadow\)\s*$", "", raw, flags=re.I).strip()
        slug = base.lower()
        slug = slug.replace("♀", "-f").replace("♂", "-m")
        slug = re.sub(r"[^a-z0-9-]+", "-", slug).strip("-")
        aliases = {
            "farfetch-d": "farfetchd",
            "sirfetch-d": "sirfetchd",
            "mr-mime": "mr-mime",
            "mime-jr": "mime-jr",
            "nidoran-f": "nidoran-f",
            "nidoran-m": "nidoran-m",
            "ho-oh": "ho-oh",
        }
        slug = aliases.get(slug, slug)
        return f"https://img.pokemondb.net/sprites/home/normal/{slug}.png"

    def pokeapi_sprite_url(self, name):
        raw = re.sub(r"\s*\(shadow\)\s*$", "", str(name).strip(), flags=re.I)
        slug = raw.lower().replace("♀", "-f").replace("♂", "-m")
        slug = re.sub(r"[^a-z0-9-]+", "-", slug).strip("-")
        aliases = {
            "farfetch-d": "farfetchd",
            "sirfetch-d": "sirfetchd",
            "mr-mime": "mr-mime",
            "mime-jr": "mime-jr",
        }
        slug = aliases.get(slug, slug)
        return f"https://pokeapi.co/api/v2/pokemon/{slug}"

    def _download_pokeapi_artwork(self, name):
        try:
            req = urllib.request.Request(
                self.pokeapi_sprite_url(name),
                headers={"User-Agent": "SinKa-PvP-100/14.0"}
            )
            with urllib.request.urlopen(req, timeout=5) as r:
                data = json.loads(r.read().decode("utf-8"))

            sprite = (
                data.get("sprites", {})
                .get("other", {})
                .get("official-artwork", {})
                .get("front_default")
            )
            if not sprite:
                return None

            req2 = urllib.request.Request(
                sprite, headers={"User-Agent": "SinKa-PvP-100/14.0"}
            )
            with urllib.request.urlopen(req2, timeout=5) as r:
                return r.read()
        except Exception:
            return None

    def _hybrid_pokemon_image_url(self, name, gigantamax=False):
        """HybridShivam/Pokemon deposundan görsel URL'si üretir."""
        try:
            base = self._gigamax_base_name(name)
            sid, poke = self._find_pokemon_data_by_name(base)
            dex = (poke or {}).get("dex") if isinstance(poke, dict) else None
            if dex is None and isinstance(poke, dict):
                dex = poke.get("pokedex")
            dex = int(dex) if dex is not None else None
            if not dex:
                return None
            suffix = "-Gmax" if gigantamax else ""
            return ("https://raw.githubusercontent.com/HybridShivam/Pokemon/master/"
                    f"assets/images/{dex:04d}{suffix}.png")
        except Exception:
            return None

    def download_image_bytes(self, name):
        if name in self.image_bytes:
            return self.image_bytes[name]

        data = None
        raw_name = str(name).strip()
        is_gmax = bool(re.match(r"^(?:Gigantamax|G-Max)\s+", raw_name, flags=re.I))
        is_mega = bool(re.match(r"^Mega\s+", raw_name, flags=re.I))

        # GigaMax için gerçek form görselini doğrudan form dosyasından dene.
        if is_gmax and not data:
            try:
                hybrid_url = self._hybrid_pokemon_image_url(raw_name, gigantamax=True)
                if hybrid_url:
                    req = urllib.request.Request(hybrid_url, headers={"User-Agent": "SinKa-PvP/1.0"})
                    with urllib.request.urlopen(req, timeout=7) as r:
                        data = r.read()
            except Exception:
                data = None

        # Mega liginde gerçek Mega form görselini önce dene.
        if is_mega:
            try:
                req = urllib.request.Request(
                    self.mega_image_url(raw_name),
                    headers={"User-Agent": "SinKa-PvP-100/15.0"}
                )
                with urllib.request.urlopen(req, timeout=7) as r:
                    data = r.read()
            except Exception:
                data = None

        # Gigamax liginde gerçek Gigantamax görselini dene.
        if is_gmax and not data:
            try:
                req = urllib.request.Request(
                    self.gigamax_image_url(raw_name),
                    headers={"User-Agent": "SinKa-PvP-100/15.0"}
                )
                with urllib.request.urlopen(req, timeout=7) as r:
                    data = r.read()
            except Exception:
                data = None

        # Normal Pokémon görsel sistemi: Grid/Ultra/Master davranışı aynen korunur.
        if not data:
            try:
                base_name = re.sub(
                    r"^(?:Gigantamax|G-Max|Dynamax)\s+",
                    "",
                    raw_name,
                    flags=re.I
                ).strip()
                req = urllib.request.Request(
                    self.pokemon_image_url(base_name),
                    headers={"User-Agent": "SinKa-PvP-100/15.0"}
                )
                with urllib.request.urlopen(req, timeout=7) as r:
                    data = r.read()
            except Exception:
                data = None

        # PokémonDB başarısız olursa normal form için HybridShivam deposunu dene.
        if not data:
            try:
                base_name = re.sub(
                    r"^(?:Gigantamax|G-Max|Dynamax)\s+", "", raw_name, flags=re.I
                ).strip()
                hybrid_url = self._hybrid_pokemon_image_url(base_name, gigantamax=False)
                if hybrid_url:
                    req = urllib.request.Request(hybrid_url, headers={"User-Agent": "SinKa-PvP/1.0"})
                    with urllib.request.urlopen(req, timeout=7) as r:
                        data = r.read()
            except Exception:
                data = None

        if not data:
            base_name = re.sub(
                r"^(?:Gigantamax|G-Max|Dynamax)\s+",
                "",
                raw_name,
                flags=re.I
            ).strip()
            data = self._download_pokeapi_artwork(base_name)

        if data:
            self.image_bytes[name] = data
        return data

    def load_pokemon_image(self, name):
        if not PIL_OK:
            return None
        if name in self.image_cache:
            return self.image_cache[name]

        try:
            from io import BytesIO
            data = self.download_image_bytes(name)
            if not data:
                return None

            img = Image.open(BytesIO(data)).convert("RGBA")
            img.thumbnail((44, 44), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.image_cache[name] = photo
            return photo
        except Exception:
            return None

    def add_images_async(self, rows, gen):
        if not PIL_OK:
            self.root.after(
                0,
                lambda: self.loading_var.set(
                    "Görseller için Pillow eksik: CMD'de  py -m pip install pillow  yaz."
                )
            )
            return

        from concurrent.futures import ThreadPoolExecutor, as_completed
        from io import BytesIO

        total = len(rows)
        completed = 0

        def fetch(row):
            name = str(row[3])
            if getattr(self, "loaded_league", "") == "Gigamax":
                # Dittobase iki tip kayıt verebilir:
                # Gigantamax X -> gerçek G-Max görseli
                # Eternatus / Zacian vb. -> normal Max/Dynamax görseli
                base = self._gigamax_base_name(name)
                if re.match(r"^(?:Gigantamax|G-Max)\s+", name, flags=re.I):
                    source_name = "Gigantamax " + base
                else:
                    source_name = base
            else:
                source_name = getattr(self, "_main_image_aliases", {}).get(name, name)
            return name, self.download_image_bytes(source_name)

        with ThreadPoolExecutor(max_workers=10) as pool:
            futures = [pool.submit(fetch, row) for row in rows]

            for future in as_completed(futures):
                if gen != self.generation:
                    return

                name, data = future.result()
                completed += 1

                def finish_one(n=name, d=data, done=completed):
                    if d:
                        try:
                            img = Image.open(BytesIO(d)).convert("RGBA")
                            img.thumbnail((44, 44), Image.LANCZOS)
                            photo = ImageTk.PhotoImage(img)
                            self.image_cache[n] = photo
                            self.photo_refs.append(photo)
                        except Exception:
                            pass

                    if d:
                        self._set_tree_image(n, photo)
                    self.root.after_idle(self.redraw_cell_images)
                    self.update_progress(
                        (done / max(1, total)) * 100,
                        f"Pokémon görselleri yükleniyor... {done}/{total}"
                    )

                self.root.after(0, finish_one)

        if gen == self.generation:
            self.root.after(
                0,
                lambda: self.loading_var.set(f"✓ {len(rows)} Pokémon ve görseller hazır")
            )

    def set_row_image(self, name, photo):
        self.image_cache[name] = photo
        self._set_tree_image(name, photo)

    def mark_image_failed(self, name):
        pass

    def _set_tree_image(self, name, photo):
        if photo is None:
            return
        self.image_cache[name] = photo
        iid = self.tree_items_by_name.get(str(name))
        if iid:
            try:
                self.tree.item(iid, image=photo)
                if not hasattr(self, "tree_photo_refs"):
                    self.tree_photo_refs = []
                self.tree_photo_refs.append(photo)
            except Exception:
                pass

    def _update_tree_images(self):
        for name, iid in getattr(self, "tree_items_by_name", {}).items():
            photo = self.image_cache.get(name)
            if photo is not None:
                try: self.tree.item(iid, image=photo)
                except Exception: pass

    def schedule_overlay_redraw(self):
        if self._overlay_job is not None:
            try: self.root.after_cancel(self._overlay_job)
            except Exception: pass
        self._overlay_job = self.root.after_idle(self.redraw_cell_images)

    def _overlay_bg(self, idx):
        return "#eeeeee" if idx % 2 else "#ffffff"

    def _clear_overlay_widgets(self):
        for w in self.overlay_widgets:
            try: w.destroy()
            except Exception: pass
        self.overlay_widgets = []

    def _make_overlay_label(self, photo, x, y, w, h, bg, image_only=True):
        if photo is None: return None
        lbl = tk.Label(self.tree, image=photo, bg=bg, bd=0, highlightthickness=0)
        lbl.place(x=int(x), y=int(y), width=max(1,int(w)), height=max(1,int(h)))
        self.overlay_widgets.append(lbl)
        lbl.bind("<MouseWheel>", lambda e: self._overlay_wheel(e))
        lbl.bind("<Button-4>", lambda e: self._overlay_wheel(e))
        lbl.bind("<Button-5>", lambda e: self._overlay_wheel(e))
        return lbl

    def _overlay_wheel(self, event):
        delta = 0
        if getattr(event, "num", None) == 4: delta = -3
        elif getattr(event, "num", None) == 5: delta = 3
        else: delta = -3 if event.delta > 0 else 3
        self.tree.yview_scroll(delta, "units")
        self.schedule_overlay_redraw()
        return "break"

    def redraw_cell_images(self):
        self._overlay_job = None
        if not PIL_OK or not hasattr(self, "tree"):
            return
        self._clear_overlay_widgets()
        display = list(self.tree["displaycolumns"])
        try:
            image_col = display.index("image") + 1
            type_col = display.index("type") + 1
            evo_col = display.index("evolution") + 1
            move_cols = {c: display.index(c) + 1 for c in ("move1", "move2", "move3") if c in display}
        except ValueError:
            return

        for idx, item in enumerate(self.tree.get_children()):
            bg = self._overlay_bg(idx)
            name = str(self.tree.set(item, "pokemon"))
            bb = self.tree.bbox(item, f"#{image_col}")
            if bb:
                x,y,w,h = bb
                photo = self.image_cache.get(name)
                if photo:
                    self._make_overlay_label(photo, x+4, y+4, w-8, h-8, bg)

            bb = self.tree.bbox(item, f"#{type_col}")
            if bb:
                x,y,w,h = bb
                keys = self.type_keys.get(name, [])[:2]
                ix = x + w - (35 * len(keys)) - 4
                for key in keys:
                    photo = self.type_icon_cache.get(key)
                    if photo:
                        self._make_overlay_label(photo, ix, y+5, 32, h-10, bg)
                        ix += 35

            bb = self.tree.bbox(item, f"#{evo_col}")
            if bb:
                x,y,w,h = bb
                evo_list = self.evolution_names.get(name, [])[:10]
                ex = x + w - (43 * len(evo_list)) - 4
                for evo_name in evo_list:
                    photo = self.evolution_photo_cache.get(evo_name) or self.image_cache.get(evo_name)
                    if photo:
                        self._make_overlay_label(photo, ex, y+4, 40, h-8, bg)
                        ex += 43

            # Güç 1/2/3 sütunlarına Pokémon GO tip ikonunu metnin hemen sağına yerleştir.
            move_keys = getattr(self, "_visible_move_type_keys", {}).get(name) or self._main_move_type_keys(name)
            for mi, col_name in enumerate(("move1", "move2", "move3")):
                if col_name not in move_cols or mi >= len(move_keys):
                    continue
                bb = self.tree.bbox(item, f"#{move_cols[col_name]}")
                if not bb:
                    continue
                x, y, w, h = bb
                key = move_keys[mi]
                if not key:
                    # Bazı satırlarda cache sadece çevrilmiş görünen isim üzerinden bulunabilir.
                    try:
                        key = self._move_type_key_from_display_name(self.tree.set(item, col_name))
                    except Exception:
                        key = ""
                photo = self.type_icon_cache.get(key) if key else None
                if photo:
                    # Metin soldan başladığı için ikon hücrenin sağında, sabit boyutta tutulur.
                    icon_size = min(24, max(18, h-10))
                    ix = x + w - icon_size - 5
                    self._make_overlay_label(photo, ix, y + (h-icon_size)//2, icon_size, icon_size, bg)

    def reposition_image_canvas(self, event=None):
        self.schedule_overlay_redraw()

    def load_type_icons_async(self, gen, needed_keys=None):
        if not PIL_OK:
            return
        keys = list(needed_keys) if needed_keys else list(self.TYPE_ICON_ID.keys())
        from concurrent.futures import ThreadPoolExecutor, as_completed
        def fetch(key):
            return key, self.download_type_icon_bytes(key)
        with ThreadPoolExecutor(max_workers=min(10, max(1, len(keys)))) as pool:
            futures = [pool.submit(fetch, k) for k in keys]
            for f in as_completed(futures):
                if gen != self.generation:
                    return
                try:
                    key, data = f.result()
                except Exception:
                    continue
                if data:
                    def finish(k=key, d=data):
                        if gen != self.generation:
                            return
                        try:
                            img = Image.open(BytesIO(d)).convert("RGBA")
                            img.thumbnail((28, 28), Image.LANCZOS)
                            self.type_icon_cache[k] = ImageTk.PhotoImage(img)
                            self.root.after_idle(self.redraw_cell_images)
                        except Exception:
                            pass
                    self.root.after(0, finish)

    def load_evolution_images_async(self, rows, gen):
        if not PIL_OK:
            return
        names=[]; seen=set()
        for row in rows:
            current=str(row[3])
            for evo in self.evolution_names.get(current, []):
                if evo not in seen:
                    seen.add(evo); names.append(evo)
        if not names:
            return
        from concurrent.futures import ThreadPoolExecutor, as_completed
        def fetch(n): return n, self.download_image_bytes(n)
        with ThreadPoolExecutor(max_workers=min(10,max(1,len(names)))) as pool:
            futures=[pool.submit(fetch,n) for n in names]
            for future in as_completed(futures):
                if gen != self.generation: return
                try: n,data=future.result()
                except Exception: continue
                if not data: continue
                def finish(evo_name=n,d=data):
                    if gen != self.generation: return
                    try:
                        img=Image.open(BytesIO(d)).convert("RGBA")
                        img.thumbnail((44,44),Image.LANCZOS)
                        self.evolution_photo_cache[evo_name]=ImageTk.PhotoImage(img)
                        self.root.after_idle(self.redraw_cell_images)
                    except Exception: pass
                self.root.after(0,finish)

    def _ensure_main_move_type_cache_async(self, gen):
        """Ana lig tablosunun hareket tiplerini bir kez arka planda yükler.
        Böylece ilk liste açılışını bloklamaz; cache geldikten sonra Güç 1/2/3
        ikonları otomatik olarak görünür.
        """
        if gen != self.generation or getattr(self, "_move_type_cache_ready", False) or getattr(self, "_move_type_cache_loading", False):
            return
        self._move_type_cache_loading = True

        def worker():
            move_cache = None
            try:
                move_cache = download_json(PVPoke_MOVES_URL)
                if isinstance(move_cache, list):
                    move_cache = {str(m.get("moveId")): m for m in move_cache if isinstance(m, dict) and m.get("moveId")}
                elif not isinstance(move_cache, dict):
                    move_cache = {}
            except Exception:
                move_cache = {}

            def finish():
                if gen != self.generation:
                    self._move_type_cache_loading = False
                    return
                self._move_type_cache_loading = False
                if move_cache:
                    self._sinka_pvpoke_moves_cache = move_cache
                    self._move_type_key_cache.clear()
                    self._main_move_types_cache.clear()
                    self._move_type_cache_ready = True

                    # Görünen satırların hareket tiplerini yeniden çıkar.
                    needed = set()
                    for item in self.tree.get_children():
                        name = str(self.tree.set(item, "pokemon"))
                        keys = self._main_move_type_keys(name)
                        self._visible_move_type_keys[name] = keys
                        needed.update(k for k in keys if k)
                    if needed:
                        threading.Thread(target=self.load_type_icons_async, args=(gen, needed), daemon=True).start()
                    self.root.after_idle(self.redraw_cell_images)
            self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def show_rows(self, rows, gen, league):
        if gen != self.generation:
            return
        # Eski veri üreticileri status sütununu bilmez. Burada tek noktadan
        # mevcut satırları yeni Durum sütununa uyarlıyoruz.
        normalized_rows = []
        for row in rows:
            vals = list(row)
            name = str(vals[3]) if len(vals) > 3 else ""
            if len(vals) == 17:  # 16 tablo sütunu + tag
                vals.insert(16, self._get_owned_status(name, league))
            elif len(vals) == 16:  # eski/eksik veri: 15 sütun + tag
                # Eksik kayıtta Mevcut ve Durum hücrelerini de güvenli biçimde ekle.
                vals.insert(15, "☑" if self._is_owned(name, league) else "☐")
                vals.insert(16, self._get_owned_status(name, league))
            normalized_rows.append(tuple(vals))
        rows = normalized_rows
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree_items_by_name = {}
        needed=set()
        self._visible_move_type_keys = {}
        for row in rows:
            try:
                row_name = str(row[3])
                needed.update(self.type_keys.get(row_name, []))
                keys = self._main_move_type_keys(row_name)
                self._visible_move_type_keys[row_name] = keys
                needed.update(k for k in keys if k)
            except Exception: pass
        threading.Thread(target=self.load_type_icons_async,args=(gen,needed),daemon=True).start()
        threading.Thread(target=self.load_evolution_images_async,args=(rows,gen),daemon=True).start()
        # Ranking verisindeki moveId -> type eşlemesini bir kez arka planda
        # getir; ilk açılış beklemeden devam eder, ikonlar sonradan tamamlanır.
        self._ensure_main_move_type_cache_async(gen)

        for idx,row in enumerate(rows):
            tagname=row[-1]
            combined=f"{tagname}_{'even' if idx%2==0 else 'odd'}"
            iid=self.tree.insert("","end",text="",values=row[:-1],tags=(combined,))
            self.tree_items_by_name[str(row[3])] = iid

        filters=[]
        filters.append("Shadow dahil" if self.shadow_var.get() else "Shadow hariç")
        filters.append("Efsanevi hariç" if self.legendary_var.get() else "Efsanevi dahil")
        filters.append("Mistik hariç" if self.mythical_var.get() else "Mistik dahil")
        if self.xl_var.get(): filters.append("XL gerektirmeyen")
        if self.special_var.get(): filters.append("Özel güç gerektirmeyen")
        self.update_progress(0,f"Pokémon görselleri yükleniyor... 0/{len(rows)}")
        if PIL_OK:
            threading.Thread(target=self.add_images_async,args=(rows,gen),daemon=True).start()
        self.set_status(f"{league}: {len(rows)} / {self.get_requested_count()} Pokémon hazır • " + " • ".join(filters) + " • XL = seviye 40 üzeri • * = Özel/Legacy hareket")
        self._update_pokemon_count()
        self._refresh_status_summary()
        self._refresh_ready_pokemon_panel()
        self.root.after_idle(self.redraw_cell_images)

    def toggle_original_rank(self):
        cols=("rank","original_rank","image","pokemon","type","evolution","score","iv","level","cp","xl","special","move1","move2","move3","current","status")
        if self.original_rank_var.get():
            self.tree["displaycolumns"] = cols
        else:
            self.tree["displaycolumns"] = tuple(c for c in cols if c != "original_rank")
        self.save_preferences()
        self._refresh_selected_filters_label()
        self.schedule_overlay_redraw()

    def _update_pokemon_count(self):
        """Üst başlıkta sıralanan / mevcut / eksik Pokémon adetlerini gösterir."""
        try:
            items = self.tree.get_children("")
            total = len(items)
            mevcut = 0
            for iid in items:
                try:
                    if str(self.tree.set(iid, "current")).strip() == "☑":
                        mevcut += 1
                except Exception:
                    pass
            eksik = max(0, total - mevcut)
            if hasattr(self, "pokemon_count_label"):
                self.pokemon_count_label.config(
                    text=f"Sıralanan Pokémon: {total}  |  Mevcut: {mevcut}  |  Eksik: {eksik}"
                )
        except Exception:
            pass

    def toggle_current(self,event):
        try:
            col=self.tree.identify_column(event.x); item=self.tree.identify_row(event.y)
            if not item: return
            display=list(self.tree["displaycolumns"])
            if col != f"#{display.index('current')+1}": return
            name=str(self.tree.set(item,"pokemon"))
            league = self.league_var.get()
            state = not self._is_owned(name, league)
            self._set_owned(name, state, league)
            self.save_preferences()
            self.tree.set(item,"current","☑" if state else "☐")
            if not state:
                self.tree.set(item, "status", "")
            else:
                self.tree.set(item, "status", self._get_owned_status(name, league))
            self._update_pokemon_count()
            return "break"
        except Exception:
            return

    def set_status(self, text):
        self.root.after(0, lambda: self.status_var.set(text))

    def _ensure_openpyxl(self):
        global OPENPYXL_OK, Workbook, XLImage, Font, PatternFill, Alignment, Border, Side, DataValidation
        if OPENPYXL_OK:
            return True
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "openpyxl"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=120,
            )
            from openpyxl import Workbook as _Workbook
            from openpyxl.drawing.image import Image as _XLImage
            from openpyxl.styles import Font as _Font, PatternFill as _PatternFill, Alignment as _Alignment, Border as _Border, Side as _Side
            from openpyxl.worksheet.datavalidation import DataValidation as _DataValidation
            Workbook, XLImage, Font, PatternFill, Alignment, Border, Side, DataValidation = (
                _Workbook, _XLImage, _Font, _PatternFill, _Alignment, _Border, _Side, _DataValidation
            )
            OPENPYXL_OK = True
            return True
        except Exception:
            return False

    def _excel_image_file(self, src_path, out_dir, safe, suffix=""):
        if not src_path or not Path(src_path).exists():
            return None
        try:
            if PIL_OK:
                img = Image.open(src_path)
                if img.mode not in ("RGB", "RGBA"):
                    img = img.convert("RGBA")
                out = Path(out_dir) / f"{safe}{suffix}_excel.png"
                img.save(out, format="PNG")
                return out
        except Exception:
            pass
        return Path(src_path)

    def save_a4_png(self):
        items = self.tree.get_children()
        if not items:
            messagebox.showinfo("SinKA PvP", "Önce listeyi yükle.")
            return
        if not PIL_OK:
            messagebox.showerror("SinKA PvP", "Resim kaydı için Pillow gerekli. CMD'de: py -m pip install pillow")
            return
        # Varsayılan isim: ilk açılışta Great Meta öner. Kullanıcı değiştirebilir.
        path = filedialog.asksaveasfilename(
            title="A4 yatay pafta PNG kaydet",
            defaultextension=".png",
            initialfile="Great Meta.png",
            filetypes=[("PNG görseli", "*.png"), ("Tüm dosyalar", "*.*")]
        )
        if not path:
            return

        # Uzantı yazılmadıysa güvenli biçimde .png ekle.
        if not Path(path).suffix:
            path = str(Path(path).with_suffix(".png"))

        # Aynı dosya adına tekrar kayıt yapılırken kesinlikle worker başlatma;
        # önce kullanıcıdan açıkça üzerine yazma onayı al. Bu, özellikle aynı
        # dosyayı art arda kaydederken oluşan donma/çakışma sorununu önler.
        if Path(path).exists():
            overwrite = messagebox.askyesno(
                "Dosya zaten mevcut",
                f"Bu dosya zaten mevcut:\n\n{Path(path).name}\n\nMevcut dosya değiştirilsin mi?",
                parent=self.root
            )
            if not overwrite:
                self.set_status("Resim kaydı iptal edildi.")
                return

        # ÖNEMLİ: Kaydetme işlemi başladıktan sonra Tkinter'a ve ağ bağlantısına
        # arka thread'den kesinlikle dokunmuyoruz. Aksi halde Windows'ta ana pencere
        # kilitlenmiş gibi davranabiliyor. Mevcut önbellekleri ana thread'de kopyala.
        display_cols = list(self.tree["displaycolumns"])
        row_snapshot = []
        needed_names = set()
        for item in items:
            name = str(self.tree.set(item, "pokemon"))
            row_values = {c: self.tree.set(item, c) for c in display_cols}
            # Durum hücresi herhangi bir nedenle displaycolumns dışında kalmışsa
            # kayıt sırasında tekrar lig bazlı kaynaktan üret. Böylece PNG'de
            # Durum bilgisi kaybolmaz.
            if "status" in display_cols:
                row_values["status"] = str(row_values.get("status", "") or self._get_owned_status(name, self.league_var.get()) or "")
            row_snapshot.append({
                "pokemon": name,
                "values": row_values,
                "move_type_keys": (getattr(self, "_visible_move_type_keys", {}).get(name) or self._main_move_type_keys(name)),
                "tag": str(self.tree.item(item, "tags")),
            })
            if name:
                needed_names.add(name)

        # Evrim adlarını da snapshot'a al; worker içinde self.evolution_names okunmayacak.
        evo_snapshot = {}
        for name in list(needed_names):
            vals = self.evolution_names.get(name, [])
            evo_snapshot[name] = list(vals) if isinstance(vals, (list, tuple)) else []
            needed_names.update(evo_snapshot[name])

        # Yalnızca daha önce indirilmiş görselleri kullan. PNG oluştururken yeni HTTP
        # isteği yapılmayacak; böylece kaydetme işlemi saatlerce beklemeyecek.
        image_snapshot = {}
        for name in list(needed_names):
            data = self.image_bytes.get(name)
            if data:
                image_snapshot[name] = data
        type_icon_snapshot = dict(self.type_icon_bytes)
        # PNG kaydını başlatmadan önce ana thread'de ağ isteği yapma.
        # Eksik ikonlar varsa worker bunları atlar; böylece aynı dosyaya
        # tekrar kayıt sırasında arayüzün donması engellenir.
        for rr in row_snapshot:
            for key in rr.get("move_type_keys", []) or []:
                if key and key in self.type_icon_bytes and key not in type_icon_snapshot:
                    type_icon_snapshot[key] = self.type_icon_bytes[key]
        type_keys_snapshot = dict(self.type_keys)
        format_types_snapshot = dict(self.TYPE_TR)

        snapshot = {
            "rows": row_snapshot,
            "display_cols": display_cols,
            "league": self.league_var.get(),
            "cp_max": LEAGUES.get(self.league_var.get(), 1500),
            "xl": bool(self.xl_var.get()),
            "shadow": bool(self.shadow_var.get()),
            "legendary": bool(self.legendary_var.get()),
            "mythical": bool(self.mythical_var.get()),
            "special": bool(self.special_var.get()),
            "original_rank": bool(self.original_rank_var.get()),
            "hero_names": [r["pokemon"] for r in row_snapshot[:3] if r.get("pokemon")],
            "image_bytes": image_snapshot,
            "evolution_names": evo_snapshot,
            "type_icon_bytes": type_icon_snapshot,
            "type_keys": type_keys_snapshot,
            "format_types": format_types_snapshot,
        }
        self.set_status("A4 pafta oluşturuluyor... Program kilitlenmeden çalışacak.")
        threading.Thread(target=self._save_a4_png_worker, args=(path, snapshot), daemon=True).start()

    def _save_a4_png_worker(self, path, snapshot=None):
        try:
            snapshot = snapshot or {"rows": [], "display_cols": [], "league": "Great League", "cp_max": 1500, "xl": False, "shadow": False, "legendary": False, "mythical": False, "special": False}
            row_snapshot = snapshot.get("rows", [])
            display_cols = snapshot.get("display_cols", [])
            from PIL import ImageDraw, ImageFont
            from datetime import datetime

            W, H = 4200, 2976
            margin = 42
            header_h = 455
            table_header_h = 110
            # Alt bölümde Durum açıklamaları için ayrıca alan bırak.
            footer_h = 155
            usable_w = W - 2 * margin
            img = Image.new("RGBA", (W, H), (248, 250, 253, 255))
            draw = ImageDraw.Draw(img)
            league = snapshot.get("league", "Great League")
            cp_max = snapshot.get("cp_max", LEAGUES.get(league, 1500))

            def load_font(size, bold=False):
                candidates = [
                    r"C:\\Windows\\Fonts\\segoeuib.ttf" if bold else r"C:\\Windows\\Fonts\\segoeui.ttf",
                    "segoeuib.ttf" if bold else "segoeui.ttf",
                    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                ]
                for f in candidates:
                    try:
                        return ImageFont.truetype(f, size)
                    except Exception:
                        pass
                return ImageFont.load_default()

            f_brand = load_font(118, True)
            f_pvp = load_font(62, False)
            f_league = load_font(122, True)
            f_cp = load_font(44, True)
            f_head = load_font(38, True)
            f_cell = load_font(42, False)
            f_cell_bold = load_font(42, True)
            f_small = load_font(31, False)
            f_small_bold = load_font(31, True)
            f_footer = load_font(30, False)

            NAVY = (9, 31, 78)
            NAVY2 = (14, 48, 105)
            BLUE = (29, 101, 218)
            CYAN = (30, 194, 180)
            GOLD = (245, 178, 42)
            RED = (218, 35, 48)
            GREEN = (0, 145, 76)
            TEXT = (18, 34, 58)
            LIGHT = (247, 250, 253)
            GRID = (211, 220, 231)

            hx0, hy0 = margin, margin
            hx1, hy1 = W - margin, margin + header_h
            draw.rectangle((hx0, hy0, hx1, hy1), fill=(255,255,255,255), outline=NAVY, width=5)

            draw.polygon([
                (780, hy0), (1030, hy0), (1280, hy1), (1000, hy1)
            ], fill=(37, 112, 222))
            draw.polygon([
                (680, hy0), (850, hy0), (1100, hy1), (930, hy1)
            ], fill=(21, 69, 155))
            draw.polygon([
                (3220, hy0), (3510, hy0), (3280, hy1), (3020, hy1)
            ], fill=(31, 193, 178))
            draw.polygon([
                (3490, hy0), (3740, hy0), (3510, hy1), (3260, hy1)
            ], fill=(14, 75, 153))

            draw.text((95, 92), "SINKA", fill=NAVY, font=f_brand)
            draw.line((105, 218, 560, 218), fill=CYAN, width=8)

            cx, cy, r = 920, 220, 72
            draw.ellipse((cx-r, cy-r, cx+r, cy+r), outline=(255,255,255), width=8)
            draw.line((cx-r, cy, cx+r, cy), fill=(255,255,255), width=7)
            draw.ellipse((cx-18, cy-18, cx+18, cy+18), fill=(255,255,255), outline=NAVY, width=5)

            star_y = 82
            for sx, sr in [(1710, 26), (1800, 40), (1890, 26)]:
                draw.regular_polygon((sx, star_y, sr), n_sides=5, rotation=0, fill=GOLD)
            draw.line((1550, 140, 2050, 140), fill=GOLD, width=5)
            draw.text((1800, 226), league.upper(), fill=NAVY, font=f_league, anchor="mm")
            draw.rounded_rectangle((1630, 305, 1970, 375), radius=35, fill=NAVY)
            cp_badge_text = "SINIRSIZ CP MAX" if league == "Master League" else f"{cp_max} CP MAX"
            draw.text((1800, 340), cp_badge_text, fill=(255,255,255), font=f_cp, anchor="mm")

            hero_x = 2200
            hero_names = []
            for nm in snapshot.get("hero_names", []):
                if nm and nm not in hero_names:
                    hero_names.append(nm)
            for j, nm in enumerate(hero_names[:3]):
                data = snapshot.get("image_bytes", {}).get(nm)
                if data:
                    try:
                        himg = Image.open(BytesIO(data)).convert("RGBA")
                        himg.thumbnail((430, 430), Image.LANCZOS)
                        xx = hero_x + j * 290
                        yy = hy0 + 15 + (430-himg.height)//2
                        img.alpha_composite(himg, (xx, yy))
                    except Exception:
                        pass

            rx = W - 105
            draw.text((rx, 92), datetime.now().strftime("%d %B %Y"), fill=TEXT, font=f_small, anchor="ra")
            draw.line((rx-330, 125, rx, 125), fill=CYAN, width=5)
            draw.text((rx, 162), f"{len(row_snapshot)} Pokémon", fill=NAVY, font=f_small_bold, anchor="ra")

            xl_on = snapshot.get("xl", False)
            shadow_on = snapshot.get("shadow", False)
            legendary_on = snapshot.get("legendary", False)
            mythical_on = snapshot.get("mythical", False)
            special_on = snapshot.get("special", False)
            xl_status = "UYGULANDI" if xl_on else "UYGULANMADI"
            shadow_status = "DAHİL EDİLDİ" if shadow_on else "DAHİL EDİLMEDİ"
            legendary_status = "DAHİL EDİLMEDİ" if legendary_on else "DAHİL EDİLDİ"
            mythical_status = "DAHİL EDİLMEDİ" if mythical_on else "DAHİL EDİLDİ"
            special_status = "UYGULANDI" if special_on else "UYGULANMADI"
            draw.text((rx, 198), f"XL Filtresi: {xl_status}", fill=GREEN if xl_on else TEXT, font=f_small_bold, anchor="ra")
            draw.text((rx, 231), f"Shadow: {shadow_status}", fill=RED if shadow_on else TEXT, font=f_small_bold, anchor="ra")
            draw.text((rx, 264), f"Efsanevi: {legendary_status}", fill=GOLD if legendary_on else TEXT, font=f_small_bold, anchor="ra")
            draw.text((rx, 297), f"Mistik: {mythical_status}", fill=GOLD if mythical_on else TEXT, font=f_small_bold, anchor="ra")
            draw.text((rx, 330), f"Özel Güç Filtresi: {special_status}", fill=GREEN if special_on else TEXT, font=f_small_bold, anchor="ra")

            draw.rectangle((hx0, hy1-52, hx1, hy1), fill=NAVY)
            draw.text((95, hy1-26), "PvPoke Overall Sıralaması  •  Rank 1 IV  •  XL  •  Özel Güç", fill=(255,255,255), font=f_small, anchor="lm")
            draw.text((W-100, hy1-26), "SINKA", fill=(255,255,255), font=f_small_bold, anchor="rm")

            # PNG sütunlarını Treeview'ın gerçek displaycolumns sırasından üret.
            # Böylece O-R-J-R açılıp/kapatıldığında sütun sayısı ile x_positions
            # hiçbir zaman birbirinden kopmaz ve IndexError oluşmaz.
            label_map = {
                "rank": "Rank", "original_rank": "O-R-J-R", "image": "Görsel",
                "pokemon": "Pokémon", "type": "Tip", "evolution": "Önceki Evrimler",
                "score": "PvPoke", "iv": "Rank 1 IV", "level": "Lvl", "cp": "CP",
                "xl": "XL", "special": "Özel Güç", "move1": "Güç 1",
                "move2": "Güç 2", "move3": "Güç 3", "current": "Mevcut",
                "status": "Durum"
            }
            width_map = {
                "rank": 80, "original_rank": 100, "image": 120, "pokemon": 350,
                "type": 230, "evolution": 720, "score": 120, "iv": 190,
                "level": 90, "cp": 100, "xl": 90, "special": 135,
                "move1": 270, "move2": 270, "move3": 270, "current": 115,
                "status": 125
            }
            # Sadece bilinen/geçerli sütunları kullan; eksik sütunlar sessizce
            # atlanır, row çizimi de aynı liste üzerinden yürür.
            safe_display_cols = [c for c in display_cols if c in label_map]
            if not safe_display_cols:
                safe_display_cols = ["rank", "pokemon", "current", "status"]
            display_cols = safe_display_cols
            cols = [(label_map[c], width_map.get(c, 120)) for c in display_cols]
            scale = usable_w / max(1, sum(w for _, w in cols))
            cols = [(name, max(1, int(w * scale))) for name, w in cols]
            x_positions = [margin]
            for _, w in cols:
                x_positions.append(x_positions[-1] + w)

            table_top = hy1 + 24
            available_h = H - table_top - footer_h - margin
            row_h = max(74, int((available_h - table_header_h) / max(1, len(row_snapshot))))
            if row_h < 74:
                row_h = 74

            draw.rectangle((margin, table_top, W-margin, table_top+table_header_h), fill=NAVY, outline=NAVY, width=2)
            for i, (label, _) in enumerate(cols):
                x0, x1 = x_positions[i], x_positions[i+1]
                draw.line((x1, table_top, x1, table_top+table_header_h), fill=(80,105,145), width=2)
                draw.text(((x0+x1)//2, table_top+table_header_h/2), label, fill=(255,255,255), font=f_head, anchor="mm")

            def fit_text(text, max_px, font):
                text = str(text or "")
                if draw.textbbox((0,0), text, font=font)[2] <= max_px:
                    return text
                while len(text) > 1 and draw.textbbox((0,0), text + "…", font=font)[2] > max_px:
                    text = text[:-1]
                return text + "…"

            def fit_font(text, max_px, start_size, min_size=12, bold=False):
                text = str(text or "")
                for size in range(int(start_size), int(min_size)-1, -1):
                    font = load_font(size, bold)
                    if draw.textbbox((0, 0), text, font=font)[2] <= max_px:
                        return font
                return load_font(int(min_size), bold)

            y = table_top + table_header_h
            for idx, row_data in enumerate(row_snapshot):
                row_map = row_data.get("values", {})
                row_bg = (255,255,255) if idx % 2 == 0 else (245,248,252)
                draw.rectangle((margin, y, W-margin, y+row_h), fill=row_bg, outline=GRID, width=1)
                tag = str(row_data.get("tag", ""))
                red = "red" in tag
                text_fill = RED if red else GREEN
                name = str(row_data.get("pokemon", ""))

                image_skip = {"image", "type", "evolution", "xl", "special", "current"}
                left_cols = {"pokemon", "move1", "move2", "move3"}
                for ci, col_name in enumerate(display_cols):
                    # display_cols ve x_positions aynı kaynaktan üretildiği için
                    # normalde bu sınır aşılmaz; yine de worker tarafında güvenli kal.
                    if ci + 1 >= len(x_positions):
                        break
                    x0, x1 = x_positions[ci], x_positions[ci+1]
                    draw.line((x1, y, x1, y+row_h), fill=GRID, width=1)
                    if col_name in image_skip:
                        continue
                    val = row_map.get(col_name, "")
                    left = col_name in left_cols
                    if col_name == "pokemon":
                        txt = val
                        font = fit_font(txt, x1-x0-24, 42, 22, True)
                    elif col_name in {"move1", "move2", "move3"}:
                        txt = fit_text(val, x1-x0-18, f_small)
                        font = f_small
                    else:
                        txt = fit_text(val, x1-x0-22, f_cell)
                        font = f_cell
                    if left:
                        draw.text((x0+12, y+row_h/2), txt, fill=text_fill, font=font, anchor="lm")
                    else:
                        draw.text(((x0+x1)/2, y+row_h/2), txt, fill=text_fill if col_name not in ("rank", "original_rank") else TEXT, font=font, anchor="mm")

                # Güç 1/2/3 hücrelerinde gerçek Pokémon GO tip ikonları.
                move_type_keys = row_data.get("move_type_keys", ["", "", ""])
                for mi, col_name in enumerate(("move1", "move2", "move3")):
                    if col_name not in display_cols or mi >= len(move_type_keys):
                        continue
                    key = str(move_type_keys[mi] or "").lower()
                    if not key:
                        continue
                    mci = display_cols.index(col_name)
                    mx0, mx1 = x_positions[mci], x_positions[mci+1]
                    mdata = snapshot.get("type_icon_bytes", {}).get(key)
                    if mdata:
                        try:
                            mimg = Image.open(BytesIO(mdata)).convert("RGBA")
                            mimg.thumbnail((30, 30), Image.LANCZOS)
                            # İkonu hücrenin sağında tut; metinle çakışmaz.
                            img.alpha_composite(mimg, (int(mx1 - mimg.width - 8), int(y + (row_h-mimg.height)/2)))
                        except Exception:
                            pass

                image_ci = display_cols.index("image") if "image" in display_cols else -1
                data = snapshot.get("image_bytes", {}).get(name)
                if data and image_ci >= 0 and image_ci + 1 < len(x_positions):
                    try:
                        pimg = Image.open(BytesIO(data)).convert("RGBA")
                        pimg.thumbnail((row_h-8, row_h-8), Image.LANCZOS)
                        img.alpha_composite(pimg, (int((x_positions[image_ci]+x_positions[image_ci+1]-pimg.width)/2), int(y+(row_h-pimg.height)/2)))
                    except Exception:
                        pass

                keys = snapshot.get("type_keys", {}).get(name, [])
                type_ci = display_cols.index("type") if "type" in display_cols else -1
                tx = x_positions[type_ci] + 10 if type_ci >= 0 and type_ci + 1 < len(x_positions) else 0
                for key in keys[:2]:
                    tdata = snapshot.get("type_icon_bytes", {}).get(key)
                    if tdata:
                        try:
                            timg = Image.open(BytesIO(tdata)).convert("RGBA")
                            timg.thumbnail((38,38), Image.LANCZOS)
                            img.alpha_composite(timg, (int(tx), int(y+(row_h-timg.height)/2)))
                            tx += 44
                        except Exception:
                            pass
                type_map = snapshot.get("format_types", {})
                type_text = " / ".join(type_map.get(str(k).lower(), str(k).title()) for k in keys) if keys else "-"
                if type_ci >= 0 and type_ci + 1 < len(x_positions):
                    type_text = fit_text(type_text, max(20, x_positions[type_ci+1]-tx-12), f_small)
                    draw.text((tx+5, y+row_h/2), type_text, fill=text_fill, font=f_small, anchor="lm")

                evo_names = snapshot.get("evolution_names", {}).get(name, [])
                evo_ci = display_cols.index("evolution") if "evolution" in display_cols else -1
                if evo_ci < 0 or evo_ci + 1 >= len(x_positions):
                    evo_names = []
                ex0 = x_positions[evo_ci] + 10 if evo_ci >= 0 else 0
                cell_w = x_positions[evo_ci+1] - ex0 - 12 if evo_ci >= 0 and evo_ci + 1 < len(x_positions) else 0
                ey = y + row_h/2
                if not evo_names:
                    draw.text((ex0, ey), "Evrim yok", fill=text_fill, font=f_small_bold, anchor="lm")
                else:
                    target_icon_size = max(44, row_h - 8)
                    min_icon_size = 18
                    icon_gap = 9
                    arrow_gap = 20
                    chosen = None
                    for icon_size in range(target_icon_size, min_icon_size - 1, -2):
                        for font_size in range(42, 9, -1):
                            candidate_font = load_font(font_size, False)
                            total = 0
                            for j, evo_name in enumerate(evo_names):
                                total += icon_size + icon_gap
                                total += draw.textbbox((0, 0), evo_name, font=candidate_font)[2]
                                if j < len(evo_names) - 1:
                                    total += arrow_gap
                            if total <= cell_w:
                                chosen = (icon_size, candidate_font)
                                break
                        if chosen:
                            break

                    if chosen is None:
                        icon_size = min_icon_size
                        evo_font = load_font(9, False)
                    else:
                        icon_size, evo_font = chosen

                    ex = ex0
                    for j, evo_name in enumerate(evo_names):
                        edata = snapshot["image_bytes"].get(evo_name)
                        if edata:
                            try:
                                eimg = Image.open(BytesIO(edata)).convert("RGBA")
                                eimg.thumbnail((icon_size, icon_size), Image.LANCZOS)
                                img.alpha_composite(eimg, (int(ex), int(y+(row_h-eimg.height)/2)))
                            except Exception:
                                pass
                        ex += icon_size + icon_gap
                        draw.text((ex, ey), evo_name, fill=text_fill, font=evo_font, anchor="lm")
                        ex += draw.textbbox((0, 0), evo_name, font=evo_font)[2]
                        if j < len(evo_names)-1:
                            draw.text((ex, ey), "→", fill=(85,95,110), font=f_small_bold, anchor="lm")
                            ex += arrow_gap

                for col_name in ("xl", "special"):
                    if col_name not in display_cols:
                        continue
                    ci = display_cols.index(col_name)
                    if ci + 1 >= len(x_positions):
                        continue
                    x0, x1 = x_positions[ci], x_positions[ci+1]
                    val = row_map.get(col_name, "")
                    cx, cy = (x0+x1)/2, y+row_h/2
                    if val == "✓":
                        draw.line((cx-15,cy,cx-4,cy+13), fill=GREEN, width=7)
                        draw.line((cx-4,cy+13,cx+18,cy-15), fill=GREEN, width=7)
                    else:
                        draw.line((cx-13,cy-13,cx+13,cy+13), fill=RED, width=7)
                        draw.line((cx+13,cy-13,cx-13,cy+13), fill=RED, width=7)
                    draw.line((x1,y,x1,y+row_h), fill=GRID, width=1)

                if "current" in display_cols:
                    ci = display_cols.index("current")
                    if ci + 1 >= len(x_positions):
                        ci = -1
                else:
                    ci = -1
                checked = row_map.get("current", "") == "☑"
                if ci >= 0:
                    x0, x1 = x_positions[ci], x_positions[ci+1]
                    cx, cy = (x0+x1)/2, y+row_h/2
                    if checked:
                        draw.line((cx-18, cy, cx-5, cy+15), fill=GREEN, width=7)
                        draw.line((cx-5, cy+15, cx+21, cy-18), fill=GREEN, width=7)

                if "status" in display_cols:
                    sci = display_cols.index("status")
                    if sci + 1 < len(x_positions):
                        sx0, sx1 = x_positions[sci], x_positions[sci+1]
                        sval = str(row_map.get("status", "") or "")
                        # Durum, Mevcut işaretli satırda doğrudan ve ayrı
                        # sütunda yazılır: ✕ / + / −.
                        if checked and sval in ("✕", "X", "x", "+", "−", "-"):
                            if sval == "-":
                                sval = "−"
                            elif sval in ("✕", "x"):
                                sval = "X"
                            sfill = RED if sval == "X" else GREEN if sval == "+" else (190, 120, 0)
                            draw.text(((sx0+sx1)/2, y+row_h/2), sval, fill=sfill, font=f_cell_bold, anchor="mm")

                y += row_h

            # ------------------------------------------------------------
            # ALT FOOTER
            # Yeni bölüm açmadan, çizginin ALTINDA tek satırlık footer kullanılır.
            # Durum açıklamaları + site + uygulama adı + tarih aynı satırdadır.
            footer_top = H - margin - footer_h
            footer_font = load_font(23, False)
            footer_bold = load_font(23, True)
            legend_font = load_font(23, False)
            legend_bold = load_font(23, True)

            fy = footer_top + 18
            draw.line((margin+20, fy, W-margin-20, fy), fill=CYAN, width=5)
            draw.ellipse((W//2-10, fy-10, W//2+10, fy+10), fill=CYAN)

            # Her şey tek satırda: Durum açıklamaları -> site -> SinKA PvP -> tarih.
            info_y = fy + 58
            x = margin + 30
            legend_items = [
                ("+", "Güçler tamam veya evrimleştirilebilir", GREEN),
                ("−", "Evrimleştirmek için event bekle", (190, 120, 0)),
                ("X", "Event dışı evrimleştirilmiş — yeniden yakala", RED),
            ]
            for idx, (symbol, description, scolor) in enumerate(legend_items):
                draw.text((x, info_y), symbol, fill=scolor, font=legend_bold, anchor="lm")
                x += draw.textbbox((0, 0), symbol, font=legend_bold)[2] + 10
                draw.text((x, info_y), description, fill=TEXT, font=legend_font, anchor="lm")
                x += draw.textbbox((0, 0), description, font=legend_font)[2] + 38

            # Site, uygulama ve tarih durum açıklamalarının devamında aynı satırda.
            draw.text((x, info_y), "|", fill=CYAN, font=footer_bold, anchor="lm")
            x += draw.textbbox((0, 0), "|", font=footer_bold)[2] + 18
            draw.text((x, info_y), "www.sinka.com", fill=NAVY, font=footer_bold, anchor="lm")
            x += draw.textbbox((0, 0), "www.sinka.com", font=footer_bold)[2] + 28
            draw.text((x, info_y), "|", fill=CYAN, font=footer_bold, anchor="lm")
            x += draw.textbbox((0, 0), "|", font=footer_bold)[2] + 18
            draw.text((x, info_y), "SinKA PvP", fill=NAVY, font=footer_bold, anchor="lm")
            x += draw.textbbox((0, 0), "SinKA PvP", font=footer_bold)[2] + 28
            draw.text((x, info_y), "|", fill=CYAN, font=footer_bold, anchor="lm")
            x += draw.textbbox((0, 0), "|", font=footer_bold)[2] + 18
            date_text = datetime.now().strftime("%d %B %Y")
            draw.text((x, info_y), date_text, fill=NAVY, font=footer_font, anchor="lm")

            # Doğrudan hedef dosyanın üzerine yazmak yerine önce aynı klasörde
            # geçici PNG oluşturup işlem sonunda atomik olarak hedefe taşı.
            # Böylece mevcut dosya için verilen "Evet" kararından sonra da
            # arayüz kilitlenmez ve yarım/bozuk dosya bırakılmaz.
            import os, tempfile
            target = Path(path)
            target.parent.mkdir(parents=True, exist_ok=True)
            fd, temp_path = tempfile.mkstemp(prefix=".sinka_tmp_", suffix=".png", dir=str(target.parent))
            os.close(fd)
            try:
                img.convert("RGB").save(temp_path, "PNG", dpi=(180,180))
                os.replace(temp_path, str(target))
            finally:
                try:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                except Exception:
                    pass
            def _done_ok(p=path):
                self.set_status(f"Tasarımlı A4 PNG kaydedildi: {p}")
                self._show_a4_result_dialog("Resim Kaydedildi", f"A4 görsel başarıyla kaydedildi.\n\n{p}")
            self.root.after(0, _done_ok)
        except Exception as e:
            err = f"{type(e).__name__}: {e}"
            self.root.after(0, lambda err=err: self._show_a4_result_dialog("Resim Kayıt Hatası", f"PNG oluşturulamadı.\n\n{err}"))

    def _show_a4_result_dialog(self, title, text):
        # Windows'ta messagebox bazen ana pencerenin arkasında kalabiliyor.
        # Bu yüzden sonucu kendi görünür penceremizle gösteriyoruz.
        win = tk.Toplevel(self.root)
        win.title(title)
        win.transient(self.root)
        win.resizable(False, False)
        win.configure(padx=24, pady=20)

        try:
            self.root.update_idletasks()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
            ww, wh = 520, 210
            x = rx + max(0, (rw - ww) // 2)
            y = ry + max(0, (rh - wh) // 2)
            win.geometry(f"{ww}x{wh}+{x}+{y}")
        except Exception:
            win.geometry("520x210")

        is_ok = "Hatası" not in title
        icon = "✓" if is_ok else "✖"
        icon_label = tk.Label(win, text=icon, font=("Segoe UI", 28, "bold"))
        icon_label.pack(pady=(0, 6))
        tk.Label(win, text=text, font=("Segoe UI", 10), justify="center", wraplength=470).pack(fill="x")

        def close_dialog():
            try:
                win.grab_release()
            except Exception:
                pass
            try:
                win.destroy()
            finally:
                try:
                    self.root.lift()
                    self.root.focus_force()
                except Exception:
                    pass

        btn = tk.Button(win, text="Tamam", width=12, command=close_dialog)
        btn.pack(pady=(18, 0))
        win.protocol("WM_DELETE_WINDOW", close_dialog)
        win.attributes("-topmost", True)
        win.lift()
        win.focus_force()
        btn.focus_set()
        # Sadece bu sonuç penceresi modal olsun; ana pencere gizli bir
        # messagebox yüzünden kilitlenmiş gibi görünmesin.
        win.grab_set()
        win.after(250, lambda: win.attributes("-topmost", False) if win.winfo_exists() else None)

    def save_csv(self):
        items = self.tree.get_children()
        if not items:
            messagebox.showinfo("SinKA PvP", "Önce listeyi yükle.")
            return

        path = filedialog.asksaveasfilename(
            title="SinKA PvP kaydet",
            defaultextension=".csv",
            filetypes=[("CSV dosyası", "*.csv"), ("Tüm dosyalar", "*.*")]
        )
        if not path:
            return

        try:
            path = os.path.abspath(path)
            base = Path(path).with_suffix("")
            image_dir = Path(str(base) + "_gorseller")
            type_dir = Path(str(base) + "_tip_gorseller")
            evolution_dir = Path(str(base) + "_evrim_gorseller")
            image_dir.mkdir(parents=True, exist_ok=True)
            type_dir.mkdir(parents=True, exist_ok=True)
            evolution_dir.mkdir(parents=True, exist_ok=True)

            headers = ["Rank"]
            if self.original_rank_var.get():
                headers.append("Orijinal Rank")
            headers += [
                "Görsel", "Pokemon", "Tip", "Önceki Evrimler", "PvPoke",
                "Rank 1 IV", "Lvl", "CP", "XL", "Özel Güç", "Güç 1", "Güç 2", "Güç 3", "Mevcut", "Durum"
            ]

            saved_images = 0
            saved_type_icons = 0
            rows_for_excel = []

            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(headers)

                for item in items:
                    vals = [self.tree.set(item, c) for c in self.tree["displaycolumns"]]
                    name = str(self.tree.set(item, "pokemon"))
                    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", name).strip("_") or "pokemon"
                    image_name = safe + ".png"

                    data = self.image_bytes.get(name)
                    if not data:
                        data = self.download_image_bytes(name)
                    image_path = None
                    if data:
                        try:
                            image_path = image_dir / image_name
                            if PIL_OK:
                                img = Image.open(BytesIO(data))
                                if img.mode not in ("RGB", "RGBA"):
                                    img = img.convert("RGBA")
                                img.save(image_path, format="PNG")
                            else:
                                image_path.write_bytes(data)
                            self.image_files[name] = str(image_path)
                            saved_images += 1
                        except Exception:
                            image_path = None
                            image_name = ""

                    type_names = self.type_keys.get(name, [])
                    type_file_paths = []
                    for key in type_names:
                        tdata = self.type_icon_bytes.get(key)
                        if not tdata:
                            tdata = self.download_type_icon_bytes(key)
                        if tdata:
                            tid = self.TYPE_ICON_ID.get(key)
                            tf = f"{tid}_{key}.png"
                            try:
                                tp = type_dir / tf
                                if PIL_OK:
                                    Image.open(BytesIO(tdata)).convert("RGBA").save(tp, format="PNG")
                                else:
                                    tp.write_bytes(tdata)
                                type_file_paths.append(tp)
                                saved_type_icons += 1
                            except Exception:
                                pass

                    evolution_paths = []
                    for evo_name in self.evolution_names.get(name, []):
                        edata = self.image_bytes.get(evo_name)
                        if not edata:
                            edata = self.download_image_bytes(evo_name)
                        if edata:
                            esafe = re.sub(r"[^A-Za-z0-9_-]+", "_", evo_name).strip("_") or "evrim"
                            ep = evolution_dir / f"{esafe}.png"
                            try:
                                if PIL_OK:
                                    Image.open(BytesIO(edata)).convert("RGBA").save(ep, format="PNG")
                                else:
                                    ep.write_bytes(edata)
                                evolution_paths.append(ep)
                            except Exception:
                                pass

                    out_vals = vals[:]
                    try:
                        image_pos = list(self.tree["displaycolumns"]).index("image")
                        out_vals[image_pos] = image_name
                    except Exception:
                        pass
                    writer.writerow(out_vals)
                    rows_for_excel.append((out_vals, name, image_path, type_file_paths, safe, type_names, evolution_paths))

            zip_path = Path(str(base) + "_paket.zip")
            try:
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
                    z.write(path, arcname=Path(path).name)
                    for folder, arc in [(image_dir, "gorseller"), (type_dir, "tip_gorseller"), (evolution_dir, "evrim_gorseller")]:
                        if folder.exists():
                            for img in folder.iterdir():
                                if img.is_file():
                                    z.write(img, arcname=f"{arc}/{img.name}")
            except Exception:
                zip_path = None

            xlsx_path = Path(str(base) + ".xlsx")
            excel_error = None
            saved_xlsx = False

            if self._ensure_openpyxl():
                try:
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "SinKA PvP"
                    ws.freeze_panes = "A2"
                    from openpyxl.utils import get_column_letter
                    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows_for_excel)+1}"

                    for col_idx, header in enumerate(headers, 1):
                        cell = ws.cell(1, col_idx, header)
                        cell.font = Font(bold=True)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    ws.row_dimensions[1].height = 28

                    thin = Side(style="thin")
                    for row_idx, (vals, name, image_path, type_paths, safe, type_names, evolution_paths) in enumerate(rows_for_excel, start=2):
                        for col_idx in range(1, len(headers) + 1):
                            value = vals[col_idx-1] if col_idx-1 < len(vals) else ""
                            ws.cell(row_idx, col_idx, value)

                    wb.save(xlsx_path)
                    saved_xlsx = True
                except Exception as ex:
                    excel_error = str(ex)

            msg = f"Dosya kaydedildi:\n{path}\n\nİndirilen Görseller: {saved_images}"
            if saved_xlsx:
                msg += f"\nExcel (.xlsx): {xlsx_path.name}"
            elif excel_error:
                msg += f"\nExcel Hatası: {excel_error}"
            if zip_path:
                msg += f"\nZIP Paketi: {zip_path.name}"

            messagebox.showinfo("SinKA PvP", msg)
            self.set_status("Kayıt tamamlandı.")
        except Exception as e:
            messagebox.showerror("Kayıt Hatası", f"Dosya kaydedilemedi:\n\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    try: root._sinka_app = app
    except Exception: pass
    root.mainloop()