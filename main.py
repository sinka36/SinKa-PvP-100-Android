# -*- coding: utf-8 -*-
"""
SinKa PvP 100 - Android
Android/Kivy port of the SinKa PvP 100 Windows V24 logic.

Data sources:
- PvPoke Overall rankings
- PvPoke gamemaster
- WatWowMap Turkish move translations
- WatWowMap type icons
- PokemonDB / PokeAPI artwork
"""
import json
import math
import os
import re
import ssl
import threading
from pathlib import Path
from io import BytesIO

import requests
from PIL import Image as PILImage

from kivy.app import App
from kivy.clock import Clock
from kivy.core.image import Image as CoreImage
from kivy.graphics import Color, RoundedRectangle
from kivy.metrics import dp
from kivy.properties import BooleanProperty, NumericProperty, StringProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.gridlayout import GridLayout
from kivy.uix.image import Image
from kivy.uix.label import Label
from kivy.uix.progressbar import ProgressBar
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.uix.widget import Widget


RANKING_URL = "https://raw.githubusercontent.com/pvpoke/pvpoke/master/src/data/rankings/all/overall/rankings-{cp}.json"
POKEMON_URL = "https://raw.githubusercontent.com/pvpoke/pvpoke/master/src/data/gamemaster/pokemon.json"
TR_MOVES_URL = "https://raw.githubusercontent.com/WatWowMap/pogo-data-api/main/data/v1/translations/tr/moves.json"
POGO_MOVES_URL = "https://raw.githubusercontent.com/WatWowMap/pogo-data-api/main/data/v1/moves.json"
TYPE_ICON_BASE = "https://raw.githubusercontent.com/WatWowMap/wwm-uicons/main/type/"

LEAGUES = {"Great League": 1500, "Ultra League": 2500, "Master League": 10000}

TYPE_TR = {
    "normal": "Normal", "fire": "Ateş", "water": "Su", "electric": "Elektrik",
    "grass": "Çim", "ice": "Buz", "fighting": "Dövüş", "poison": "Zehir",
    "ground": "Yer", "flying": "Uçan", "psychic": "Psişik", "bug": "Böcek",
    "rock": "Kaya", "ghost": "Hayalet", "dragon": "Ejderha", "dark": "Karanlık",
    "steel": "Çelik", "fairy": "Peri"
}
TYPE_ICON_ID = {
    "normal": 1, "fighting": 2, "flying": 3, "poison": 4, "ground": 5,
    "rock": 6, "bug": 7, "ghost": 8, "steel": 9, "fire": 10, "water": 11,
    "grass": 12, "electric": 13, "psychic": 14, "ice": 15, "dragon": 16,
    "dark": 17, "fairy": 18
}

# Pokemon GO CP multipliers, half-levels 1.0 through 51.0.
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
# The source V24 uses its full CPM table.  To avoid accidental indexing errors,
# the app derives the actual half-level list from this table length.
LEVELS = [1.0 + i * 0.5 for i in range(len(CPM_VALUES))]
CPM = dict(zip(LEVELS, CPM_VALUES))


def http_json(url, timeout=30):
    r = requests.get(url, headers={"User-Agent": "SinKa-PvP-100/Android"}, timeout=timeout)
    r.raise_for_status()
    return r.json()


def build_move_translator(pogo_moves, tr_moves):
    proto_to_id = {}
    for m in pogo_moves:
        proto = m.get("proto")
        mid = m.get("moveId")
        if proto is not None and mid is not None:
            try:
                proto_to_id[str(proto).upper()] = int(mid)
            except Exception:
                pass

    def translate(move_id):
        mid = proto_to_id.get(str(move_id).upper())
        if mid is not None:
            value = tr_moves.get("move_" + str(mid))
            if value and not str(value).startswith("<<"):
                return str(value)
        return str(move_id).replace("_", " ").title()
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
    # Same exact search strategy as the Windows V24 source:
    # all 4096 IV combinations + binary search for highest legal half-level.
    if cp_limit >= 10000:
        ivs = (15, 15, 15)
        return ivs, 50.0, calc_cp(base, ivs, 50.0)

    best_key = None
    best_ivs = (15, 15, 15)
    best_level = 1.0
    best_cp = 0

    for a in range(16):
        for d in range(16):
            for h in range(16):
                ivs = (a, d, h)
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
                eff_def = (base["def"] + d) * cpm
                eff_hp = math.floor((base["hp"] + h) * cpm)
                key = (product, cp, eff_def, eff_hp, -a)
                if best_key is None or key > best_key:
                    best_key = key
                    best_ivs, best_level, best_cp = ivs, level, cp
    return best_ivs, best_level, best_cp


def is_shadow(item):
    sid = str(item.get("speciesId", ""))
    return sid.endswith("_shadow") or "(Shadow)" in str(item.get("speciesName", ""))


def slugify(name):
    raw = re.sub(r"\s*\(shadow\)\s*$", "", str(name), flags=re.I).strip()
    s = raw.lower().replace("♀", "-f").replace("♂", "-m")
    s = re.sub(r"[^a-z0-9-]+", "-", s).strip("-")
    return {"farfetch-d": "farfetchd", "sirfetch-d": "sirfetchd"}.get(s, s)


def image_url(name):
    return f"https://img.pokemondb.net/sprites/home/normal/{slugify(name)}.png"


def type_icon_url(type_key):
    tid = TYPE_ICON_ID.get(type_key)
    return f"{TYPE_ICON_BASE}{tid}.png" if tid else ""


class HeaderLabel(Label):
    pass


class SinKaApp(App):
    league = StringProperty("Great League")
    show_shadow = BooleanProperty(False)
    no_xl = BooleanProperty(False)
    no_special = BooleanProperty(False)
    requested_count = NumericProperty(36)
    progress = NumericProperty(0)

    def build(self):
        self.title = "SinKa PvP 100"
        self.pokemon = {}
        self.ranking_data = []
        self.loaded_league = ""
        self.translate_move = lambda x: str(x).replace("_", " ").title()
        self.owned = {}
        self.type_keys = {}
        self.evolution_names = {}
        self.generation = 0
        self.data_dir = Path(self.user_data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.state_file = self.data_dir / "state.json"
        self.load_state()

        root = BoxLayout(orientation="vertical", spacing=dp(5), padding=dp(7))
        root.add_widget(self.build_top())
        self.progressbar = ProgressBar(max=100, value=0, size_hint_y=None, height=dp(7))
        root.add_widget(self.progressbar)

        self.status = Label(text="Hazır.", size_hint_y=None, height=dp(30),
                            halign="left", valign="middle")
        self.status.bind(size=lambda inst, val: setattr(inst, "text_size", (inst.width, None)))
        root.add_widget(self.status)

        self.scroll = ScrollView(do_scroll_x=True, do_scroll_y=True)
        self.table = GridLayout(cols=15, spacing=dp(1), size_hint=(None, None),
                                padding=dp(1))
        self.scroll.add_widget(self.table)
        root.add_widget(self.scroll)

        self.start_background(self.initial_load)
        return root

    def build_top(self):
        box = BoxLayout(orientation="vertical", spacing=dp(4), size_hint_y=None, height=dp(155))

        title = Label(text="[b]SINKA[/b] [color=#1d3b63]PvP 100[/color]",
                      markup=True, font_size="25sp", size_hint_y=None, height=dp(34))
        box.add_widget(title)

        controls = BoxLayout(spacing=dp(5), size_hint_y=None, height=dp(42))
        self.spinner = Spinner(text=self.league, values=list(LEAGUES.keys()), size_hint_x=None, width=dp(145))
        self.spinner.bind(text=self.on_league)
        controls.add_widget(self.spinner)

        self.count_input = TextInput(text=str(int(self.requested_count)), multiline=False,
                                     input_filter="int", size_hint_x=None, width=dp(65))
        self.count_input.bind(on_text_validate=self.on_count)
        controls.add_widget(self.count_input)

        apply_btn = Button(text="Uygula", size_hint_x=None, width=dp(75))
        apply_btn.bind(on_release=lambda *_: self.on_count())
        controls.add_widget(apply_btn)

        reload_btn = Button(text="Yenile", size_hint_x=None, width=dp(75))
        reload_btn.bind(on_release=lambda *_: self.reload())
        controls.add_widget(reload_btn)
        box.add_widget(controls)

        filters = BoxLayout(spacing=dp(10), size_hint_y=None, height=dp(45))
        self.shadow_cb = CheckBox(active=self.show_shadow, size_hint_x=None, width=dp(32))
        self.shadow_cb.bind(active=lambda *_: self.filter_changed())
        filters.add_widget(self.shadow_cb)
        filters.add_widget(Label(text="Shadow'ları göster.", halign="left"))

        self.xl_cb = CheckBox(active=self.no_xl, size_hint_x=None, width=dp(32))
        self.xl_cb.bind(active=lambda *_: self.filter_changed())
        filters.add_widget(self.xl_cb)
        filters.add_widget(Label(text="XL gerektirmeyenleri göster.", halign="left"))

        self.special_cb = CheckBox(active=self.no_special, size_hint_x=None, width=dp(32))
        self.special_cb.bind(active=lambda *_: self.filter_changed())
        filters.add_widget(self.special_cb)
        filters.add_widget(Label(text="Özel güç (*) gerektirmeyenleri göster.", halign="left"))
        box.add_widget(filters)

        self.info = Label(text="", font_size="12sp", size_hint_y=None, height=dp(28),
                          halign="left")
        box.add_widget(self.info)
        return box

    def load_state(self):
        try:
            data = json.loads(self.state_file.read_text("utf-8"))
            self.owned = data.get("owned", {})
            self.league = data.get("league", "Great League")
            self.show_shadow = bool(data.get("shadow", False))
            self.no_xl = bool(data.get("no_xl", False))
            self.no_special = bool(data.get("no_special", False))
            self.requested_count = max(1, min(100, int(data.get("count", 36))))
        except Exception:
            pass

    def save_state(self):
        try:
            self.state_file.write_text(json.dumps({
                "owned": self.owned, "league": self.league,
                "shadow": self.show_shadow, "no_xl": self.no_xl,
                "no_special": self.no_special,
                "count": int(self.requested_count)
            }, ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass

    def set_status(self, text):
        self.status.text = text

    def start_background(self, fn, *args):
        threading.Thread(target=fn, args=args, daemon=True).start()

    def initial_load(self):
        try:
            self.set_status("Pokémon ve hareket verileri indiriliyor...")
            pokemon_list = http_json(POKEMON_URL)
            tr_moves = http_json(TR_MOVES_URL)
            pogo_moves = http_json(POGO_MOVES_URL)
            self.pokemon = {p.get("speciesId"): p for p in pokemon_list if p.get("speciesId")}
            self.translate_move = build_move_translator(pogo_moves, tr_moves)
            self.load_league(self.league)
        except Exception as e:
            Clock.schedule_once(lambda *_: self.set_status("Veri yükleme hatası: " + str(e)[:160]), 0)

    def load_league(self, league):
        try:
            Clock.schedule_once(lambda *_: self.set_status(f"{league} sıralaması indiriliyor..."), 0)
            data = http_json(RANKING_URL.format(cp=LEAGUES[league]))
            self.ranking_data = data
            self.loaded_league = league
            Clock.schedule_once(lambda *_: self.refresh(), 0)
        except Exception as e:
            Clock.schedule_once(lambda *_: self.set_status("Lig verisi indirilemedi: " + str(e)[:160]), 0)

    def on_league(self, spinner, value):
        self.league = value
        self.save_state()
        if self.loaded_league == value:
            self.refresh()
        else:
            self.start_background(self.load_league, value)

    def on_count(self, *_):
        try:
            n = int(self.count_input.text)
        except Exception:
            n = 36
        self.requested_count = max(1, min(100, n))
        self.count_input.text = str(int(self.requested_count))
        self.save_state()
        self.refresh()

    def filter_changed(self):
        self.show_shadow = self.shadow_cb.active
        self.no_xl = self.xl_cb.active
        self.no_special = self.special_cb.active
        self.save_state()
        self.refresh()

    def reload(self):
        self.start_background(self.initial_load)

    def refresh(self):
        if not self.ranking_data or not self.pokemon:
            return
        self.generation += 1
        gen = self.generation
        candidates = [p for p in self.ranking_data if self.show_shadow or not is_shadow(p)]
        # Keep ranking order; filters are evaluated after exact Rank-1 data is known.
        self.table.clear_widgets()
        self.add_headers()
        self.progressbar.value = 0
        self.info.text = f"{self.league} • {'Shadow dahil' if self.show_shadow else 'Shadow hariç'}"
        self.set_status("Rank-1 IV hesaplanıyor...")
        self.start_background(self.calculate_rows, candidates, gen)

    def add_headers(self):
        headers = ["Rank","Görsel","Pokémon","Tip","Önceki Evrimler","PvPoke",
                   "Rank 1 IV","Lvl","CP","XL","Özel Güç","Güç 1","Güç 2","Güç 3","Mevcut"]
        widths = [55,65,125,120,210,70,90,55,55,55,75,150,150,150,70]
        for h, w in zip(headers, widths):
            lab = Label(text=f"[b]{h}[/b]", markup=True, size_hint=(None,None),
                        size=(dp(w),dp(42)), halign="center", valign="middle")
            lab.text_size = lab.size
            with lab.canvas.before:
                Color(0.11,0.23,0.39,1)
                RoundedRectangle(pos=lab.pos, size=lab.size, radius=[dp(2)])
            lab.bind(pos=lambda inst, *_: self.sync_bg(inst), size=lambda inst, *_: self.sync_bg(inst))
            self.table.add_widget(lab)

    def sync_bg(self, widget):
        pass

    def calculate_rows(self, candidates, gen):
        rows = []
        target = 0
        total = len(candidates)
        for idx, p in enumerate(candidates):
            if gen != self.generation:
                return
            sid = p.get("speciesId", "")
            name = p.get("speciesName", sid)
            score = float(p.get("score", 0))
            poke = self.pokemon.get(sid) or self.pokemon.get(str(sid).replace("_shadow",""))
            if not poke or "baseStats" not in poke:
                continue
            ivs, level, cp = rank1_iv(poke["baseStats"], LEAGUES[self.league])
            moveset = p.get("moveset") or []
            special_moves = set(poke.get("eliteMoves", [])) | set(poke.get("legacyMoves", []))
            move_names = []
            has_special = False
            for mv in moveset[:3]:
                special = mv in special_moves
                has_special |= special
                move_names.append(self.translate_move(mv) + ("*" if special else ""))
            while len(move_names) < 3:
                move_names.append("-")
            xl = level > 40.0
            if self.no_xl and xl:
                continue
            if self.no_special and has_special:
                continue

            target += 1
            evolution = self.evolution_stage(poke)
            types = self.get_types(poke)
            row = {
                "rank": target, "name": name, "types": types,
                "evolution": evolution, "score": f"{score:.1f}",
                "iv": f"{ivs[0]}/{ivs[1]}/{ivs[2]}",
                "level": f"{level:g}", "cp": str(cp), "xl": xl,
                "special": has_special, "moves": move_names,
                "owned": bool(self.owned.get(name, False))
            }
            rows.append(row)
            pct = (idx + 1) / max(1,total) * 100
            Clock.schedule_once(lambda dt, p=pct, i=idx+1,t=total:
                                self.update_progress(p, f"Veriler işleniyor... {i}/{t}"), 0)
            if target >= int(self.requested_count):
                break

        Clock.schedule_once(lambda *_: self.render_rows(rows, gen), 0)

    def update_progress(self, value, text):
        self.progressbar.value = value
        self.set_status(text)

    def get_types(self, poke):
        vals = poke.get("types") or poke.get("type") or poke.get("typing") or []
        if isinstance(vals, str):
            vals = re.split(r"[/,]", vals)
        keys = []
        for t in vals:
            if isinstance(t, dict):
                t = t.get("name") or t.get("type", {}).get("name")
            t = str(t).lower()
            if t in TYPE_TR:
                keys.append(t)
        return keys

    def evolution_stage(self, poke):
        parent = (poke.get("family") or {}).get("parent")
        previous = []
        seen = set()
        while parent and parent not in seen:
            seen.add(parent)
            pp = self.pokemon.get(parent) or self.pokemon.get(str(parent).replace("_shadow",""))
            if not pp:
                break
            previous.append(pp.get("name") or pp.get("speciesName") or parent)
            parent = (pp.get("family") or {}).get("parent")
        return " → ".join(reversed(previous)) if previous else "Evrim yok"

    def render_rows(self, rows, gen):
        if gen != self.generation:
            return
        for r in rows:
            self.add_row(r)
        self.progressbar.value = 100
        filters = []
        filters.append("Shadow dahil" if self.show_shadow else "Shadow hariç")
        if self.no_xl: filters.append("XL gerektirmeyen")
        if self.no_special: filters.append("Özel güç gerektirmeyen")
        self.info.text = f"{self.league} • {len(rows)} Pokémon • " + " • ".join(filters)
        self.set_status("Hazır.")
        self.load_row_images(rows, gen)

    def cell(self, text, width, row_idx, bold=False):
        bg = (0.93,0.94,0.96,1) if row_idx % 2 == 0 else (1,1,1,1)
        color = (0.75,0,0,1) if self._row_red else (0,0.48,0.05,1)
        lab = Label(text=str(text), markup=bold, size_hint=(None,None),
                    size=(dp(width),dp(64)), halign="center", valign="middle",
                    color=color)
        lab.text_size = lab.size
        with lab.canvas.before:
            Color(*bg)
            RoundedRectangle(pos=lab.pos, size=lab.size, radius=[dp(1)])
        return lab

    def add_row(self, r):
        row_idx = int(r["rank"]) - 1
        self._row_red = bool(r["xl"] or r["special"])
        widths = [55,65,125,120,210,70,90,55,55,55,75,150,150,150,70]

        self.table.add_widget(self.cell(r["rank"], widths[0], row_idx))
        img_box = BoxLayout(size_hint=(None,None), size=(dp(widths[1]),dp(64)))
        img_box.add_widget(Image(source="", allow_stretch=True))
        self.table.add_widget(img_box)
        self.table.add_widget(self.cell(r["name"], widths[2], row_idx, bold=True))

        type_box = BoxLayout(size_hint=(None,None), size=(dp(widths[3]),dp(64)), spacing=dp(2))
        for t in r["types"]:
            im = Image(size_hint=(None,None), size=(dp(25),dp(25)))
            type_box.add_widget(im)
            self.start_background(self.load_image_widget, im, type_icon_url(t))
        type_box.add_widget(Label(text=" / ".join(TYPE_TR.get(t,t) for t in r["types"]),
                                  halign="left", valign="middle"))
        self.table.add_widget(type_box)

        evo_box = BoxLayout(size_hint=(None,None), size=(dp(widths[4]),dp(64)), spacing=dp(3))
        evo_box.add_widget(Label(text=r["evolution"], halign="left", valign="middle"))
        self.table.add_widget(evo_box)

        self.table.add_widget(self.cell(r["score"], widths[5], row_idx))
        self.table.add_widget(self.cell(r["iv"], widths[6], row_idx))
        self.table.add_widget(self.cell(r["level"], widths[7], row_idx))
        self.table.add_widget(self.cell(r["cp"], widths[8], row_idx))
        self.table.add_widget(self.cell("✖" if r["xl"] else "✓", widths[9], row_idx))
        self.table.add_widget(self.cell("✖" if r["special"] else "✓", widths[10], row_idx))
        for mv,w in zip(r["moves"], widths[11:14]):
            self.table.add_widget(self.cell(mv,w,row_idx))
        owned_btn = Button(text="☑" if r["owned"] else "☐", size_hint=(None,None),
                           size=(dp(widths[14]),dp(64)))
        owned_btn.bind(on_release=lambda btn, n=r["name"]: self.toggle_owned(btn,n))
        self.table.add_widget(owned_btn)

    def toggle_owned(self, btn, name):
        self.owned[name] = not self.owned.get(name, False)
        btn.text = "☑" if self.owned[name] else "☐"
        self.save_state()

    def load_row_images(self, rows, gen):
        # Images are loaded after the table is visible so the UI stays responsive.
        for r, child_index in zip(rows, range(len(rows))):
            self.start_background(self.load_pokemon_into_row, r, gen, child_index)

    def load_pokemon_into_row(self, r, gen, row_index):
        # Grid index of the image cell: header row has 15 widgets; each row 15 widgets.
        data = self.download_bytes(image_url(r["name"]))
        if not data:
            return
        Clock.schedule_once(lambda *_: self.apply_image_by_grid_index(data, 15 + row_index*15 + 1), 0)

    def apply_image_by_grid_index(self, data, index):
        if index >= len(self.table.children):
            return
        # GridLayout children are reverse ordered; locate via widget index from bottom.
        try:
            widgets = list(reversed(self.table.children))
            box = widgets[index]
            if isinstance(box, BoxLayout) and box.children:
                img = box.children[0]
                img.texture = CoreImage(BytesIO(data), ext="png").texture
        except Exception:
            pass

    def load_image_widget(self, widget, url):
        if not url:
            return
        data = self.download_bytes(url)
        if not data:
            return
        try:
            tex = CoreImage(BytesIO(data), ext="png").texture
            Clock.schedule_once(lambda *_: setattr(widget, "texture", tex), 0)
        except Exception:
            pass

    def download_bytes(self, url):
        try:
            r = requests.get(url, headers={"User-Agent":"SinKa-PvP-100/Android"}, timeout=10)
            r.raise_for_status()
            return r.content
        except Exception:
            return None


if __name__ == "__main__":
    SinKaApp().run()
