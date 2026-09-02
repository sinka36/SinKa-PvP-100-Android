import json
import math
import os
import threading
import urllib.request
import urllib.parse
from pathlib import Path

from kivy.app import App
from kivy.clock import Clock
from kivy.core.image import Image as CoreImage
from kivy.io import Loader
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.properties import BooleanProperty, NumericProperty, StringProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.image import AsyncImage, Image
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView

# API Bağlantıları
RANKING_URL = "https://raw.githubusercontent.com/pvpoke/pvpoke/master/src/data/rankings/all/overall/rankings-{cp}.json"
POKEMON_URL = "https://raw.githubusercontent.com/pvpoke/pvpoke/master/src/data/gamemaster/pokemon.json"
TR_MOVES_URL = "https://raw.githubusercontent.com/WatWowMap/pogo-data-api/main/data/v1/translations/tr/moves.json"
POGO_MOVES_URL = "https://raw.githubusercontent.com/WatWowMap/pogo-data-api/main/data/v1/moves.json"

LEAGUES = {
    "Great League": 1500,
    "Ultra League": 2500,
    "Master League": 10000,
}

CPM_VALUES = [
    0.094, 0.135137432, 0.16639787, 0.192650919, 0.21573247, 0.236572661,
    0.25572005, 0.273530381, 0.29024988, 0.306057377, 0.3210876, 0.335445036,
    0.34921268, 0.362457751, 0.37523559, 0.387592406, 0.39956728, 0.411193551,
    0.42250001, 0.432926419, 0.44310755, 0.453059958, 0.46279839, 0.472336083,
    0.48168495, 0.4908558, 0.49985844, 0.508701765, 0.51739395, 0.525942511,
    0.53435433, 0.542635767, 0.55079269, 0.558830576, 0.56675452, 0.574569153,
    0.58227891, 0.589887917, 0.59740001, 0.604818814, 0.61215729, 0.619399365,
    0.62656713, 0.633644533, 0.64065295, 0.647576426, 0.65443563, 0.661214806,
    0.667934, 0.674577537, 0.68116492, 0.687680648, 0.69414365, 0.700538673,
    0.70688421, 0.713164996, 0.71939909, 0.725571552, 0.7317, 0.734741009,
    0.73776948, 0.740785574, 0.74378943, 0.746781211, 0.74976104, 0.752729087,
    0.75568551, 0.758630378, 0.76156384, 0.764486065, 0.76739717, 0.770297266,
    0.7731865, 0.776064962, 0.77893275, 0.781790055, 0.78463697, 0.787473578,
    0.79030001, 0.79280395, 0.79530001, 0.79780391, 0.80030001, 0.80280389,
    0.80530001, 0.80780390, 0.81030001, 0.81280390, 0.81530001, 0.81780390,
    0.82030001, 0.82280390, 0.82530001, 0.82780390, 0.83030001, 0.83280390,
    0.83530003, 0.83780375, 0.84030003, 0.84280373, 0.84530002
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

def download_json(url):
    req = urllib.request.Request(url, headers={"User-Agent": "SinKa-PvP-100-Android/2.0"})
    with urllib.request.urlopen(req, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))

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


class MainApp(App):
    show_original_rank = BooleanProperty(False)
    show_shadows = BooleanProperty(False)
    exclude_legendary = BooleanProperty(False)
    exclude_mythical = BooleanProperty(False)
    no_xl = BooleanProperty(False)
    no_special = BooleanProperty(False)
    
    current_league = StringProperty("Great League")
    count_requested = NumericProperty(36)
    status_text = StringProperty("Hazırlanıyor...")

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.ranking_data = []
        self.pokemon = {}
        self.owned = {}
        self.generation = 0
        self.pref_file = Path(self.user_data_dir) / "preferences.json"

    def build(self):
        self.load_preferences()
        self.root_ui = Builder.load_string('''
BoxLayout:
    orientation: 'vertical'
    
    # Üst Panel
    BoxLayout:
        size_hint_y: None
        height: '48dp'
        padding: '5dp'
        spacing: '5dp'
        
        Spinner:
            id: league_spinner
            text: app.current_league
            values: ["Great League", "Ultra League", "Master League"]
            on_text: app.on_league_change(self.text)
            
        TextInput:
            id: count_input
            text: str(app.count_requested)
            input_filter: 'int'
            multiline: False
            size_hint_x: 0.3
            on_text_validate: app.on_count_change(self.text)

        Button:
            text: "Filtreler"
            size_hint_x: 0.4
            on_release: app.open_filters_popup()

        Button:
            text: "↻"
            size_hint_x: 0.2
            on_release: app.reload_data()

    # Durum Bแถı
    Label:
        text: app.status_text
        size_hint_y: None
        height: '24dp'
        font_size: '12sp'

    # Ana Tablo
    ScrollView:
        do_scroll_x: True
        do_scroll_y: True
        
        GridLayout:
            id: table_grid
            cols: 16 if app.show_original_rank else 15
            size_hint: None, None
            height: self.minimum_height
            width: self.minimum_width
            row_default_height: '52dp'
''')
        threading.Thread(target=self.load_initial, daemon=True).start()
        return self.root_ui

    def load_preferences(self):
        try:
            if self.pref_file.exists():
                with open(self.pref_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.owned = data.get("owned", {})
                self.current_league = data.get("league", "Great League")
                self.count_requested = data.get("count", 36)
                self.show_shadows = data.get("shadow", False)
                self.exclude_legendary = data.get("legendary", False)
                self.exclude_mythical = data.get("mythical", False)
                self.no_xl = data.get("xl", False)
                self.no_special = data.get("special", False)
                self.show_original_rank = data.get("original_rank", False)
        except Exception:
            pass

    def save_preferences(self):
        try:
            data = {
                "owned": self.owned,
                "league": self.current_league,
                "count": self.count_requested,
                "shadow": self.show_shadows,
                "legendary": self.exclude_legendary,
                "mythical": self.exclude_mythical,
                "xl": self.no_xl,
                "special": self.no_special,
                "original_rank": self.show_original_rank
            }
            with open(self.pref_file, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
        except Exception:
            pass

    def load_initial(self):
        try:
            self.status_text = "Veriler indiriliyor..."
            pokemon_list = download_json(POKEMON_URL)
            self.pokemon = {p.get("speciesId"): p for p in pokemon_list if p.get("speciesId")}
            
            cp = LEAGUES[self.current_league]
            self.ranking_data = download_json(RANKING_URL.format(cp=cp))
            Clock.schedule_once(lambda dt: self.refresh_ui())
        except Exception as e:
            self.status_text = f"Hata: {str(e)}"

    def refresh_ui(self):
        grid = self.root_ui.ids.table_grid
        grid.clear_widgets()
        
        # Sütun Başlıkları
        headers = ["Rank"]
        if self.show_original_rank:
            headers.append("O-R-J-R")
        headers.extend(["Görsel", "Pokémon", "Tip", "Evrim", "Score", "IV", "Lvl", "CP", "XL", "Özel", "G1", "G2", "G3", "Mevcut"])
        
        widths = [50]
        if self.show_original_rank:
            widths.append(60)
        widths.extend([50, 120, 100, 120, 60, 90, 50, 60, 40, 40, 100, 100, 100, 50])
        
        grid.width = sum(widths)
        
        for idx, h in enumerate(headers):
            grid.add_widget(Label(text=f"[b]{h}[/b]", markup=True, size_hint_x=None, width=dp(widths[idx])))

        self.generation += 1
        gen = self.generation
        
        threading.Thread(target=self.process_rows, args=(gen,), daemon=True).start()

    def process_rows(self, gen):
        cp_limit = LEAGUES[self.current_league]
        candidates = [
            p for p in self.ranking_data
            if (self.show_shadows or not is_shadow(p))
        ]
        
        rows = []
        count = 0
        
        for idx, p in enumerate(candidates):
            if gen != self.generation:
                return
            
            sid = p.get("speciesId", "")
            name = p.get("speciesName", sid)
            poke = self.pokemon.get(sid) or self.pokemon.get(sid[:-7] if sid.endswith("_shadow") else "")
            
            if not poke:
                continue
                
            orig_rank = p.get("rank", idx + 1)
            ivs, level, cp = rank1_iv(poke.get("baseStats", {"atk":100, "def":100, "hp":100}), cp_limit)
            
            xl = level > 40.0
            if self.no_xl and xl:
                continue
                
            count += 1
            row_data = {
                "rank": count,
                "orig_rank": orig_rank,
                "name": name,
                "score": f"{float(p.get('score', 0)):.1f}",
                "iv": f"{ivs[0]}/{ivs[1]}/{ivs[2]}",
                "level": f"{level:g}",
                "cp": cp,
                "xl": "✖" if xl else "✓",
                "special": "✓",
                "owned": "☑" if self.owned.get(name, False) else "☐"
            }
            rows.append(row_data)
            
            if count >= self.count_requested:
                break

        Clock.schedule_once(lambda dt: self.populate_grid(rows, gen))

    def populate_grid(self, rows, gen):
        if gen != self.generation:
            return
            
        grid = self.root_ui.ids.table_grid
        for r in rows:
            grid.add_widget(Label(text=str(r["rank"]), size_hint_x=None, width=dp(50)))
            if self.show_original_rank:
                grid.add_widget(Label(text=str(r["orig_rank"]), size_hint_x=None, width=dp(60)))
                
            # Önizleme Görseli
            slug = r["name"].lower().replace(" ", "-").replace("♀", "-f").replace("♂", "-m")
            img_url = f"https://img.pokemondb.net/sprites/home/normal/{slug}.png"
            grid.add_widget(AsyncImage(source=img_url, size_hint_x=None, width=dp(50)))
            
            grid.add_widget(Label(text=r["name"], size_hint_x=None, width=dp(120)))
            grid.add_widget(Label(text="-", size_hint_x=None, width=dp(100))) # Tip
            grid.add_widget(Label(text="-", size_hint_x=None, width=dp(120))) # Evrim
            grid.add_widget(Label(text=r["score"], size_hint_x=None, width=dp(60)))
            grid.add_widget(Label(text=r["iv"], size_hint_x=None, width=dp(90)))
            grid.add_widget(Label(text=r["level"], size_hint_x=None, width=dp(50)))
            grid.add_widget(Label(text=str(r["cp"]), size_hint_x=None, width=dp(60)))
            grid.add_widget(Label(text=r["xl"], size_hint_x=None, width=dp(40)))
            grid.add_widget(Label(text=r["special"], size_hint_x=None, width=dp(40)))
            grid.add_widget(Label(text="-", size_hint_x=None, width=dp(100)))
            grid.add_widget(Label(text="-", size_hint_x=None, width=dp(100)))
            grid.add_widget(Label(text="-", size_hint_x=None, width=dp(100)))
            
            # Sahip olunma durumu butonu/etiketi
            btn = Label(text=r["owned"], size_hint_x=None, width=dp(50))
            grid.add_widget(btn)

        self.status_text = f"Tamamlandı: {len(rows)} Pokémon"

    def open_filters_popup(self):
        # Popup ile filtre anahtarları (Original Rank dahil) eklendi
        box = BoxLayout(orientation='vertical', padding='10dp', spacing='10dp')
        
        # Orijinal Rank Kutusu
        btn_orig = Button(text=f"Orijinal Rank Göster: {'AÇIK' if self.show_original_rank else 'KAPALI'}")
        def toggle_orig(instance):
            self.show_original_rank = not self.show_original_rank
            instance.text = f"Orijinal Rank Göster: {'AÇIK' if self.show_original_rank else 'KAPALI'}"
            self.save_preferences()
        btn_orig.bind(on_release=toggle_orig)
        box.add_widget(btn_orig)

        # Shadow Kutusu
        btn_shadow = Button(text=f"Shadow Göster: {'AÇIK' if self.show_shadows else 'KAPALI'}")
        def toggle_shadow(instance):
            self.show_shadows = not self.show_shadows
            instance.text = f"Shadow Göster: {'AÇIK' if self.show_shadows else 'KAPALI'}"
            self.save_preferences()
        btn_shadow.bind(on_release=toggle_shadow)
        box.add_widget(btn_shadow)

        popup = Popup(title='Filtreler', content=box, size_hint=(0.8, 0.6))
        popup.bind(on_dismiss=lambda x: self.refresh_ui())
        popup.open()

    def on_league_change(self, text):
        self.current_league = text
        self.save_preferences()
        threading.Thread(target=self.load_initial, daemon=True).start()

    def on_count_change(self, text):
        if text.isdigit():
            self.count_requested = int(text)
            self.save_preferences()
            self.refresh_ui()

    def reload_data(self):
        threading.Thread(target=self.load_initial, daemon=True).start()

if __name__ == '__main__':
    MainApp().run()