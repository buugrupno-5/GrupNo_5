"""
============================================================
  Proje 8: Tasarruf Planı ve Borç Azaltma Modeli
  Gerçek Takvim, Act/365, Kesin Hassasiyet (Decimal)
  GÜNCEL: UI Sabitlemesi, BDDK Limitleri ve
  TAM İZOLE ALT SEKMELER + Auto-Snapping & Boş=0 Mantığı
============================================================
"""

import tkinter as tk
from tkinter import ttk
import math
import datetime
from decimal import Decimal, ROUND_HALF_UP, getcontext
import calendar

getcontext().prec = 28


def stopaj_hesapla(vade_gun):
    """Vade günü sayısına göre yasal stopaj oranını döndürür."""
    if vade_gun < 32:
        return Decimal('17.5')
    elif vade_gun <= 192:
        return Decimal('17.5')
    elif vade_gun <= 384:
        return Decimal('15.0')
    else:
        return Decimal('10.0')


# ─────────────────────────── Renk Paletleri ───────────────────────────
TEMA_KARANLIK = {
    "DARK_BG": "#0D1117", "PANEL_BG": "#161B22", "BORDER": "#30363D",
    "ACCENT_BLUE": "#58A6FF", "ACCENT_DARK_BLUE": "#1F6FEB", "ACCENT_GREEN": "#3FB950", "ACCENT_RED": "#F85149",
    "ACCENT_GOLD": "#D29922", "ACCENT_ORANGE": "#E8820C", "TEXT_PRIMARY": "#E6EDF3",
    "TEXT_MUTED": "#8B949E", "CHART_BG": "#0D1117", "ENTRY_BG": "#1C2128",
    "ACCENT_TAHAKKUK": "#58A6FF", "ROW_EVEN": "#161B22", "ROW_ODD": "#0D1117"
}

TEMA_AYDINLIK = {
    "DARK_BG": "#F0F4F8", "PANEL_BG": "#FFFFFF", "BORDER": "#CBD5E1",
    "ACCENT_BLUE": "#1A73E8", "ACCENT_DARK_BLUE": "#0D47A1", "ACCENT_GREEN": "#035F46", "ACCENT_RED": "#B91C1C",
    "ACCENT_GOLD": "#B45309", "ACCENT_ORANGE": "#EA580C", "TEXT_PRIMARY": "#0F172A",
    "TEXT_MUTED": "#475569", "CHART_BG": "#F0F4F8", "ENTRY_BG": "#F1F5F9",
    "ACCENT_TAHAKKUK": "#083D7F", "ROW_EVEN": "#FFFFFF", "ROW_ODD": "#F8FAFC"
}


def para_format(deger):
    return Decimal(str(deger)).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)


# ══════════════════════════════════════════════════════════════════════
#  TOOLTIP SINIFI
# ══════════════════════════════════════════════════════════════════════
class Tooltip:
    def __init__(self, widget, text, r):
        self.widget = widget
        self.text = text
        self.r = r
        self.tip_window = None
        self.id = None
        self.widget.bind("<Enter>", self.enter, add="+")
        self.widget.bind("<Leave>", self.leave, add="+")

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hide_tip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(300, self.show_tip)

    def unschedule(self):
        id_ = self.id
        self.id = None
        if id_: self.widget.after_cancel(id_)

    def show_tip(self, event=None):
        if self.tip_window or not self.text: return
        x, y, _, _ = self.widget.bbox("insert") if hasattr(self.widget, "bbox") else (0, 0, 0, 0)
        x += self.widget.winfo_rootx() + 25;
        y += self.widget.winfo_rooty() + 20
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        underline_words = getattr(self, "underline_words", [])
        if underline_words:
            font_normal = ("Courier New", 10)
            font_underline = ("Courier New", 10, "underline")
            txt = tk.Text(tw, background=self.r["ENTRY_BG"], foreground=self.r["TEXT_PRIMARY"], relief="solid",
                          borderwidth=1, font=font_normal, wrap="none", state="normal", cursor="arrow", padx=6, pady=4)
            txt.tag_configure("underline", font=font_underline)
            remaining = self.text
            while remaining:
                earliest_idx = len(remaining);
                earliest_word = None
                for w in underline_words:
                    idx = remaining.find(w)
                    if idx != -1 and idx < earliest_idx: earliest_idx = idx; earliest_word = w
                if earliest_word is None: txt.insert("end", remaining); break
                txt.insert("end", remaining[:earliest_idx]);
                txt.insert("end", earliest_word, "underline")
                remaining = remaining[earliest_idx + len(earliest_word):]
            txt.config(state="disabled")
            lines = self.text.split("\n");
            max_chars = max(len(l) for l in lines)
            txt.config(width=max_chars, height=len(lines))
            txt.pack()
        else:
            label = tk.Label(tw, text=self.text, justify="left", background=self.r["ENTRY_BG"],
                             foreground=self.r["TEXT_PRIMARY"], relief="solid", borderwidth=1, font=("Courier New", 10))
            label.pack(ipadx=6, ipady=4)

    def hide_tip(self):
        tw = self.tip_window
        self.tip_window = None
        if tw: tw.destroy()


# ══════════════════════════════════════════════════════════════════════
#  MATEMATİK KATMANI
# ══════════════════════════════════════════════════════════════════════
def hesapla_mevduat(baslangic, yillik_faiz, aylik_yatirim, donem_sayisi, stopaj_orani, vade_gun, baslangic_tarihi,
                    erken_cekim_var=False, tek_seferlik=False, ozel_islemler=None, enflasyon=0.0):
    if ozel_islemler is None: ozel_islemler = {}
    A0 = para_format(baslangic);
    F = Decimal(str(yillik_faiz)) / Decimal('100')
    oto_stopaj = stopaj_hesapla(vade_gun);
    SC = Decimal('1') - (oto_stopaj / Decimal('100'))
    duzenli_yatirim = para_format(aylik_yatirim);
    gun = Decimal(str(vade_gun));
    enf_orani = float(enflasyon) / 100.0
    r_donem = F * gun / Decimal('365');
    p_min = para_format(A0 * r_donem * SC)

    bakiyeler_nominal = [float(A0)];
    bakiyeler_reel = [float(A0)];
    ekstre = []
    bakiye = A0;
    aktif_tarih = baslangic_tarihi;
    vade_bozuldu = False;
    vade_bozulan_donemler = set()
    yanan_faiz_toplami = Decimal('0.00');
    toplam_brut_faiz = Decimal('0.00')
    toplam_stopaj_kesinti = Decimal('0.00');
    toplam_net_faiz = Decimal('0.00')

    for n in range(donem_sayisi):
        donem_no = n + 1;
        aktif_tarih += datetime.timedelta(days=int(vade_gun))
        brut_faiz = para_format(bakiye * r_donem);
        net_faiz = para_format(brut_faiz * SC)
        stopaj_kesinti = para_format(brut_faiz - net_faiz)
        ozel_tutar = ozel_islemler.get(donem_no, Decimal('0.00'));
        uygulanacak_islem = Decimal('0.00')
        if not tek_seferlik or (tek_seferlik and donem_no == 1): uygulanacak_islem = duzenli_yatirim
        guncel_islem = uygulanacak_islem + ozel_tutar;
        donem_vade_bozuldu = False

        if erken_cekim_var and guncel_islem < Decimal('0'):
            vade_bozuldu = True;
            donem_vade_bozuldu = True;
            vade_bozulan_donemler.add(donem_no)
            yanan_faiz_toplami += net_faiz;
            net_faiz = Decimal('0.00');
            stopaj_kesinti = Decimal('0.00');
            brut_faiz = Decimal('0.00')

        toplam_brut_faiz += brut_faiz;
        toplam_stopaj_kesinti += stopaj_kesinti;
        toplam_net_faiz += net_faiz
        bakiye += net_faiz;
        bakiye += guncel_islem
        if bakiye < Decimal('0'): bakiye = Decimal('0')

        bakiyeler_nominal.append(float(para_format(bakiye)))
        gecen_gun = donem_no * float(vade_gun)
        reel_bakiye = bakiye * Decimal(str(math.pow(1.0 + enf_orani, -gecen_gun / 365.0))) if enf_orani > 0 else bakiye
        bakiyeler_reel.append(float(para_format(reel_bakiye)))
        ekstre.append(
            {"donem": donem_no, "tarih": aktif_tarih.strftime("%d.%m.%Y"), "faiz": net_faiz, "islem": guncel_islem,
             "bakiye": bakiye, "reel_bakiye": para_format(reel_bakiye), "vade_bozuldu": donem_vade_bozuldu})

    karakteristik = "YAKINSAK (Sıfıra Eriyip Bitti)" if bakiye <= Decimal('0') else (
        "IRAKSAK (Sürekli Büyür)" if (Decimal('0.00') if tek_seferlik else duzenli_yatirim) >= Decimal('0') else
        "YAKINSAK (Zamanla Eriyip Biter)" if abs(Decimal('0.00') if tek_seferlik else duzenli_yatirim) > para_format(
            bakiye * r_donem * SC) else "IRAKSAK (Faiz Çekimi Karşılıyor, Para Bitmez)")

    return {"nominal": bakiyeler_nominal, "reel": bakiyeler_reel, "p_min": p_min, "bitis_tarihi": aktif_tarih,
            "vade_bozuldu": vade_bozuldu, "vade_bozulan_donemler": vade_bozulan_donemler,
            "yanan_faiz": yanan_faiz_toplami, "ekstre": ekstre, "karakteristik": karakteristik,
            "toplam_brut_faiz": para_format(toplam_brut_faiz),
            "toplam_stopaj_kesinti": para_format(toplam_stopaj_kesinti),
            "toplam_net_faiz": para_format(toplam_net_faiz), "oto_stopaj": oto_stopaj}


def hesapla_borc(baslangic, yillik_faiz, aylik_odeme, vergi_orani, baslangic_tarihi, ara_odemeler=None,
                 hedef_vade=None):
    if ara_odemeler is None: ara_odemeler = {}
    B0 = para_format(baslangic);
    F = Decimal(str(yillik_faiz)) / Decimal('100');
    vergi_orani_d = Decimal(str(vergi_orani)) / Decimal('100')
    r_aylik = (F / Decimal('12')) * (Decimal('1') + vergi_orani_d);
    p_min = para_format(B0 * r_aylik)

    P = para_format(aylik_odeme);
    bakiyeler_nom = [float(B0)];
    ekstre = [];
    borc = B0
    aktif_tarih = baslangic_tarihi;
    toplam_odenen_nom = Decimal('0.00');
    iraksar = False
    dongu_limiti = hedef_vade if hedef_vade is not None else 600

    if P <= p_min and not ara_odemeler and hedef_vade is None: iraksar = True; dongu_limiti = 24

    n = 0
    while n < dongu_limiti and borc > Decimal('0'):
        eski_tarih = aktif_tarih;
        yeni_ay = eski_tarih.month % 12 + 1;
        yeni_yil = eski_tarih.year + (eski_tarih.month // 12)
        try:
            aktif_tarih = datetime.date(yeni_yil, yeni_ay, eski_tarih.day)
        except ValueError:
            aktif_tarih = datetime.date(yeni_yil, yeni_ay, calendar.monthrange(yeni_yil, yeni_ay)[1])

        donem_faizi = para_format(borc * r_aylik);
        ekstra = ara_odemeler.get(n + 1, Decimal('0.00'))
        bu_ay_odenen_plan = P + ekstra;
        faizli_borc = borc + donem_faizi

        if (hedef_vade is not None and (n + 1) == hedef_vade) or faizli_borc <= bu_ay_odenen_plan:
            odenen_anapara = borc;
            toplam_odenen_nom += faizli_borc;
            gercek_odeme = faizli_borc;
            borc = Decimal('0.00')
        else:
            odenen_anapara = bu_ay_odenen_plan - donem_faizi;
            toplam_odenen_nom += bu_ay_odenen_plan
            gercek_odeme = bu_ay_odenen_plan;
            borc = faizli_borc - bu_ay_odenen_plan

        bakiyeler_nom.append(float(borc))
        ekstre.append(
            {"taksit": n + 1, "tarih": aktif_tarih.strftime("%d.%m.%Y"), "kalan_borc": borc, "faiz": donem_faizi,
             "anapara": odenen_anapara, "ara_odeme": ekstra, "tutar": gercek_odeme})
        n += 1

    vade_kurtarmiyor = False
    if hedef_vade is not None and not iraksar:
        if borc > Decimal('0.05'): vade_kurtarmiyor = True

    karakteristik = "IRAKSAK (Borç Asla Kapanmaz, Büyür)" if iraksar else "YAKINSAK (Limit Sınırında Borç Sıfırlanır)"
    return {"nominal": bakiyeler_nom, "iraksar": iraksar, "kapanma_ay": n, "p_min": p_min,
            "toplam_odenen_nom": toplam_odenen_nom, "ekstre": ekstre, "bitis_tarihi": aktif_tarih,
            "karakteristik": karakteristik, "vade_kurtarmiyor": vade_kurtarmiyor, "kalan_borc": borc}


# ══════════════════════════════════════════════════════════════════════
#  TAKVİM POPUP & GRAFİK KATMANI
# ══════════════════════════════════════════════════════════════════════
class TakvimPopup(tk.Toplevel):
    def __init__(self, parent, tarih_var, renkler, entry_widget):
        super().__init__(parent);
        self.tarih_var = tarih_var;
        self.renkler = renkler;
        self.overrideredirect(True);
        self.configure(bg=renkler["BORDER"])
        try:
            mevcut = datetime.datetime.strptime(tarih_var.get(), "%d.%m.%Y").date()
        except:
            mevcut = datetime.date.today()
        self.goruntulenen_yil, self.goruntulenen_ay, self.secili_tarih = mevcut.year, mevcut.month, mevcut
        self._aralik_guncelle();
        self.update_idletasks()
        self.geometry(f"+{entry_widget.winfo_rootx()}+{entry_widget.winfo_rooty() + entry_widget.winfo_height() + 2}")
        self.bind("<FocusOut>", lambda e: self.after(100, self.destroy));
        self.focus_set()

    def _aralik_guncelle(self):
        for w in self.winfo_children(): w.destroy()
        r = self.renkler;
        f = tk.Frame(self, bg=r["PANEL_BG"], padx=4, pady=4);
        f.pack(fill="both", expand=True, padx=1, pady=1)
        baslik = tk.Frame(f, bg=r["PANEL_BG"]);
        baslik.pack(fill="x", pady=(0, 4))
        tk.Button(baslik, text="◀", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], relief="flat",
                  command=self._onceki_ay).pack(side="left")
        ay_isim = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım",
                   "Aralık"]
        tk.Label(baslik, text=f"{ay_isim[self.goruntulenen_ay - 1]} {self.goruntulenen_yil}", bg=r["PANEL_BG"],
                 fg=r["TEXT_PRIMARY"], font=("Courier New", 10, "bold")).pack(side="left", expand=True)
        tk.Button(baslik, text="▶", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], relief="flat",
                  command=self._sonraki_ay).pack(side="right")
        takvim_frame = tk.Frame(f, bg=r["PANEL_BG"]);
        takvim_frame.pack()
        for i, g in enumerate(["Pt", "Sa", "Ça", "Pe", "Cu", "Ct", "Pa"]): tk.Label(takvim_frame, text=g,
                                                                                    bg=r["PANEL_BG"],
                                                                                    fg=r["ACCENT_RED"] if i == 6 else r[
                                                                                        "TEXT_MUTED"],
                                                                                    font=("Courier New", 8, "bold"),
                                                                                    width=3).grid(row=0, column=i)
        ilk_gun, ay_gunu = calendar.monthrange(self.goruntulenen_yil, self.goruntulenen_ay)
        bugun, satir, sutun = datetime.date.today(), 1, ilk_gun
        for gun in range(1, ay_gunu + 1):
            tarih = datetime.date(self.goruntulenen_yil, self.goruntulenen_ay, gun)
            bg = r["ACCENT_BLUE"] if tarih == self.secili_tarih else (
                r["ACCENT_GOLD"] if tarih == bugun else r["PANEL_BG"])
            fg = r["DARK_BG"] if tarih in (self.secili_tarih, bugun) else (
                r["ACCENT_RED"] if sutun == 6 else r["TEXT_PRIMARY"])
            tk.Button(takvim_frame, text=str(gun), width=3, bg=bg, fg=fg, relief="flat", font=("Courier New", 9),
                      command=lambda t=tarih: self._sec(t)).grid(row=satir, column=sutun, padx=1, pady=1)
            sutun += 1
            if sutun == 7: sutun, satir = 0, satir + 1

    def _onceki_ay(self):
        if self.goruntulenen_ay == 1:
            self.goruntulenen_ay, self.goruntulenen_yil = 12, self.goruntulenen_yil - 1
        else:
            self.goruntulenen_ay -= 1
        self._aralik_guncelle()

    def _sonraki_ay(self):
        if self.goruntulenen_ay == 12:
            self.goruntulenen_ay, self.goruntulenen_yil = 1, self.goruntulenen_yil + 1
        else:
            self.goruntulenen_ay += 1
        self._aralik_guncelle()

    def _sec(self, tarih):
        self.tarih_var.set(tarih.strftime("%d.%m.%Y")); self.destroy()


def ciz_grafik(canvas: tk.Canvas, veriler_nom: list, veriler_reel: list, renk_nom: str, baslik: str, r: dict,
               birim: str = "₺", tree=None, vade_bozulan_donemler=None, log_olcek=False):
    if vade_bozulan_donemler is None: vade_bozulan_donemler = set()
    canvas.delete("all");
    W = canvas.winfo_width() if canvas.winfo_width() > 50 else 750;
    H = canvas.winfo_height() if canvas.winfo_height() > 50 else 300
    pad_left, pad_right, pad_top, pad_bottom = 85, 30, 50, 60;
    iw, ih = W - pad_left - pad_right, H - pad_top - pad_bottom
    canvas.configure(bg=r["CHART_BG"])
    canvas.create_rectangle(pad_left, pad_top, pad_left + iw, pad_top + ih, fill=r["PANEL_BG"], outline=r["BORDER"])

    tum_veriler = veriler_nom + (veriler_reel if veriler_reel else [])
    v_min, v_max = min(0, min(tum_veriler)), max(tum_veriler)
    if v_max == v_min: v_max = v_min + 1

    kullan_log = log_olcek and v_max > 0
    if kullan_log: v_min_log = 0; v_max_log = math.log10(v_max) if v_max > 1 else 1

    for i in range(7):
        y = pad_top + ih - int(i / 6 * ih)
        if kullan_log:
            val_log = v_min_log + i / 6 * (v_max_log - v_min_log);
            val = 10 ** val_log;
            eksen_metin = f"{val:,.0f} {birim} (L)"
        else:
            val = v_min + i / 6 * (v_max - v_min);
            eksen_metin = f"{val:,.0f} {birim}"
        canvas.create_line(pad_left, y, pad_left + iw, y, fill=r["BORDER"], dash=(4, 4))
        canvas.create_text(pad_left - 8, y, anchor="e", text=eksen_metin, fill=r["TEXT_MUTED"],
                           font=("Courier New", 10))

    n = len(veriler_nom);
    adim = max(1, n // 10)
    for i in range(0, n, adim):
        x = pad_left + int(i / (n - 1) * iw) if n > 1 else pad_left
        canvas.create_line(x, pad_top, x, pad_top + ih, fill=r["BORDER"], dash=(4, 4))
        canvas.create_text(x, pad_top + ih + 15, anchor="n", text=str(i), fill=r["TEXT_MUTED"],
                           font=("Courier New", 10))

    alt_metin = "Süre (Dönem)" if not kullan_log else "Süre (Dönem) - Logaritmik Ölçek Aktif"
    canvas.create_text(W // 2, H - 20, text=alt_metin, fill=r["TEXT_MUTED"], font=("Courier New", 11))
    canvas.create_text(W // 2, 20, text=baslik, fill=r["TEXT_PRIMARY"], font=("Courier New", 13, "bold"))

    if n < 2: return

    def veri2pix(idx, val):
        px = pad_left + int(idx / (n - 1) * iw)
        if kullan_log:
            val_l = math.log10(val) if val > 1 else 0; py = pad_top + ih - int(
                (val_l - v_min_log) / (v_max_log - v_min_log) * ih)
        else:
            py = pad_top + ih - int((val - v_min) / (v_max - v_min) * ih)
        return px, py

    def cizgi_olustur(veriler, renk, kalinlik, is_dashed=False):
        if not veriler: return
        noktalar = []
        for i, v in enumerate(veriler):
            px, py = veri2pix(i, v);
            noktalar.extend([px, py])
        if len(noktalar) >= 4: canvas.create_line(noktalar, fill=renk, width=kalinlik, smooth=False,
                                                  dash=(4, 4) if is_dashed else None)
        son_px, son_py = veri2pix(n - 1, veriler[-1])
        canvas.create_oval(son_px - 5, son_py - 5, son_px + 5, son_py + 5, fill=renk, outline=r["PANEL_BG"], width=2)

    if veriler_reel: cizgi_olustur(veriler_reel, r["ACCENT_ORANGE"], 2, True)
    cizgi_olustur(veriler_nom, renk_nom, 3)

    canvas.noktalar = []
    for i, v in enumerate(veriler_nom):
        px, py = veri2pix(i, v);
        canvas.noktalar.append((px, py, v, i))

    tooltip_bg = canvas.create_rectangle(0, 0, 0, 0, fill=r["PANEL_BG"], outline=r["ACCENT_BLUE"], width=2,
                                         state="hidden", tags="tooltip_bg")
    tooltip_text = canvas.create_text(0, 0, text="", fill=r["TEXT_PRIMARY"], font=("Courier New", 11, "bold"),
                                      state="hidden", tags="tooltip_text")
    vurgulu_nokta = [None]

    def on_hover(event):
        x, y = event.x, event.y;
        canvas.delete("crosshair")
        if x < pad_left or x > pad_left + iw: on_leave(event); return
        idx = int(round((x - pad_left) / iw * (n - 1))) if n > 1 else 0
        idx = max(0, min(idx, len(veriler_nom) - 1))
        val_nom = veriler_nom[idx];
        px_nom, py_nom = veri2pix(idx, val_nom)
        canvas.create_line(px_nom, pad_top + ih, px_nom, pad_top, dash=(2, 2), fill=r["TEXT_MUTED"], tags="crosshair")
        canvas.create_oval(px_nom - 5, py_nom - 5, px_nom + 5, py_nom + 5, fill=renk_nom, outline=r["PANEL_BG"],
                           width=2, tags="crosshair")
        metin = f"Dönem: {idx}\nNominal: ₺{val_nom:,.2f}"
        if veriler_reel:
            val_reel = veriler_reel[idx];
            _, py_reel = veri2pix(idx, val_reel)
            metin += f"\nReel: ₺{val_reel:,.2f}"
            canvas.create_oval(px_nom - 5, py_reel - 5, px_nom + 5, py_reel + 5, fill=r["ACCENT_ORANGE"],
                               outline=r["PANEL_BG"], width=2, tags="crosshair")
        canvas.itemconfig(tooltip_text, text=metin, state="normal")
        canvas.coords(tooltip_text, px_nom + (-70 if px_nom > W - 150 else 30), py_nom - 20)
        bbox = canvas.bbox(tooltip_text)
        if bbox: canvas.coords(tooltip_bg, bbox[0] - 6, bbox[1] - 6, bbox[2] + 6, bbox[3] + 6); canvas.itemconfig(
            tooltip_bg, state="normal")
        canvas.tag_raise("crosshair");
        canvas.tag_raise("tooltip_bg");
        canvas.tag_raise("tooltip_text")
        if tree is not None and idx > 0:
            cocuklar = tree.get_children()
            if cocuklar and idx <= len(cocuklar):
                hedef = cocuklar[idx - 1]
                if vurgulu_nokta[0] != hedef: vurgulu_nokta[0] = hedef; tree.selection_set(hedef); tree.see(hedef)

    def on_leave(event):
        canvas.delete("crosshair");
        canvas.itemconfig(tooltip_text, state="hidden");
        canvas.itemconfig(tooltip_bg, state="hidden")
        if tree is not None: tree.selection_remove(tree.selection())
        vurgulu_nokta[0] = None

    canvas.bind("<Motion>", on_hover);
    canvas.bind("<Leave>", on_leave)


# ══════════════════════════════════════════════════════════════════════
#  GUI SINIFI (Otomatik Max & Graceful 0 Handling)
# ══════════════════════════════════════════════════════════════════════
class UygulamaGUI:
    def __init__(self, kok: tk.Tk):
        self.kok = kok
        kok.title("Tasarruf Planı ve Borç Azaltma Modeli: Finansal Projeksiyon")
        kok.geometry("1280x850")
        kok.minsize(1000, 700)

        self.kredi_durumlari = {
            "İhtiyaç Kredisi": {"teminat": "", "baslangic": "", "faiz": "", "vergi": "30", "odeme": "",
                                "vade_var": False, "vade": "", "ara_odemeler": [], "hesaplandi": False},
            "Taşıt Kredisi": {"teminat": "", "baslangic": "", "faiz": "", "vergi": "30", "odeme": "", "vade_var": False,
                              "vade": "", "ara_odemeler": [], "hesaplandi": False},
            "Konut Kredisi": {"teminat": "", "baslangic": "", "faiz": "", "vergi": "0", "odeme": "", "vade_var": False,
                              "vade": "", "ara_odemeler": [], "hesaplandi": False},
            "Özel Kredi": {"teminat": "", "baslangic": "", "faiz": "", "vergi": "", "odeme": "", "vade_var": False,
                           "vade": "", "ara_odemeler": [], "hesaplandi": False}
        }

        bugun = datetime.date.today().strftime("%d.%m.%Y")
        self.m_tarih = tk.StringVar(value=bugun)
        self.m_baslangic = tk.StringVar(value="")
        self.m_faiz = tk.StringVar(value="")
        self.m_stopaj = tk.StringVar(value="")
        self.m_enflasyon = tk.StringVar(value="")
        self.m_enflasyon_aktif = tk.BooleanVar(value=False)
        self.m_vade_gun = tk.StringVar(value="")
        self.m_erken_cekim_var = tk.BooleanVar(value=False)
        self.m_duzenli_islem_var = tk.BooleanVar(value=False)
        self.m_yatirim = tk.StringVar(value="")
        self.m_sure = tk.StringVar(value="")
        self.m_erken_cekim_liste = []
        self.m_ozel_islemler_liste = []

        self.b_tarih = tk.StringVar(value=bugun)
        self.b_kredi_tipi = tk.StringVar(value="İhtiyaç Kredisi")
        self.b_teminat = tk.StringVar(value="")
        self.b_baslangic = tk.StringVar(value="")
        self.b_faiz = tk.StringVar(value="")
        self.b_vergi = tk.StringVar(value="30")
        self.b_odeme = tk.StringVar(value="")
        self.b_vade_var = tk.BooleanVar(value=False)
        self.b_vade = tk.StringVar(value="")
        self.b_ara_odemeler_liste = []

        self.karanlik_mod = False
        self.renkler = TEMA_AYDINLIK
        self.m_hesaplandi = self.b_hesaplandi = False
        self._init_tamamlandi = False

        self.ana_cerceve = tk.Frame(self.kok);
        self.ana_cerceve.pack(fill="both", expand=True)
        self._tema_uygula()

        self.b_teminat.trace_add("write", self._bddk_guncelle)
        self.b_baslangic.trace_add("write", self._bddk_guncelle)
        self.b_kredi_tipi.trace_add("write", self._bddk_guncelle)

        self.b_baslangic.trace_add("write", self._dinamik_min_odeme)
        self.b_faiz.trace_add("write", self._dinamik_min_odeme)
        self.b_vergi.trace_add("write", self._dinamik_min_odeme)
        self.b_vade.trace_add("write", self._dinamik_min_odeme)
        self._dinamik_min_odeme()

        self.m_vade_gun.trace_add("write", self._oto_stopaj_guncelle)

        self.kok.bind('<Return>', self._enter_basildi)
        self.kok.bind('<Configure>', self._pencere_boyutu_degisti)
        self.kok.bind_all('<Key>', self._virgul_nokta_cevir, add='+')
        self._init_tamamlandi = True

    def _bddk_sinirlari_hesapla(self):
        tip = self.b_kredi_tipi.get()
        teminat_str = self.b_teminat.get().replace('.', '').replace(',', '.')
        borc_str = self.b_baslangic.get().replace('.', '').replace(',', '.')
        teminat = float(teminat_str) if teminat_str not in ["", ".", "-"] else 0.0
        borc = float(borc_str) if borc_str not in ["", ".", "-"] else 0.0
        maks_kredi = float('inf');
        maks_vade = 1200;
        mesaj = "";
        renk = self.renkler["ACCENT_BLUE"]

        if tip == "İhtiyaç Kredisi":
            if borc > 0:
                if borc <= 125000:
                    maks_vade = 36
                elif borc <= 250000:
                    maks_vade = 24
                else:
                    maks_vade = 12
                mesaj = f"BDDK İhtiyaç Limiti: ₺{borc:,.2f} için Maksimum Vade {maks_vade} Ay"
            else:
                mesaj = "BDDK İhtiyaç: 0-125k (36 Ay), 125k-250k (24 Ay), >250k (12 Ay)"

        elif tip == "Taşıt Kredisi":
            if teminat > 0:
                if teminat <= 2500000:
                    orani, maks_vade = 0.70, 48
                elif teminat <= 5000000:
                    orani, maks_vade = 0.50, 36
                elif teminat <= 6500000:
                    orani, maks_vade = 0.30, 24
                elif teminat <= 7500000:
                    orani, maks_vade = 0.20, 12
                else:
                    orani, maks_vade = 0.0, 0
                maks_kredi = teminat * orani
                if maks_vade > 0:
                    mesaj = f"BDDK Taşıt (LTV %{int(orani * 100)}): Çekilebilir Maks. Kredi ₺{maks_kredi:,.2f} | Maks. Vade {maks_vade} Ay"
                else:
                    mesaj = "BDDK Taşıt: 7.5M TL üzeri araçlar için kredi kullanımı yasal olarak engellenmiştir!"; renk = \
                    self.renkler["ACCENT_RED"]
            else:
                mesaj = "Lütfen önce araç fatura değerini giriniz."

        elif tip == "Konut Kredisi":
            if teminat > 0:
                if teminat <= 5000000:
                    orani = 0.90
                elif teminat <= 10000000:
                    orani = 0.80
                elif teminat <= 20000000:
                    orani = 0.70
                else:
                    orani = 0.50
                maks_vade = 120;
                maks_kredi = teminat * orani
                mesaj = f"BDDK Konut (LTV %{int(orani * 100)}): Çekilebilir Maks. Kredi ₺{maks_kredi:,.2f} | Maks. Vade {maks_vade} Ay"
            else:
                mesaj = "Lütfen önce konut ekspertiz değerini giriniz."
        else:
            mesaj = "Özel Kredi: BDDK limit kısıtlaması uygulanmaz."

        return maks_kredi, maks_vade, mesaj, renk

    def _bddk_guncelle(self, *args):
        try:
            maks_kredi, maks_vade, mesaj, renk = self._bddk_sinirlari_hesapla()
            self.lbl_bddk_bilgi.config(text=mesaj, fg=renk)
        except Exception:
            pass

    def _oto_stopaj_guncelle(self, *args):
        try:
            vade = int(self.m_vade_gun.get())
            oran = stopaj_hesapla(vade);
            s = f"{float(oran):g}"
            self.m_stopaj.set(s)
        except (ValueError, Exception):
            self.m_stopaj.set("")

    def _dinamik_yazdir(self, lbl, metin, base_size, is_bold=False):
        satir_sayisi = metin.count('\n') + sum(len(satir) // 45 for satir in metin.split('\n'))
        if satir_sayisi > 8:
            size = base_size - 2
        elif satir_sayisi > 5:
            size = base_size - 1
        else:
            size = base_size
        weight = "bold" if is_bold else "normal"
        lbl.config(text=metin, font=("Courier New", max(8, size), weight))

    def _ikon_olustur(self, parent, tooltip_text, underline_words=None):
        icon_lbl = tk.Label(parent, text="ⓘ", fg=self.renkler["ACCENT_BLUE"], bg=self.renkler["PANEL_BG"],
                            font=("Segoe UI Symbol", 10, "bold"), cursor="question_arrow")
        t = Tooltip(icon_lbl, tooltip_text, self.renkler)
        if underline_words: t.underline_words = underline_words
        icon_lbl.tooltip = t;
        return icon_lbl

    def _virgul_nokta_cevir(self, event):
        if event.char == ',' and isinstance(event.widget, tk.Entry) and not getattr(event.widget,
                                                                                    'is_currency_formatted', False):
            try:
                event.widget.insert(event.widget.index(tk.INSERT), '.'); return 'break'
            except:
                pass

    def _para_formatla_event(self, event):
        if event.keysym in ('Left', 'Right', 'Up', 'Down', 'Home', 'End', 'Shift_L', 'Shift_R'): return
        w = event.widget;
        raw_val = w.get()
        if not raw_val: return
        c_before = len(raw_val[:w.index(tk.INSERT)].replace('.', ''))
        parts = raw_val.replace('.', '').split(',');
        int_part = ''.join(c for c in parts[0] if c.isdigit())
        f_int = f"{int(int_part):,}".replace(',', '.') if int_part else ("0" if len(parts) > 1 else "")
        new_val = f"{f_int},{''.join(c for c in parts[1] if c.isdigit())}" if len(parts) > 1 else (
            f"{f_int}," if ',' in raw_val else f_int)
        if raw_val != new_val:
            w.delete(0, tk.END);
            w.insert(0, new_val)
            non_dot = 0
            for i, c in enumerate(new_val):
                if non_dot == c_before: w.icursor(i); break
                if c != '.': non_dot += 1
            else:
                w.icursor(len(new_val))

    def _pencere_boyutu_degisti(self, event):
        if event.widget == self.kok: self.kok.after(50, self._grafikleri_yenile)

    def _sekme_degisti(self, event):
        self.kok.after(50, self._grafikleri_yenile)

    def _grafikleri_yenile(self):
        if self.m_hesaplandi:
            self._mevduat_hesapla(sessiz=True)
        else:
            self._bosh_grafik_ciz(self.m_canvas, self.renkler, "Tasarruf Eğrisi")
        if self.b_hesaplandi:
            self._borc_hesapla(sessiz=True)
        else:
            self._bosh_grafik_ciz(self.b_canvas, self.renkler, "Borç Eğrisi")

    def _ilk_cizimleri_yap(self):
        self.kok.update_idletasks()
        if not self.m_hesaplandi: self._bosh_grafik_ciz(self.m_canvas, self.renkler, "Tasarruf Eğrisi")
        if not self.b_hesaplandi: self._bosh_grafik_ciz(self.b_canvas, self.renkler, "Borç Eğrisi")

    def _bosh_grafik_ciz(self, canvas, r, baslik):
        canvas.delete("all")
        W = canvas.winfo_width() if canvas.winfo_width() > 50 else self.kok.winfo_width() // 2
        H = canvas.winfo_height() if canvas.winfo_height() > 50 else 300
        pad_left, pad_right, pad_top, pad_bottom = 85, 30, 50, 60
        iw, ih = W - pad_left - pad_right, H - pad_top - pad_bottom
        canvas.configure(bg=r["CHART_BG"])
        canvas.create_rectangle(pad_left, pad_top, pad_left + iw, pad_top + ih, fill=r["PANEL_BG"], outline=r["BORDER"])
        canvas.create_text(pad_left + (iw // 2), pad_top + (ih // 2),
                           text="Grafiği görmek için 'Hesapla' butonuna basınız.", fill=r["TEXT_MUTED"],
                           font=("Courier New", 12, "italic"), justify="center")
        canvas.create_text(W // 2, 20, text=baslik, fill=r["TEXT_PRIMARY"], font=("Courier New", 13, "bold"))

    def _genel_dogrulama(self, P, tip):
        if P == "": return True
        if tip == "yuzde": return (P == ".") or (
                    P.count('.') <= 1 and '-' not in P and all(c in "0123456789." for c in P) and (
                        '.' not in P or len(P.split('.')[1]) <= 2) and len(P.split('.')[0]) <= 3)
        if tip == "ondalik": return P in ["-", ".", "-."] or (
                    P.count('.') <= 1 and P.count('-') <= 1 and (P.find('-') in [-1, 0]) and all(
                c in "0123456789.-" for c in P) and ('.' not in P or len(P.split('.')[1]) <= 2))
        if tip == "tamsayi": return all(c in "0123456789" for c in P)
        if tip == "tarih": return len(P) <= 10 and all(c in "0123456789." for c in P)
        if tip == "para_gorsel":
            c = P.replace('.', '')
            if c.count(',') <= 1 and all(ch in "0123456789," for ch in c):
                parts = c.split(',')
                # Virgülden önceki kısım (tam sayı) maksimum 10 basamak
                if len(parts[0]) <= 10 and (len(parts) == 1 or len(parts[1]) <= 2):
                    return True
            return False
        return True

    def _enter_basildi(self, event):
        if self.nb.index(self.nb.select()) == 0:
            self._mevduat_hesapla(sessiz=False)
        else:
            self._borc_hesapla(sessiz=False)

    def _dinamik_min_odeme(self, *args):
        try:
            b_bas_str = self.b_baslangic.get().replace('.', '').replace(',', '.')
            B0 = Decimal(b_bas_str) if b_bas_str not in ["", ".", "-"] else Decimal('0')

            faiz_str = self.b_faiz.get();
            vergi_str = self.b_vergi.get()
            F = Decimal(faiz_str) / Decimal('100') if faiz_str not in ["", ".", "-"] else Decimal('0')
            vergi_orani_d = Decimal(vergi_str) / Decimal('100') if vergi_str not in ["", ".", "-"] else Decimal('0')
            r_aylik = (F / Decimal('12')) * (Decimal('1') + vergi_orani_d)

            tahakkuk = (B0 * r_aylik).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
            maks_kredi, maks_vade, mesaj, renk = self._bddk_sinirlari_hesapla()
            tip = self.b_kredi_tipi.get()

            if tip != "Özel Kredi" and maks_vade > 0 and maks_vade < 1200:
                if r_aylik > 0:
                    yasal_min_odeme = (B0 * (r_aylik / (1 - (1 + r_aylik) ** (-maks_vade)))).quantize(Decimal('.01'),
                                                                                                      rounding=ROUND_HALF_UP)
                else:
                    yasal_min_odeme = (B0 / Decimal(maks_vade)).quantize(Decimal('.01'),
                                                                         rounding=ROUND_HALF_UP) if maks_vade > 0 else Decimal(
                        '0')
                self.lbl_dinamik_min.config(
                    text=f"(Yasal sınır, {maks_vade} ayı karşılamak için gereken minimum ödeme tutarı: ₺{yasal_min_odeme:,.2f})")
            elif maks_vade == 0:
                self.lbl_dinamik_min.config(text="(Kredi kullanımı yasal olarak engellenmiştir.)")
            else:
                self.lbl_dinamik_min.config(text=f"Tahakkuk Eden Faiz (Min. Ödeme): ₺{tahakkuk:,.2f}")
        except:
            self.lbl_dinamik_min.config(text="Bekleniyor...")

        if not hasattr(self, 'lbl_vade_min_odeme'): return
        try:
            if not self.b_vade_var.get() or int(self.b_vade.get()) < 1: raise ValueError
            B0 = Decimal(self.b_baslangic.get().replace('.', '').replace(',', '.')) if self.b_baslangic.get() not in [
                "", ".", "-"] else Decimal('0')
            F = Decimal(self.b_faiz.get()) / Decimal('100') if self.b_faiz.get() not in ["", ".", "-"] else Decimal('0')
            vergi_orani_d = Decimal(self.b_vergi.get()) / Decimal('100') if self.b_vergi.get() not in ["", ".",
                                                                                                       "-"] else Decimal(
                '0')
            r_aylik = (F / Decimal('12')) * (Decimal('1') + vergi_orani_d)

            if r_aylik > 0:
                min_odeme_vade = (B0 * (r_aylik / (1 - (1 + r_aylik) ** (-int(self.b_vade.get()))))).quantize(
                    Decimal('.01'), rounding=ROUND_HALF_UP)
            else:
                min_odeme_vade = (B0 / Decimal(self.b_vade.get())).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)

            self.lbl_vade_min_odeme.config(
                text=f"Vade Hedefi ({int(self.b_vade.get())} ay) Taksit Tutarı: ₺{min_odeme_vade:,.2f}")
            if self.b_vade_var.get(): self.b_odeme.set(
                f"{min_odeme_vade:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        except:
            self.lbl_vade_min_odeme.config(text="")

    def _tema_degistir(self):
        if hasattr(self, 'kredi_durumlari') and getattr(self, '_init_tamamlandi', False):
            eski_tip = self.b_kredi_tipi.get()
            self.kredi_durumlari[eski_tip] = {
                "teminat": self.b_teminat.get(), "baslangic": self.b_baslangic.get(), "faiz": self.b_faiz.get(),
                "vergi": self.b_vergi.get(), "odeme": self.b_odeme.get(), "vade_var": self.b_vade_var.get(),
                "vade": self.b_vade.get(), "ara_odemeler": [(d.get(), t.get()) for d, t in self.b_ara_odemeler_liste],
                "hesaplandi": self.b_hesaplandi
            }

        m_erken_kayit = [(d.get(), t.get()) for d, t in self.m_erken_cekim_liste]
        m_ozel_kayit = [(d.get(), t.get()) for d, t in self.m_ozel_islemler_liste]

        self._init_tamamlandi = False
        aktif_sekme = self.nb.index(self.nb.select());
        self.karanlik_mod = not self.karanlik_mod
        self.renkler = TEMA_KARANLIK if self.karanlik_mod else TEMA_AYDINLIK

        for widget in self.ana_cerceve.winfo_children(): widget.destroy()

        self.m_erken_cekim_liste.clear();
        self.m_ozel_islemler_liste.clear();
        self.b_ara_odemeler_liste.clear()
        self._tema_uygula();
        self.nb.select(aktif_sekme)

        for d, t in m_erken_kayit:
            if self.m_erken_cekim_var.get(): self._erken_cekim_ekle(d, t)
        for d, t in m_ozel_kayit: self._ozel_islem_ekle(d, t)

        self.kok.update_idletasks();
        self._init_tamamlandi = True

        self.kok.after(100, lambda: self._mevduat_hesapla(sessiz=True) if self.m_hesaplandi else self._bosh_grafik_ciz(
            self.m_canvas, self.renkler, "Tasarruf Eğrisi"))
        self.kok.after(100, lambda: self._borc_hesapla(sessiz=True) if self.b_hesaplandi else self._bosh_grafik_ciz(
            self.b_canvas, self.renkler, "Borç Eğrisi"))

    def _tema_uygula(self):
        r = self.renkler
        self.kok.configure(bg=r["DARK_BG"]);
        self.ana_cerceve.configure(bg=r["DARK_BG"])
        stil = ttk.Style();
        stil.theme_use("clam")
        stil.configure("TNotebook", background=r["DARK_BG"], borderwidth=0)
        stil.configure("TNotebook.Tab", background=r["PANEL_BG"], foreground=r["TEXT_MUTED"], padding=[20, 8],
                       font=("Courier New", 12, "bold"), borderwidth=0)
        stil.map("TNotebook.Tab", background=[("selected", r["PANEL_BG"])], foreground=[("selected", r["ACCENT_BLUE"])])
        stil.configure("TFrame", background=r["DARK_BG"])
        stil.configure("Treeview", background=r["ENTRY_BG"], foreground=r["TEXT_PRIMARY"],
                       fieldbackground=r["ENTRY_BG"], rowheight=28, font=("Courier New", 10))
        stil.configure("Treeview.Heading", background=r["PANEL_BG"], foreground=r["ACCENT_BLUE"],
                       font=("Courier New", 11, "bold"))
        stil.map('Treeview', background=[('selected', r["ACCENT_BLUE"])], foreground=[('selected', r["DARK_BG"])])
        self._arayuz_olustur()

    def _arayuz_olustur(self):
        r = self.renkler
        baslik_frame = tk.Frame(self.ana_cerceve, bg=r["DARK_BG"]);
        baslik_frame.pack(fill="x", padx=24, pady=(12, 0))
        tk.Label(baslik_frame, text="Tasarruf Planı ve Borç Azaltma Modeli: Finansal Projeksiyon", bg=r["DARK_BG"],
                 fg=r["ACCENT_BLUE"], font=("Courier New", 15, "bold")).pack(side="left")
        tk.Button(baslik_frame, text="☀️/🌙 Temayı Değiştir", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                  font=("Courier New", 10, "bold"), relief="solid", bd=1, cursor="hand2",
                  command=self._tema_degistir).pack(side="right")

        self.nb = ttk.Notebook(self.ana_cerceve);
        self.nb.pack(fill="both", expand=True, padx=16, pady=8)
        self.nb.bind("<<NotebookTabChanged>>", self._sekme_degisti)

        sekme_mevduat = ttk.Frame(self.nb);
        sekme_borc = ttk.Frame(self.nb)
        self.nb.add(sekme_mevduat, text="  🏦  Tasarruf Planı  ");
        self.nb.add(sekme_borc, text="  📉  Borç Ekstresi  ")

        self._mevduat_sekmesi(sekme_mevduat);
        self._borc_sekmesi(sekme_borc)
        self._dinamik_min_odeme();
        self.kok.after(200, self._ilk_cizimleri_yap)

    def _tarih_entry_olustur(self, parent, tarih_var, satir):
        r = self.renkler
        cerceve = tk.Frame(parent, bg=r["PANEL_BG"]);
        cerceve.grid(row=satir, column=1, sticky="w", pady=2)
        vcmd = (self.kok.register(lambda P: self._genel_dogrulama(P, "tarih")), '%P')
        entry = tk.Entry(cerceve, textvariable=tarih_var, width=12, bg=r["ENTRY_BG"], fg=r["TEXT_PRIMARY"],
                         insertbackground=r["TEXT_PRIMARY"], relief="solid", font=("Courier New", 12), bd=1,
                         validate="key", validatecommand=vcmd)
        entry.pack(side="left")

        def _otomatik_nokta(event):
            if event.keysym in ('BackSpace', 'Delete', 'Left', 'Right', 'Home', 'End'): return
            deger = tarih_var.get()
            if len(deger) == 2 and deger.count('.') == 0:
                tarih_var.set(deger + '.'); entry.icursor(3)
            elif len(deger) == 5 and deger.count('.') == 1:
                tarih_var.set(deger + '.'); entry.icursor(6)

        entry.bind('<KeyRelease>', _otomatik_nokta)
        tk.Button(cerceve, text="📅", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], relief="flat", font=("Courier New", 11),
                  cursor="hand2", command=lambda: self._takvim_ac(tarih_var, entry)).pack(side="left", padx=(2, 0))
        return entry

    def _takvim_ac(self, tarih_var, entry_widget):
        TakvimPopup(self.kok, tarih_var, self.renkler, entry_widget)

    def _erken_cekim_toggle(self):
        if self.m_erken_cekim_var.get():
            self.btn_erken_ekle.pack(side="right")
        else:
            self.btn_erken_ekle.pack_forget()
            for w in self.erken_cekim_icerik.winfo_children(): w.destroy()
            self.m_erken_cekim_liste.clear();
            self.erken_cekim_icerik.configure(height=1)

    def _erken_cekim_ekle(self, d_val="", t_val=""):
        r = self.renkler;
        f = tk.Frame(self.erken_cekim_icerik, bg=r["PANEL_BG"]);
        f.pack(fill="x", pady=1)
        d_var = tk.StringVar(value=str(d_val));
        t_var = tk.StringVar(value=str(t_val))
        vcmd_int = (self.kok.register(lambda P: self._genel_dogrulama(P, "tamsayi")), '%P')
        vcmd_dec = (self.kok.register(lambda P: self._genel_dogrulama(P, "ondalik")), '%P')
        tk.Label(f, text="Dönem:", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], font=("Courier New", 10)).pack(side="left")
        tk.Entry(f, textvariable=d_var, width=5, bg=r["ENTRY_BG"], fg=r["TEXT_PRIMARY"], relief="solid",
                 font=("Courier New", 11), bd=1, validate="key", validatecommand=vcmd_int).pack(side="left",
                                                                                                padx=(2, 6))
        tk.Label(f, text="Çekim(₺):", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], font=("Courier New", 10)).pack(
            side="left")
        tk.Entry(f, textvariable=t_var, width=10, bg=r["ENTRY_BG"], fg=r["TEXT_PRIMARY"], relief="solid",
                 font=("Courier New", 11), bd=1, validate="key", validatecommand=vcmd_dec).pack(side="left",
                                                                                                padx=(2, 6))
        tk.Button(f, text="✕", bg=r["ACCENT_RED"], fg="#FFFFFF", relief="flat", font=("Courier New", 9, "bold"),
                  cursor="hand2", command=lambda frm=f, v=(d_var, t_var): self._erken_cekim_sil(frm, v)).pack(
            side="left")
        self.m_erken_cekim_liste.append((d_var, t_var))

    def _erken_cekim_sil(self, frame, var_tuple):
        if var_tuple in self.m_erken_cekim_liste: self.m_erken_cekim_liste.remove(var_tuple)
        frame.pack_forget();
        frame.destroy()
        self.erken_cekim_icerik.update_idletasks();
        self.erken_cekim_icerik.configure(height=1)

    def _duzenli_islem_toggle(self):
        if self.m_duzenli_islem_var.get():
            self.m_yatirim_entry.pack(side="left", padx=(10, 0))
        else:
            self.m_yatirim_entry.pack_forget(); self.m_yatirim.set("")

    def _enflasyon_toggle(self):
        if self.m_enflasyon_aktif.get():
            self.enf_entry.pack(side="left", padx=(10, 0))
        else:
            self.enf_entry.pack_forget(); self.m_enflasyon.set("")

    def _b_vade_toggle(self):
        if self.b_vade_var.get():
            self.b_vade_entry.pack(side="left", padx=(10, 0))
            self.entry_b_odeme.config(state="disabled", disabledbackground="#D9D9D9", disabledforeground="#3A3A3A")
            self._dinamik_min_odeme()
        else:
            self.b_vade_entry.pack_forget();
            self.b_vade.set("")
            self.entry_b_odeme.config(state="normal")

    def _etiket_giris(self, parent, metin, attr, satir, v_tipi, tooltip=""):
        r = self.renkler
        tk.Label(parent, text=metin, bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], font=("Courier New", 11)).grid(row=satir,
                                                                                                            column=0,
                                                                                                            sticky="w",
                                                                                                            pady=2,
                                                                                                            padx=(0, 8))
        f = tk.Frame(parent, bg=r["PANEL_BG"]);
        f.grid(row=satir, column=1, sticky="w", pady=2)
        vcmd = (self.kok.register(lambda P: self._genel_dogrulama(P, v_tipi)), '%P')
        entry = tk.Entry(f, textvariable=getattr(self, attr), width=15, bg=r["ENTRY_BG"], fg=r["TEXT_PRIMARY"],
                         insertbackground=r["TEXT_PRIMARY"], relief="solid", font=("Courier New", 12), bd=1,
                         validate="key", validatecommand=vcmd)
        entry.pack(side="left")
        if tooltip: self._ikon_olustur(f, tooltip).pack(side="left", padx=(4, 0))
        return entry

    def _mevduat_sekmesi(self, parent):
        r = self.renkler
        icerik = tk.Frame(parent, bg=r["DARK_BG"]);
        icerik.pack(fill="both", expand=True, padx=8, pady=4)
        icerik.grid_columnconfigure(1, weight=1);
        icerik.grid_rowconfigure(0, weight=1)

        sol_panel = tk.Frame(icerik, bg=r["DARK_BG"]);
        sol_panel.grid(row=0, column=0, sticky="ns", padx=(0, 16))
        sag_panel = tk.Frame(icerik, bg=r["DARK_BG"]);
        sag_panel.grid(row=0, column=1, sticky="nsew")

        giris = tk.Frame(sol_panel, bg=r["PANEL_BG"], padx=16, pady=6);
        giris.pack(fill="x", pady=(0, 8))
        giris.grid_columnconfigure(0, minsize=260);
        giris.grid_columnconfigure(1, weight=1)

        tk.Label(giris, text="Başlangıç Tarihi:", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                 font=("Courier New", 11)).grid(row=0, column=0, sticky="w", pady=2, padx=(0, 8))
        self._tarih_entry_olustur(giris, self.m_tarih, 0)

        entry_m_bas = self._etiket_giris(giris, "Başlangıç Bakiye (₺):", "m_baslangic", 1, "para_gorsel",
                                         "Başlangıçtaki anapara tutarını giriniz.\nMaksimum 10 basamaklı bir değer girin.")
        entry_m_bas.is_currency_formatted = True;
        entry_m_bas.bind('<KeyRelease>', self._para_formatla_event, add='+')

        self._etiket_giris(giris, "Brüt Faiz Oranı (%):", "m_faiz", 2, "yuzde",
                           "Yıllık brüt banka faiz oranını giriniz.\n%0.0 - %100.0 arasında bir değer girin.")
        self._etiket_giris(giris, "Vade Günü (Tamsayı):", "m_vade_gun", 3, "tamsayi",
                           "Faizin tahakkuk edeceği gün sayısı (Örn: 32).\n1 - 999 arasında gün sayısını girin.")
        entry_m_stopaj = self._etiket_giris(giris, "Stopaj (Vergi) (%):", "m_stopaj", 4, "yuzde",
                                            "Faizden kesilen yasal vergi.\nVade sürenize göre otomatik olarak hesaplanır.")
        entry_m_stopaj.config(state="disabled", disabledbackground="#D9D9D9", disabledforeground="#3A3A3A")

        erken_cekim_baslik = tk.Frame(giris, bg=r["PANEL_BG"])
        erken_cekim_baslik.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(4, 2))
        tk.Label(erken_cekim_baslik, text="Erken Çekimler:", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                 font=("Courier New", 11)).pack(side="left")

        # YENİ: 3. satır eklendi
        self._ikon_olustur(erken_cekim_baslik,
                           "Dönem bitmeden önce para çekme veya yatırma yapacaksanız belirtiniz.\nDönem bitmeden para çekmek vadeyi bozar.\nPara çekimi işlemleriniz için tutarın başına \"-\" koyunuz. (Örn: -500)").pack(
            side="left", padx=(4, 8))

        tk.Checkbutton(erken_cekim_baslik, text="Aktif (Vade Bozar)", variable=self.m_erken_cekim_var,
                       command=self._erken_cekim_toggle, bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                       selectcolor=r["ENTRY_BG"], activebackground=r["PANEL_BG"], activeforeground=r["TEXT_PRIMARY"],
                       font=("Courier New", 10)).pack(side="left", padx=(8, 0))
        self.btn_erken_ekle = tk.Button(erken_cekim_baslik, text="＋ Ekle", bg=r["ACCENT_BLUE"], fg=r["DARK_BG"],
                                        relief="flat", font=("Courier New", 9, "bold"), cursor="hand2",
                                        command=self._erken_cekim_ekle)

        self.erken_cekim_icerik = tk.Frame(giris, bg=r["PANEL_BG"])
        self.erken_cekim_icerik.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(0, 2))

        self.enf_frame = tk.Frame(giris, bg=r["PANEL_BG"]);
        self.enf_frame.grid(row=7, column=0, columnspan=2, sticky="w", pady=4)
        tk.Checkbutton(self.enf_frame, text="Reel Alım Gücü (Enflasyon):", variable=self.m_enflasyon_aktif,
                       command=self._enflasyon_toggle, bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                       selectcolor=r["ENTRY_BG"], activebackground=r["PANEL_BG"], activeforeground=r["TEXT_PRIMARY"],
                       font=("Courier New", 11)).pack(side="left", padx=(0, 8))
        self.enf_entry = tk.Entry(self.enf_frame, textvariable=self.m_enflasyon, width=8, bg=r["ENTRY_BG"],
                                  fg=r["TEXT_PRIMARY"], insertbackground=r["TEXT_PRIMARY"], relief="solid",
                                  font=("Courier New", 12), bd=1, validate="key",
                                  validatecommand=(self.kok.register(lambda P: self._genel_dogrulama(P, "yuzde")),
                                                   '%P'))
        if self.m_enflasyon_aktif.get(): self.enf_entry.pack(side="left", padx=(10, 0))
        enf_icon = self._ikon_olustur(self.enf_frame,
                                      "Mevcut/beklenen yıllık enflasyon oranını girin.\n0 - 100 arasında bir değer girin.");
        enf_icon.pack(side="left", padx=(4, 0))

        self.duzenli_islem_frame = tk.Frame(giris, bg=r["PANEL_BG"]);
        self.duzenli_islem_frame.grid(row=8, column=0, columnspan=2, sticky="w", pady=4)
        tk.Checkbutton(self.duzenli_islem_frame, text="Düzenli İşlem (Her Dönem):", variable=self.m_duzenli_islem_var,
                       command=self._duzenli_islem_toggle, bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                       selectcolor=r["ENTRY_BG"], activebackground=r["PANEL_BG"], activeforeground=r["TEXT_PRIMARY"],
                       font=("Courier New", 11)).pack(side="left", padx=(0, 8))
        self._ikon_olustur(self.duzenli_islem_frame,
                           "Vade sonunda her dönem yapacağınız para çekme\nveya yatırma işlemlerini giriniz.",
                           underline_words=["her dönem"]).pack(side="left", padx=(0, 6))
        self.m_yatirim_entry = tk.Entry(self.duzenli_islem_frame, textvariable=self.m_yatirim, width=12,
                                        bg=r["ENTRY_BG"], fg=r["TEXT_PRIMARY"], insertbackground=r["TEXT_PRIMARY"],
                                        relief="solid", font=("Courier New", 12), bd=1, validate="key",
                                        validatecommand=(
                                            self.kok.register(lambda P: self._genel_dogrulama(P, "ondalik")), '%P'))
        if self.m_duzenli_islem_var.get(): self.m_yatirim_entry.pack(side="left", padx=(10, 0))

        ozel_baslik_frame = tk.Frame(giris, bg=r["PANEL_BG"]);
        ozel_baslik_frame.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(4, 2))
        tk.Label(ozel_baslik_frame, text="Özel İşlemler:", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                 font=("Courier New", 11)).pack(side="left")
        self._ikon_olustur(ozel_baslik_frame,
                           "Vade sonunda yapacağınız düzensiz para çekme\nve yatırma işlemlerini giriniz.").pack(
            side="left", padx=(4, 0))
        tk.Button(ozel_baslik_frame, text="＋ Ekle", bg=r["ACCENT_BLUE"], fg=r["DARK_BG"], relief="flat",
                  font=("Courier New", 9, "bold"), cursor="hand2", command=self._ozel_islem_ekle).pack(side="right")

        self.ozel_islem_frame = tk.Frame(giris, bg=r["PANEL_BG"]);
        self.ozel_islem_frame.grid(row=10, column=0, columnspan=2, sticky="ew", pady=(0, 2))
        self._etiket_giris(giris, "Tekrar Edecek Dönem (1-120):", "m_sure", 11, "tamsayi",
                           "Simülasyonun toplam uzunluğunu giriniz.\n1 - 120 Dönem arasında bir değer girin.")

        btn_frame_m = tk.Frame(giris, bg=r["PANEL_BG"]);
        btn_frame_m.grid(row=12, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        btn_frame_m.grid_columnconfigure(0, weight=4);
        btn_frame_m.grid_columnconfigure(1, weight=1)
        tk.Button(btn_frame_m, text="  Hesapla  ", bg=r["ACCENT_BLUE"], fg="#FFFFFF", font=("Courier New", 11, "bold"),
                  relief="flat", cursor="hand2", command=self._mevduat_hesapla).grid(row=0, column=0, sticky="ew")
        tk.Button(btn_frame_m, text="Temizle", bg=r["ACCENT_RED"], fg="#FFFFFF", font=("Courier New", 11, "bold"),
                  relief="flat", cursor="hand2", command=self._mevduat_temizle).grid(row=0, column=1, sticky="ew",
                                                                                     padx=(4, 0))

        self.m_sonuc_frame = tk.Frame(sol_panel, bg=r["PANEL_BG"], padx=16, pady=8);
        self.m_sonuc_frame.pack(fill="both", expand=True)
        self.m_sonuc_lbl = tk.Label(self.m_sonuc_frame, text="", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                                    font=("Courier New", 11), justify="left", wraplength=380);
        self.m_sonuc_lbl.pack(anchor="w")
        self.m_uyari_lbl = tk.Label(self.m_sonuc_frame, text="", bg=r["PANEL_BG"], fg=r["ACCENT_ORANGE"],
                                    font=("Courier New", 11, "bold"), justify="left", wraplength=380);
        self.m_uyari_lbl.pack(anchor="w", pady=(2, 0), fill="both", expand=True)

        sag_panel.grid_rowconfigure(0, weight=1);
        sag_panel.grid_rowconfigure(1, weight=2);
        sag_panel.grid_columnconfigure(0, weight=1)
        self.m_canvas = tk.Canvas(sag_panel, bg=r["CHART_BG"], highlightthickness=0);
        self.m_canvas.grid(row=0, column=0, sticky="nsew", pady=(0, 10))

        tablo_frame_m = tk.Frame(sag_panel, bg=r["ENTRY_BG"]);
        tablo_frame_m.grid(row=1, column=0, sticky="nsew")
        tablo_frame_m.grid_rowconfigure(0, weight=1);
        tablo_frame_m.grid_columnconfigure(0, weight=1)

        self.tree_m = ttk.Treeview(tablo_frame_m, columns=("donem", "tarih", "faiz", "islem", "bakiye", "reel_bakiye"),
                                   show="headings")
        for c, n, w in [("donem", "Dönem", 50), ("tarih", "Tarih", 90), ("faiz", "Net Faiz", 110),
                        ("islem", "Nakit Akışı", 110), ("bakiye", "Nominal Bakiye", 130),
                        ("reel_bakiye", "Reel Bakiye", 130)]:
            self.tree_m.heading(c, text=n);
            self.tree_m.column(c, width=w, anchor="center" if c in ("donem", "tarih") else "e")
        self.tree_m.tag_configure("evenrow", background=r["ROW_EVEN"]);
        self.tree_m.tag_configure("oddrow", background=r["ROW_ODD"])
        self.tree_m.tag_configure("vade_bozuldu", foreground=r["ACCENT_ORANGE"])

        scroll_m = ttk.Scrollbar(tablo_frame_m, orient="vertical", command=self.tree_m.yview)
        self.tree_m.configure(yscrollcommand=scroll_m.set);
        self.tree_m.grid(row=0, column=0, sticky="nsew");
        scroll_m.grid(row=0, column=1, sticky="ns")
        self.tree_m.bind("<<TreeviewSelect>>", lambda e: self._tablo_secim_grafik(self.tree_m, self.m_canvas, "m"))

    def _ozel_islem_ekle(self, donem_deger="", tutar_deger=""):
        r = self.renkler;
        satir_frame = tk.Frame(self.ozel_islem_frame, bg=r["PANEL_BG"]);
        satir_frame.pack(fill="x", pady=1)
        donem_var = tk.StringVar(value=str(donem_deger));
        tutar_var = tk.StringVar(value=str(tutar_deger))
        vcmd_int = (self.kok.register(lambda P: self._genel_dogrulama(P, "tamsayi")), '%P')
        vcmd_dec = (self.kok.register(lambda P: self._genel_dogrulama(P, "ondalik")), '%P')
        tk.Label(satir_frame, text="Dönem:", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], font=("Courier New", 10)).pack(
            side="left")
        tk.Entry(satir_frame, textvariable=donem_var, width=5, bg=r["ENTRY_BG"], fg=r["TEXT_PRIMARY"], relief="solid",
                 font=("Courier New", 11), bd=1, validate="key", validatecommand=vcmd_int).pack(side="left",
                                                                                                padx=(2, 6))
        tk.Label(satir_frame, text="Tutar (₺):", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], font=("Courier New", 10)).pack(
            side="left")
        tk.Entry(satir_frame, textvariable=tutar_var, width=10, bg=r["ENTRY_BG"], fg=r["TEXT_PRIMARY"], relief="solid",
                 font=("Courier New", 11), bd=1, validate="key", validatecommand=vcmd_dec).pack(side="left",
                                                                                                padx=(2, 6))
        tk.Button(satir_frame, text="✕", bg=r["ACCENT_RED"], fg="#FFFFFF", relief="flat",
                  font=("Courier New", 9, "bold"), cursor="hand2",
                  command=lambda f=satir_frame, v=(donem_var, tutar_var): self._ozel_islem_sil(f, v)).pack(side="left")
        self.m_ozel_islemler_liste.append((donem_var, tutar_var))

    def _ozel_islem_sil(self, frame, var_tuple):
        if var_tuple in self.m_ozel_islemler_liste: self.m_ozel_islemler_liste.remove(var_tuple)
        frame.pack_forget();
        frame.destroy()
        self.ozel_islem_frame.update_idletasks();
        self.ozel_islem_frame.configure(height=1);
        self.ozel_islem_frame.grid_configure(pady=(0, 0) if not self.ozel_islem_frame.winfo_children() else (0, 4))

    def _borc_sekmesi(self, parent):
        r = self.renkler
        icerik = tk.Frame(parent, bg=r["DARK_BG"]);
        icerik.pack(fill="both", expand=True, padx=8, pady=4)
        icerik.grid_columnconfigure(1, weight=1);
        icerik.grid_rowconfigure(0, weight=1)

        sol_panel = tk.Frame(icerik, bg=r["DARK_BG"]);
        sol_panel.grid(row=0, column=0, sticky="ns", padx=(0, 16))
        sag_panel = tk.Frame(icerik, bg=r["DARK_BG"]);
        sag_panel.grid(row=0, column=1, sticky="nsew")

        giris = tk.Frame(sol_panel, bg=r["PANEL_BG"], padx=16, pady=6);
        giris.pack(fill="x", pady=(0, 8))
        giris.grid_columnconfigure(0, minsize=260);
        giris.grid_columnconfigure(1, weight=1)

        self.kredi_tip_frame = tk.Frame(giris, bg=r["ENTRY_BG"], highlightbackground=r["BORDER"], highlightthickness=1)
        self.kredi_tip_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        self.kredi_sekmeleri = {}

        def _kredi_sekme_sec(tip_id, aktif_renk):
            eski_tip = self.b_kredi_tipi.get()

            if getattr(self, '_init_tamamlandi', False):
                if eski_tip in self.kredi_durumlari:
                    self.kredi_durumlari[eski_tip] = {
                        "teminat": self.b_teminat.get(), "baslangic": self.b_baslangic.get(),
                        "faiz": self.b_faiz.get(), "vergi": self.b_vergi.get(),
                        "odeme": self.b_odeme.get(), "vade_var": self.b_vade_var.get(),
                        "vade": self.b_vade.get(),
                        "ara_odemeler": [(d.get(), t.get()) for d, t in self.b_ara_odemeler_liste],
                        "hesaplandi": self.b_hesaplandi
                    }

            self.b_kredi_tipi.set(tip_id)
            for k, d in self.kredi_sekmeleri.items():
                d["btn"].config(bg=aktif_renk if k == tip_id else r["ENTRY_BG"],
                                fg=r["DARK_BG"] if k == tip_id else r["TEXT_MUTED"])

            # YENİ: Dinamik Tooltip Güncellemesi
            if tip_id == "İhtiyaç Kredisi" or tip_id == "Özel Kredi":
                self.lbl_teminat_widget.grid_remove();
                self.frame_b_teminat.grid_remove()
            else:
                if tip_id == "Taşıt Kredisi":
                    self.lbl_teminat_text.set("Araç Değeri (₺):")
                    self.ikon_teminat.tooltip.text = "BDDK'nın LTV (Kredi/Değer) oranını hesaplayabilmesi için\naraç fatura değerini giriniz.\nMaksimum 10 basamaklı bir değer girin."
                elif tip_id == "Konut Kredisi":
                    self.lbl_teminat_text.set("Konut Değeri (₺):")
                    self.ikon_teminat.tooltip.text = "BDDK'nın LTV (Kredi/Değer) oranını hesaplayabilmesi için\nkonut ekspertiz değerini giriniz.\nMaksimum 10 basamaklı bir değer girin."
                self.lbl_teminat_widget.grid();
                self.frame_b_teminat.grid()

            if tip_id in ["İhtiyaç Kredisi", "Taşıt Kredisi", "Konut Kredisi"]:
                self.entry_b_vergi.config(state="disabled", disabledbackground="#D9D9D9", disabledforeground="#3A3A3A")
            else:
                self.entry_b_vergi.config(state="normal")

            yeni_durum = self.kredi_durumlari[tip_id]
            self.b_teminat.set(yeni_durum["teminat"])
            self.b_baslangic.set(yeni_durum["baslangic"])
            self.b_faiz.set(yeni_durum["faiz"])
            self.b_vergi.set(yeni_durum["vergi"])
            self.b_odeme.set(yeni_durum["odeme"])
            self.b_vade_var.set(yeni_durum["vade_var"]);
            self._b_vade_toggle()
            self.b_vade.set(yeni_durum["vade"])

            for w in self.b_ara_odeme_frame.winfo_children(): w.destroy()
            self.b_ara_odemeler_liste.clear()
            for d_val, t_val in yeni_durum["ara_odemeler"]: self._b_ara_odeme_ekle(d_val, t_val)

            self.b_hesaplandi = yeni_durum["hesaplandi"]

            if self.b_hesaplandi:
                self._borc_hesapla(sessiz=True)
            else:
                self.b_sonuc_lbl.config(text="");
                self.b_uyari_lbl.config(text="")
                self.tree.delete(*self.tree.get_children())
                self._bosh_grafik_ciz(self.b_canvas, self.renkler, "Borç Eğrisi")

            self._bddk_guncelle();
            self._dinamik_min_odeme()

        OZEL_RENK = "#C9A000" if not self.karanlik_mod else "#D4B000"
        for tip_id, text, renk in [("İhtiyaç Kredisi", "💸 İhtiyaç", r["ACCENT_GREEN"]),
                                   ("Taşıt Kredisi", "🚗 Taşıt", r["ACCENT_ORANGE"]),
                                   ("Konut Kredisi", "🏠 Konut", r["ACCENT_BLUE"]),
                                   ("Özel Kredi", "🛠️ Özel", OZEL_RENK)]:
            btn = tk.Button(self.kredi_tip_frame, text=text, font=("Courier New", 11, "bold"), bg=r["ENTRY_BG"],
                            fg=r["TEXT_MUTED"], relief="flat", cursor="hand2", bd=0,
                            command=lambda t=tip_id, c=renk: _kredi_sekme_sec(t, c))
            btn.pack(side="left", fill="x", expand=True, ipady=6)
            self.kredi_sekmeleri[tip_id] = {"btn": btn, "renk": renk}

        tk.Label(giris, text="Başlangıç Tarihi:", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                 font=("Courier New", 11)).grid(row=1, column=0, sticky="w", pady=2, padx=(0, 8))
        self._tarih_entry_olustur(giris, self.b_tarih, 1)

        self.lbl_teminat_text = tk.StringVar(value="Teminat Değeri (₺):")
        self.lbl_teminat_widget = tk.Label(giris, textvariable=self.lbl_teminat_text, bg=r["PANEL_BG"],
                                           fg=r["TEXT_PRIMARY"], font=("Courier New", 11))
        self.lbl_teminat_widget.grid(row=2, column=0, sticky="w", pady=2, padx=(0, 8))

        self.frame_b_teminat = tk.Frame(giris, bg=r["PANEL_BG"]);
        self.frame_b_teminat.grid(row=2, column=1, sticky="w", pady=2)
        vcmd_gorsel = (self.kok.register(lambda P: self._genel_dogrulama(P, "para_gorsel")), '%P')
        self.entry_b_teminat = tk.Entry(self.frame_b_teminat, textvariable=self.b_teminat, width=15, bg=r["ENTRY_BG"],
                                        fg=r["TEXT_PRIMARY"], insertbackground=r["TEXT_PRIMARY"], relief="solid",
                                        font=("Courier New", 12), bd=1, validate="key", validatecommand=vcmd_gorsel)
        self.entry_b_teminat.pack(side="left")
        self.entry_b_teminat.is_currency_formatted = True;
        self.entry_b_teminat.bind('<KeyRelease>', self._para_formatla_event, add='+')
        self.ikon_teminat = self._ikon_olustur(self.frame_b_teminat,
                                               "BDDK'nın LTV (Kredi/Değer) oranını hesaplayabilmesi için\naraç fatura değeri veya konut ekspertiz değerini girin.")
        self.ikon_teminat.pack(side="left", padx=(4, 0))

        entry_b_bas = self._etiket_giris(giris, "Talep Edilen Kredi (₺):", "b_baslangic", 3, "para_gorsel",
                                         "Çekmek istediğiniz borç/kredi tutarını giriniz.\nMaksimum 10 basamaklı bir değer girin.")
        entry_b_bas.is_currency_formatted = True;
        entry_b_bas.bind('<KeyRelease>', self._para_formatla_event, add='+')

        self.lbl_bddk_bilgi = tk.Label(giris, text="", bg=r["PANEL_BG"], fg=r["ACCENT_BLUE"],
                                       font=("Courier New", 10, "bold"), justify="left", wraplength=480)
        self.lbl_bddk_bilgi.grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 6), padx=(0, 8))

        self.entry_b_faiz = self._etiket_giris(giris, "Kredi Faiz Oranı (Yıllık %):", "b_faiz", 5, "yuzde",
                                               "Kredinin yıllık brüt faiz oranını giriniz.\n%0.0 - %100.0 arasında bir değer girin.")
        self.entry_b_vergi = self._etiket_giris(giris, "Vergi (KKDF+BSMV %):", "b_vergi", 6, "yuzde",
                                                "Faize eklenen yasal vergi oranını giriniz.\n%0.0 - %100.0 arasında bir değer girin.")
        self.entry_b_odeme = self._etiket_giris(giris, "Aylık Ödeme (₺):", "b_odeme", 7, "para_gorsel",
                                                "Aylık sabit taksit tutarını giriniz.\nAylık ödeme en fazla toplam borca eşit olabilir.")
        self.entry_b_odeme.is_currency_formatted = True;
        self.entry_b_odeme.bind('<KeyRelease>', self._para_formatla_event, add='+')

        f_ara_btn = tk.Frame(giris, bg=r["PANEL_BG"]);
        f_ara_btn.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(6, 2))
        tk.Label(f_ara_btn, text="Ara Ödemeler:", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                 font=("Courier New", 11)).pack(side="left")
        self._ikon_olustur(f_ara_btn, "Aylık ödemeler dışında yapacağınız\nara ödemeleri varsa belirtiniz.").pack(
            side="left", padx=(4, 0))
        tk.Button(f_ara_btn, text="＋ Ekle", bg=r["ACCENT_BLUE"], fg=r["DARK_BG"], relief="flat",
                  font=("Courier New", 9, "bold"), cursor="hand2", command=self._b_ara_odeme_ekle).pack(side="right")

        self.b_ara_odeme_frame = tk.Frame(giris, bg=r["PANEL_BG"]);
        self.b_ara_odeme_frame.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(0, 4))

        self.b_vade_frame = tk.Frame(giris, bg=r["PANEL_BG"]);
        self.b_vade_frame.grid(row=10, column=0, columnspan=2, sticky="w", pady=4)
        tk.Checkbutton(self.b_vade_frame, text="Vade Hedefi (Ay):", variable=self.b_vade_var,
                       command=self._b_vade_toggle, bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], selectcolor=r["ENTRY_BG"],
                       activebackground=r["PANEL_BG"], activeforeground=r["TEXT_PRIMARY"],
                       font=("Courier New", 11)).pack(side="left", padx=(0, 8))
        vcmd_int = (self.kok.register(lambda P: self._genel_dogrulama(P, "tamsayi")), '%P')
        self.b_vade_entry = tk.Entry(self.b_vade_frame, textvariable=self.b_vade, width=8, bg=r["ENTRY_BG"],
                                     fg=r["TEXT_PRIMARY"], insertbackground=r["TEXT_PRIMARY"], relief="solid",
                                     font=("Courier New", 12), bd=1, validate="key", validatecommand=vcmd_int)
        if self.b_vade_var.get(): self.b_vade_entry.pack(side="left", padx=(10, 0))
        self._ikon_olustur(self.b_vade_frame,
                           "Borcunuzu hesaplanan dönemden farklı bir dönemde bitirmek\nistiyorsanız belirtiniz.\nAylık ödemeniz bu döneme göre hesaplanır.").pack(
            side="left", padx=(6, 0))

        self.lbl_dinamik_min = tk.Label(giris, text="", bg=r["PANEL_BG"], fg=r["ACCENT_TAHAKKUK"],
                                        font=("Courier New", 10, "bold"), justify="left", wraplength=480)
        self.lbl_dinamik_min.grid(row=11, column=0, columnspan=2, sticky="w", pady=(2, 0), padx=(0, 8))
        self.lbl_vade_min_odeme = tk.Label(giris, text="", bg=r["PANEL_BG"], fg=r["ACCENT_BLUE"],
                                           font=("Courier New", 10, "bold"), justify="left", wraplength=480)
        self.lbl_vade_min_odeme.grid(row=12, column=0, columnspan=2, sticky="w", pady=(0, 8), padx=(0, 8))

        btn_frame_b = tk.Frame(giris, bg=r["PANEL_BG"]);
        btn_frame_b.grid(row=13, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        btn_frame_b.grid_columnconfigure(0, weight=4);
        btn_frame_b.grid_columnconfigure(1, weight=1)
        tk.Button(btn_frame_b, text="  Hesapla  ", bg=r["ACCENT_BLUE"], fg="#FFFFFF", font=("Courier New", 11, "bold"),
                  relief="flat", cursor="hand2", command=self._borc_hesapla).grid(row=0, column=0, sticky="ew")
        tk.Button(btn_frame_b, text="Temizle", bg=r["ACCENT_RED"], fg="#FFFFFF", font=("Courier New", 11, "bold"),
                  relief="flat", cursor="hand2", command=self._borc_temizle).grid(row=0, column=1, sticky="ew",
                                                                                  padx=(4, 0))

        self.b_sonuc_frame = tk.Frame(sol_panel, bg=r["PANEL_BG"], padx=16, pady=8);
        self.b_sonuc_frame.pack(fill="both", expand=True)
        self.b_sonuc_lbl = tk.Label(self.b_sonuc_frame, text="", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"],
                                    font=("Courier New", 11), justify="left", wraplength=380);
        self.b_sonuc_lbl.pack(anchor="w")
        self.b_uyari_lbl = tk.Label(self.b_sonuc_frame, text="", bg=r["PANEL_BG"], fg=r["ACCENT_ORANGE"],
                                    font=("Courier New", 11, "bold"), justify="left", wraplength=380);
        self.b_uyari_lbl.pack(anchor="w", pady=(2, 0), fill="both", expand=True)

        sag_panel.grid_rowconfigure(0, weight=1);
        sag_panel.grid_rowconfigure(1, weight=2);
        sag_panel.grid_columnconfigure(0, weight=1)
        self.b_canvas = tk.Canvas(sag_panel, bg=r["CHART_BG"], highlightthickness=0);
        self.b_canvas.grid(row=0, column=0, sticky="nsew", pady=(0, 10))

        tablo_frame = tk.Frame(sag_panel, bg=r["ENTRY_BG"]);
        tablo_frame.grid(row=1, column=0, sticky="nsew")
        tablo_frame.grid_rowconfigure(0, weight=1);
        tablo_frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tablo_frame,
                                 columns=("taksit", "tarih", "kalan", "faiz", "anapara", "ara_odeme", "tutar"),
                                 show="headings")
        for c, n, w in [("taksit", "No", 50), ("tarih", "Tarih", 100), ("kalan", "Kalan Anapara", 120),
                        ("faiz", "Faiz Yükü", 100), ("anapara", "Anapara Öd.", 100), ("ara_odeme", "Ara Ödeme", 100),
                        ("tutar", "Toplam Taksit", 120)]:
            self.tree.heading(c, text=n);
            self.tree.column(c, width=w, anchor="center" if c in ("taksit", "tarih") else "e")
        self.tree.tag_configure("evenrow", background=r["ROW_EVEN"]);
        self.tree.tag_configure("oddrow", background=r["ROW_ODD"])

        scroll = ttk.Scrollbar(tablo_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set);
        self.tree.grid(row=0, column=0, sticky="nsew");
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._tablo_secim_grafik(self.tree, self.b_canvas, "b"))

        mevcut_tip = self.b_kredi_tipi.get()
        renk = self.kredi_sekmeleri.get(mevcut_tip, {}).get("renk", r["ACCENT_GREEN"])
        _kredi_sekme_sec(mevcut_tip, renk)

    def _b_ara_odeme_sil(self, frame, var_tuple):
        if var_tuple in self.b_ara_odemeler_liste: self.b_ara_odemeler_liste.remove(var_tuple)
        frame.pack_forget();
        frame.destroy()
        self.b_ara_odeme_frame.update_idletasks();
        self.b_ara_odeme_frame.configure(height=1)

    def _b_ara_odeme_ekle(self, d_val="", t_val=""):
        r = self.renkler;
        f = tk.Frame(self.b_ara_odeme_frame, bg=r["PANEL_BG"]);
        f.pack(fill="x", pady=1)
        d_var = tk.StringVar(value=str(d_val));
        t_var = tk.StringVar(value=str(t_val))
        tk.Label(f, text="Ay:", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], font=("Courier New", 10)).pack(side="left")
        tk.Entry(f, textvariable=d_var, width=5, bg=r["ENTRY_BG"], fg=r["TEXT_PRIMARY"], relief="solid").pack(
            side="left", padx=(2, 6))
        tk.Label(f, text="Tutar (₺):", bg=r["PANEL_BG"], fg=r["TEXT_PRIMARY"], font=("Courier New", 10)).pack(
            side="left")
        tk.Entry(f, textvariable=t_var, width=10, bg=r["ENTRY_BG"], fg=r["TEXT_PRIMARY"], relief="solid").pack(
            side="left", padx=(2, 6))
        tk.Button(f, text="✕", bg=r["ACCENT_RED"], fg="#FFFFFF", relief="flat", font=("Courier New", 9, "bold"),
                  command=lambda frm=f, v=(d_var, t_var): self._b_ara_odeme_sil(frm, v)).pack(side="left")
        self.b_ara_odemeler_liste.append((d_var, t_var))

    def _mevduat_temizle(self):
        for var in [self.m_baslangic, self.m_faiz, self.m_stopaj, self.m_enflasyon, self.m_vade_gun, self.m_yatirim,
                    self.m_sure]: var.set("")
        self.m_erken_cekim_var.set(False);
        self._erken_cekim_toggle()
        self.m_duzenli_islem_var.set(False);
        self._duzenli_islem_toggle()
        self.m_enflasyon_aktif.set(False);
        self._enflasyon_toggle()
        for w in self.ozel_islem_frame.winfo_children(): w.destroy()
        self.m_ozel_islemler_liste.clear()

        self.m_sonuc_lbl.config(text="");
        self.m_uyari_lbl.config(text="")
        self.tree_m.delete(*self.tree_m.get_children())
        self._bosh_grafik_ciz(self.m_canvas, self.renkler, "Tasarruf Eğrisi")
        self.m_hesaplandi = False

    def _borc_temizle(self):
        tip_id = self.b_kredi_tipi.get()
        for var in [self.b_teminat, self.b_baslangic, self.b_faiz, self.b_odeme, self.b_vade]: var.set("")

        if tip_id == "Özel Kredi":
            self.b_vergi.set("")
        elif tip_id in ["İhtiyaç Kredisi", "Taşıt Kredisi"]:
            self.b_vergi.set("30")
        elif tip_id == "Konut Kredisi":
            self.b_vergi.set("0")

        self.b_vade_var.set(False);
        self._b_vade_toggle()
        for w in self.b_ara_odeme_frame.winfo_children(): w.destroy()
        self.b_ara_odemeler_liste.clear()

        self.b_sonuc_lbl.config(text="");
        self.b_uyari_lbl.config(text="")
        self.tree.delete(*self.tree.get_children())
        self._bosh_grafik_ciz(self.b_canvas, self.renkler, "Borç Eğrisi")
        self.b_hesaplandi = False

        if getattr(self, '_init_tamamlandi', False) and tip_id in self.kredi_durumlari:
            self.kredi_durumlari[tip_id] = {"teminat": "", "baslangic": "", "faiz": "", "vergi": self.b_vergi.get(),
                                            "odeme": "", "vade_var": False, "vade": "", "ara_odemeler": [],
                                            "hesaplandi": False}
        self._bddk_guncelle();
        self._dinamik_min_odeme()

    def _tablo_secim_grafik(self, tree, canvas, mod):
        secim = tree.selection()
        if not secim: return
        cocuklar = tree.get_children()
        try:
            idx = list(cocuklar).index(secim[0])
        except ValueError:
            return
        canvas.delete("tablo_vurgu");
        canvas.itemconfig("tooltip_text", state="hidden");
        canvas.itemconfig("tooltip_bg", state="hidden")
        if not hasattr(canvas, "noktalar") or not canvas.noktalar: return
        hedef_ay = idx + 1;
        hedef_nokta = next((p for p in canvas.noktalar if p[3] == hedef_ay), None)
        if hedef_nokta:
            px, py, val, _ = hedef_nokta
            H = canvas.winfo_height() if canvas.winfo_height() > 50 else 300
            canvas.create_line(px, 50, px, H - 60, fill=self.renkler["ACCENT_GOLD"], width=2, dash=(4, 2),
                               tags="tablo_vurgu")
            canvas.create_oval(px - 6, py - 6, px + 6, py + 6, fill=self.renkler["ACCENT_GOLD"],
                               outline=self.renkler["PANEL_BG"], width=2, tags="tablo_vurgu")
            canvas.itemconfig("tooltip_text", text=f"Dönem: {hedef_ay}\nDeğer: ₺{val:,.2f}", state="normal")
            W = canvas.winfo_width() if canvas.winfo_width() > 50 else 750
            canvas.coords("tooltip_text", px + (-60 if px > W - 150 else 30), py - 20)
            bbox = canvas.bbox("tooltip_text")
            if bbox: canvas.coords("tooltip_bg", bbox[0] - 6, bbox[1] - 6, bbox[2] + 6, bbox[3] + 6); canvas.itemconfig(
                "tooltip_bg", state="normal")
            canvas.tag_raise("tablo_vurgu");
            canvas.tag_raise("tooltip_bg");
            canvas.tag_raise("tooltip_text")

    def _hata_ve_temizle(self, var, baslik, mesaj, tab_lbl=None):
        if tab_lbl:
            self._dinamik_yazdir(tab_lbl, f"❌ {baslik}: {mesaj}", base_size=11, is_bold=True)
            tab_lbl.config(fg=self.renkler["ACCENT_RED"])
        if var: var.set("")

    def _dogrula(self, deger_str, min_val, max_val, ad, lbl_uyari=None, sessiz=False):
        try:
            if deger_str in ["", ".", "-"]: raise ValueError
            val = float(deger_str)
            if not (min_val <= val <= max_val): raise ValueError
            return val
        except ValueError:
            if not sessiz and lbl_uyari: self._hata_ve_temizle(None, "Geçersiz Veri",
                                                               f"'{ad}' alanına {min_val} ile {max_val} arası eksiksiz sayı girin.",
                                                               lbl_uyari)
            return None

    def _dogrula_int(self, deger_str, min_val, max_val, ad, lbl_uyari=None, sessiz=False):
        try:
            if deger_str == "": raise ValueError
            val = int(deger_str)
            if not (min_val <= val <= max_val): raise ValueError
            return val
        except ValueError:
            if not sessiz and lbl_uyari: self._hata_ve_temizle(None, "Geçersiz Veri",
                                                               f"'{ad}' alanına {min_val}-{max_val} arası geçerli TAM SAYI girin.",
                                                               lbl_uyari)
            return None

    def _tarih_dogrula(self, tarih_str, lbl_uyari=None, sessiz=False):
        try:
            return datetime.datetime.strptime(tarih_str, "%d.%m.%Y").date()
        except ValueError:
            if not sessiz and lbl_uyari: self._hata_ve_temizle(None, "Tarih Hatası",
                                                               "Tarihi GG.AA.YYYY formatında eksiksiz girin.",
                                                               lbl_uyari)
            return None

    def _parse_ozel_islemler(self, lbl_uyari, sessiz):
        islemler = {}
        for donem_var, tutar_var in self.m_ozel_islemler_liste:
            d_str = donem_var.get().strip();
            t_str = tutar_var.get().strip()
            if not d_str: continue
            if d_str and not t_str: t_str = "0"
            try:
                d = int(d_str);
                t = float(t_str)
                if d < 1: raise ValueError
                islemler[d] = Decimal(str(t))
            except (ValueError, Exception):
                if not sessiz and lbl_uyari: self._hata_ve_temizle(None, "Özel İşlem Hatası",
                                                                   "İşlem satırlarını doğru doldurun (Dönem: tam sayı).",
                                                                   lbl_uyari)
                return None
        return islemler

    def _mevduat_hesapla(self, sessiz=False):
        r = self.renkler;
        self.m_uyari_lbl.config(text="");
        lbl = self.m_uyari_lbl

        # BOŞ = 0 Mantığı
        faiz_str = self.m_faiz.get()
        if faiz_str in ["", ".", "-"]: faiz_str = "0"
        faiz = self._dogrula(faiz_str, 0, 100, "Brüt Faiz", lbl, sessiz)
        if faiz is None: return False

        vade_str = self.m_vade_gun.get()
        if vade_str == "" or vade_str == "0":
            if not sessiz: self._hata_ve_temizle(self.m_vade_gun, "Eksik Veri", "'Vade Günü' 1'den küçük olamaz.", lbl);
            return False
        vade = self._dogrula_int(vade_str, 1, 999, "Vade Günü", lbl, sessiz)
        if vade is None: return False

        stopaj = float(stopaj_hesapla(vade));
        self.m_stopaj.set(str(stopaj).rstrip('0').rstrip('.') if '.' in str(stopaj) else str(stopaj))

        sure_str = self.m_sure.get()
        if sure_str == "" or sure_str == "0":
            if not sessiz: self._hata_ve_temizle(self.m_sure, "Eksik Veri", "'Tekrar Edecek Dönem' 1'den küçük olamaz.",
                                                 lbl);
            return False
        N = self._dogrula_int(sure_str, 1, 120, "Dönem", lbl, sessiz)
        if N is None: return False

        tarih = self._tarih_dogrula(self.m_tarih.get(), lbl, sessiz)
        if tarih is None: return False

        baslangic_str = self.m_baslangic.get().replace('.', '').replace(',', '.')
        if baslangic_str in ["", ".", "-", "-."]: baslangic_str = "0"
        try:
            A0 = float(baslangic_str)
        except:
            if not sessiz: self._hata_ve_temizle(self.m_baslangic, "Hata", "Geçersiz Başlangıç Bakiye formati.", lbl)
            return False

        if self.m_duzenli_islem_var.get():
            yatirim_str = self.m_yatirim.get().replace('.', '').replace(',', '.')
            if yatirim_str in ["", ".", "-", "-."]: yatirim_str = "0"
            try:
                P = float(yatirim_str)
            except:
                if not sessiz: self._hata_ve_temizle(self.m_yatirim, "Hata", "Geçersiz Düzenli İşlem formati.", lbl)
                return False
        else:
            P = 0.0

        enf_str = self.m_enflasyon.get()
        if enf_str in ["", ".", "-"]: enf_str = "0"
        try:
            enf = float(enf_str) if self.m_enflasyon_aktif.get() else 0.0
        except ValueError:
            if not sessiz: self._hata_ve_temizle(self.m_enflasyon, "Hata", "Enflasyon formatı hatalı.",
                                                 lbl); return False

        ozel_islem_dict = self._parse_ozel_islemler(lbl, sessiz)
        if ozel_islem_dict is None: return False

        erken_cekim_dict = {}
        if self.m_erken_cekim_var.get():
            for d_var, t_var in self.m_erken_cekim_liste:
                d_str = d_var.get().strip();
                t_str = t_var.get().strip()
                if not d_str: continue
                if d_str and not t_str: t_str = "0"
                try:
                    d = int(d_str);
                    t = float(t_str)
                    if d < 1: raise ValueError
                    val = Decimal(str(-abs(t)))
                    if d in erken_cekim_dict:
                        erken_cekim_dict[d] += val
                    else:
                        erken_cekim_dict[d] = val
                except:
                    if not sessiz: self._hata_ve_temizle(None, "Hata", "Erken çekim satırlarını doğru doldurun.", lbl)
                    return False

        birlesik_islemler = dict(ozel_islem_dict)
        for k, v in erken_cekim_dict.items():
            if k in birlesik_islemler:
                birlesik_islemler[k] += v
            else:
                birlesik_islemler[k] = v

        if not sessiz and birlesik_islemler:
            buyuk_donemler = [d for d in birlesik_islemler if d > N]
            if buyuk_donemler:
                self._hata_ve_temizle(None, "Özel İşlem Hatası",
                                      f"İşlem dönemi ({', '.join(map(str, sorted(buyuk_donemler)))}), belirlenen süreyi ({N}) aşıyor!",
                                      lbl)
                return False

        erken_cekim_var = self.m_erken_cekim_var.get();
        tek_seferlik_mi = False
        sonuc = hesapla_mevduat(A0, faiz, P, N, stopaj, vade, tarih, erken_cekim_var=erken_cekim_var,
                                tek_seferlik=tek_seferlik_mi, ozel_islemler=birlesik_islemler, enflasyon=enf)
        nom = sonuc["nominal"];
        reel = sonuc["reel"];
        p_min = sonuc["p_min"];
        bitis = sonuc["bitis_tarihi"].strftime("%d.%m.%Y")
        renk = r["ACCENT_GREEN"] if nom[-1] > 0 else r["ACCENT_RED"]

        t_brut = sonuc['toplam_brut_faiz'];
        t_stop = sonuc['toplam_stopaj_kesinti'];
        t_net = sonuc['toplam_net_faiz'];
        oto_st = sonuc['oto_stopaj']
        net_son = Decimal(str(nom[-1]));
        brut_son_bakiye = net_son + t_stop
        metin = (
            f"Son Bakiye Tarihi : {bitis}\n"
            f"Dizi Durumu       : {sonuc['karakteristik']}\n"
            f"Kritik Sınır P_min: ₺{p_min:,.2f} /Dönem\n\n"
            f"Uygulanan Stopaj  : %{oto_st}\n"
            f"Toplam Brüt Faiz  : ₺{t_brut:,.2f}\n"
            f"Toplam Stopaj Kes.: ₺{t_stop:,.2f}\n"
            f"Toplam Net Faiz   : ₺{t_net:,.2f}\n\n"
            f"Brüt Son Bakiye   : ₺{brut_son_bakiye:,.2f}\n"
            f"Net Son Bakiye    : ₺{net_son:,.2f}\n"
        )
        if enf > 0: metin += f"Reel Alım Gücü    : ₺{reel[-1]:,.2f} (Enflasyon İskontolu)\n"

        if sonuc["vade_bozuldu"]:
            vade_bozulanlar = sorted(list(sonuc["vade_bozulan_donemler"]))
            donemler_str = f"{vade_bozulanlar[0]}, {vade_bozulanlar[1]}, {vade_bozulanlar[2]}... dahil toplam {len(vade_bozulanlar)} farklı" if len(
                vade_bozulanlar) > 4 else ", ".join(map(str, vade_bozulanlar))
            uyari_metin = f"UYARI! {donemler_str}. dönemde vade bozuldu,\n₺{sonuc['yanan_faiz']:,.2f} faiz YANDI! ❌"
            metin += "(Hesaplamalarda yanan faizler düşülmüştür)"
            self._dinamik_yazdir(self.m_uyari_lbl, uyari_metin, base_size=11, is_bold=True)
            self.m_uyari_lbl.config(fg=r["ACCENT_ORANGE"])
        elif "YAKINSAK" in sonuc["karakteristik"]:
            uyari_metin = "⚠️ DİKKAT: Paranız zamanla eriyip tükenecek!\nÇekim miktarınız faiz getirinizi aşıyor."
            self._dinamik_yazdir(self.m_uyari_lbl, uyari_metin, base_size=11, is_bold=True)
            self.m_uyari_lbl.config(fg=r["ACCENT_ORANGE"])
            metin += "(Sıfır Hata Toleransıyla Hesaplanmıştır) ✓"
        else:
            self.m_uyari_lbl.config(text="")
            metin += "(Sıfır Hata Toleransıyla Hesaplanmıştır) ✓"

        self._dinamik_yazdir(self.m_sonuc_lbl, metin, base_size=13)
        self.m_sonuc_lbl.config(fg=renk)

        self.tree_m.delete(*self.tree_m.get_children())
        for i, satir in enumerate(sonuc["ekstre"]):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            tgs = (tag, "vade_bozuldu") if satir.get("vade_bozuldu") else (tag,)
            self.tree_m.insert("", "end", values=(satir["donem"], satir["tarih"], f"₺{satir['faiz']:,.2f}",
                                                  f"₺{satir['islem']:,.2f}", f"₺{satir['bakiye']:,.2f}",
                                                  f"₺{satir['reel_bakiye']:,.2f}"), tags=tgs)

        ciz_grafik(self.m_canvas, nom, reel if enf > 0 else [], renk, "Tasarruf Eğrisi", r, tree=self.tree_m,
                   vade_bozulan_donemler=sonuc.get("vade_bozulan_donemler", set()))
        self.m_hesaplandi = True;
        return True

    def _borc_hesapla(self, sessiz=False):
        r = self.renkler;
        self.b_uyari_lbl.config(text="");
        lbl = self.b_uyari_lbl
        maks_kredi, maks_vade, mesaj, renk = self._bddk_sinirlari_hesapla()

        # BOŞ = 0 Mantığı
        faiz_str = self.b_faiz.get()
        if faiz_str in ["", ".", "-"]: faiz_str = "0"
        faiz = self._dogrula(faiz_str, 0, 100, "Kredi Faiz Oranı", lbl, sessiz)
        if faiz is None: return False

        vergi_str = self.b_vergi.get()
        if vergi_str in ["", ".", "-"]: vergi_str = "0"
        vergi = self._dogrula(vergi_str, 0, 100, "Vergi (KKDF)", lbl, sessiz)
        if vergi is None: return False

        tarih = self._tarih_dogrula(self.b_tarih.get(), lbl, sessiz)
        if tarih is None: return False

        baslangic_str = self.b_baslangic.get().replace('.', '').replace(',', '.')
        if baslangic_str in ["", ".", "-"]: baslangic_str = "0"
        try:
            A0 = float(baslangic_str)

            # YENİ: Otomatik Max Geri Çekme
            if maks_kredi != float('inf') and A0 > maks_kredi:
                A0 = float(maks_kredi)
                # Formata uygun UI Güncellemesi
                A0_dec = Decimal(str(maks_kredi)).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
                A0_int = int(A0_dec)
                formatted_val = f"{A0_int:,}".replace(',', '.')
                decimal_part = str(A0_dec).split('.')[1]
                if decimal_part != "00": formatted_val += f",{decimal_part}"
                self.b_baslangic.set(formatted_val)

        except ValueError:
            if not sessiz: self._hata_ve_temizle(self.b_baslangic, "Hata", "'Toplam Borç' alanını sayısal girin.", lbl)
            return False

        odeme_str = self.b_odeme.get().replace('.', '').replace(',', '.')
        if odeme_str in ["", ".", "-"]: odeme_str = "0"
        try:
            P = float(odeme_str)
        except ValueError:
            if not sessiz: self._hata_ve_temizle(self.b_odeme, "Hata", "'Aylık Ödeme' formatı hatalı.",
                                                 lbl); return False

        if not sessiz and P > A0 and A0 > 0:
            self._hata_ve_temizle(self.b_odeme, "Geçersiz Değer",
                                  f"Aylık ödeme (₺{P:,.2f}), toplam borcu (₺{A0:,.2f}) aşamaz.", lbl)
            return False

        hedef_vade = None
        if self.b_vade_var.get():
            vade_str = self.b_vade.get()
            if vade_str == "" or vade_str == "0":
                if not sessiz: self._hata_ve_temizle(self.b_vade, "Geçersiz Değer", "Vade Hedefi 1'den küçük olamaz.",
                                                     lbl);
                return False
            hedef_vade = self._dogrula_int(vade_str, 1, 120, "Vade Hedefi", lbl, sessiz)
            if hedef_vade is None: return False

            if not sessiz and hedef_vade > maks_vade:
                self._hata_ve_temizle(None, "BDDK Vade Limiti",
                                      f"Girilen vade ({hedef_vade} ay), BDDK yasal sınırını ({maks_vade} ay) aşıyor.",
                                      lbl)
                return False

        ara = {}
        for dv, tv in self.b_ara_odemeler_liste:
            d_val = dv.get().strip()
            t_val = tv.get().replace('.', '').replace(',', '.').strip()
            if d_val:
                if not t_val or t_val in [".", "-"]: t_val = "0"
                ara[int(d_val)] = Decimal(t_val)

        sonuc = hesapla_borc(A0, faiz, P, vergi, tarih, ara, hedef_vade=hedef_vade)
        kapanma_ay = sonuc["kapanma_ay"]

        if not self.b_vade_var.get() and not sonuc["iraksar"] and kapanma_ay > maks_vade:
            if not sessiz: self._hata_ve_temizle(None, "BDDK Vade Limiti",
                                                 f"Mevcut taksitle kredi {kapanma_ay} ayda bitiyor. Yasal sınır {maks_vade} aydır. Taksiti artırın.",
                                                 lbl)
            return False

        if sonuc["iraksar"]:
            renk = r["ACCENT_RED"]
            if abs(float(P) - float(sonuc['p_min'])) < 0.01:
                metin = (
                    f"DİKKAT: Dizi {sonuc['karakteristik']}\nAylık Ödemeniz   : ₺{P:,.2f}\nAylık Faiz (Min) : ₺{float(sonuc['p_min']):,.2f}\nDurum            : Yalnızca faiz ödeniyor,\n                   anapara eksilmiyor.\n\nToplam Ödenen    : ₺{sonuc['toplam_odenen_nom']:,.2f}")
            else:
                metin = (
                    f"DİKKAT: Dizi {sonuc['karakteristik']}\nAylık Ödemeniz   : ₺{P:,.2f}\nAylık Faiz (Min) : ₺{float(sonuc['p_min']):,.2f}\nDurum            : Ödeme faizi karşılamıyor,\n                   borç büyümeye devam eder.\n\nToplam Ödenen    : ₺{sonuc['toplam_odenen_nom']:,.2f}")
            uyari_metin = "⚠️ Aşağıda borcun 2 yılda (24 ay) nasıl büyüdüğünün logaritmik simülasyonu gösterilmiştir."
            self._dinamik_yazdir(self.b_uyari_lbl, uyari_metin, base_size=11, is_bold=True);
            self.b_uyari_lbl.config(fg=r["ACCENT_RED"])
        else:
            renk = r["ACCENT_GREEN"];
            bitis = sonuc["bitis_tarihi"].strftime("%d.%m.%Y")
            metin = (
                f"Borç Bitiş Tarihi : {bitis}\nDizi Durumu       : {sonuc['karakteristik']}\n\nToplam Anapara    : ₺{A0:,.2f}\nToplam Ödenen     : ₺{sonuc['toplam_odenen_nom']:,.2f}\nToplam Faiz Yükü  : ₺{float(sonuc['toplam_odenen_nom']) - A0:,.2f}\n")
            if sonuc.get("vade_kurtarmiyor"):
                uyari_metin = f"DİKKAT: Ödeme borcu hedeflenen sürede kapatmıyor!\n{hedef_vade}. ay sonunda ₺{sonuc['kalan_borc']:,.2f} borç kalıyor."
                self._dinamik_yazdir(self.b_uyari_lbl, uyari_metin, base_size=11, is_bold=True);
                self.b_uyari_lbl.config(fg=r["ACCENT_ORANGE"])
            elif hedef_vade is None:
                uyari_metin = f"✅ Borç {kapanma_ay} ayda ({kapanma_ay // 12} yıl {kapanma_ay % 12} ay) ödenerek kapatıldı."
                self._dinamik_yazdir(self.b_uyari_lbl, uyari_metin, base_size=11, is_bold=True);
                self.b_uyari_lbl.config(fg=r["ACCENT_GREEN"])
            else:
                uyari_metin = f"✅ Borç tam hedeflenen {hedef_vade}. ayda sıfırlanarak kapatıldı."
                self._dinamik_yazdir(self.b_uyari_lbl, uyari_metin, base_size=11, is_bold=True);
                self.b_uyari_lbl.config(fg=r["ACCENT_GREEN"])

        self._dinamik_yazdir(self.b_sonuc_lbl, metin, base_size=13);
        self.b_sonuc_lbl.config(fg=renk)

        self.tree.delete(*self.tree.get_children())
        for i, satir in enumerate(sonuc["ekstre"]):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=(satir["taksit"], satir["tarih"], f"₺{satir['kalan_borc']:,.2f}",
                                                f"₺{satir['faiz']:,.2f}", f"₺{satir['anapara']:,.2f}",
                                                f"₺{satir.get('ara_odeme', Decimal('0.00')):,.2f}",
                                                f"₺{satir['tutar']:,.2f}"), tags=(tag,))

        grafik_baslik = "Borç Eğrisi (Iraksak Durum için Logaritmik)" if sonuc["iraksar"] else "Borç Eğrisi"
        ciz_grafik(self.b_canvas, sonuc.get("nominal", [A0]), [], renk, grafik_baslik, r, tree=self.tree,
                   log_olcek=sonuc["iraksar"])
        self.b_hesaplandi = True;
        return True


if __name__ == "__main__":
    kok = tk.Tk()
    uygulama = UygulamaGUI(kok)
    kok.mainloop()
