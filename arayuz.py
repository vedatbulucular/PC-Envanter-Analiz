# -*- coding: utf-8 -*-
"""
Staj Projesi – 5. Asama
PC Envanter Analiz Arayuzu  (customtkinter  |  Dark Mode)

Baslangic: python arayuz.py
"""

import threading
import datetime
import os
import sys
import io
try:
    import pythoncom
    _PYTHONCOM = True
except ImportError:
    _PYTHONCOM = False  # wmi kullanilamayan ortamlarda guvenli devam

# ── Windowed .exe uyumluluğu ────────────────────────────────────────────────
# PyInstaller --windowed modunda sys.stdout ve sys.stderr None olarak gelir;
# herhangi bir print() veya .buffer erişimi NoneType hatasına neden olur.
# Aşağıdaki blok bu kanalları os.devnull'a yönlendirerek sessizce yok sayar.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8", errors="replace")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8", errors="replace")

import customtkinter as ctk
from tkinter import messagebox

# sistem_bilgisi.py'deki tum fonksiyonlari ic aktar
from sistem_bilgisi import (
    bilgisayar_adi,
    isletim_sistemi_bilgisi,
    cpu_bilgisi,
    ram_bilgisi,
    disk_bilgisi,
    ip_durumu,
    ag_uyeligi,
    yuklu_program_kontrol,
    standart_kontrol,
    excel_raporu_kaydet,
)

# ──────────────────────────────────────────────────────────────────────────────
# Renk & Stil Sabitleri
# ──────────────────────────────────────────────────────────────────────────────
RENK = {
    "bg"          : "#0f1117",   # ana zemin
    "panel"       : "#1a1d27",   # kart / panel zemini
    "border"      : "#2a2d3e",   # kenarlik
    "accent"      : "#4f8ef7",   # mavi vurgu
    "accent_hover": "#3a7de8",
    "success"     : "#2ecc71",
    "warning"     : "#f39c12",
    "danger"      : "#e74c3c",
    "text"        : "#e8eaf0",
    "text_dim"    : "#8b8fa8",
    "green"       : "#00d97e",
    "orange"      : "#ff8c00",   # statik IP kart vurgu rengi
    "purple"      : "#a855f7",   # IP DHCP kart vurgu rengi
    "card_1"      : "#1e2235",
    "card_2"      : "#1b2032",
}

FONT = {
    "title"  : ("Segoe UI", 22, "bold"),
    "sub"    : ("Segoe UI", 12),
    "card_v" : ("Segoe UI", 18, "bold"),
    "card_l" : ("Segoe UI", 10),
    "btn"    : ("Segoe UI", 13, "bold"),
    "mono"   : ("Consolas", 12),
    "mono_hd": ("Consolas", 11, "bold"),
    "status" : ("Segoe UI", 10),
}

# ──────────────────────────────────────────────────────────────────────────────
# Yardimci – analiz sonucu metin formatla
# ──────────────────────────────────────────────────────────────────────────────

def _analiz_yap() -> tuple:
    """Tum kontrolleri calistirir; ham veri demetini dondurur."""
    ram        = ram_bilgisi()
    cpu        = cpu_bilgisi()
    isletim    = isletim_sistemi_bilgisi()
    diskler    = disk_bilgisi()
    programlar = yuklu_program_kontrol()
    ip         = ip_durumu()
    uyelik     = ag_uyeligi()
    disk_yuzde = next(
        (d["Doluluk Oranı (%)"] for d in diskler if d["Bağlama Noktası"] == "C:\\"),
        0.0,
    )
    uyarilar = standart_kontrol(
        ram_gb           = ram["Toplam RAM (GB)"],
        office_durum     = programlar["Microsoft Office"],
        disk_yuzde       = disk_yuzde,
        teamviewer_durum = programlar.get("TeamViewer", ""),
        ip               = ip,
        uyelik           = uyelik,
    )
    return ram, cpu, isletim, diskler, programlar, ip, uyelik, uyarilar


def _metni_olustur(ram, cpu, isletim, diskler, programlar, ip, uyelik, uyarilar) -> str:
    """Analiz sonuclarini formatlı metin olarak dondurur."""
    zaman = datetime.datetime.now().strftime("%d.%m.%Y  %H:%M:%S")
    s = []
    sep  = "=" * 56
    sep2 = "-" * 56
    sep3 = "*" * 56

    s.append(sep)
    s.append(f"  PC ENVANTER RAPORU   |   {zaman}")
    s.append(f"  Bilgisayar : {bilgisayar_adi()}")
    s.append(sep)

    # ── YUKLU PROGRAM KONTROLU (en uste, vurgulu)
    s.append("")
    s.append(sep3)
    s.append("  *** YUKLU PROGRAM KONTROLU ***")
    s.append(sep3)
    for prog, durum in programlar.items():
        isaretci = "[+] VAR " if "VAR" in durum else "[-] YOK "
        s.append(f"  {isaretci}  {prog:<20}  {durum}")
    s.append(sep3)

    # ── IP / AG DURUMU
    s.append("\n  [AG / IP DURUMU]")
    s.append(sep2)
    s.append(f"  IP Yapilandirmasi   : {ip}")
    s.append(f"  Ag Uyeligi          : {uyelik}")

    # ── Isletim Sistemi
    s.append("\n  [ISLETIM SISTEMI]")
    s.append(sep2)
    for k, v in isletim.items():
        s.append(f"  {k:<24}: {v}")

    # ── CPU
    s.append("\n  [CPU]")
    s.append(sep2)
    for k, v in cpu.items():
        s.append(f"  {k:<30}: {v}")

    # ── RAM
    s.append("\n  [RAM]")
    s.append(sep2)
    for k, v in ram.items():
        s.append(f"  {k:<24}: {v}")

    # ── Disk
    s.append("\n  [DISK]")
    s.append(sep2)
    for i, d in enumerate(diskler, 1):
        s.append(f"  [Disk {i}]")
        for k, v in d.items():
            s.append(f"    {k:<22}: {v}")

    # ── Degerlendirme
    s.append("\n  [DEGERLENDIRME / UYARILAR]")
    s.append(sep2)
    if uyarilar:
        for u in uyarilar:
            s.append(f"  {u}")
    else:
        s.append("  [OK] Tum kriterler karsilandi – uyari yok.")

    s.append("\n" + sep)
    return "\n".join(s)


# ──────────────────────────────────────────────────────────────────────────────
# Ozet Kart Bileşeni
# ──────────────────────────────────────────────────────────────────────────────

class OzetKart(ctk.CTkFrame):
    """Ust satirdaki kucuk bilgi karti."""

    def __init__(self, parent, etiket: str, deger: str = "—",
                 renk: str = RENK["accent"], **kwargs):
        super().__init__(
            parent,
            fg_color=RENK["card_1"],
            corner_radius=12,
            border_width=1,
            border_color=RENK["border"],
            **kwargs,
        )
        self._renk = renk

        # Renkli sol serit (canvas ile)
        self._serit = ctk.CTkCanvas(
            self, width=4, bg=renk, highlightthickness=0
        )
        self._serit.pack(side="left", fill="y", padx=(6, 0), pady=8)

        ic = ctk.CTkFrame(self, fg_color="transparent")
        ic.pack(side="left", fill="both", expand=True, padx=10, pady=8)

        self._deger_lbl = ctk.CTkLabel(
            ic, text=deger, font=FONT["card_v"],
            text_color=renk, anchor="w",
        )
        self._deger_lbl.pack(anchor="w")

        ctk.CTkLabel(
            ic, text=etiket, font=FONT["card_l"],
            text_color=RENK["text_dim"], anchor="w",
        ).pack(anchor="w")

        self._alt_lbl = ctk.CTkLabel(
            ic, text="", font=("Segoe UI", 10),
            text_color=RENK["text_dim"], anchor="w", justify="left"
        )
        self._alt_lbl.pack(anchor="w", pady=(2, 0))

    def guncelle(self, deger: str):
        self._deger_lbl.configure(text=deger)


# ──────────────────────────────────────────────────────────────────────────────
# Ana Pencere
# ──────────────────────────────────────────────────────────────────────────────

class PCAnalizApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        # ── customtkinter global ayarları
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title("PC Envanter Analiz Sistemi  –  Staj Projesi")
        self.geometry("900x680")
        self.minsize(760, 540)
        self.configure(fg_color=RENK["bg"])

        # ── analiz sonuclari (paylasimli durum)
        self._son_ram: dict        = {}
        self._son_cpu: dict        = {}
        self._son_isletim: dict    = {}
        self._son_diskler: list    = []
        self._son_programlar: dict = {}
        self._son_ip: str          = ""
        self._son_uyelik: str      = ""
        self._son_uyarilar: list   = []
        self._analiz_yapildi       = False
        self._ip_blink_aktif       = False   # Statik IP yanip sonme durumu

        self._arayuz_olustur()

    # ── Layout inşa ─────────────────────────────────────────────────────────

    def _arayuz_olustur(self):
        # Ana ızgara: sol kenar boşluğu + içerik
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._baslik_olustur()
        self._icerik_olustur()
        self._durum_cubugu_olustur()

    def _baslik_olustur(self):
        """Ust baslik baneri."""
        banner = ctk.CTkFrame(self, fg_color=RENK["panel"],
                              corner_radius=0, height=70)
        banner.grid(row=0, column=0, sticky="ew")
        banner.grid_propagate(False)
        banner.grid_columnconfigure(1, weight=1)

        # Sol – ikon + baslik
        sol = ctk.CTkFrame(banner, fg_color="transparent")
        sol.grid(row=0, column=0, padx=24, pady=12, sticky="w")

        ctk.CTkLabel(
            sol, text="⬡", font=("Segoe UI", 26, "bold"),
            text_color=RENK["accent"],
        ).pack(side="left", padx=(0, 10))

        baslik_ic = ctk.CTkFrame(sol, fg_color="transparent")
        baslik_ic.pack(side="left")
        ctk.CTkLabel(
            baslik_ic, text="PC Envanter Analiz Sistemi",
            font=FONT["title"], text_color=RENK["text"],
        ).pack(anchor="w")
        ctk.CTkLabel(
            baslik_ic, text="Staj Projesi  •  Aşama 5",
            font=FONT["sub"], text_color=RENK["text_dim"],
        ).pack(anchor="w")

        # Sag – zaman etiketi (her saniye guncellenir)
        self._saat_lbl = ctk.CTkLabel(
            banner, text="", font=FONT["sub"],
            text_color=RENK["text_dim"],
        )
        self._saat_lbl.grid(row=0, column=1, sticky="e", padx=24)
        self._saati_guncelle()

    def _icerik_olustur(self):
        """Ozet kartlar + metin kutusu + butonlar."""
        ana = ctk.CTkFrame(self, fg_color="transparent")
        ana.grid(row=1, column=0, sticky="nsew", padx=20, pady=(16, 8))
        ana.grid_columnconfigure(0, weight=1)
        ana.grid_rowconfigure(1, weight=1)

        # ── Ozet kart satiri ────────────────────────────────────────────────
        kart_satiri = ctk.CTkFrame(ana, fg_color="transparent")
        kart_satiri.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        for i in range(5):
            kart_satiri.grid_columnconfigure(i, weight=1, uniform="kart")

        self._kart_bilgisayar = OzetKart(
            kart_satiri, "Bilgisayar Adı", renk=RENK["accent"])
        self._kart_bilgisayar.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self._kart_ram = OzetKart(
            kart_satiri, "Toplam RAM", renk=RENK["green"])
        self._kart_ram.grid(row=0, column=1, sticky="ew", padx=5)

        self._kart_disk = OzetKart(
            kart_satiri, "Disk Doluluk", renk=RENK["warning"])
        self._kart_disk.grid(row=0, column=2, sticky="ew", padx=5)

        self._kart_ip = OzetKart(
            kart_satiri, "IP Durumu", deger="—", renk=RENK["purple"])
        self._kart_ip.grid(row=0, column=3, sticky="ew", padx=5)

        self._kart_uyari = OzetKart(
            kart_satiri, "Uyarı Sayısı", renk=RENK["danger"])
        self._kart_uyari.grid(row=0, column=4, sticky="ew", padx=(5, 0))

        # ── Metin kutusu paneli ──────────────────────────────────────────────
        metin_panel = ctk.CTkFrame(
            ana, fg_color=RENK["panel"],
            corner_radius=14, border_width=1,
            border_color=RENK["border"],
        )
        metin_panel.grid(row=1, column=0, sticky="nsew")
        metin_panel.grid_columnconfigure(0, weight=1)
        metin_panel.grid_rowconfigure(1, weight=1)

        # Panel başlığı
        panel_baslik = ctk.CTkFrame(metin_panel, fg_color="transparent")
        panel_baslik.grid(row=0, column=0, sticky="ew", padx=16, pady=(12, 4))

        ctk.CTkLabel(
            panel_baslik, text="Analiz Çıktısı",
            font=("Segoe UI", 13, "bold"), text_color=RENK["text"],
        ).pack(side="left")

        self._rozet = ctk.CTkLabel(
            panel_baslik, text="  Hazır  ",
            font=("Segoe UI", 10, "bold"),
            text_color="#ffffff",
            fg_color=RENK["border"],
            corner_radius=6,
        )
        self._rozet.pack(side="left", padx=10)

        # Metin kutusu
        self._metin = ctk.CTkTextbox(
            metin_panel,
            font=FONT["mono"],
            fg_color="#0b0e18",
            text_color=RENK["text"],
            border_width=0,
            corner_radius=0,
            wrap="none",
            activate_scrollbars=True,
        )
        self._metin.grid(row=1, column=0, sticky="nsew",
                         padx=(12, 4), pady=(0, 12))

        self._metin.insert("end",
            "  Analiz başlatmak için  'Sistemi Analiz Et'  butonuna tıklayın.\n\n"
            "  Tüm donanım, yazılım ve güvenlik kontrolleri arka planda\n"
            "  çalıştırılacak ve sonuçlar burada görüntülenecektir.\n"
        )
        self._metin.configure(state="disabled")

        # ── Buton satiri ─────────────────────────────────────────────────────
        btn_satiri = ctk.CTkFrame(ana, fg_color="transparent")
        btn_satiri.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        btn_satiri.grid_columnconfigure((0, 1, 2), weight=1)

        self._btn_analiz = ctk.CTkButton(
            btn_satiri,
            text="  ▶   Sistemi Analiz Et",
            font=FONT["btn"],
            fg_color=RENK["accent"],
            hover_color=RENK["accent_hover"],
            height=46,
            corner_radius=12,
            command=self._analiz_baslat,
        )
        self._btn_analiz.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self._btn_sifirla = ctk.CTkButton(
            btn_satiri,
            text="  ⟳   Sistemi Sıfırla",
            font=FONT["btn"],
            fg_color="#4f5b66",
            hover_color="#343d46",
            height=46,
            corner_radius=12,
            command=self._sistemi_sifirla,
        )
        self._btn_sifirla.grid(row=0, column=1, sticky="ew", padx=4)

        self._btn_excel = ctk.CTkButton(
            btn_satiri,
            text="  ⬇   Excel'e Aktar",
            font=FONT["btn"],
            fg_color="#1e6c3e",
            hover_color="#155230",
            height=46,
            corner_radius=12,
            state="disabled",
            command=self._excel_aktar,
        )
        self._btn_excel.grid(row=0, column=2, sticky="ew", padx=(4, 0))

    def _durum_cubugu_olustur(self):
        """Alt durum cubugu."""
        bar = ctk.CTkFrame(
            self, fg_color=RENK["panel"],
            corner_radius=0, height=28,
        )
        bar.grid(row=2, column=0, sticky="ew")
        bar.grid_propagate(False)
        bar.grid_columnconfigure(0, weight=1)

        self._durum_lbl = ctk.CTkLabel(
            bar, text="Hazır.",
            font=FONT["status"],
            text_color=RENK["text_dim"],
            anchor="w",
        )
        self._durum_lbl.grid(row=0, column=0, sticky="w", padx=16)

        ctk.CTkLabel(
            bar, text="customtkinter  |  psutil  |  pandas  |  openpyxl",
            font=FONT["status"],
            text_color=RENK["border"],
        ).grid(row=0, column=1, sticky="e", padx=16)

    # ── Yardimci metodlar ───────────────────────────────────────────────────

    def _saati_guncelle(self):
        now = datetime.datetime.now().strftime("%d.%m.%Y  %H:%M:%S")
        self._saat_lbl.configure(text=now)
        self.after(1000, self._saati_guncelle)

    def _durum_guncelle(self, mesaj: str, renk: str = RENK["text_dim"]):
        self._durum_lbl.configure(text=mesaj, text_color=renk)

    def _rozet_guncelle(self, metin: str, renk: str):
        self._rozet.configure(text=f"  {metin}  ", fg_color=renk)

    def _metin_yaz(self, icerik: str):
        """Metin kutusunu temizleyip yeni icerik yazar."""
        self._metin.configure(state="normal")
        self._metin.delete("1.0", "end")
        self._metin.insert("end", icerik)
        self._metin.configure(state="disabled")

    def _sistemi_sifirla(self):
        """Arayuzu varsayilan haline dondurur."""
        # Yanip sonme varsa durdur
        self._ip_blink_aktif = False

        self._analiz_yapildi = False
        self._son_ram.clear()
        self._son_cpu.clear()
        self._son_isletim.clear()
        self._son_diskler.clear()
        self._son_programlar.clear()
        self._son_ip = ""
        self._son_uyelik = ""
        self._son_uyarilar.clear()

        # Metin kutusu sifirlama
        self._metin_yaz(
            "  Analiz başlatmak için  'Sistemi Analiz Et'  butonuna tıklayın.\n\n"
            "  Tüm donanım, yazılım ve güvenlik kontrolleri arka planda\n"
            "  çalıştırılacak ve sonuçlar burada görüntülenecektir.\n"
        )
        
        # Kartlari sifirlama
        self._kart_bilgisayar.guncelle("—")
        self._kart_bilgisayar._alt_lbl.configure(text="")
        
        self._kart_ram.guncelle("—")
        self._kart_ram._alt_lbl.configure(text="")
        
        self._kart_disk.guncelle("—")
        self._kart_disk._deger_lbl.configure(text_color=RENK["warning"])
        self._kart_disk._alt_lbl.configure(text="")
        
        # IP kartini sifirla ve cerceve rengini orijinaline dondur
        self._kart_ip.guncelle("—")
        self._kart_ip._deger_lbl.configure(text_color=RENK["purple"])
        self._kart_ip._serit.configure(bg=RENK["purple"])
        self._kart_ip.configure(border_color=RENK["border"])
        self._kart_ip._alt_lbl.configure(text="")
        
        self._kart_uyari.guncelle("—")
        self._kart_uyari._deger_lbl.configure(text_color=RENK["danger"])
        self._kart_uyari._alt_lbl.configure(text="")

        # Butonlar ve Durum
        self._btn_analiz.configure(state="normal", text="  ▶   Sistemi Analiz Et")
        self._btn_excel.configure(state="disabled", text="  ⬇   Excel'e Aktar")
        self._rozet_guncelle("Hazır", RENK["border"])
        self._durum_guncelle("Hazır.", RENK["text_dim"])

    def _ip_kart_yanip_sonsun(self, acik: bool = True):
        """Statik IP algilandiğında IP kartinin cercevesini 500ms'de bir kirmiziyla yanip sondurur."""
        if not self._ip_blink_aktif:
            # Yanip sonme durduruldu, kartı normal renge geri al
            self._kart_ip.configure(border_color=RENK["border"])
            return
        if acik:
            self._kart_ip.configure(border_color=RENK["danger"])
        else:
            self._kart_ip.configure(border_color=RENK["border"])
        self.after(500, self._ip_kart_yanip_sonsun, not acik)

    # ── Analiz ──────────────────────────────────────────────────────────────

    def _analiz_baslat(self):
        """Analizi arka planda (thread) baslatir."""
        self._btn_analiz.configure(state="disabled", text="  ⏳  Analiz ediliyor…")
        self._btn_excel.configure(state="disabled")
        self._rozet_guncelle("Çalışıyor", RENK["warning"])
        self._durum_guncelle("Sistem bilgileri toplanıyor, lütfen bekleyin…",
                             RENK["warning"])
        self._metin_yaz(
            "  Analiz başlatıldı…\n\n"
            "  • Donanım bilgileri alınıyor\n"
            "  • Yüklü programlar kontrol ediliyor\n"
            "  • Değerlendirme kuralları uygulanıyor\n\n"
            "  Lütfen birkaç saniye bekleyin.\n"
        )
        t = threading.Thread(target=self._analiz_thread, daemon=True)
        t.start()

    def _analiz_thread(self):
        """Arka plan is parcacigi – UI'yi dogrudan guncellememeli."""
        if _PYTHONCOM:
            pythoncom.CoInitialize()
        try:
            ram, cpu, isletim, diskler, programlar, ip, uyelik, uyarilar = _analiz_yap()
            self._son_ram        = ram
            self._son_cpu        = cpu
            self._son_isletim    = isletim
            self._son_diskler    = diskler
            self._son_programlar = programlar
            self._son_ip         = ip
            self._son_uyelik     = uyelik
            self._son_uyarilar   = uyarilar
            self._analiz_yapildi = True
            metin = _metni_olustur(ram, cpu, isletim, diskler, programlar, ip, uyelik, uyarilar)
            self.after(0, self._analiz_bitti, metin, ram, diskler, ip, uyelik, uyarilar)
        except Exception as hata:
            self.after(0, self._analiz_hata, str(hata))
        finally:
            if _PYTHONCOM:
                pythoncom.CoUninitialize()

    def _analiz_bitti(self, metin: str, ram: dict,
                      diskler: list, ip: str, uyelik: str, uyarilar: list):
        """Ana thread'de UI'yi gunceller."""
        self._metin_yaz(metin)

        # Ozet kartlari guncelle
        # Bilgisayar kartı: ad + ag uyeligi alt bilgisi
        pc_adi = bilgisayar_adi()
        # Uyeligin kisa halini al (ilk kelime yeterli degil; tum degeri goster)
        uyelik_kisa = uyelik.split(":", 1)[-1].strip() if ":" in uyelik else uyelik
        self._kart_bilgisayar.guncelle(pc_adi)
        # Alt etiket guncelle (varsa)
        if hasattr(self._kart_bilgisayar, "_alt_lbl"):
            self._kart_bilgisayar._alt_lbl.configure(text=uyelik_kisa)
        self._kart_ram.guncelle(f"{ram.get('Toplam RAM (GB)', '?')} GB")

        disk_yuzde = next(
            (d["Doluluk Oranı (%)"] for d in diskler
             if d["Bağlama Noktası"] == "C:\\"), 0.0)
        disk_renk = RENK["danger"] if disk_yuzde > 85 else \
                    RENK["warning"] if disk_yuzde > 70 else RENK["green"]
        self._kart_disk.guncelle(f"%{disk_yuzde}")
        self._kart_disk._deger_lbl.configure(text_color=disk_renk)

        # IP kartı – Statik ise turuncu + yanip sonme, DHCP ise mor
        if "Statik" in ip:
            ip_renk = RENK["orange"]
            ip_kisa = "Statik IP"
        elif "DHCP" in ip:
            ip_renk = RENK["purple"]
            ip_kisa = "Otomatik (DHCP)"
        else:
            ip_renk = RENK["text_dim"]
            ip_kisa = "Bilinmiyor"
        self._kart_ip.guncelle(ip_kisa)
        self._kart_ip._deger_lbl.configure(text_color=ip_renk)
        self._kart_ip._serit.configure(bg=ip_renk)

        # Statik IP ise yanip sonmeyi baslat
        if "Statik" in ip:
            self._ip_blink_aktif = True
            self._ip_kart_yanip_sonsun(acik=True)
        else:
            self._ip_blink_aktif = False
            self._kart_ip.configure(border_color=RENK["border"])

        uyari_renk = RENK["danger"] if uyarilar else RENK["green"]
        self._kart_uyari.guncelle(str(len(uyarilar)))
        self._kart_uyari._deger_lbl.configure(text_color=uyari_renk)
        
        # Uyari karti: tum uyarilari kisalt ve listele (sinir yok)
        if uyarilar:
            def kisa_uyari(u: str) -> str:
                u = u.replace("[-]", "").replace("[!]", "").strip()
                if ":" in u:
                    u = u.split(":")[0].strip()
                return f"• {u}"
            ozet = "\n".join([kisa_uyari(u) for u in uyarilar])
            self._kart_uyari._alt_lbl.configure(
                text=ozet, text_color=RENK["danger"],
                font=("Segoe UI", 9),
            )
        else:
            self._kart_uyari._alt_lbl.configure(
                text="Sorun Yok", text_color=RENK["green"],
                font=("Segoe UI", 9),
            )


        # Buton ve rozet
        self._btn_analiz.configure(state="normal",
                                   text="  ▶   Sistemi Analiz Et")
        self._btn_excel.configure(state="normal")
        self._rozet_guncelle("Tamamlandı", RENK["success"])
        self._durum_guncelle(
            f"Analiz tamamlandı  •  {datetime.datetime.now().strftime('%H:%M:%S')}",
            RENK["success"],
        )

    def _analiz_hata(self, hata_mesaj: str):
        self._metin_yaz(f"  [HATA] Analiz sirasinda bir hata olustu:\n\n  {hata_mesaj}")
        self._btn_analiz.configure(state="normal",
                                   text="  ▶   Sistemi Analiz Et")
        self._rozet_guncelle("Hata", RENK["danger"])
        self._durum_guncelle("Analiz basarisiz oldu.", RENK["danger"])

    # ── Excel Aktar ─────────────────────────────────────────────────────────

    def _excel_aktar(self):
        """pc_rapor_{bilgisayar_adi}.xlsx dosyasini olusturur ve kullaniciya bilgi verir."""
        if not self._analiz_yapildi:
            messagebox.showwarning(
                "Uyarı",
                "Lütfen önce 'Sistemi Analiz Et' butonuna tıklayın.",
            )
            return

        self._btn_excel.configure(state="disabled", text="  ⏳  Kaydediliyor…")
        self._durum_guncelle("Excel dosyası oluşturuluyor…", RENK["warning"])
        t = threading.Thread(target=self._excel_thread, daemon=True)
        t.start()

    def _excel_thread(self):
        try:
            yol = excel_raporu_kaydet(
                ram              = self._son_ram,
                cpu              = self._son_cpu,
                isletim_sistemi  = self._son_isletim,
                diskler          = self._son_diskler,
                programlar       = self._son_programlar,
                uyarilar         = self._son_uyarilar,
                ip               = self._son_ip,
                ag_uyeligi_bilgi = self._son_uyelik,
            )
            self.after(0, self._excel_bitti, yol)
        except Exception as hata:
            self.after(0, self._excel_hata, str(hata))

    def _excel_bitti(self, yol: str):
        self._btn_excel.configure(state="normal", text="  ⬇   Excel'e Aktar")
        self._durum_guncelle(
            f"Rapor kaydedildi → {yol}",
            RENK["success"],
        )
        dosya_adi = os.path.basename(yol)
        messagebox.showinfo(
            "Rapor Başarıyla Kaydedildi",
            f"{dosya_adi} dosyası oluşturuldu:\n\n{yol}",
        )

    def _excel_hata(self, hata_mesaj: str):
        self._btn_excel.configure(state="normal", text="  ⬇   Excel'e Aktar")
        self._durum_guncelle("Excel kaydedilemedi.", RENK["danger"])
        messagebox.showerror("Hata", f"Excel dosyası kaydedilemedi:\n\n{hata_mesaj}")


# ──────────────────────────────────────────────────────────────────────────────
# Giris noktasi
# ──────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # Windows'ta UTF-8 stdout – sadece .buffer mevcutsa yonlendir
    # (PyInstaller --windowed modunda sys.stdout None ya da buffer'siz olabilir)
    if sys.platform == "win32" and hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace"
        )
    app = PCAnalizApp()
    app.mainloop()
