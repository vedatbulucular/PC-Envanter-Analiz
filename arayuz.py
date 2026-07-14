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
import concurrent.futures
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
    chrome_kurulu_mu,
    office_kurulu_mu,
    office_lisans_durumu,
    teamviewer_kurulu_mu,
    bitdefender_kurulu_mu,
    acrobat_reader_kurulu_mu,
    sikistirma_araci_kurulu_mu,
    windows_lisans_durumu,
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


def _metni_olustur(ram, cpu, isletim, diskler, programlar, ip, uyelik, uyarilar) -> str:
    """Analiz sonuclarini formatli metin olarak dondurur.
    Sadece program kontrolu, lisans ve ag/IP durumu gosterilir."""
    s = []
    sep2 = "-" * 48

    # ── YUKLU PROGRAM KONTROLU
    s.append("[YUKLU PROGRAM KONTROLU]")
    s.append(sep2)

    # Her program icin ozel etiket ve durum gosterimi
    for prog, durum in programlar.items():
        if prog == "Microsoft Office":
            # Durum formati: 'VAR|Lisansli' veya 'YOK|Kurulu Degil'
            parcalar = durum.split("|", 1)
            kurulum  = parcalar[0].strip()  # VAR / YOK
            lisans   = parcalar[1].strip() if len(parcalar) > 1 else ""
            if "VAR" in kurulum:
                s.append(f"[\u2713] Microsoft Office (Lisans Durumu: {lisans})")
            else:
                s.append("[X] Microsoft Office (Bulunamadi)")
        elif prog == "Antivirus":
            if "VAR" in durum:
                s.append(f"[\u2713] Antivirus: Bitdefender Endpoint Security Tools")
            else:
                s.append("[X] Antivirus: Bitdefender Endpoint Security Tools (Bulunamadi)")
        elif prog == "Acrobat Reader":
            if "VAR" in durum:
                s.append(f"[\u2713] Adobe Acrobat Reader")
            else:
                s.append("[X] Adobe Acrobat Reader (Bulunamadi)")
        elif prog == "Sikistirma Araci":
            if "VAR" in durum:
                # durum = 'VAR (WinRAR)' gibi; arac adini goster
                arac = durum.replace("VAR", "").strip().strip("()").strip()
                etiket = f"Sikistirma Araci ({arac})" if arac else "Sikistirma Araci"
                s.append(f"[\u2713] {etiket}")
            else:
                s.append("[X] Sikistirma Araci (WinRAR / 7-Zip) (Bulunamadi)")
        elif prog == "Windows Lisans":
            # Bu satiri atliyoruz — asagida ayri bolumde gosterilecek
            continue
        elif prog == "Google Chrome":
            if "VAR" in durum:
                s.append(f"[\u2713] Google Chrome")
            else:
                s.append("[X] Google Chrome (Bulunamadi)")
        elif prog == "TeamViewer":
            if "VAR" in durum:
                s.append(f"[\u2713] TeamViewer")
            else:
                s.append("[X] TeamViewer (Bulunamadi)")
        else:
            # Bilinmeyen programlar icin genel gosterim
            if "VAR" in durum:
                s.append(f"[\u2713] {prog}")
            else:
                s.append(f"[X] {prog} (Bulunamadi)")

    # ── SISTEM VE LISANS DURUMU
    s.append("")
    s.append("[SISTEM VE LISANS DURUMU]")
    s.append(sep2)
    win_lisans = programlar.get("Windows Lisans", "Bilinmiyor")
    s.append(f"Windows Isletim Sistemi : {win_lisans}")

    # ── AG / IP DURUMU
    s.append("")
    s.append("[AG / IP DURUMU]")
    s.append(sep2)
    s.append(f"IP Yapilandirmasi  : {ip}")
    s.append(f"Ag Uyeligi         : {uyelik}")

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

        # ── Ozet kart satiri ─────────────────────────────────────────────────────────
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
            bar, text="customtkinter  |  psutil  |  openpyxl",
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
        self._ip_yanip_sonme_iptal()
        self._analiz_yapildi = False
        self._son_ram.clear()
        self._son_cpu.clear()
        self._son_isletim.clear()
        self._son_diskler.clear()
        self._son_programlar.clear()
        self._son_ip = ""
        self._son_uyelik = ""
        self._son_uyarilar.clear()

        self._metin_yaz(
            "  Analiz başlatmak için  'Sistemi Analiz Et'  butonuna tıklayın.\n\n"
            "  Tüm donanım, yazılım ve güvenlik kontrolleri arka planda\n"
            "  çalıştırılacak ve sonuçlar burada görüntülenecektir.\n"
        )

        self._kart_bilgisayar.guncelle("—")
        if hasattr(self._kart_bilgisayar, "_alt_lbl"):
            self._kart_bilgisayar._alt_lbl.configure(text="")
        
        self._kart_ram.guncelle("—")
        if hasattr(self._kart_ram, "_alt_lbl"):
            self._kart_ram._alt_lbl.configure(text="")
        
        self._kart_disk.guncelle("—")
        self._kart_disk._deger_lbl.configure(text_color=RENK["warning"])
        if hasattr(self._kart_disk, "_alt_lbl"):
            self._kart_disk._alt_lbl.configure(text="")
        
        self._kart_ip.guncelle("—")
        self._kart_ip._deger_lbl.configure(text_color=RENK["purple"])
        self._kart_ip._serit.configure(bg=RENK["purple"])
        self._kart_ip.configure(border_color=RENK["border"])
        if hasattr(self._kart_ip, "_alt_lbl"):
            self._kart_ip._alt_lbl.configure(text="")

        self._kart_uyari.guncelle("—")
        self._kart_uyari._deger_lbl.configure(text_color=RENK["danger"])
        self._kart_uyari._alt_lbl.configure(text="")

        self._btn_analiz.configure(state="normal", text="  ▶   Sistemi Analiz Et")
        self._btn_excel.configure(state="disabled", text="  ⬇   Excel'e Aktar")
        self._rozet_guncelle("Hazır", RENK["border"])
        self._durum_guncelle("Hazır.", RENK["text_dim"])

    def _ip_yanip_sonme_iptal(self):
        self._ip_blink_aktif = False
        if hasattr(self, "_ip_blink_id") and self._ip_blink_id:
            self.after_cancel(self._ip_blink_id)
            self._ip_blink_id = None
        self._kart_ip.configure(border_color=RENK["border"])

    def _ip_kart_yanip_sonsun(self, acik: bool = True):
        """Statik IP algilandiginda IP kartinin cercevesini 500ms'de bir kirmiziyla yanip sondurur."""
        if not self._ip_blink_aktif:
            self._kart_ip.configure(border_color=RENK["border"])
            return
        if acik:
            self._kart_ip.configure(border_color=RENK["danger"])
        else:
            self._kart_ip.configure(border_color=RENK["border"])
        self._ip_blink_id = self.after(500, self._ip_kart_yanip_sonsun, not acik)

    # ── Analiz ──────────────────────────────────────────────────────────────

    def _analiz_baslat(self):
        """Analizi arka planda (thread) baslatir."""
        self._ip_yanip_sonme_iptal()
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
        """Arka plan is parcacigi – concurrent.futures ile paralel calisir."""
        if _PYTHONCOM:
            pythoncom.CoInitialize()
        try:
            with concurrent.futures.ThreadPoolExecutor(max_workers=14) as executor:
                futures = {
                    executor.submit(ram_bilgisi): "ram",
                    executor.submit(cpu_bilgisi): "cpu",
                    executor.submit(isletim_sistemi_bilgisi): "os",
                    executor.submit(disk_bilgisi): "disk",
                    executor.submit(ip_durumu): "ip",
                    executor.submit(ag_uyeligi): "uyelik",
                    executor.submit(chrome_kurulu_mu): "chrome",
                    executor.submit(office_kurulu_mu): "office_k",
                    executor.submit(office_lisans_durumu): "office_l",
                    executor.submit(teamviewer_kurulu_mu): "teamviewer",
                    executor.submit(bitdefender_kurulu_mu): "antivirus",
                    executor.submit(acrobat_reader_kurulu_mu): "acrobat",
                    executor.submit(sikistirma_araci_kurulu_mu): "zip",
                    executor.submit(windows_lisans_durumu): "win_lisans",
                }
                
                results = {}
                for f in concurrent.futures.as_completed(futures):
                    name = futures[f]
                    try:
                        res = f.result()
                        results[name] = res
                        
                        if name == "ram":
                            self.after(0, lambda r=res: self._kart_ram.guncelle(f"{r.get('Toplam RAM (GB)', '?')} GB"))
                        elif name == "disk":
                            dy = next((d["Doluluk Oranı (%)"] for d in res if d["Bağlama Noktası"] == "C:\\"), 0.0)
                            dr = RENK["danger"] if dy > 85 else RENK["warning"] if dy > 70 else RENK["green"]
                            def up_disk(y=dy, c=dr):
                                self._kart_disk.guncelle(f"%{y}")
                                self._kart_disk._deger_lbl.configure(text_color=c)
                            self.after(0, up_disk)
                        elif name == "ip":
                            def up_ip(ir=res):
                                cr = RENK["orange"] if "Statik" in ir else RENK["purple"] if "DHCP" in ir else RENK["text_dim"]
                                kisa = "Statik IP" if "Statik" in ir else "Otomatik (DHCP)" if "DHCP" in ir else "Bilinmiyor"
                                self._kart_ip.guncelle(kisa)
                                self._kart_ip._deger_lbl.configure(text_color=cr)
                                self._kart_ip._serit.configure(bg=cr)
                                if "Statik" in ir:
                                    self._ip_blink_aktif = True
                                    self._ip_kart_yanip_sonsun(acik=True)
                                else:
                                    self._ip_yanip_sonme_iptal()
                            self.after(0, up_ip)
                        elif name == "uyelik":
                            def up_uy(ur=res):
                                uk = ur.split(":", 1)[-1].strip() if ":" in ur else ur
                                pc = bilgisayar_adi()
                                self._kart_bilgisayar.guncelle(pc)
                                if hasattr(self._kart_bilgisayar, "_alt_lbl"):
                                    self._kart_bilgisayar._alt_lbl.configure(text=uk)
                            self.after(0, up_uy)
                    except Exception as e:
                        print(f"Hata ({name}): {e}")

            programlar = {
                "Google Chrome": results.get("chrome", "YOK"),
                "Microsoft Office": f"{results.get('office_k', 'YOK')}|{results.get('office_l', 'Kurulu Degil') if 'VAR' in results.get('office_k', '') else 'Kurulu Degil'}",
                "TeamViewer": results.get("teamviewer", "YOK"),
                "Antivirus": results.get("antivirus", "YOK"),
                "Acrobat Reader": results.get("acrobat", "YOK"),
                "Sikistirma Araci": results.get("zip", "YOK"),
                "Windows Lisans": results.get("win_lisans", "Sorgulanamadi")
            }
            
            disk_yuzde = next((d["Doluluk Oranı (%)"] for d in results.get("disk", []) if d["Bağlama Noktası"] == "C:\\"), 0.0)
            uyarilar = standart_kontrol(
                ram_gb           = results.get("ram", {}).get("Toplam RAM (GB)", 0),
                office_durum     = programlar["Microsoft Office"],
                disk_yuzde       = disk_yuzde,
                teamviewer_durum = programlar["TeamViewer"],
                ip               = results.get("ip", ""),
                uyelik           = results.get("uyelik", ""),
                antivirus_durum  = programlar["Antivirus"],
                acrobat_durum    = programlar["Acrobat Reader"],
                sikistirma_durum = programlar["Sikistirma Araci"],
                windows_lisans   = programlar["Windows Lisans"],
            )

            self._son_ram        = results.get("ram", {})
            self._son_cpu        = results.get("cpu", {})
            self._son_isletim    = results.get("os", {})
            self._son_diskler    = results.get("disk", [])
            self._son_programlar = programlar
            self._son_ip         = results.get("ip", "")
            self._son_uyelik     = results.get("uyelik", "")
            self._son_uyarilar   = uyarilar
            self._analiz_yapildi = True

            metin = _metni_olustur(self._son_ram, self._son_cpu, self._son_isletim, self._son_diskler, programlar, self._son_ip, self._son_uyelik, uyarilar)
            self.after(0, self._analiz_bitti_guncelle, metin, uyarilar)
            
        except Exception as hata:
            self.after(0, self._analiz_hata, str(hata))
        finally:
            if _PYTHONCOM:
                pythoncom.CoUninitialize()

    def _analiz_bitti_guncelle(self, metin: str, uyarilar: list):
        """Tum gorevler bitince Log (Textbox) ve Uyari panelini gunceller."""
        self._metin_yaz(metin)

        uyari_renk = RENK["danger"] if uyarilar else RENK["green"]
        self._kart_uyari.guncelle(str(len(uyarilar)))
        self._kart_uyari._deger_lbl.configure(text_color=uyari_renk)

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
