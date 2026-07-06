# -*- coding: utf-8 -*-
"""
Sistem Bilgisi Toplayici
Staj Projesi - 1., 2., 3. ve 4. Asama (WMI ve WinReg Optimizasyonlu)

Bu script;
  - Bilgisayar adi, isletim sistemi, CPU, RAM ve disk bilgilerini,
  - Yuklu program kontrolunu (Chrome, Office, Antivirus) (winreg & wmi ile),
  - IP durumu (DHCP / Statik) sorgusunu (wmi ile),
  - Standart kontrol mantigina gore degerlendirme uyarilarini
ekrana yazdirir ve pc_rapor_{bilgisayar_adi}.xlsx olarak disa aktarir.
"""

import datetime
import os
import platform
import socket
import sys
import io
import psutil
import pandas as pd
import wmi
import winreg

# Not: stdout UTF-8 yonlendirmesi yalnizca dogrudan calistirildigi zaman
# etkinlestirilir (asagida __main__ blogu icinde). Boylece arayuz.py
# tarafindan import edildiginde tkinter boru hattini bozmaz.

def bytes_to_gb(bayt: int) -> float:
    """Bayt cinsinden gelen değeri GB'a çevirir."""
    return round(bayt / (1024 ** 3), 2)

def bilgisayar_adi() -> str:
    """Bilgisayarın ağ adını döndürür."""
    return socket.gethostname()

def isletim_sistemi_bilgisi() -> dict:
    """İşletim sistemi bilgilerini döndürür."""
    return {
        "Sistem"   : platform.system(),
        "Sürüm"    : platform.version(),
        "Mimari"   : platform.machine(),
        "İşlemci"  : platform.processor(),
    }

def cpu_bilgisi() -> dict:
    """CPU çekirdek sayısı ve anlık kullanım yüzdesini döndürür."""
    return {
        "Fiziksel Çekirdek Sayısı" : psutil.cpu_count(logical=False),
        "Mantıksal Çekirdek Sayısı": psutil.cpu_count(logical=True),
        "Mevcut Kullanım (%)"      : psutil.cpu_percent(interval=1),
        "Maksimum Frekans (MHz)"   : psutil.cpu_freq().max if psutil.cpu_freq() else "Bilinmiyor",
        "Mevcut Frekans (MHz)"     : psutil.cpu_freq().current if psutil.cpu_freq() else "Bilinmiyor",
    }

def ram_bilgisi() -> dict:
    """Toplam, kullanılan ve boş RAM miktarını döndürür."""
    ram = psutil.virtual_memory()
    return {
        "Toplam RAM (GB)"   : bytes_to_gb(ram.total),
        "Kullanılan (GB)"   : bytes_to_gb(ram.used),
        "Boş (GB)"          : bytes_to_gb(ram.available),
        "Kullanım Oranı (%)": ram.percent,
    }

def disk_bilgisi() -> list[dict]:
    """Tüm disk bölümlerinin doluluk bilgisini döndürür."""
    disk_listesi = []
    for bolum in psutil.disk_partitions(all=False):
        try:
            kullanim = psutil.disk_usage(bolum.mountpoint)
            disk_listesi.append({
                "Bağlama Noktası"    : bolum.mountpoint,
                "Dosya Sistemi"      : bolum.fstype,
                "Toplam (GB)"        : bytes_to_gb(kullanim.total),
                "Kullanılan (GB)"    : bytes_to_gb(kullanim.used),
                "Boş (GB)"           : bytes_to_gb(kullanim.free),
                "Doluluk Oranı (%)"  : kullanim.percent,
            })
        except PermissionError:
            continue
    return disk_listesi

def ip_durumu() -> str:
    """
    Aktif ag bagdastiricisinin DHCP durumunu WMI ile sorgular.
    """
    try:
        c = wmi.WMI()
        for interface in c.Win32_NetworkAdapterConfiguration(IPEnabled=True):
            if interface.DHCPEnabled:
                return "Otomatik IP (DHCP)"
            else:
                return "Statik IP (Manuel)"
        return "Bilinmiyor"
    except Exception as e:
        return f"Hata Olustu {str(e)}"


def ag_uyeligi() -> str:
    """
    WMI Win32_ComputerSystem ile bilgisayarin Domain veya Workgroup
    uyeligini sorgular.
    """
    try:
        c = wmi.WMI()
        for sistem in c.Win32_ComputerSystem():
            if getattr(sistem, "PartOfDomain", False):
                return f"Etki Alani: {getattr(sistem, 'Domain', 'Bilinmiyor')}"
            else:
                return f"Calisma Grubu: {getattr(sistem, 'Workgroup', 'Bilinmiyor')}"
        return "Bilinmiyor"
    except Exception as e:
        return f"Hata Olustu: {str(e)}"


# ──────────────────────────────────────────────────────────────────────────────
# AŞAMA 2 – YÜKLÜ PROGRAM KONTROLÜ (WinReg & WMI)
# ──────────────────────────────────────────────────────────────────────────────

def _kayit_defterinde_ara(aranacak_kelimeler: list[str]) -> bool:
    """
    Kayıt defterindeki Uninstall dizinlerinde verilen kelimeleri arar.
    """
    yollar = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    hives = [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]
    
    for hive in hives:
        for yol in yollar:
            try:
                key = winreg.OpenKey(hive, yol)
                for i in range(0, winreg.QueryInfoKey(key)[0]):
                    try:
                        subkey_name = winreg.EnumKey(key, i)
                        subkey = winreg.OpenKey(key, subkey_name)
                        name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                        
                        for kelime in aranacak_kelimeler:
                            if kelime.lower() in str(name).lower():
                                winreg.CloseKey(subkey)
                                winreg.CloseKey(key)
                                return True
                        winreg.CloseKey(subkey)
                    except EnvironmentError:
                        continue
                winreg.CloseKey(key)
            except Exception:
                continue
    return False

def chrome_kurulu_mu() -> str:
    """Google Chrome'un kurulu olup olmadigini kontrol eder."""
    return "VAR" if _kayit_defterinde_ara(["Google Chrome"]) else "YOK"

def office_kurulu_mu() -> str:
    """Microsoft Office'in kurulu olup olmadigini kontrol eder."""
    return "VAR" if _kayit_defterinde_ara(["Microsoft Office", "Microsoft 365", "Office 365"]) else "YOK"

def teamviewer_kurulu_mu() -> str:
    """
    TeamViewer'in kurulu olup olmadigini kayit defteri ve
    varsayilan kurulum yolu uzerinden kontrol eder.
    """
    # Oncelikle kayit defterinde ara
    if _kayit_defterinde_ara(["TeamViewer"]):
        return "VAR"
    # Ikincil kontrol: varsayilan kurulum dizinleri
    kurulum_yollari = [
        r"C:\Program Files\TeamViewer\TeamViewer.exe",
        r"C:\Program Files (x86)\TeamViewer\TeamViewer.exe",
    ]
    for yol in kurulum_yollari:
        if os.path.exists(yol):
            return "VAR"
    return "YOK"

def antivirusler() -> dict:
    """
    WMI üzerinden kayitli antivirusleri sorgular.
    """
    sonuclar: dict = {}

    # -- Kayitli 3. parti antivirusler (SecurityCenter2 namespace) ----------
    try:
        c_sc = wmi.WMI(namespace=r"root\SecurityCenter2")
        for av in c_sc.AntiVirusProduct():
            ad = av.displayName.strip()
            if ad and ad.lower() != "windows defender":
                sonuclar[ad] = "VAR"
    except Exception:
        pass  # Admin yetkisi yoksa veya WMI bozuksa atla

    # -- Windows Defender durumu ---------------------------------------------
    try:
        c_def = wmi.WMI(namespace=r"root\Microsoft\Windows\Defender")
        defender_status = c_def.MSFT_MpComputerStatus()[0]
        if getattr(defender_status, "AntivirusEnabled", False):
            sonuclar["Windows Defender"] = "VAR (Aktif)"
        else:
            sonuclar["Windows Defender"] = "VAR (Devre Disi)"
    except Exception:
        sonuclar["Windows Defender"] = "YOK / Erisilemedi (Yetki Eksik)"

    return sonuclar

def yuklu_program_kontrol() -> dict:
    av = antivirusler()
    if av:
        av_ozet = ", ".join(f"{k}: {v}" for k, v in av.items())
    else:
        av_ozet = "YOK / Tespit Edilemedi"

    return {
        "Google Chrome"    : chrome_kurulu_mu(),
        "Microsoft Office" : office_kurulu_mu(),
        "TeamViewer"       : teamviewer_kurulu_mu(),
        "Antivirus"        : av_ozet,
    }


# ──────────────────────────────────────────────────────────────────────────────
# ASAMA 3 – STANDART KONTROL MANTIGI (DEGERLENDIRME)
# ──────────────────────────────────────────────────────────────────────────────

RAM_ESIK_GB       = 8.0
DISK_ESIK_YUZDE  = 85.0

def standart_kontrol(
    ram_gb: float,
    office_durum: str,
    disk_yuzde: float,
    teamviewer_durum: str = "",
    ip: str = "",
    uyelik: str = "",
) -> list[str]:
    """
    Donanim, yazilim ve ag kriterlerini degerlendirerek uyari listesi dondurur.

    Mevcut kurallar:
      - RAM < 8 GB
      - Microsoft Office yok
      - Disk dolulugu > %85
    Yeni eklenen kurallar:
      - TeamViewer yok
      - Statik / Manuel IP konfigurasyonu
      - Cihaz etki alani (Domain) disinda (Workgroup)
    """
    uyarilar: list[str] = []

    # ── Mevcut Donanim & Yazilim Kurallari ───────────────────────────────────
    if ram_gb < RAM_ESIK_GB:
        uyarilar.append(f"[!] RAM Yukseltme Gerekli  : Mevcut {ram_gb} GB < Esik {RAM_ESIK_GB} GB")
    if "YOK" in office_durum.upper():
        uyarilar.append("[!] Eksik Yazilim          : Microsoft Office kurulu degil")
    if disk_yuzde > DISK_ESIK_YUZDE:
        uyarilar.append(f"[!] Disk Dikkat            : Doluluk {disk_yuzde}% > Esik {DISK_ESIK_YUZDE}%")

    # ── Yeni Yazilim Kurali: TeamViewer ──────────────────────────────────────
    if teamviewer_durum and "YOK" in teamviewer_durum.upper():
        uyarilar.append("[-] TeamViewer yüklü değil!")

    # ── Yeni Ag Kurallari ────────────────────────────────────────────────────
    if ip and "Statik" in ip:
        uyarilar.append("[-] Cihazda Statik IP yapılandırması mevcut!")
    if uyelik and ("Calisma Grubu" in uyelik or "Workgroup" in uyelik.lower()):
        uyarilar.append("[-] Cihaz etki alanına (Domain) bağlı değil!")

    return uyarilar


# ──────────────────────────────────────────────────────────────────────────────
# ASAMA 4 – EXCEL RAPORU
# ──────────────────────────────────────────────────────────────────────────────

def excel_raporu_kaydet(
    ram: dict, cpu: dict, isletim_sistemi: dict, diskler: list[dict],
    programlar: dict, uyarilar: list[str], ip: str = "",
    ag_uyeligi_bilgi: str = "",
    cikti_dizini: str = ".", dosya_adi: str = ""
) -> str:
    zaman_damgasi = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    bilgisayar = bilgisayar_adi()

    if not dosya_adi:
        temiz_ad = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in bilgisayar)
        dosya_adi = f"pc_rapor_{temiz_ad}.xlsx"
        
    ozet_satirlar = [
        {"Kategori": "Rapor Zamani",       "Bilgi": zaman_damgasi},
        {"Kategori": "Bilgisayar Adi",     "Bilgi": bilgisayar},
        {"Kategori": "IP Durumu",          "Bilgi": ip if ip else "Bilinmiyor"},
        {"Kategori": "Ag Uyeligi",         "Bilgi": ag_uyeligi_bilgi if ag_uyeligi_bilgi else "Bilinmiyor"},
        {"Kategori": "Isletim Sistemi",    "Bilgi": isletim_sistemi.get("Sistem", "")},
        {"Kategori": "OS Surumu",          "Bilgi": isletim_sistemi.get("Sürüm", "")},
        {"Kategori": "Mimari",             "Bilgi": isletim_sistemi.get("Mimari", "")},
        {"Kategori": "Islemci",            "Bilgi": isletim_sistemi.get("İşlemci", "")},
        {"Kategori": "Fiziksel Cekirdek",  "Bilgi": cpu.get("Fiziksel Çekirdek Sayısı", "")},
        {"Kategori": "Mantiksal Cekirdek", "Bilgi": cpu.get("Mantıksal Çekirdek Sayısı", "")},
        {"Kategori": "CPU Kullanim (%)",   "Bilgi": cpu.get("Mevcut Kullanım (%)", "")},
        {"Kategori": "Toplam RAM (GB)",    "Bilgi": ram.get("Toplam RAM (GB)", "")},
        {"Kategori": "Kullanilan RAM (GB)","Bilgi": ram.get("Kullanılan (GB)", "")},
        {"Kategori": "Bos RAM (GB)",       "Bilgi": ram.get("Boş (GB)", "")},
        {"Kategori": "RAM Kullanim (%)",   "Bilgi": ram.get("Kullanım Oranı (%)", "")},
    ]
    df_ozet = pd.DataFrame(ozet_satirlar)

    df_disk = pd.DataFrame(diskler) if diskler else pd.DataFrame(
        columns=["Baglama Noktasi", "Dosya Sistemi", "Toplam (GB)",
                 "Kullanilan (GB)", "Bos (GB)", "Doluluk Orani (%)"]
    )

    program_satirlar = [{"Program": prog, "Durum": durum} for prog, durum in programlar.items()]
    df_program = pd.DataFrame(program_satirlar)

    if uyarilar:
        uyari_satirlar = [{"Uyari": u.lstrip("[!] ").strip()} for u in uyarilar]
    else:
        uyari_satirlar = [{"Uyari": "Tum kriterler karsilandi – uyari yok."}]
    df_uyari = pd.DataFrame(uyari_satirlar)

    dosya_yolu = os.path.join(cikti_dizini, dosya_adi)
    with pd.ExcelWriter(dosya_yolu, engine="openpyxl") as writer:
        df_ozet.to_excel(writer,    sheet_name="Sistem Ozeti",      index=False)
        df_disk.to_excel(writer,    sheet_name="Disk Bilgileri",     index=False)
        df_program.to_excel(writer, sheet_name="Program Durumu",     index=False)
        df_uyari.to_excel(writer,   sheet_name="Degerlendirme",      index=False)

        for sayfa in writer.sheets.values():
            for sutun_hucreleri in sayfa.columns:
                max_uzunluk = max(
                    len(str(hucre.value)) if hucre.value is not None else 0
                    for hucre in sutun_hucreleri
                )
                sayfa.column_dimensions[sutun_hucreleri[0].column_letter].width = min(max_uzunluk + 4, 60)

    return os.path.abspath(dosya_yolu)


def raporu_yazdir() -> None:
    ayrac = "=" * 50
    print(ayrac)
    print("  SİSTEM BİLGİSİ RAPORU")
    print(ayrac)
    print(f"\n  Bilgisayar Adi : {bilgisayar_adi()}\n")

    print("[ISLETIM SISTEMI]")
    print("-" * 40)
    for anahtar, deger in isletim_sistemi_bilgisi().items():
        print(f"  {anahtar:<25}: {deger}")

    print("\n[CPU]")
    print("-" * 40)
    for anahtar, deger in cpu_bilgisi().items():
        print(f"  {anahtar:<30}: {deger}")

    print("\n[RAM]")
    print("-" * 40)
    for anahtar, deger in ram_bilgisi().items():
        print(f"  {anahtar:<25}: {deger}")

    print("\n[DISK]")
    print("-" * 40)
    for i, disk in enumerate(disk_bilgisi(), start=1):
        print(f"\n  [Disk {i}]")
        for anahtar, deger in disk.items():
            print(f"    {anahtar:<22}: {deger}")

    print("\n[YUKLU PROGRAM KONTROLU]")
    print("-" * 40)
    print("  (Kayıt defteri ve WMI taranıyor, lutfen bekleyin...)")
    programlar = yuklu_program_kontrol()
    for program, durum in programlar.items():
        isaretci = "[+]" if "VAR" in durum else "[-]"
        print(f"  {isaretci} {program:<20}: {durum}")

    print("\n[AG / IP DURUMU]")
    print("-" * 40)
    ip = ip_durumu()
    uyelik = ag_uyeligi()
    print(f"  IP Yapılandırması : {ip}")
    print(f"  Ag Uyeligi        : {uyelik}")

    print("\n[DEGERLENDIRME / UYARILAR]")
    print("-" * 40)
    ram_veri    = ram_bilgisi()
    disk_veri   = disk_bilgisi()
    disk_yuzde  = next((d["Doluluk Oranı (%)"] for d in disk_veri if d["Bağlama Noktası"] == "C:\\"), 0.0)
    
    uyarilar = standart_kontrol(
        ram_gb           = ram_veri["Toplam RAM (GB)"],
        office_durum     = programlar["Microsoft Office"],
        disk_yuzde       = disk_yuzde,
        teamviewer_durum = programlar.get("TeamViewer", ""),
        ip               = ip,
        uyelik           = uyelik,
    )
    if uyarilar:
        for uyari in uyarilar:
            print(f"  {uyari}")
    else:
        print("  [OK] Tum kriterler karsilandi – uyari yok.")

    print("\n[EXCEL RAPORU]")
    print("-" * 40)
    print("  Rapor olusturuluyor...")
    kaydedilen_yol = excel_raporu_kaydet(
        ram              = ram_veri,
        cpu              = cpu_bilgisi(),
        isletim_sistemi  = isletim_sistemi_bilgisi(),
        diskler          = disk_veri,
        programlar       = programlar,
        uyarilar         = uyarilar,
        ip               = ip,
        ag_uyeligi_bilgi = uyelik,
    )
    print(f"  [OK] Rapor kaydedildi: {kaydedilen_yol}")
    print(f"\n{ayrac}\n")


if __name__ == "__main__":
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    raporu_yazdir()