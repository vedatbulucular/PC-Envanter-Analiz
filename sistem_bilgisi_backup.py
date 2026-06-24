# -*- coding: utf-8 -*-
"""
Sistem Bilgisi Toplayici
Staj Projesi - 1., 2., 3. ve 4. Asama

Bu script;
  - Bilgisayar adi, isletim sistemi, CPU, RAM ve disk bilgilerini,
  - Yuklu program kontrolunu (Chrome, Office, Antivirus),
  - IP durumu (DHCP / Statik) sorgusunu,
  - Standart kontrol mantigina gore degerlendirme uyarilarini
ekrana yazdirir ve pc_rapor_{bilgisayar_adi}.xlsx olarak disa aktarir.
"""

import datetime
import os
import platform
import socket
import subprocess
import sys
import io
import psutil
import pandas as pd

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
        "Kullanım Oranı (%)" : ram.percent,
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
            # Erişim izni olmayan bölümleri atla
            continue
    return disk_listesi


def ip_durumu() -> str:
    """
    Aktif ag bagdastiricisinin DHCP durumunu PowerShell ile sorgular.

    Get-NetIPConfiguration komutu ile tum bagdastiricilari tarar;
    bir IPv4 adresi atanmis ilk interface'i bulup DHCP durumunu dondurur.

    Donus
    -----
    'Otomatik IP (DHCP)'  – DHCPEnabled True  ise
    'Statik IP (Manuel)'  – DHCPEnabled False ise
    'Bilinmiyor'          – sorgu basarisiz olursa
    """
    komut = (
        "$ag = Get-NetIPConfiguration -ErrorAction SilentlyContinue "
        "| Where-Object { $_.IPv4Address -ne $null } "
        "| Select-Object -First 1; "
        "if ($ag) { "
        "  $dhcp = (Get-NetIPInterface -InterfaceIndex $ag.InterfaceIndex "
        "           -AddressFamily IPv4 -ErrorAction SilentlyContinue).Dhcp; "
        "  if ($dhcp -eq 'Enabled') { 'DHCP' } else { 'Statik' } "
        "} else { 'Bilinmiyor' }"
    )
    try:
        sonuc = subprocess.run(
            ["powershell", "-NonInteractive", "-Command", komut],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        ).stdout.strip()
    except Exception:
        return "Bilinmiyor"

    if sonuc == "DHCP":
        return "Otomatik IP (DHCP)"
    elif sonuc == "Statik":
        return "Statik IP (Manuel)"
    return "Bilinmiyor"


# ──────────────────────────────────────────────────────────────────────────────
# AŞAMA 2 – YÜKLÜ PROGRAM KONTROLÜ
# ──────────────────────────────────────────────────────────────────────────────

def _powershell_calistir(komut: str) -> str:
    """
    Verilen PowerShell komutunu calistirir ve ciktiyi string olarak dondurur.
    Hata durumunda bos string dondurur.
    """
    try:
        sonuc = subprocess.run(
            ["powershell", "-NonInteractive", "-Command", komut],
            capture_output=True,
            text=True,
            timeout=15,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return sonuc.stdout.strip()
    except Exception:
        return ""


def chrome_kurulu_mu() -> str:
    """
    Google Chrome'un kurulu olup olmadigini kayit defteri uzerinden kontrol eder.
    HKLM ve HKCU altindaki Uninstall anahtarlarini tarar.
    Donus: 'VAR' veya 'YOK'
    """
    komut = (
        "$yollar = @("
        "  'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',"
        "  'HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',"
        "  'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*'"
        "); "
        "$bulundu = Get-ItemProperty $yollar -ErrorAction SilentlyContinue "
        "  | Where-Object { $_.DisplayName -like '*Google Chrome*' }; "
        "if ($bulundu) { 'VAR' } else { 'YOK' }"
    )
    cikti = _powershell_calistir(komut)
    return "VAR" if "VAR" in cikti else "YOK"


def office_kurulu_mu() -> str:
    """
    Microsoft Office'in (herhangi bir sürüm/365) kurulu olup olmadigini
    kayit defteri ve WMI üzerinden kontrol eder.
    Donus: 'VAR' veya 'YOK'
    """
    komut = (
        "$yollar = @("
        "  'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',"
        "  'HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*',"
        "  'HKCU:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*'"
        "); "
        "$bulundu = Get-ItemProperty $yollar -ErrorAction SilentlyContinue "
        "  | Where-Object { $_.DisplayName -match 'Microsoft Office|Microsoft 365|Office 365' }; "
        "if ($bulundu) { 'VAR' } else { 'YOK' }"
    )
    cikti = _powershell_calistir(komut)
    return "VAR" if "VAR" in cikti else "YOK"


def antivirusler() -> dict:
    """
    Windows Security Center (WMI/CIM) üzerinden kayitli antivirusleri sorgular.
    Windows Defender her zaman ayrica kontrol edilir.
    Donus: {'Antivirus Adi': 'VAR / YOK'} seklinde sozluk
    """
    sonuclar: dict = {}

    # -- Kayitli 3. parti antivirusler (SecurityCenter2 namespace) ----------
    kayitli_komut = (
        "Get-CimInstance -Namespace 'root/SecurityCenter2' -ClassName AntiVirusProduct "
        "-ErrorAction SilentlyContinue | Select-Object -ExpandProperty displayName"
    )
    kayitli_cikti = _powershell_calistir(kayitli_komut)
    if kayitli_cikti:
        for satir in kayitli_cikti.splitlines():
            ad = satir.strip()
            if ad and ad.lower() != "windows defender":
                sonuclar[ad] = "VAR"

    # -- Windows Defender durumu (MpComputerStatus) --------------------------
    defender_komut = (
        "$d = Get-MpComputerStatus -ErrorAction SilentlyContinue; "
        "if ($d) { $d.AntivirusEnabled } else { 'HATA' }"
    )
    defender_cikti = _powershell_calistir(defender_komut)
    if defender_cikti == "True":
        sonuclar["Windows Defender"] = "VAR (Aktif)"
    elif defender_cikti == "False":
        sonuclar["Windows Defender"] = "VAR (Devre Disi)"
    else:
        sonuclar["Windows Defender"] = "YOK / Erisilemedi"

    return sonuclar


def yuklu_program_kontrol() -> dict:
    """
    Tüm program kontrollerini bir araya getirir ve tek bir sozluk dondurur.
    """
    av = antivirusler()
    # Birden fazla antivirus olabilir; raporlama icin tek satira indiriyoruz
    if av:
        av_ozet = ", ".join(f"{k}: {v}" for k, v in av.items())
    else:
        av_ozet = "YOK / Tespit Edilemedi"

    return {
        "Google Chrome"    : chrome_kurulu_mu(),
        "Microsoft Office" : office_kurulu_mu(),
        "Antivirus"        : av_ozet,
    }


# ──────────────────────────────────────────────────────────────────────────────
# ASAMA 3 – STANDART KONTROL MANTIGI (DEGERLENDIRME)
# ──────────────────────────────────────────────────────────────────────────────

# Esik degerleri – ileride kolayca degistirilebilsin diye sabite alindi
RAM_ESIK_GB       = 8.0    # Bu degerin altinda → Yükseltme gerekli
DISK_ESIK_YUZDE  = 85.0   # Bu degerin üstünde → Dikkat


def standart_kontrol(ram_gb: float, office_durum: str, disk_yuzde: float) -> list[str]:
    """
    Standart kontrol mantigini uygular ve uyari listesi dondurur.

    Parametreler
    ------------
    ram_gb       : Toplam RAM miktari (GB)
    office_durum : office_kurulu_mu() ciktisi ('VAR' veya 'YOK')
    disk_yuzde   : C:\\ surucusunun doluluk orani (yuzde)

    Donus
    -----
    Uyari mesajlarindan olusan liste; hicbir kural tetiklenmezse bos liste.
    """
    uyarilar: list[str] = []

    if ram_gb < RAM_ESIK_GB:
        uyarilar.append(
            f"[!] RAM Yukseltme Gerekli  : Mevcut {ram_gb} GB < Esik {RAM_ESIK_GB} GB"
        )

    if "YOK" in office_durum.upper():
        uyarilar.append(
            "[!] Eksik Yazilim          : Microsoft Office kurulu degil"
        )

    if disk_yuzde > DISK_ESIK_YUZDE:
        uyarilar.append(
            f"[!] Disk Dikkat            : Doluluk {disk_yuzde}% > Esik {DISK_ESIK_YUZDE}%"
        )

    return uyarilar


# ──────────────────────────────────────────────────────────────────────────────
# ASAMA 4 – EXCEL RAPORU (pandas + openpyxl)
# ──────────────────────────────────────────────────────────────────────────────

def excel_raporu_kaydet(
    ram: dict,
    cpu: dict,
    isletim_sistemi: dict,
    diskler: list[dict],
    programlar: dict,
    uyarilar: list[str],
    ip: str = "",
    cikti_dizini: str = ".",
    dosya_adi: str = "",          # bos birakilirsa bilgisayar adina gore otomatik uretilir
) -> str:
    """
    Tum sistem bilgilerini, program durumlarini ve degerlendirme uyarilarini
    pc_rapor_{bilgisayar_adi}.xlsx adinda bir Excel dosyasina yazar.
    Farkli makinelerde calistirildiginda eski raporlarin uzerine yazilmaz.

    Donus: Kaydedilen dosyanin tam yolu.
    """
    zaman_damgasi = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    bilgisayar = bilgisayar_adi()

    # Dosya adini dinamik olustur: bos gelirse bilgisayar adini kullan
    if not dosya_adi:
        # Dosya sisteminde gecersiz olabilecek karakterleri temizle
        temiz_ad = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in bilgisayar)
        dosya_adi = f"pc_rapor_{temiz_ad}.xlsx"
    # ── Sayfa 1: Sistem Ozeti ────────────────────────────────────────────────
    ozet_satirlar = [
        {"Kategori": "Rapor Zamani",       "Bilgi": zaman_damgasi},
        {"Kategori": "Bilgisayar Adi",     "Bilgi": bilgisayar},
        {"Kategori": "IP Durumu",          "Bilgi": ip if ip else "Bilinmiyor"},
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

    # ── Sayfa 2: Disk Bilgileri ──────────────────────────────────────────────
    df_disk = pd.DataFrame(diskler) if diskler else pd.DataFrame(
        columns=["Baglama Noktasi", "Dosya Sistemi", "Toplam (GB)",
                 "Kullanilan (GB)", "Bos (GB)", "Doluluk Orani (%)"]
    )

    # ── Sayfa 3: Program Durumu ──────────────────────────────────────────────
    program_satirlar = [
        {"Program": prog, "Durum": durum}
        for prog, durum in programlar.items()
    ]
    df_program = pd.DataFrame(program_satirlar)

    # ── Sayfa 4: Degerlendirme / Uyarilar ───────────────────────────────────
    if uyarilar:
        uyari_satirlar = [{"Uyari": u.lstrip("[!] ").strip()} for u in uyarilar]
    else:
        uyari_satirlar = [{"Uyari": "Tum kriterler karsilandi – uyari yok."}]
    df_uyari = pd.DataFrame(uyari_satirlar)

    # ── Excel'e Yaz ─────────────────────────────────────────────────────────
    dosya_yolu = os.path.join(cikti_dizini, dosya_adi)
    with pd.ExcelWriter(dosya_yolu, engine="openpyxl") as writer:
        df_ozet.to_excel(writer,    sheet_name="Sistem Ozeti",      index=False)
        df_disk.to_excel(writer,    sheet_name="Disk Bilgileri",     index=False)
        df_program.to_excel(writer, sheet_name="Program Durumu",     index=False)
        df_uyari.to_excel(writer,   sheet_name="Degerlendirme",      index=False)

        # Sutun genisliklerini otomatik ayarla
        for sayfa in writer.sheets.values():
            for sutun_hucreleri in sayfa.columns:
                max_uzunluk = max(
                    len(str(hucre.value)) if hucre.value is not None else 0
                    for hucre in sutun_hucreleri
                )
                sayfa.column_dimensions[
                    sutun_hucreleri[0].column_letter
                ].width = min(max_uzunluk + 4, 60)

    return os.path.abspath(dosya_yolu)



def raporu_yazdir() -> None:
    """Tüm sistem bilgilerini ekrana formatlanmış şekilde yazdırır."""
    ayrac = "=" * 50

    # ── Bilgisayar Adı ──────────────────────────────────
    print(ayrac)
    print("  SİSTEM BİLGİSİ RAPORU")
    print(ayrac)
    print(f"\n  Bilgisayar Adi : {bilgisayar_adi()}\n")

    # ── İşletim Sistemi ─────────────────────────────────
    print("[ISLETIM SISTEMI]")
    print("-" * 40)
    for anahtar, deger in isletim_sistemi_bilgisi().items():
        print(f"  {anahtar:<25}: {deger}")

    # ── CPU ─────────────────────────────────────────────
    print("\n[CPU]")
    print("-" * 40)
    for anahtar, deger in cpu_bilgisi().items():
        print(f"  {anahtar:<30}: {deger}")

    # ── RAM ─────────────────────────────────────────────
    print("\n[RAM]")
    print("-" * 40)
    for anahtar, deger in ram_bilgisi().items():
        print(f"  {anahtar:<25}: {deger}")

    # ── Disk ────────────────────────────────────────────
    print("\n[DISK]")
    print("-" * 40)
    for i, disk in enumerate(disk_bilgisi(), start=1):
        print(f"\n  [Disk {i}]")
        for anahtar, deger in disk.items():
            print(f"    {anahtar:<22}: {deger}")

    # ── Yüklü Program Kontrolü ──────────────────────────────────────────────
    print("\n[YUKLU PROGRAM KONTROLU]")
    print("-" * 40)
    print("  (Kontrol ediliyor, lutfen bekleyin...)")
    programlar = yuklu_program_kontrol()
    for program, durum in programlar.items():
        isaretci = "[+]" if "VAR" in durum else "[-]"
        print(f"  {isaretci} {program:<20}: {durum}")

    # ── IP Durumu ────────────────────────────────────────────────────────────
    print("\n[AG / IP DURUMU]")
    print("-" * 40)
    ip = ip_durumu()
    print(f"  IP Yapılandırması : {ip}")

    # ── Degerlendirme (Asama 3) ──────────────────────────────────────────────
    print("\n[DEGERLENDIRME / UYARILAR]")
    print("-" * 40)
    ram_veri    = ram_bilgisi()
    disk_veri   = disk_bilgisi()
    # C:\ surucusunun doluluk oranini al; bulunamazsa 0
    disk_yuzde  = next(
        (d["Doluluk Oranı (%)"] for d in disk_veri if d["Bağlama Noktası"] == "C:\\"),
        0.0,
    )
    uyarilar = standart_kontrol(
        ram_gb       = ram_veri["Toplam RAM (GB)"],
        office_durum = programlar["Microsoft Office"],
        disk_yuzde   = disk_yuzde,
    )
    if uyarilar:
        for uyari in uyarilar:
            print(f"  {uyari}")
    else:
        print("  [OK] Tum kriterler karsilandi – uyari yok.")

    # ── Excel Raporu (Asama 4) ───────────────────────────────────────────────
    print("\n[EXCEL RAPORU]")
    print("-" * 40)
    print("  Rapor olusturuluyor...")
    kaydedilen_yol = excel_raporu_kaydet(
        ram             = ram_veri,
        cpu             = cpu_bilgisi(),
        isletim_sistemi = isletim_sistemi_bilgisi(),
        diskler         = disk_veri,
        programlar      = programlar,
        uyarilar        = uyarilar,
        ip              = ip,
    )
    print(f"  [OK] Rapor kaydedildi: {kaydedilen_yol}")

    print(f"\n{ayrac}\n")


if __name__ == "__main__":
    # Terminal icin UTF-8 stdout (yalnizca dogrudan calistirinca)
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    raporu_yazdir()
