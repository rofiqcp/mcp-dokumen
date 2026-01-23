# MCP DOCX Unified Server

Server Model Context Protocol (MCP) untuk membaca, membuat, dan mengedit dokumen Word (.docx) dengan 90 tool terintegrasi.

## Peringatan Versi Python

- **Windows**: Unduh dan pasang Python 3.10.11 (64-bit) dari https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe sebelum melanjutkan.
- **Ubuntu 22.04**: Contoh instalasi Python 3.10.12 beserta venv dan pip:
  ```bash
  sudo apt update
  sudo apt install -y software-properties-common
  sudo add-apt-repository -y ppa:deadsnakes/ppa
  sudo apt install -y python3.10 python3.10-venv python3.10-distutils python3.10-dev
  curl -fsSL https://bootstrap.pypa.io/get-pip.py | sudo python3.10
  ```

## Ringkas: Jalur Cepat

1. Clone atau ekstrak repo ini ke folder pilihan (misal `~/mcpdocx`).
2. Jalankan skrip otomatis sesuai OS:
   - Linux/macOS: `chmod +x install.sh && ./install.sh`
   - Windows (PowerShell): `Set-ExecutionPolicy -Scope Process RemoteSigned` lalu `.\\install.ps1`
3. Skrip akan:
   - Menemukan atau membuat `mcp.json` VS Code dan menambahkan server `mcpdocx`.
   - Menyiapkan `.venv` dan meng-install dependensi di folder clone ini.
4. Aktifkan server via MCP client (VS Code / CLI) dan pakai tool DOCX-nya.

## Persiapan

- Python 3.10.x terpasang dan ada di PATH (`python3.10 --version`).
- Git (untuk clone) atau akses file zip.
- VS Code dengan ekstensi MCP/LLM yang membaca `mcp.json`.

## Cara Clone atau Ekstrak

- Clone: `git clone https://github.com/mcpdocx/mcpdocx-unified.git` lalu `cd mcpdocx-unified`.
- Ekstrak ZIP: unduh dari GitHub, ekstrak, lalu buka folder hasil ekstraksi.

## Instalasi Otomatis (Disarankan)

### Linux / macOS

1. Pastikan skrip bisa dieksekusi:
   ```bash
   chmod +x install.sh
   ./install.sh
   ```
2. Skrip akan memeriksa Python 3.10+, membuat `.venv`, meng-install paket, lalu memperbarui `mcp.json` jika ditemukan di lokasi umum:
   - `~/.config/Code/User/globalStorage/mcp/mcp.json`
   - `~/.config/Code - OSS/User/globalStorage/mcp/mcp.json`
   - `~/.vscode-server/data/User/globalStorage/mcp/mcp.json` (VS Code Server/SSH)
   - `~/Library/Application Support/Code/User/globalStorage/mcp/mcp.json` (macOS)
   Jika tidak ada, skrip membuat file di lokasi pertama.

### Windows (PowerShell)

1. Buka PowerShell, izinkan eksekusi sementara:
   ```powershell
   Set-ExecutionPolicy -Scope Process RemoteSigned
   ```
2. Jalankan skrip:
   ```powershell
   .\install.ps1
   ```
3. Skrip akan mencari atau membuat `mcp.json` pada lokasi umum:
   - `%APPDATA%\Code\User\globalStorage\mcp\mcp.json`
   - `%APPDATA%\Code - OSS\User\globalStorage\mcp\mcp.json`
   - `%USERPROFILE%\.vscode\data\User\globalStorage\mcp\mcp.json`
   - `%APPDATA%\VSCodium\User\globalStorage\mcp\mcp.json`
   Jika tidak ditemukan, file dibuat di lokasi pertama.

## Instalasi Manual (Alternatif)

1. Cek Python 3.10+: `python3.10 --version` atau `python --version`.
2. Buat virtual environment di folder repo:
   ```bash
   python3.10 -m venv .venv
   source .venv/bin/activate  # Windows: .\.venv\Scripts\activate
   python -m pip install --upgrade pip
   ```
3. Install paket:
   ```bash
   python -m pip install -e .
   ```
4. Update `mcp.json` (buat jika belum ada) dan tambahkan server:
   ```json
   {
     "mcpServers": {
       "mcpdocx": {
         "command": "/path/ke/repo/.venv/bin/python",
         "args": ["-m", "mcpdocx"],
         "cwd": "/path/ke/repo"
       }
     }
   }
   ```

## Menjalankan Server

- Aktifkan venv, lalu jalankan modul:
  ```bash
  source .venv/bin/activate
  python -m mcpdocx
  ```
- Atau sesuai konfigurasi MCP, biarkan klien memanggil otomatis.

## Fitur Utama (Ringkas)

- Pembuatan dan manajemen dokumen, paragraf, heading, daftar.
- Operasi tabel lengkap (buat, merge, shading, zebra, lebar/tinggi, impor CSV).
- Layout halaman, section, header/footer, penomoran, watermark teks/gambar.
- Sisip gambar/file base64, hyperlink, bookmark, internal link.
- Search/replace, batch replace, redaksi, sanitasi link eksternal.
- Metadata, komentar, footnote, TOC, outline, struktur, statistik, word count.
- Track changes (enable/disable/accept/reject) dan proteksi dokumen.
- Mail merge, copy/merge dokumen, komparasi dokumen, list file .docx.
- **Automatic Tool Usage Tracking**: Setiap tool yang digunakan otomatis dicatat di `mcpdocx/data.json`.

## Tool Usage Tracking

Server ini secara otomatis melacak penggunaan setiap tool dalam file `mcpdocx/data.json`. Fitur ini berguna untuk:

- Menganalisis tool mana yang paling sering digunakan
- Memantau pola penggunaan server
- Debugging dan monitoring

### Cara Kerja

Setiap kali tool dipanggil, fungsi `track_tool_usage()` akan otomatis dijalankan dan mencatat penggunaan ke `data.json`:

```python
# Di dalam setiap tool di server.py
async def create_document(file_path: str, title: Optional[str] = None) -> str:
    """Create a new Word document."""
    track_tool_usage("create_document")  # Otomatis mencatat penggunaan
    return processor.create_document(file_path, title)
```

### Format Data

File `data.json` menyimpan data dalam format:

```json
{
  "tool_usage": {
    "create_document": 5,
    "add_paragraph": 12,
    "save_document": 5
  }
}
```

### Melihat Statistik

Untuk melihat statistik penggunaan, baca langsung file `data.json` atau gunakan script Python:

```python
import json

with open('mcpdocx/data.json', 'r') as f:
    data = json.load(f)
    usage = data.get('tool_usage', {})
    
    # Sort by usage count
    sorted_usage = sorted(usage.items(), key=lambda x: x[1], reverse=True)
    
    print("Top 10 Most Used Tools:")
    for tool, count in sorted_usage[:10]:
        print(f"{tool}: {count}")
```

## Lisensi

MIT License
