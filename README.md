# Elastic Slow Query and Nginx Export

Seluruh data diambil dari Elasticsearch/Kibana; tidak ada pembacaan slow query dari file lokal atau konfigurasi YAML.

## Sumber MySQL dari Elastic

ID data view MySQL dan Nginx dibaca dari `.env`.

Kolom Excel MySQL mengikuti urutan berikut:

`@timestamp`, `user`, `schema`, `query_time`, `lock_time`, `rows_examined`, `rows_sent`, `query`

```bash
docker compose run --rm elastic-export mysql \
  --start-time "2026-10-24T00:00:00Z" \
  --end-time "2026-10-24T23:59:59Z" \
  --output output/mysql_slow_queries.xlsx
```

Secara default, `START_TIME=auto` mencari tanggal paling awal dari kolom timestamp pada file CSV/Excel di folder `input/`, lalu mengambil data hingga `END_TIME` (default `now`). Nama kolom yang dikenali: `@timestamp`, `mysql_@timestamp`, `http_timestamp`, `timestamp`, `date`, dan `datetime`. Timestamp format ISO dan format ekspor `Oct 26, 2026 @ 08:56:50.238` didukung. Jika tidak ditemukan nilai timestamp yang bisa dibaca, pengambilan dimulai dari awal hari ini berdasarkan zona waktu Asia/Jakarta. Timestamp tanpa zona waktu pada file dianggap sebagai waktu Asia/Jakarta. Secara default semua record pada rentang tersebut diambil. Gunakan `--max-records 10000` jika ingin membatasi hingga 10.000 record per sumber; `--max-records 0` berarti tanpa batas.

## Sumber Nginx dari Elastic

```bash
docker compose run --rm elastic-export nginx \
  --start-time "2026-10-24T00:00:00Z" \
  --end-time "2026-10-24T23:59:59Z" \
  --exclude-code 404 \
  --exclude-code 499 \
  --output output/nginx_access_logs.xlsx
```

`--max-records` membatasi jumlah hasil. Nilai `0` berarti tidak dibatasi. `--exclude-code` dapat diulang; kode tersebut dikecualikan pada query Elasticsearch.

## Ambil kedua sumber dan gabungkan

Timestamp pada file Excel ditampilkan dalam zona waktu Asia/Jakarta dengan format `MMM DD, YYYY @ HH:mm:ss.SSS`, contohnya `Oct 26, 2026 @ 08:56:50.238`. Mode `merge` memasangkan record berdasarkan waktu yang sama hingga detik, sesuai tanggal tanpa milidetik; jika beberapa record Nginx memiliki waktu yang sama, semuanya menghasilkan pasangan baris.

Hasil diurutkan dari waktu terbaru ke terlama, sehingga record terbaru muncul di bagian atas.

```bash
docker compose run --rm elastic-export merge \
  --start-time "2026-10-24T00:00:00Z" \
  --end-time "2026-10-24T23:59:59Z" \
  --exclude-code 404 \
  --output output/merge_logs.xlsx
```

## Visualisasi hasil merge

Dashboard HTML tersedia di `src/index.html`. Buka file tersebut di browser, pilih `output/merge_logs.xlsx`, dan dashboard akan menampilkan jumlah baris, rata-rata waktu query dan response, error HTTP, grafik per user/schema/status, tren query time per jam, hubungan query time dengan response time, tooltip, serta unduhan PNG. Tabel menyediakan mode semua log dan mode unique yang menggabungkan path yang sama per user, dilengkapi jumlah dan rentang tanggal data yang difilter. Saat grafik user diklik, tooltip path akan tetap terbuka dan menyediakan Excel ringkas path untuk user tersebut. Tabel mencakup referer dan dapat diekspor ke Excel mengikuti filter dan mode tabel yang sedang aktif.

Data Excel dibaca dan diproses langsung di browser; tidak ada backend atau database dan file tidak diunggah. Saat ini halaman memuat library pembaca Excel dari jsDelivr, sehingga perlu koneksi internet ketika dashboard dibuka.

Status Nginx yang dikecualikan secara default diatur lewat `NGINX_EXCLUDE_CODES` di `.env`. Isi dengan daftar dipisahkan koma, misalnya `401,403,499`; jika kosong, tidak ada status yang otomatis dikecualikan. Opsi `--exclude-code` tetap bisa dipakai untuk menambahkan pengecualian saat menjalankan perintah. Pengaturan ini berlaku pada mode `nginx` dan `merge`. Hasil merge memiliki satu baris untuk setiap pasangan timestamp yang sama. Record MySQL tanpa pasangan tetap ditampilkan dengan kolom Nginx kosong. `--max-records` berlaku terpisah untuk masing-masing data view.

## Kredensial dan konfigurasi

Salin `.env.example` menjadi `.env`, lalu sesuaikan koneksi dan filter. Docker Compose membaca nilai konfigurasi dari `.env`:

Environment variable opsional:

| Nama | Default | Fungsi |
|---|---|---|
| `ELASTIC_URL` | wajib diisi | URL Kibana/Elastic |
| `ELASTIC_USER` | kosong | Username Basic Auth |
| `ELASTIC_PASS` | kosong | Password Basic Auth |
| `MYSQL_DATA_VIEW_ID` | wajib diisi | ID data view slow query MySQL |
| `NGINX_DATA_VIEW_ID` | wajib diisi | ID data view Nginx |
| `START_TIME` | `auto` | Waktu mulai, atau `auto` untuk membaca file slow query |
| `END_TIME` | `now` | Batas akhir waktu |
| `INPUT_DIR` | `input` | Folder berisi file CSV/Excel acuan tanggal |
| `MYSQL_FILTER` | kosong | Filter Elasticsearch `query_string` tambahan MySQL, misalnya `schema: shop` |
| `MYSQL_QUERY_FORMAT` | `short` | Bentuk query MySQL di Excel: `full` untuk SQL mentah, `short` untuk pola query, atau `super_short` untuk label operasi saja |
| `NGINX_FILTER` | kosong | Filter Elasticsearch `query_string` tambahan Nginx, misalnya `code: 500` |
| `NGINX_EXCLUDE_CODES` | kosong | Status HTTP Nginx yang dikecualikan otomatis, pisahkan dengan koma |

Filter dapat dikombinasikan dengan pengecualian status `--exclude-code` pada mode Nginx/merge.

`MYSQL_QUERY_FORMAT=short` mengambil `query_normalized` dari Elastic jika tersedia; jika tidak, nilai string dan angka pada SQL diganti placeholder agar query serupa memiliki pola yang sama. Query pendek dibatasi 1.200 karakter. `MYSQL_QUERY_FORMAT=super_short` hanya menyimpan perintah utama seperti `SELECT`, `INSERT`, `UPDATE`, atau `DELETE`. Gunakan `MYSQL_QUERY_FORMAT=full` untuk menyimpan SQL mentah. Pengaturan ini berlaku pada mode `mysql` dan `merge`.

Sertifikat HTTPS diverifikasi. Ekspor menggunakan pagination per halaman untuk membatasi penggunaan memori dan tidak meminta total-hit count dari Elasticsearch.

## Docker

```bash
docker compose build

# Merge dengan default env
docker compose run --build --rm elastic-export merge

# Export MySQL
docker compose run --build --rm elastic-export mysql \
  --start-time "2026-10-24T00:00:00Z" \
  --end-time "2026-10-24T23:59:59Z"

# Export Nginx, kecualikan HTTP 404
docker compose run --build --rm elastic-export nginx \
  --start-time "2026-10-24T00:00:00Z" \
  --end-time "2026-10-24T23:59:59Z" \
  --exclude-code 404

# Fetch kedua data view dan hasilkan Excel gabungan
docker compose run --build --rm elastic-export merge \
  --start-time "2026-10-24T00:00:00Z" \
  --end-time "2026-10-24T23:59:59Z"
```

Folder proyek di-mount ke container; file hasil tersedia pada folder `output/` di host. Jalankan `docker compose run --rm elastic-export --help` untuk melihat mode, atau `docker compose run --rm elastic-export mysql --help` untuk bantuan mode tertentu. Jangan jalankan `python3 main.py` langsung di host karena dependensi aplikasi disediakan di dalam image Docker.