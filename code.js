/**
 * @OnlyCurrentDoc
 */
/* 
  @customfunction 
  @ts-ignore 
*/

const SHEET_ID = '1z5pmGZ3B_bDfdON0bPU1MxqANYm5luHdeekvctrxJeQ';
const GEMINI_API_KEY = 'AQ.Ab8RN6IQSjqkG2MsFmtavCpV47nCSLBZUQunsyFYmo2ZjEnYkA';

// ==========================================
// FITUR 3: "AI CUSTOM CONTEXT" - LIBRARY TEMPLATE / STARTER KITS
// Setiap entri berisi ringkasan arsitektur + potongan kode acuan yang akan
// disisipkan ke prompt Gemini sebagai referensi, supaya hasil generate
// lebih rapi, konsisten, dan minim bug (AI tidak mulai dari nol).
// ==========================================
const TEMPLATES = {
  pos_kasir: {
    label: 'Template POS Kasir',
    description: 'Aplikasi kasir/point-of-sale dengan daftar produk, keranjang belanja, kalkulasi total & kembalian, dan riwayat transaksi sederhana.',
    referenceSnippet: `
      Struktur acuan Template POS Kasir:
      - State utama: products[], cart[], totalTransaksi.
      - Layout 2 kolom: kiri = grid produk (klik untuk masuk keranjang), kanan = ringkasan keranjang + tombol "Bayar".
      - Fungsi wajib: tambahKeranjang(id), hapusItem(id), hitungTotal(), prosesPembayaran(uangDibayar).
      - Gunakan localStorage untuk menyimpan histori transaksi harian di sisi client.
      - Tampilkan struk sederhana (nomor transaksi, tanggal, daftar item, total, kembalian) setelah pembayaran berhasil.
      - Desain sederhana, kontras tinggi, tombol besar (cocok untuk layar sentuh kasir).
    `
  },
  landing_page: {
    label: 'Template Landing Page',
    description: 'Landing page produk/jasa dengan hero section, fitur unggulan, testimoni, harga, dan call-to-action.',
    referenceSnippet: `
      Struktur acuan Template Landing Page:
      - Section wajib berurutan: Navbar sticky -> Hero (headline + CTA button) -> Fitur Unggulan (grid 3 kolom ikon+teks) -> Testimoni (carousel/grid) -> Harga (Pricing table 2-3 tier) -> CTA akhir -> Footer.
      - Gunakan smooth-scroll anchor link dari navbar ke tiap section.
      - Responsive mobile-first, gunakan CSS flexbox/grid, hindari layout yang pecah di layar kecil.
      - Warna aksen konsisten dipakai di semua tombol CTA.
      - Copywriting singkat, jelas, dan berorientasi manfaat (bukan fitur teknis).
    `
  },
  portal_berita: {
    label: 'Template Portal Berita',
    description: 'Portal berita/blog dengan daftar artikel, halaman detail artikel, kategori, dan pencarian sederhana.',
    referenceSnippet: `
      Struktur acuan Template Portal Berita:
      - State utama: articles[] (id, judul, kategori, ringkasan, isi, tanggal, gambar).
      - Halaman/komponen: Header dengan search bar & filter kategori, Daftar Artikel (card grid dengan thumbnail+judul+ringkasan), Halaman Detail Artikel (judul, tanggal, isi lengkap, tombol kembali).
      - Fungsi wajib: filterByCategory(kategori), searchArticles(keyword), openArticle(id).
      - Tampilkan artikel terbaru di atas (urut tanggal descending).
      - Layout mirip media online: 2/3 kolom konten utama, 1/3 kolom sidebar (kategori populer/artikel trending).
    `
  }
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const sheet = SpreadsheetApp.openById(SHEET_ID);
    
    if (action === 'activateAccount') {
      sheet.getSheetByName('Users').appendRow([data.email, 'Active', new Date()]);
      return successResponse({ message: 'Aktivasi berhasil. Selamat datang di AppCraft!' });
    } 
    
    if (action === 'saveProject') {
      const projSheet = sheet.getSheetByName('Projects');
      const now = new Date();
      projSheet.appendRow([data.projectId, data.email, data.code, now]);

      // FITUR 2: VERSION CONTROL - setiap kali Save diklik, simpan snapshot
      // ke sheet 'History' terpisah supaya bisa di-rollback nanti.
      const historySheet = getOrCreateHistorySheet(sheet);
      const versionLabel = data.versionLabel || ('Auto-save ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm:ss'));
      historySheet.appendRow([data.projectId, data.email, data.code, now, versionLabel]);

      return successResponse({ message: 'Kode berhasil diunggah ke Cloud & tersimpan di Riwayat.' });
    }

    if (action === 'getHistory') {
      const historySheet = getOrCreateHistorySheet(sheet);
      const rows = historySheet.getDataRange().getValues();
      const versions = [];
      // Kolom: 0=projectId, 1=email, 2=code, 3=timestamp, 4=versionLabel
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.projectId) {
          versions.push({
            rowIndex: i + 1, // 1-based index sheet, dipakai untuk rollback
            timestamp: rows[i][3],
            label: rows[i][4]
          });
        }
      }
      versions.reverse(); // versi terbaru di atas
      return successResponse({ versions: versions.slice(0, 30) });
    }

    if (action === 'rollbackVersion') {
      const historySheet = getOrCreateHistorySheet(sheet);
      const rowIndex = Number(data.rowIndex);
      if (!rowIndex) return errorResponse('Versi tidak valid.');

      const rowValues = historySheet.getRange(rowIndex, 1, 1, 5).getValues()[0];
      const code = rowValues[2];

      // Simpan kembali sebagai versi terbaru di Projects & History (rollback = restore forward)
      const projSheet = sheet.getSheetByName('Projects');
      const now = new Date();
      projSheet.appendRow([data.projectId, data.email, code, now]);
      historySheet.appendRow([data.projectId, data.email, code, now, 'Rollback ke ' + rowValues[4]]);

      return successResponse({ message: 'Rollback berhasil.', code: code });
    }

    if (action === 'deployVercel') {
      return deployToVercel(data);
    }

    if (action === 'deployNetlify') {
      return deployToNetlify(data);
    }
    
    if (action === 'askGemini') {
      const promptText = data.prompt;
      const templateId = data.templateId; // FITUR 3: id starter kit terpilih (opsional)
      // Menggunakan model Gemini terbaru (3.6 Flash) yang direkomendasikan
      const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${GEMINI_API_KEY}`;

      // Jika user memilih Starter Kit, sisipkan referensi template ke prompt
      // supaya AI tidak mulai dari nol dan hasilnya lebih konsisten.
      let templateContext = '';
      if (templateId && TEMPLATES[templateId]) {
        const tpl = TEMPLATES[templateId];
        templateContext = `
        REFERENSI STARTER KIT (WAJIB DIJADIKAN ACUAN UTAMA STRUKTUR & KOMPONEN):
        Nama Template: ${tpl.label}
        Deskripsi: ${tpl.description}
        ${tpl.referenceSnippet}
        Sesuaikan referensi di atas dengan detail permintaan user di bawah, jangan menyalin mentah-mentah,
        gunakan sebagai kerangka arsitektur agar hasil rapi & minim bug.
        `;
      }

      const systemPrompt = `Anda adalah Senior Full-Stack Architect. Buatkan aplikasi: ${promptText}.
      ${templateContext}
      WAJIB BALAS HANYA DENGAN FORMAT JSON VALID tanpa markdown, tanpa tanda backtick di luar JSON.
      Kunci (key) adalah path/nama file, dan nilai (value) adalah kodenya.
      Berdasarkan permintaan, generate file yang sesuai dari daftar ini:
      - Frontend: index.html, style.css, script.js, src/App.jsx, src/main.jsx
      - Config: package.json, vite.config.js, Dockerfile, vercel.json, .github/workflows/deploy.yml
      - Backend/DB: server.js, schema.sql, prisma/schema.prisma, routes/api.js, controllers/main.js
      
      SANGAT PENTING: Anda WAJIB men-generate file 'README.md' di dalam JSON tersebut yang berisi tutorial cara menjalankan aplikasinya.
      PENTING JIKA MEMBUAT REACT: Jangan gunakan import React dari 'react'. Langsung gunakan const App = () => {...} dan ReactDOM.createRoot(...).`;

      const payload = {
        contents: [{ parts: [{ text: systemPrompt }] }],
        // PAKSA Gemini API untuk HANYA mengeluarkan output JSON
        generationConfig: {
          responseMimeType: "application/json"
        }
      };
      
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(geminiUrl, options);
      const result = JSON.parse(response.getContentText());
      
      if (result.error) return errorResponse(`Gemini Error: ${result.error.message}`);
      if (!result.candidates || result.candidates.length === 0) return errorResponse('AI gagal merespons.');
      
      let aiResponseText = result.candidates[0].content.parts[0].text;
      
      return successResponse({ answer: aiResponseText });
    }

    return errorResponse('Aksi tidak dikenali sistem.');
  } catch (error) {
    return errorResponse(`Server Error: ${error.toString()}`);
  }
}

// SMART COMPILER ENGINE (Untuk Live Server via URL GAS)
function doGet(e) {
  const projectId = e.parameter.id;
  
  if (projectId) {
    try {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Projects');
      const data = sheet.getDataRange().getValues();
      
      for (let i = data.length - 1; i > 0; i--) {
        if (data[i][0] === projectId) {
          const rawCode = data[i][2];
          
          try {
            const files = JSON.parse(rawCode);
            
            // CEK APAKAH INI APLIKASI REACT
            if (files['src/App.jsx'] || files['App.jsx'] || files['src/main.jsx']) {
               let reactCompilerHtml = `
               <!DOCTYPE html>
               <html lang="en">
               <head>
                 <meta charset="UTF-8">
                 <meta name="viewport" content="width=device-width, initial-scale=1.0">
                 <title>React AppCraft Preview</title>
                 <script src="https://unpkg.com/react@18/umd/react.development.js" crossorigin></script>
                 <script src="https://unpkg.com/react-dom@18/umd/react-dom.development.js" crossorigin></script>
                 <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
                 <style>${files['style.css'] || ''}</style>
               </head>
               <body>
                 <div id="root"></div>
                 <script type="text/babel">
                   ${files['src/App.jsx'] || files['App.jsx'] || ''}
                   ${files['src/main.jsx'] || ''}
                 </script>
               </body>
               </html>`;
               return HtmlService.createHtmlOutput(reactCompilerHtml)
                .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
                .addMetaTag('viewport', 'width=device-width, initial-scale=1');
            } 
            else {
              let htmlCode = files['index.html'] || '<h1>File index.html tidak ditemukan</h1>';
              if (files['style.css']) htmlCode = htmlCode.replace('</head>', `<style>${files['style.css']}</style></head>`);
              if (files['script.js']) htmlCode = htmlCode.replace('</body>', `<script>${files['script.js']}</script></body>`);
              
              return HtmlService.createHtmlOutput(htmlCode)
                .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
                .addMetaTag('viewport', 'width=device-width, initial-scale=1');
            }
          } catch (jsonError) {
            return HtmlService.createHtmlOutput("Error memproses file project.")
              .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
          }
        }
      }
      return ContentService.createTextOutput("404: Project tidak ditemukan.");
    } catch (error) {
      return ContentService.createTextOutput("Error Server: " + error.toString());
    }
  }
  
  return ContentService.createTextOutput("AppCraft Compiler Engine Aktif. Gunakan parameter ?id=NamaProject");
}

// ==========================================
// FITUR 2: VERSION CONTROL - helper sheet 'History'
// ==========================================
function getOrCreateHistorySheet(spreadsheet) {
  let historySheet = spreadsheet.getSheetByName('History');
  if (!historySheet) {
    historySheet = spreadsheet.insertSheet('History');
    historySheet.appendRow(['ProjectId', 'Email', 'Code', 'Timestamp', 'VersionLabel']);
  }
  return historySheet;
}

// ==========================================
// FITUR 1: DEPLOYMENT INSTAN - Vercel
// data: { vercelToken, projectName, files: { 'index.html': '...', ... } }
// Dijalankan di server (Apps Script) via UrlFetchApp supaya token tidak
// terekspos ke CORS browser dan menghindari masalah cross-origin.
// ==========================================
function deployToVercel(data) {
  try {
    if (!data.vercelToken) return errorResponse('Vercel Access Token wajib diisi.');
    if (!data.files || Object.keys(data.files).length === 0) return errorResponse('Tidak ada file untuk dideploy.');

    const projectName = (data.projectName || 'appcraft-project').toLowerCase().replace(/[^a-z0-9-]/g, '-');

    // Vercel API butuh index.html di root untuk static deployment sederhana.
    const files = Object.entries(data.files).map(([filename, content]) => ({
      file: filename,
      data: content
    }));

    const payload = {
      name: projectName,
      files: files,
      target: 'production',
      projectSettings: { framework: null }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + data.vercelToken },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api.vercel.com/v13/deployments', options);
    const result = JSON.parse(response.getContentText());

    if (result.error) return errorResponse('Vercel Error: ' + result.error.message);

    const url = result.url ? ('https://' + result.url) : null;
    return successResponse({ message: 'Deploy ke Vercel berhasil!', url: url, raw: result.id });
  } catch (error) {
    return errorResponse('Deploy Vercel gagal: ' + error.toString());
  }
}

// ==========================================
// FITUR 1: DEPLOYMENT INSTAN - Netlify
// data: { netlifyToken, siteId (opsional), projectName, files }
// Netlify butuh file di-zip lalu di-PUT sebagai application/zip.
// Apps Script mendukung ini secara native lewat Utilities.zip().
// ==========================================
function deployToNetlify(data) {
  try {
    if (!data.netlifyToken) return errorResponse('Netlify Access Token wajib diisi.');
    if (!data.files || Object.keys(data.files).length === 0) return errorResponse('Tidak ada file untuk dideploy.');

    const blobs = Object.entries(data.files).map(([filename, content]) =>
      Utilities.newBlob(content, 'text/plain', filename)
    );
    const zipBlob = Utilities.zip(blobs, 'site.zip');

    // Jika siteId belum ada, Netlify akan otomatis membuat site baru.
    const endpoint = data.siteId
      ? `https://api.netlify.com/api/v1/sites/${data.siteId}/deploys`
      : 'https://api.netlify.com/api/v1/sites';

    const options = {
      method: 'post',
      contentType: 'application/zip',
      headers: { Authorization: 'Bearer ' + data.netlifyToken },
      payload: zipBlob.getBytes(),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(endpoint, options);
    const result = JSON.parse(response.getContentText());

    if (result.error || result.errors) return errorResponse('Netlify Error: ' + (result.error || JSON.stringify(result.errors)));

    const url = result.ssl_url || result.url || (result.deploy_ssl_url) || null;
    return successResponse({ message: 'Deploy ke Netlify berhasil!', url: url, siteId: result.site_id || result.id });
  } catch (error) {
    return errorResponse('Deploy Netlify gagal: ' + error.toString());
  }
}

function successResponse(data) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
