const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const ExcelJS = require('exceljs');
const path = require('path');
const crypto = require('crypto');
const session = require('express-session');
const SQLiteStore = require('connect-sqlite3')(session);
const { rateLimit } = require('express-rate-limit');

const app = express();
const isProduction = process.env.NODE_ENV === 'production';
const databasePath = process.env.DB_PATH || path.join(__dirname, 'rekam_medis.db');
const appPassword = process.env.APP_PASSWORD || 'dentrecap-local-2026';
const sessionSecret = process.env.SESSION_SECRET || crypto.randomBytes(32).toString('hex');

if (isProduction && (!process.env.APP_PASSWORD || !process.env.SESSION_SECRET)) {
    throw new Error('APP_PASSWORD and SESSION_SECRET must be set in production.');
}

app.disable('x-powered-by');
if (isProduction) app.set('trust proxy', 1);
app.use(express.json());
app.use(session({
    store: new SQLiteStore({
        db: 'sessions.sqlite',
        dir: path.dirname(databasePath)
    }),
    secret: sessionSecret,
    resave: false,
    saveUninitialized: false,
    cookie: {
        httpOnly: true,
        secure: isProduction,
        sameSite: 'lax',
        maxAge: 8 * 60 * 60 * 1000
    }
}));
app.use(express.static(path.join(__dirname, 'public')));

// Inisialisasi Database SQLite
const db = new sqlite3.Database(databasePath, (err) => {
    if (err) console.error(err.message);
    console.log('Terkoneksi ke database SQLite.');
});

const months = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AGST', 'SEP', 'OKT', 'NOV', 'DES'];
const getAppDate = () => new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Jakarta',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
}).format(new Date());
const isValidDate = value => {
    if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
    const date = new Date(`${value}T00:00:00Z`);
    return !Number.isNaN(date.getTime()) && date.toISOString().slice(0, 10) === value;
};
const defaultDiagnoses = [
    { kode: 'K08.3', warna: 'FFFFFF00' },
    { kode: 'K04.11', warna: 'FFFFD700' },
    { kode: 'Z01.6', warna: 'FF00FFFF' },
    { kode: 'K10', warna: 'FF00FF00' },
    { kode: 'K03.10', warna: 'FFFF00FF' }
];

const dbRun = (sql, params = []) => new Promise((resolve, reject) => {
    db.run(sql, params, function(err) {
        if (err) return reject(err);
        resolve(this);
    });
});

const dbAll = (sql, params = []) => new Promise((resolve, reject) => {
    db.all(sql, params, (err, rows) => err ? reject(err) : resolve(rows));
});

const dbGet = (sql, params = []) => new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => err ? reject(err) : resolve(row));
});

async function initializeDatabase() {
    await dbRun(`CREATE TABLE IF NOT EXISTS records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nama TEXT,
    umur INTEGER,
    bulan TEXT,
    tindakan TEXT,
    icd_kode TEXT,
    icd_warna TEXT,
    gender TEXT NOT NULL DEFAULT 'L'
)`);

    const columns = await dbAll('PRAGMA table_info(records)');
    if (!columns.some(column => column.name === 'gender')) {
        await dbRun("ALTER TABLE records ADD COLUMN gender TEXT NOT NULL DEFAULT 'L'");
    }
    if (!columns.some(column => column.name === 'tanggal')) {
        await dbRun('ALTER TABLE records ADD COLUMN tanggal TEXT');
        await dbRun('UPDATE records SET tanggal = ? WHERE tanggal IS NULL', [getAppDate()]);
    }

    await dbRun(`CREATE TABLE IF NOT EXISTS diagnoses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kode TEXT NOT NULL UNIQUE,
        warna TEXT NOT NULL
    )`);

    const diagnosisCount = await dbGet('SELECT COUNT(*) AS count FROM diagnoses');
    if (diagnosisCount.count === 0) {
        for (const diagnosis of defaultDiagnoses) {
            await dbRun('INSERT INTO diagnoses (kode, warna) VALUES (?, ?)', [diagnosis.kode, diagnosis.warna]);
        }
    }
}

function normalizeDiagnosisColor(color) {
    if (typeof color !== 'string' || !/^#[0-9a-fA-F]{6}$/.test(color)) return null;
    return `FF${color.slice(1).toUpperCase()}`;
}

function passwordMatches(candidate) {
    if (typeof candidate !== 'string') return false;
    const expected = crypto.createHash('sha256').update(appPassword).digest();
    const actual = crypto.createHash('sha256').update(candidate).digest();
    return crypto.timingSafeEqual(expected, actual);
}

const loginLimiter = rateLimit({
    windowMs: 15 * 60 * 1000,
    limit: 5,
    skipSuccessfulRequests: true,
    standardHeaders: 'draft-8',
    legacyHeaders: false,
    message: { error: 'Terlalu banyak percobaan kata sandi. Tunggu sebelum mencoba lagi.' }
});

app.get('/api/session', (req, res) => {
    res.json({ authenticated: Boolean(req.session.authenticated), today: getAppDate() });
});

app.post('/api/login', loginLimiter, (req, res, next) => {
    if (!passwordMatches(req.body.password)) {
        return res.status(401).json({ error: 'Kata sandi tidak sesuai.' });
    }

    req.session.regenerate(err => {
        if (err) return next(err);
        req.session.authenticated = true;
        req.session.save(saveErr => {
            if (saveErr) return next(saveErr);
            res.json({ authenticated: true });
        });
    });
});

app.post('/api/logout', (req, res, next) => {
    req.session.destroy(err => {
        if (err) return next(err);
        res.clearCookie('connect.sid', {
            httpOnly: true,
            secure: isProduction,
            sameSite: 'lax'
        });
        res.json({ authenticated: false });
    });
});

app.use('/api', (req, res, next) => {
    if (!req.session.authenticated) {
        return res.status(401).json({ error: 'Silakan login untuk melanjutkan.' });
    }
    next();
});

// API: Simpan Data Baru
app.post('/api/records', (req, res) => {
    const { nama, umur, bulan, gender, tindakan, icd_kode } = req.body;
    if (!months.includes(bulan) || !['L', 'P'].includes(gender)) {
        return res.status(400).json({ error: 'Bulan atau jenis kelamin tidak valid.' });
    }

    db.get('SELECT warna FROM diagnoses WHERE kode = ?', [icd_kode], (diagnosisErr, diagnosis) => {
        if (diagnosisErr) return res.status(500).json({ error: diagnosisErr.message });
        if (!diagnosis) return res.status(400).json({ error: 'Diagnosis tidak ditemukan.' });

        const query = `INSERT INTO records (nama, umur, bulan, tindakan, icd_kode, icd_warna, gender, tanggal) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`;
        db.run(query, [nama, umur, bulan, tindakan, icd_kode, diagnosis.warna, gender, getAppDate()], function(err) {
            if (err) return res.status(500).json({ error: err.message });
            res.json({ id: this.lastID, message: 'Data berhasil disimpan!' });
        });
    });
});

// API: Kelola diagnosis
app.get('/api/diagnoses', (req, res) => {
    db.all('SELECT * FROM diagnoses ORDER BY id', [], (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(rows);
    });
});

app.post('/api/diagnoses', (req, res) => {
    const kode = typeof req.body.kode === 'string' ? req.body.kode.trim().toUpperCase() : '';
    const warna = normalizeDiagnosisColor(req.body.warna);
    if (!kode || !warna) return res.status(400).json({ error: 'Kode dan warna diagnosis wajib diisi.' });

    db.run('INSERT INTO diagnoses (kode, warna) VALUES (?, ?)', [kode, warna], function(err) {
        if (err) return res.status(err.code === 'SQLITE_CONSTRAINT' ? 409 : 500).json({ error: err.code === 'SQLITE_CONSTRAINT' ? 'Kode diagnosis sudah ada.' : err.message });
        res.status(201).json({ id: this.lastID, kode, warna });
    });
});

app.put('/api/diagnoses/:id', (req, res) => {
    const id = Number(req.params.id);
    const kode = typeof req.body.kode === 'string' ? req.body.kode.trim().toUpperCase() : '';
    const warna = normalizeDiagnosisColor(req.body.warna);
    if (!Number.isSafeInteger(id) || id < 1 || !kode || !warna) {
        return res.status(400).json({ error: 'ID, kode, atau warna diagnosis tidak valid.' });
    }

    db.run('UPDATE diagnoses SET kode = ?, warna = ? WHERE id = ?', [kode, warna, id], function(err) {
        if (err) return res.status(err.code === 'SQLITE_CONSTRAINT' ? 409 : 500).json({ error: err.code === 'SQLITE_CONSTRAINT' ? 'Kode diagnosis sudah ada.' : err.message });
        if (this.changes === 0) return res.status(404).json({ error: 'Diagnosis tidak ditemukan.' });
        res.json({ id, kode, warna });
    });
});

app.delete('/api/diagnoses/:id', (req, res) => {
    const id = Number(req.params.id);
    if (!Number.isSafeInteger(id) || id < 1) return res.status(400).json({ error: 'ID diagnosis tidak valid.' });
    db.run('DELETE FROM diagnoses WHERE id = ?', [id], function(err) {
        if (err) return res.status(500).json({ error: err.message });
        if (this.changes === 0) return res.status(404).json({ error: 'Diagnosis tidak ditemukan.' });
        res.status(204).end();
    });
});

// API: Ambil Semua Data
app.get('/api/records', (req, res) => {
    const requestedDate = req.query.date || getAppDate();
    if (!isValidDate(requestedDate)) return res.status(400).json({ error: 'Tanggal tidak valid.' });
    db.all(`SELECT * FROM records WHERE tanggal = ? ORDER BY id DESC`, [requestedDate], (err, rows) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(rows);
    });
});

// API: Download Excel
app.get('/api/export', (req, res) => {
    const requestedIds = req.query.ids;
    if (typeof requestedIds !== 'string' || requestedIds.length === 0) {
        return res.status(400).json({ error: 'Pilih minimal satu data untuk diekspor.' });
    }

    const idValues = requestedIds.split(',');
    const ids = [...new Set(idValues.map(Number))];
    if (idValues.some((id, index) => !/^[1-9]\d*$/.test(id) || !Number.isSafeInteger(ids[index]))) {
        return res.status(400).json({ error: 'Daftar ID data tidak valid.' });
    }

    const placeholders = ids.map(() => '?').join(',');
    db.all(`SELECT * FROM records WHERE id IN (${placeholders})`, ids, async(err, rows) => {
        if (err) return res.status(500).send("Database error");

        const workbook = new ExcelJS.Workbook();
        months.forEach(month => {
            const sheet = workbook.addWorksheet(month);
            // Setup Kolom
            sheet.columns = [
                { header: 'NAMA PASIEN', key: 'nama', width: 30 },
                { header: 'TANGGAL DICATAT', key: 'tanggal', width: 16 },
                { header: 'UMUR LAKI-LAKI', key: 'umur_laki', width: 16 },
                { header: 'UMUR PEREMPUAN', key: 'umur_perempuan', width: 18 },
                { header: 'HASIL PEMERIKSAAN', key: 'tindakan', width: 40 },
                { header: 'KODE', key: 'icd_kode', width: 15 }
            ];

            // Filter data sesuai tab bulan ini
            const monthData = rows.filter(r => r.bulan === month);

            monthData.forEach(data => {
                const row = sheet.addRow({
                    nama: data.nama,
                    tanggal: data.tanggal,
                    umur_laki: data.gender === 'L' ? data.umur : '',
                    umur_perempuan: data.gender === 'P' ? data.umur : '',
                    tindakan: data.tindakan,
                    icd_kode: data.icd_kode
                });

                // Beri warna pada sel Kode ICD (Kolom E / ke-5)
                const icdCell = row.getCell(6);
                if (data.icd_warna) {
                    icdCell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: data.icd_warna } // Format hex ExcelJS (ex: FF00FF00)
                    };
                }

                // Tambahkan border ke semua sel di baris ini
                row.eachCell((cell) => {
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="Rekap_Medis_Gigi_Terpilih.xlsx"');
        await workbook.xlsx.write(res);
        res.end();
    });
});

const PORT = Number(process.env.PORT) || 3000;
const bindAddress = isProduction ? '0.0.0.0' : '127.0.0.1';
initializeDatabase().then(() => {
    app.listen(PORT, bindAddress, () => {
        console.log(`Server berjalan di http://localhost:${PORT}`);
    });
}).catch(err => {
    console.error('Gagal menyiapkan database:', err.message);
    process.exitCode = 1;
});