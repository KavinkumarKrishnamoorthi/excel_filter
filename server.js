const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3000;

// Create upload & output directories if not exist
const uploadDir = path.join(__dirname, "uploads");
const outputDir = path.join(__dirname, "outputs");
[uploadDir, outputDir].forEach(dir => {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// Multer storage setup
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadDir),
    filename: (req, file, cb) => cb(null, Date.now() + path.extname(file.originalname))
});
const upload = multer({ storage });

// Convert column letter to index
function columnLetterToIndex(letter) {
    let col = 0;
    for (let i = 0; i < letter.length; i++) {
        col = col * 26 + letter.charCodeAt(i) - 64;
    }
    return col - 1;
}

// Delete columns with styles preserved
function deleteColumnsWithStyles(workbook, sheetName, columnsToDelete) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;

    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false });

    const colIndexes = columnsToDelete
        .split(",")
        .map(c => columnLetterToIndex(c.trim()))
        .sort((a, b) => b - a);

    // Remove columns from data
    data.forEach(row => {
        colIndexes.forEach(idx => {
            if (idx >= 0 && idx < row.length) row.splice(idx, 1);
        });
    });

    const newSheet = XLSX.utils.aoa_to_sheet(data);

    // Preserve styles
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            if (!colIndexes.includes(C)) {
                const oldAddr = XLSX.utils.encode_cell({ r: R, c: C });
                const newC = C - colIndexes.filter(ci => ci < C).length;
                const newAddr = XLSX.utils.encode_cell({ r: R, c: newC });

                if (worksheet[oldAddr] && newSheet[newAddr]) {
                    if (worksheet[oldAddr].s) {
                        newSheet[newAddr].s = worksheet[oldAddr].s;
                    }
                }
            }
        }
    }

    // Preserve column widths
    if (worksheet["!cols"]) {
        newSheet["!cols"] = worksheet["!cols"].filter((_, idx) => !colIndexes.includes(idx));
    }

    // Preserve merged cells
    if (worksheet["!merges"]) {
        const newMerges = worksheet["!merges"]
            .map(m => {
                const nm = { ...m };
                let remove = false;

                colIndexes.forEach(ci => {
                    if (ci >= m.s.c && ci <= m.e.c) remove = true;
                    if (ci < m.s.c) { nm.s.c--; nm.e.c--; }
                    else if (ci < m.e.c) nm.e.c--;
                });

                return remove ? null : nm;
            })
            .filter(Boolean);
        newSheet["!merges"] = newMerges;
    }

    workbook.Sheets[sheetName] = newSheet;
}

// Serve upload form
app.get("/", (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Excel File Upload</title>
        </head>
        <body>
            <h2>Upload Excel File</h2>
            <form id="uploadForm" enctype="multipart/form-data" method="POST" action="/upload">
                <input type="file" name="excel" accept=".xlsx,.xls" required />
                <button type="submit">Upload</button>
            </form>
        </body>
        </html>
    `);
});

// Upload & process file
app.post("/upload", upload.single("excel"), (req, res) => {
    try {
        if (!req.file) return res.status(400).send("No file uploaded");

        const deleteConfig = {
            "invoice": "G,K,O,P,Q,R,U,Y,AC,AD,AG,AH,AI,AJ,AK,AL,AM,AN,AR,AS",
            "invoice_ims": "G,K,O,P,Q,R,U,Y,AC,AD,AG,AH,AI,AJ,AK,AL,AM,AN,AR,AS",
            "note": "G,L,Q,R,S,V,AA,AF,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX",
            "note_ims": "G,L,Q,R,S,V,AA,AF,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX"
        };

        const workbook = XLSX.readFile(req.file.path, { cellStyles: true });

        for (const [sheetName, cols] of Object.entries(deleteConfig)) {
            const targetSheet = workbook.SheetNames.find(s => s.toLowerCase() === sheetName.toLowerCase());
            if (targetSheet) {
                deleteColumnsWithStyles(workbook, targetSheet, cols);
            }
        }

        const outputPath = path.join(outputDir, `processed_${req.file.originalname}`);
        XLSX.writeFile(workbook, outputPath, { cellStyles: true });

        fs.unlinkSync(req.file.path); // delete temp upload

        res.download(outputPath, err => {
            if (err) console.error("Download error:", err);
        });

    } catch (err) {
        console.error(err);
        res.status(500).send("Error processing file");
    }
});

app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
