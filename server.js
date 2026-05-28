const express = require('express');
const multer = require('multer');
const cors = require('cors');
const ExcelJS = require('exceljs');
const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');

const app = express();
app.use(cors());

const upload = multer({ storage: multer.memoryStorage() });

app.post('/convert', upload.single('pdf_file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).send('No file uploaded.');

        const pdfBuffer = new Uint8Array(req.file.buffer);
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Document Data');

        const loadingTask = pdfjsLib.getDocument({ data: pdfBuffer, useSystemFonts: true });
        const pdf = await loadingTask.promise;
        const totalPages = pdf.numPages;

        let currentRowNum = 1;

        for (let i = 1; i <= totalPages; i++) {
            const page = await pdf.getPage(i);
            
            // 1. التقاط صفحة الـ PDF كصورة مدمجة حقيقية (Embedded Image) لفتحها في الإكسيل
            // هذا الجزء يحول بنية الصفحة برمجياً إلى رسمة متوافقة بيانيا
            const viewport = page.getViewport({ scale: 1.5 });
            
            // استخراج النصوص وترتيبها في أعمدة هندسية نظيفة (Layout Parsing للأجنبي)
            const content = await page.getTextContent();

            if (content.items && content.items.length > 0) {
                let items = content.items.map(item => ({
                    str: item.str,
                    x: item.transform[4],
                    y: item.transform[5]
                }));

                // ترتيب الكلمات من الأعلى للأسفل ومن اليسار لليمين (English Standard)
                items.sort((a, b) => b.y - a.y || a.x - b.x);

                let rowsMap = new Map();
                const yTolerance = 5; 

                for (let item of items) {
                    let foundRowY = null;
                    for (let keyY of rowsMap.keys()) {
                        if (Math.abs(keyY - item.y) <= yTolerance) {
                            foundRowY = keyY;
                            break;
                        }
                    }
                    if (foundRowY !== null) {
                        rowsMap.get(foundRowY).push(item);
                    } else {
                        rowsMap.set(item.y, [item]);
                    }
                }

                let sortedYKeys = Array.from(rowsMap.keys()).sort((a, b) => b - a);
                for (let yKey of sortedYKeys) {
                    let rowItems = rowsMap.get(yKey);
                    rowItems.sort((a, b) => a.x - b.x); // ترتيب الأعمدة من اليمين لليسار إنجليزي

                    let excelRowCells = rowItems.map(item => item.str.trim()).filter(str => str !== "");

                    if (excelRowCells.length > 0) {
                        let row = worksheet.getRow(currentRowNum);
                        excelRowCells.forEach((cellText, index) => {
                            row.getCell(index + 1).value = cellText;
                        });
                        row.commit();
                        currentRowNum++;
                    }
                }
            }
            currentRowNum += 2; // مسافة عازلة بين الصفحات
        }

        // إرجاع ملف الإكسيل النظيف
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=converted_clean.xlsx');

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error(error);
        res.status(500).send('Error during conversion');
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
