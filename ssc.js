    async function exportToExcel() {
        if (!filteredData.length) { alert("No data to export!"); return;}

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("FilteredData");

        // Header row
        worksheet.columns = [
            { header: "SL", key: "sl", width: 6 },
            { header: "Roll", key: "roll", width: 12 },
            { header: "Trade Name", key: "trade", width: 20 },
            { header: "Status", key: "status", width: 12 },
            { header: "Institute", key: "institute", width: 30 },
            { header: "Room", key: "room", width: 10 }
        ];

        // Data rows
        filteredData.forEach((row, index) => {
            worksheet.addRow({ sl: index + 1,  roll: row.Roll, trade: row.Trade,  status: row.Type, institute: row.Institute,  room: row.Room || "N/A" });
        });

        worksheet.getRow(1).font = { bold: true };
        worksheet.getRow(1).alignment = { vertical: "middle", horizontal: "center" };

        // File name fix
        const dateStr = new Date().toISOString().split("T")[0];
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = `SSC_Export_${dateStr}.xlsx`;
        link.click();
    }

        // --- Seat Labels ---
    async function generateSeatLabels() {

        const table = document.getElementById("dataTable");
        const rows = table.querySelectorAll("tbody tr");
        let data = [];
        rows.forEach(row => {
            if (row.cells.length > 1) {
                data.push({ roll: row.cells[1].innerText,  trade: row.cells[2].innerText, type: row.cells[3].innerText.trim()});
            }
        });

        if (!data.length) return alert("No data!");
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("Seat Labels");

        sheet.columns = [
            { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, // label 1
            { width: 3 }, // gap
            { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, // label 2
            { width: 3 }, // gap
            { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 } // label 3
        ];
        let startRow = 1;

        for (let i = 0; i < data.length; i += 3) {
            for (let j = 0; j < 3; j++) {

                const item = data[i + j];
                if (!item) continue;

                let col = (j === 0) ? 1 : (j === 1 ? 6 : 11);

                // ===== Row 1: Title (Nikosh) =====
                sheet.mergeCells(startRow, col, startRow, col + 3);
                let c1 = sheet.getCell(startRow, col);
                c1.value = "এসএসসি (ভোক) পরীক্ষা - ২০২৬";
                styleHeader(c1);

                // ===== Row 2: Center (Nikosh) =====
                sheet.mergeCells(startRow + 1, col, startRow + 1, col + 3);
                let c2 = sheet.getCell(startRow + 1, col);
                c2.value = "কেন্দ্র: পীরগঞ্জ সরকারি টেকনিক্যাল স্কুল ও কলেজ";
                styleHeader(c2);

                // ===== Type (Arial) =====
                sheet.mergeCells(startRow + 2, col, startRow + 2, col + 1);
                let typeCell = sheet.getCell(startRow + 2, col);
                typeCell.value = item.type;
                styleEnglish(typeCell);

                // ===== Roll (Arial Big) =====
                sheet.mergeCells(startRow + 2, col + 2, startRow + 3, col + 3);
                let rollCell = sheet.getCell(startRow + 2, col + 2);
                rollCell.value = item.roll;
                rollCell.font = { name: "Arial", size: 20, bold: true };
                rollCell.alignment = centerAlign();

                // ===== Trade (Arial) =====
                sheet.mergeCells(startRow + 3, col, startRow + 3, col + 1);
                let tradeCell = sheet.getCell(startRow + 3, col);
                tradeCell.value = item.trade;
                styleEnglish(tradeCell);

                // ===== Border apply =====
                for (let r = 0; r < 4; r++) {
                    for (let c = 0; c < 4; c++) {
                        let cell = sheet.getCell(startRow + r, col + c);
                        applyBorder(cell);
                        cell.alignment = centerAlign();
                    }
                }
            }

            startRow += 6;
        }

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer]);
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "Seat_Labels.xlsx";
        link.click();
    }

    // ===== Alignment =====
    function centerAlign() { return { vertical: "middle", horizontal: "center", wrapText: true };}

    // ===== Bangla Header (Nikosh) =====
    function styleHeader(cell) {cell.font = { name: "Nikosh", bold: true, size: 14 };  cell.alignment = centerAlign();}

    // ===== English Font (Arial) =====
    function styleEnglish(cell) { cell.font = { name: "Arial", size: 11 }; cell.alignment = centerAlign(); }

    // ===== Border =====
    function applyBorder(cell) { cell.border = { top: { style: "thin" },  left: { style: "thin" },  bottom: { style: "thin" },  right: { style: "thin" }};}


    // --- Seat Plan Logic ---
    function generateSeatPlan() {
        if (!excelData.length) return alert("Upload database first!");
        
        document.getElementById('main-ui').style.display = 'none';
        const ui = document.getElementById('secondary-ui');
        ui.style.display = 'block';

        const grouped = excelData.reduce((acc, s) => {
            const r = s.Room || "Unassigned";
            if (!acc[r]) acc[r] = [];
            acc[r].push(s);
            return acc;
        }, {});

        roomList = Object.keys(grouped).sort();
        renderRoom(grouped);
    }

    // প্রতিষ্ঠানের জন্য ডাইনামিক কালার জেনারেটর
    function getInstColor(instName) {
        let hash = 0;
        for (let i = 0; i < instName.length; i++) {
            hash = instName.charCodeAt(i) + ((hash << 5) - hash);
        }
        // হালকা কালার (Pastel colors) জেনারেট করার জন্য hsl ব্যবহার করা হয়েছে
        return `hsl(${hash % 360}, 70%, 92%)`;
    }

    function renderRoom(data) {
        const roomName = roomList[currentRoomIndex];
        // প্রথমে রোল অনুযায়ী ছোট থেকে বড় (Ascending) সাজিয়ে নেয়া
        const students = data[roomName].sort((a, b) => a.Roll - b.Roll);
        
        let tableHTML = "";
        const rowsPerCol = 10;
        const totalCols = 6;

        for (let r = 0; r < rowsPerCol; r++) {
            tableHTML += "<tr>";
            for (let c = 0; c < totalCols; c++) {
                const idx = c * rowsPerCol + r;
                const student = students[idx];
                
                if (student) {
                    const bgColor = getInstColor(student.Institute);
                    const instShort = student.Institute.substring(0, 3).toUpperCase();
                    const serialNo = idx + 1; // রুমে কততম ছাত্র

                    tableHTML += `
                        <td style="border:1px solid black; padding:5px; background-color: ${bgColor}; height:55px; vertical-align: middle;">
                            <div style="font-size: 1.1rem; font-weight: bold; line-height: 1;">${student.Roll}</div>
                            <div style="font-size: 0.65rem; color: #444; margin-top: 3px; font-weight: normal;">
                                #${serialNo} | ${instShort}
                            </div>
                        </td>`;
                } else {
                    tableHTML += `<td style="border:1px solid black; background: #fff;"></td>`;
                }
            }
            tableHTML += "</tr>";
        }

        const instSum = students.reduce((acc, s) => { acc[s.Institute] = (acc[s.Institute] || 0)+1; return acc; }, {});
        const summary = Object.entries(instSum).map(([n, c]) => `${n}-${c}`).join(" + ");

        document.getElementById('secondary-ui').innerHTML = `
            <div class="p-4 text-center">
                <div class="no-print mb-4 d-flex justify-content-center gap-3">
                    <button class="btn btn-dark" onclick="changeRoom(-1, ${JSON.stringify(data).replace(/"/g, '&quot;')})"><i class="fa fa-chevron-left"></i> Prev Room</button>
                    <span class="badge bg-primary fs-5 px-4">Room: ${roomName}</span>
                    <button class="btn btn-dark" onclick="changeRoom(1, ${JSON.stringify(data).replace(/"/g, '&quot;')})">Next Room <i class="fa fa-chevron-right"></i></button>
                </div>
                <div class="seat-plan-wrapper" style="background:white; border:2px solid #000; padding:30px; max-width:950px; margin:auto; text-align:center;">
                    <h5 class="fw-bold">এসএসসি (ভোক.) বোর্ড ফাইনাল পরীক্ষা ২০২৬ এর আসন বিন্যাস</h5>
                    <p class="mb-1">কেন্দ্র: পীরগঞ্জ সরকারি টেকনিক্যাল স্কুল ও কলেজ, ঠাকুরগাঁও।</p>
                    <h4 class="fw-bold py-2" style="background:#f1f1f1; border:1px solid black; margin-bottom:10px;">কক্ষ নং- ${roomName}</h4>
                    <div class="d-flex justify-content-between px-4 small fw-bold"><span>শিক্ষক টেবিল</span><span>দরজা ⬅</span></div>
                    <table style="width:100%; border-collapse:collapse; margin:10px 0;">${tableHTML}</table>
                    <div class="d-flex justify-content-between px-4 small fw-bold"><span>${summary}</span><span>দরজা ⬅</span></div>
                    <div class="border-top mt-3 pt-2 fw-bold text-start">মোট পরীক্ষার্থী সংখ্যা = ${students.length} জন।</div>
                </div>
            </div>`;
    }

    function changeRoom(dir) {
        currentRoomIndex = (currentRoomIndex + dir + roomList.length) % roomList.length;
        generateSeatPlan();
    }

    function instituteSummary() {
        const instQ = document.getElementById('instSearch').value;
        if (!excelData || excelData.length === 0) { Swal.fire({icon: "warning", title: "Oops!", text: "Please! upload excel file first."}); return;}
        if (instQ === "") { Swal.fire({icon: "warning", title: "Oops!", text: "অনুগ্রহ করে একটি Institute select করুন!"}); return;}

        const filteredData = excelData.filter(row => String(row.Institute || "").trim() === instQ.trim());
        const summary = {};
        const subjects = ["Bangla", "English", "Math", "BDGS", "Physics", "Chemistry", "Self Emp", "Religion", "Trade1", "Trade2"];

        // সকল কলামের সমষ্টি রাখার জন্য grandTotal অবজেক্ট
        let grandTotal = {
            total: 0, regular: 0, irregular: 0,
            Bangla: 0, English: 0, Math: 0, BDGS: 0, 
            Physics: 0, Chemistry: 0, "Self Emp": 0, 
            Religion: 0, Trade1: 0, Trade2: 0
        };

        filteredData.forEach(row => {
            const trade = String(row.Trade || "Unknown Trade").trim();
            if (!summary[trade]) {
                summary[trade] = { regular: 0, irregular: 0 };
                subjects.forEach(sub => summary[trade][sub] = 0);
            }
            const type = String(row.Type || "").trim().toLowerCase();
            if (type === 'regular') {
                summary[trade].regular++;
                grandTotal.regular++;
                grandTotal.total++;
            } else if (type === 'irregular') {
                summary[trade].irregular++;
                grandTotal.irregular++;
                grandTotal.total++;
                
                subjects.forEach(sub => {
                    if (row[sub] != null && (row[sub] == 1 || String(row[sub]).trim() == "1")) {
                        summary[trade][sub]++;
                        grandTotal[sub]++; 
                    }
                });
            }
        });

        document.getElementById('main-ui').style.display = 'none';
        const summaryUI = document.getElementById('secondary-ui');
        summaryUI.style.display = 'block';

        let tableHTML = `
            <div class="p-4">
                <div class="text-center mb-4 no-print">
                    <h4 class="fw-bold">Institute Name: ${instQ}</h4>
                    <p class="text-muted">Students Statistics Summary</p>
                </div>
            
                    <div class="table-responsive">
                        <table class="table table-bordered align-middle text-center">
                            <thead>
                                <tr class="table-primary">
                                    <th rowspan="2">SL</th>
                                    <th rowspan="2">Trade</th>
                                    <th rowspan="2">Total</th>
                                    <th rowspan="2">Regular</th>
                                    <th rowspan="2">Irregular</th>
                                    <th colspan="10">অনিয়মিত পরীক্ষার্থী সংখ্যা</th>
                                </tr>
                                <tr class="table-primary"> ${subjects.map(sub => `<th>${sub}</th>`).join('')}</tr>
                            </thead>
                            <tbody>`;

        let sl = 1;
        for (let trade in summary) {
            const d = summary[trade];
            const rowTotal = d.regular + d.irregular;
            tableHTML += `
                <tr><td>${sl++}</td>
                    <td class="text-start fw-bold">${trade}</td>
                    <td class="bg-light">${rowTotal}</td>
                    <td class="text-success">${d.regular}</td>
                    <td class="text-danger">${d.irregular}</td>
                    ${subjects.map(sub => `<td>${d[sub] > 0 ? d[sub] : '-'}</td>`).join('')}
                </tr>`;
        }

        tableHTML += `
            </tbody>
            <tfoot class="table-secondary fw-bold">
                <tr>
                    <td colspan="2" class="text-end">Grand Total:</td>
                    <td>${grandTotal.total}</td>
                    <td class="text-success">${grandTotal.regular}</td>
                    <td class="text-danger">${grandTotal.irregular}</td>
                    ${subjects.map(sub => `<td>${grandTotal[sub]}</td>`).join('')}
                </tr>
            </tfoot>
        </table>`;

        tableHTML += `</div>
        <button class="btn btn-secondary mt-3 no-print" onclick="showUI('main-ui')"><i class="fa-solid fa-arrow-left me-2"></i>Back to Dashboard </button>
        </div>`;

        summaryUI.innerHTML = tableHTML;
    }

    function generatePracticalList() {
        if (!filteredData.length) { alert("ড্যাশবোর্ডে কোন ডাটা নেই! আগে ডাটা ফিল্টার করুন।"); return;}
        const practicalSubs = ["Physics", "Chemistry", "Trade1", "Trade2"];
        // Practical students filter
        const practicalStudents = filteredData.filter(row => {
            const type = String(row.Type || "").trim().toLowerCase();
            if (type === 'regular') return true;
            if (type === 'irregular') {
                return practicalSubs.some(sub =>
                    row[sub] != null &&
                    (row[sub] == 1 || String(row[sub]).trim() == "1")
                );
            }
            return false;
        });

        // ৩. UI জেনারেশন
        document.getElementById('main-ui').style.display = 'none';
        const displayArea = document.getElementById('secondary-ui');
        displayArea.style.display = 'block';

        let tableHTML = `
            <div class="p-4">
                <div class="d-flex justify-content-between align-items-center mb-2">
                    <h6 class="fw-bold text-primary">
                        <i class="fa-solid fa-flask-vial me-2"></i>
                        ব্যবহারিক পরীক্ষার্থীর তালিকা (ফিল্টার অনুযায়ী)
                    </h6>
                    <span class="badge bg-primary px-4 py-2">
                        Total: ${practicalStudents.length}
                    </span>
                </div>

                <div class="data-card shadow-sm border-0">
                    <div class="table-responsive">
                        <table class="table table-bordered align-middle text-center">
                            <thead class="table-dark">
                                <tr>
                                    <th style="width: 80px;">SL</th>
                                    <th>Roll No</th>
                                    <th>Trade Name</th>
                                    <th>Status</th>
                                    <th>Institute Name</th>
                                </tr>
                            </thead>
                            <tbody>
        `;

        practicalStudents.forEach((row, index) => {
            const type = String(row.Type || "").trim();
            tableHTML += `
                <tr>
                    <td>${index + 1}</td>
                    <td class="fw-bold">${row.Roll}</td>
                    <td>${row.Trade}</td>
                    <td>
                        <span class="badge ${
                            type.toLowerCase() === 'regular' ? 'badge-reg' : 'badge-irreg'
                        } rounded-pill px-3">
                            ${type}
                        </span>
                    </td>
                    <td class="text-secondary small">${row.Institute}</td>
                </tr>
            `;
        });

        // যদি কোনো ডাটা না থাকে
        if (practicalStudents.length === 0) {
            tableHTML += `
                <tr>
                    <td colspan="5" class="py-5 text-muted">
                        কোন ব্যবহারিক পরীক্ষার্থী পাওয়া যায়নি
                    </td>
                </tr>
            `;
        }
        tableHTML += `
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="no-print mt-4">
                    <button class="btn btn-outline-secondary me-2" onclick="showUI('main-ui')">
                        <i class="fa-solid fa-arrow-left me-2"></i>
                        Back to Filtered List
                    </button>
                </div>
            </div>
        `;
        displayArea.innerHTML = tableHTML;
    }

    function calculateInstituteFee() {
        const instQ = document.getElementById('instSearch').value;

        if (!excelData || excelData.length === 0) { Swal.fire({icon: "warning", title: "Oops!", text: "Please upload Excel file first."}); return;}
        if (!instQ) { Swal.fire({ icon: "warning", title: "Missing Selection", text: "অনুগ্রহ করে একটি Institute select করুন!"}); return; }

        const data = excelData.filter(row => row.Institute === instQ);
        if (data.length === 0) {Swal.fire({ icon: "error", title: "No Data Found", text: "এই Institute-এর কোনো ডাটা পাওয়া যায়নি!"}); return;}

        // ===== Fee Setup (SINGLE CENTER FEE FOR ALL) =====
        const FEES = {CENTER: 550, PRAC_CENTER: 150, PRAC_PER_SUB: 35, PRACTICAL_SUBJECTS: ["Physics", "Chemistry", "Trade1", "Trade2"]};

        const getPracCount = (row) => {
            let count = 0;
            FEES.PRACTICAL_SUBJECTS.forEach(sub => {if (row[sub] == 1) count++; });
            return count;
        };

        // ===== Regular =====
        const regStudents = data.filter(r => String(r.Type).toLowerCase() === "regular");
        const regCount = regStudents.length;

        const regTotal = regCount * (FEES.CENTER + FEES.PRAC_CENTER + (6 * FEES.PRAC_PER_SUB)
        );

        // ===== Irregular without practical =====
        const irregNoPrac = data.filter(r => {
            const type = String(r.Type).toLowerCase();
            return type === "irregular" && getPracCount(r) === 0;
        });

        const irregNoPracCount = irregNoPrac.length;
        const irregNoPracTotal = irregNoPracCount * FEES.CENTER;

        // ===== Irregular with practical =====
        const irregWithPrac = data.filter(r => {
            const type = String(r.Type).toLowerCase();
            return type === "irregular" && getPracCount(r) > 0;
        });

        let summary = {1:0, 2:0, 3:0, 4:0};

        irregWithPrac.forEach(r => { const c = getPracCount(r); if (c > 0) summary[c]++; });

        let rows = "";
        let irregTotal = 0;

        [1,2,3,4].forEach(num => {
            const count = summary[num]; if (count === 0) return;
            const rowTotal = count * (FEES.CENTER + FEES.PRAC_CENTER + (num * FEES.PRAC_PER_SUB));
            irregTotal += rowTotal;

            rows += `
                <tr>
                    <td>${num}</td>
                    <td>${count}</td>
                    <td>${FEES.CENTER}</td>
                    <td>${FEES.PRAC_CENTER}</td>
                    <td>${num * FEES.PRAC_PER_SUB}</td>
                    <td>${rowTotal}</td>
                </tr>`;
        });

        if (!rows) { rows = `<tr><td colspan="6" class="text-muted">No practical students</td></tr>`;}

        const grandTotal = regTotal + irregNoPracTotal + irregTotal;
        document.getElementById('main-ui').style.display = 'none';
        const display = document.getElementById('secondary-ui');
        display.style.display = 'block';

        display.innerHTML = `
            <div class="p-5 bg-white shadow mx-auto" style="max-width:950px">

                <div class="text-center mb-4">
                    <h4>${instQ}</h4>
                    <h5 class="text-secondary">Center Fee Statement</h5>
                </div>

                <h6>1. Regular Students</h6>
                <table class="table table-bordered text-center">
                    <tr>
                        <th>Students</th><th>Center</th><th>Prac Center</th>
                        <th>Subjects</th><th>Prac Fee</th><th>Total</th>
                    </tr>
                    <tr>
                        <td>${regCount}</td>
                        <td>${FEES.CENTER}</td>
                        <td>${FEES.PRAC_CENTER}</td>
                        <td>6</td>
                        <td>${6 * FEES.PRAC_PER_SUB}</td>
                        <td>${regTotal}</td>
                    </tr>
                </table>

                <h6 class="mt-4">2. Irregular Students</h6>

                <p>Without Practical</p>
                <table class="table table-bordered text-center">
                    <tr><th>Students</th><th>Center</th><th>Total</th></tr>
                    <tr>
                        <td>${irregNoPracCount}</td>
                        <td>${FEES.CENTER}</td>
                        <td>${irregNoPracTotal}</td>
                    </tr>
                </table>

                <p class="mt-4">With Practical</p>
                <table class="table table-bordered text-center">
                    <tr>
                        <th>Subjects</th><th>Students</th><th>Center</th>
                        <th>Prac Center</th><th>Prac Fee</th><th>Total</th>
                    </tr>
                    ${rows}
                </table>

                <div class="text-end mt-4"><h5>Grand Total = ${grandTotal} /-</h5></div>

                <div class="mt-4 no-print">
                    <button class="btn btn-secondary" onclick="showUI('main-ui')"> <i class="fa-solid fa-arrow-left me-2"></i>Back to Dashboard</button>
                    <button class="btn btn-primary" onclick="downloadInstituteFeeExcel()"> <i class="fa-solid fa-download me-2"></i>Download Institute Fee (Excel)</button>
                </div>

            </div>`;
    }

    async function downloadInstituteFeeExcel() {
        const instQ = document.getElementById('instSearch').value;
        const data = excelData.filter(row => row.Institute === instQ);
        const FEES = {CENTER: 550, PRAC_CENTER: 150, PRAC_PER_SUB: 35, PRACTICAL_SUBJECTS: ["Physics", "Chemistry", "Trade1", "Trade2"]};

        const getPracCount = (row) => {
            let count = 0;
            FEES.PRACTICAL_SUBJECTS.forEach(sub => {if (row[sub] == 1) count++; });
            return count;
        };

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("Institute Fee");
        sheet.addRow([ "SL", "Roll", "Type", "Center Fee", "Prac Center Fee", "Practical Subjects", "Practical Fee", "Total Fee"]);

        let sl = 1;
        let grandTotal = 0;
        data.forEach(row => {
            const type = String(row.Type).toLowerCase();
            const prac = getPracCount(row);

            let center = FEES.CENTER;
            let pracCenter = 0;
            let pracFee = 0;

            if (type === "regular") { pracCenter = FEES.PRAC_CENTER; pracFee = 6 * FEES.PRAC_PER_SUB;}
            if (type === "irregular" && prac > 0) { pracCenter = FEES.PRAC_CENTER; pracFee = prac * FEES.PRAC_PER_SUB;}

            const total = center + pracCenter + pracFee;
            grandTotal += total;

            sheet.addRow([ sl++, row.Roll, row.Type, center, pracCenter, prac, pracFee, total]);
        });

        sheet.addRow([]);
        sheet.addRow(["", "", "", "", "", "", "", "Grand Total", grandTotal]);

        sheet.columns.forEach(c => c.width = 18);
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});

        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = `${instQ}_Institute_Fee.xlsx`;
        a.click();

        Swal.fire({icon: "success", title: "Downloaded", text: "Excel file ready!"});
    }

    async function makeRoomSeatPlan() {
        if (!filteredData.length) {
            alert("No data found!");
            return;
        }

        if (filteredData.length > 60) {
            alert("Maximum 60 students allowed for one room!");
            return;
        }

        let useColumns = false;
        const { value: formValues } = await Swal.fire({
            title: 'Room তথ্য দিন',
            html: `
            <div class="container-fluid text-start">

                <div class="mb-3">
                    <label class="form-label fw-semibold">🏢 তলা (Floor)</label>
                    <input id="floor" type="number" class="form-control" placeholder="e.g. 3">
                </div>

                <div class="mb-3">
                    <label class="form-label fw-semibold">🚪 রুম নম্বর</label>
                    <input id="room" type="text" class="form-control" placeholder="e.g. 101">
                </div>

                ${filteredData.length <= 30 ? `
                <div class="mb-3">
                    <label class="form-label fw-semibold">📊 কলাম সেট</label>
                    <select id="columnSet" class="form-select">
                        <option value="">Select column type</option>
                        <option value="all">All Columns</option>
                        <option value="odd">Odd Columns (1,3,5)</option>
                        <option value="even">Even Columns (2,4,6)</option>
                    </select>
                </div>
                ` : ''}

            </div>
            `,
            focusConfirm: false,
            showCancelButton: true,
            confirmButtonText: 'Generate',
            preConfirm: () => {
                const floor = document.getElementById('floor').value;
                const room = document.getElementById('room').value;
                const colSet = document.getElementById('columnSet')?.value;

                if (!floor || !room) {
                    Swal.showValidationMessage('তলা এবং রুম নম্বর দিতে হবে!');
                    return false;
                }

                if (filteredData.length <= 30 && !colSet) {
                    Swal.showValidationMessage('কলাম সেট নির্বাচন করুন!');
                    return false;
                }

                return { floor, room, colSet };
            }
        });

        if (!formValues) return;

        const { floor, room, colSet } = formValues;

        if (filteredData.length <= 30) {
            if (colSet === "all") useColumns = [0,1,2,3,4,5];
            else if (colSet === "odd") useColumns = [0,2,4];
            else if (colSet === "even") useColumns = [1,3,5];
        }

        // ✅ Sort
        const students = [...filteredData].sort((a, b) => Number(a.Roll) - Number(b.Roll));

        const rowsPerCol = 10;
        const totalCols = 6;

        let tableHTML = "";

        for (let r = 0; r < rowsPerCol; r++) {
            tableHTML += "<tr>";

            for (let c = 0; c < totalCols; c++) {

                let idx;

                if (filteredData.length <= 30) {
                    // selected columns only
                    const colIndex = useColumns.indexOf(c);
                    if (colIndex === -1) {
                        tableHTML += `<td class="seat-cell empty"></td>`;
                        continue;
                    }
                    idx = colIndex * rowsPerCol + r;
                } else {
                    idx = c * rowsPerCol + r;
                }

                const student = students[idx];

                if (student) {
                    const instShort = (student.Institute || "").substring(0, 3).toUpperCase();
                    const serialNo = idx + 1;

                    tableHTML += `
                        <td class="seat-cell">
                            <div class="roll">${student.Roll}</div>
                            <div class="meta">#${serialNo} | ${instShort}</div>
                        </td>`;
                } else {
                    tableHTML += `<td class="seat-cell empty"></td>`;
                }
            }

            tableHTML += "</tr>";
        }

        // ✅ Summary
        const instSum = students.reduce((acc, s) => {
            acc[s.Institute] = (acc[s.Institute] || 0) + 1;
            return acc;
        }, {});

        const summary = Object.entries(instSum)
            .map(([n, c]) => `${n}-${c}`)
            .join(" + ");

        // UI
        document.getElementById('main-ui').style.display = 'none';
        const container = document.getElementById('secondary-ui');
        container.style.display = 'block';

        container.innerHTML = `
            <div class="p-3 text-center">

                <div class="no-print mb-3 d-flex justify-content-center gap-2">
                    <button class="btn btn-dark" onclick="showUI('main-ui')">⬅ Back</button>
                <button class="btn btn-success" onclick="downloadPDF()">📄 Download PDF</button>
                </div>

                <div id="printArea" class="seat-plan-wrapper">

                    <div class="seat-title">
                        এসএসসি (ভোক.) বোর্ড ফাইনাল পরীক্ষা ২০২৬ এর আসন বিন্যাস
                    </div>

                    <div class="seat-subtitle">
                        কেন্দ্র: পীরগঞ্জ সরকারি টেকনিক্যাল স্কুল ও কলেজ, ঠাকুরগাঁও।
                    </div>

                    <div class="header-box">
                        কক্ষ নং- ${room} (${floor} তলা)
                    </div>

                    <div class="top-line">
                        <span>শিক্ষক টেবিল</span>
                        <span>দরজা ⬅</span>
                    </div>

                    <table class="seat-table">
                        <colgroup>
                            <col><col><col><col><col><col>
                        </colgroup>
                        ${tableHTML}
                    </table>

                    <div class="bottom-line">
                        <span>${summary}</span>
                        <span>দরজা ⬅</span>
                    </div>

                    <div class="footer">
                        মোট পরীক্ষার্থী সংখ্যা = ${students.length} জন।
                    </div>

                </div>
            </div>
        `;
    }

    function downloadPDF() {
        const element = document.getElementById("printArea");

        html2pdf().set({
            margin: 0.2,
            filename: "seat-plan.pdf",

            image: { type: "jpeg", quality: 1 },

            html2canvas: {
                scale: 2,
                useCORS: true
            },

            jsPDF: {
                unit: "mm",
                format: "a4",
                orientation: "portrait"
            },

            pagebreak: {
                mode: ["avoid-all"]
            }
        }).from(element).save();
    }

    async function downloadSortedRolls() {
        if (!filteredData.length) { Swal.fire({  icon: 'warning', title: 'No Data', text: 'No filtered data to export!' }); return;}
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Roll List');
        const rowsPerColumn = 50;
        const maxBlocksPerRow = 4;

        let col = 1;
        let row = 2;
        let baseRow = 2;
        let blockCount = 0;

        //First header
        sheet.getCell(1, 1).value = "SL";
        sheet.getCell(1, 2).value = "Roll";
        sheet.getCell(1, 3).value = "Room";
        sheet.getColumn(1).width = 4;
        sheet.getColumn(2).width = 10;
        sheet.getColumn(3).width = 6;

        //Sort data
        const sortedData = [...filteredData].sort((a, b) => a.Roll - b.Roll);
        sortedData.forEach((item, index) => {
            const sl = index + 1;
            sheet.getCell(row, col).value = sl;
            sheet.getCell(row, col + 1).value = item.Roll;
            sheet.getCell(row, col + 2).value = item.Room || "";
            row++;

            if (row > baseRow + rowsPerColumn - 1) {
                blockCount++;

                // 4 টি ব্লক পূর্ণ হলে নিচে নাম্ববে
                if (blockCount % maxBlocksPerRow === 0) {
                    baseRow += rowsPerColumn + 2; // Two row gap
                    row = baseRow;
                    col = 1;
                } else { col += 3; row = baseRow;}

                // নতুন header বসানো
                sheet.getCell(baseRow - 1, col).value = "SL";
                sheet.getCell(baseRow - 1, col + 1).value = "Roll";
                sheet.getCell(baseRow - 1, col + 2).value = "Room";
                sheet.getColumn(col).width = 6;
                sheet.getColumn(col + 1).width = 15;
                sheet.getColumn(col + 2).width = 10;
            }
        });

        // DOWNLOAD
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "Roll_List.xlsx";
        link.click();
    }


    let lastBackup = null;
    function resetDatabase() {
        if (!excelData.length) {
            Swal.fire({
                icon: 'info',
                title: 'No Data',
                text: 'Database is already empty!'
            });
            return;
        }

        Swal.fire({
            title: 'Reset Database?',
            html: 'You can <b>download backup</b> before reset.<br>Also you can undo within 5 seconds!',
            icon: 'warning',
            showCancelButton: true,
            showDenyButton: true,
            confirmButtonText: 'Reset',
            denyButtonText: 'Download Backup',
            cancelButtonText: 'Cancel',
            confirmButtonColor: '#d33'
        }).then((result) => {

            // 🔽 Backup Download
            if (result.isDenied) {
                const dataStr = JSON.stringify(excelData, null, 2);
                const blob = new Blob([dataStr], { type: "application/json" });
                const url = URL.createObjectURL(blob);

                const a = document.createElement("a");
                a.href = url;
                a.download = "ssc_backup.json";
                a.click();

                URL.revokeObjectURL(url);
                return;
            }

            // 🔥 Reset
            if (result.isConfirmed) {

                // Save backup for undo
                lastBackup = [...excelData];

                // Clear storage
                localStorage.removeItem("sscData");
                excelData = [];
                filteredData = [];
                roomList = [];

                // Reset UI
                document.getElementById('tableBody').innerHTML = `
                    <tr>
                        <td colspan="6" class="text-center py-5 text-muted">
                            <i class="fa-solid fa-file-excel fa-3x mb-3 text-secondary opacity-50"></i>
                            <p class="mb-0 fw-bold">Please upload an Excel database ("All" sheet)</p>
                        </td>
                    </tr>
                `;

                document.getElementById('stats').innerText = 'Total: 0';
                document.getElementById('tradeSearch').innerHTML = '<option value="">All Trades</option>';
                document.getElementById('instSearch').innerHTML = '<option value="">All Institutes</option>';
                document.getElementById('excelFile').value = "";

                const icon = document.getElementById('uploadIcon');
                icon.classList.replace('text-success', 'text-danger');

                // 🔁 Undo option (5 sec)
                let timerInterval;
                Swal.fire({
                    title: 'Database Reset!',
                    html: 'Undo available for <b>5</b> seconds...',
                    timer: 5000,
                    timerProgressBar: true,
                    showConfirmButton: true,
                    confirmButtonText: 'Undo',
                    didOpen: () => {
                        const b = Swal.getHtmlContainer().querySelector('b');
                        timerInterval = setInterval(() => {
                            b.textContent = Math.ceil(Swal.getTimerLeft() / 1000);
                        }, 200);
                    },
                    willClose: () => {
                        clearInterval(timerInterval);
                    }
                }).then((res) => {
                    if (res.isConfirmed && lastBackup) {

                        // Restore backup
                        excelData = [...lastBackup];
                        localStorage.setItem("sscData", JSON.stringify(excelData));

                        autoPopulateDropdowns();
                        applyFilters();

                        document.getElementById('uploadIcon')
                            .classList.replace('text-danger', 'text-success');

                        Swal.fire({
                            icon: 'success',
                            title: 'Restored!',
                            text: 'Database has been recovered.'
                        });
                    }
                });
            }
        });
    }
