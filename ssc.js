    // --- Export Logic ---
    function exportToExcel() {
        const table = document.getElementById("dataTable");
        const bodyRows = table.querySelectorAll("tbody tr");
        let exportData = [["SL", "Roll", "Trade Name", "Status", "Institute", "Room"]];

        bodyRows.forEach(row => {
            if (row.cells.length > 1) {
                let rowData = Array.from(row.cells).map(cell => cell.innerText);
                exportData.push(rowData);
            }
        });

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(exportData);
        XLSX.utils.book_append_sheet(wb, ws, "FilteredData");
        XLSX.writeFile(wb, `SSC_Export_${new Date().toLocaleDateString()}.xlsx`);
    }

    // --- Seat Labels ---
    function generateSeatLabels() {
        const table = document.getElementById("dataTable");
        const bodyRows = table.querySelectorAll("tbody tr");
        let currentData = [];

        bodyRows.forEach(row => {
            if (row.cells.length > 1) {
                currentData.push({ roll: row.cells[1].innerText, trade: row.cells[2].innerText, type: row.cells[3].innerText });
            }
        });

        if (!currentData.length) return alert("No data to labels!");

        let ws_data = [];
        for (let i = 0; i < currentData.length; i += 3) {
            let r1=[], r2=[], r3=[], r4=[], r5=["","","","","",""];
            for (let j=0; j<3; j++) {
                if (currentData[i+j]) {
                    const item = currentData[i+j];
                    r1.push("এসএসসি (ভোক) সমাপনী পরীক্ষা - ২০২৬", "");
                    r2.push("কেন্দ্র: পীরগঞ্জ সরকারি টেকনিক্যাল স্কুল ও কলেজ", "");
                    r3.push(item.type, item.roll);
                    r4.push(item.trade, "");
                }
            }
            ws_data.push(r1, r2, r3, r4, r5);
        }
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(ws_data);
        XLSX.utils.book_append_sheet(wb, ws, "Labels");
        XLSX.writeFile(wb, "Seat_Labels.xlsx");
    }

    // --- Seat Plan Logic ---
    function generateSeatPlan() {
        if (!excelData.length) return alert("Upload database first!");
        
        document.getElementById('main-ui').style.display = 'none';
        const ui = document.getElementById('seat-plan-ui');
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

        document.getElementById('seat-plan-ui').innerHTML = `
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

function showSummary() {
    const instQ = document.getElementById('instSearch').value;
    if (!excelData || excelData.length === 0) {
        alert("Please upload excel first!");
        return;
    }
    if (instQ === "") {
        alert("অনুগ্রহ করে আগে একটি Institute সিলেক্ট করুন!");
        return;
    }

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
                    grandTotal[sub]++; // গ্র্যান্ড টোটালে যোগ করা হচ্ছে
                }
            });
        }
    });

    document.getElementById('main-ui').style.display = 'none';
    const summaryUI = document.getElementById('seat-plan-ui');
    summaryUI.style.display = 'block';

    let tableHTML = `
        <div class="p-4">
            <div class="text-center mb-4 no-print">
                <h4 class="fw-bold">Institute Name: ${instQ}</h4>
                <p class="text-muted">Students Statistics Summary</p>
            </div>
            <div class="data-card shadow-sm">
                <div class="table-responsive">
                    <table class="table table-bordered align-middle text-center">
                        <thead class="table-dark">
                            <tr>
                                <th rowspan="2">SL</th>
                                <th rowspan="2">Trade</th>
                                <th rowspan="2">Total</th>
                                <th rowspan="2">Regular</th>
                                <th rowspan="2">Irregular</th>
                                <th colspan="10">অনিয়মিত বিষয় (সংখ্যা)</th>
                            </tr>
                            <tr>
                                ${subjects.map(sub => `<th>${sub}</th>`).join('')}
                            </tr>
                        </thead>
                        <tbody>`;

    let sl = 1;
    for (let trade in summary) {
        const d = summary[trade];
        const rowTotal = d.regular + d.irregular;
        tableHTML += `
            <tr>
                <td>${sl++}</td>
                <td class="text-start fw-bold">${trade}</td>
                <td class="bg-light">${rowTotal}</td>
                <td class="text-success">${d.regular}</td>
                <td class="text-danger">${d.irregular}</td>
                ${subjects.map(sub => `<td>${d[sub] > 0 ? d[sub] : '-'}</td>`).join('')}
            </tr>`;
    }

    // --- এখানে সমষ্টির Row (Footer) যোগ করা হয়েছে ---
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

    tableHTML += `</div></div>
    <button class="btn btn-secondary mt-3 no-print" onclick="backToMain()">
        <i class="fa-solid fa-arrow-left me-2"></i>Back to Dashboard
    </button>
    </div>`;

    summaryUI.innerHTML = tableHTML;
}


function backToMain() {
    // সামারি বা সিট প্ল্যান হাইড করুন
    document.getElementById('seat-plan-ui').style.display = 'none';
    // মেইন ড্যাশবোর্ড বা টেবিল শো করুন
    document.getElementById('main-ui').style.display = 'block';
}

function generatePracticalList() {
    // ১. মেইন টেবিলের বর্তমান বডি থেকে ডাটা নেওয়া (ফিল্টার করা ডাটা)
    const tableRows = document.querySelectorAll("#tableBody tr");
    
    // যদি টেবিলে কোন ডাটা না থাকে বা ডিফল্ট মেসেজ থাকে
    if (tableRows.length === 0 || tableRows[0].cells.length < 2) {
        alert("ড্যাশবোর্ডে কোন ডাটা নেই! আগে ডাটা ফিল্টার করুন।");
        return;
    }

    let practicalStudents = [];

    // ২. বর্তমান ফিল্টার করা ডাটা থেকে ব্যবহারিক পরীক্ষার্থী খুঁজে বের করা
    // আমরা সরাসরি excelData থেকে চেক করবো কিন্তু কেবল সেই রোলগুলোকে নিবো যা ড্যাশবোর্ডে দৃশ্যমান
    const visibleRolls = Array.from(tableRows).map(row => row.cells[1].innerText.trim());

    practicalStudents = excelData.filter(row => {
        // রোলটি কি বর্তমানে ফিল্টার করা তালিকায় আছে?
        if (!visibleRolls.includes(String(row.Roll).trim())) return false;

        const type = String(row.Type || "").trim().toLowerCase();
        
        // শর্ত ১: সকল Regular পরীক্ষার্থী
        if (type === 'regular') return true;

        // শর্ত ২: Irregular দের মধ্যে নির্দিষ্ট ৪টি বিষয় চেক
        if (type === 'irregular') {
            const practicalSubs = ["Physics", "Chemistry", "Trade1", "Trade2"];
            return practicalSubs.some(sub => row[sub] != null && (row[sub] == 1 || String(row[sub]).trim() == "1"));
        }
        return false;
    });

    // ৩. UI জেনারেশন
    document.getElementById('main-ui').style.display = 'none';
    const displayArea = document.getElementById('seat-plan-ui');
    displayArea.style.display = 'block';

    let tableHTML = `
        <div class="p-4">
            <div class="d-flex justify-content-between align-items-center mb-4">
                <h4 class="fw-bold text-primary">
                    <i class="fa-solid fa-flask-vial me-2"></i>ব্যবহারিক পরীক্ষার্থীর তালিকা (ফিল্টার অনুযায়ী)
                </h4>
                <span class="badge bg-primary px-4 py-2">Total: ${practicalStudents.length}</span>
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
                        <tbody>`;

    practicalStudents.forEach((row, index) => {
        const type = String(row.Type || "").trim();
        tableHTML += `
            <tr>
                <td>${index + 1}</td>
                <td class="fw-bold">${row.Roll}</td>
                <td>${row.Trade}</td>
                <td>
                    <span class="badge ${type.toLowerCase() === 'regular' ? 'badge-reg' : 'badge-irreg'} rounded-pill px-3">
                        ${type}
                    </span>
                </td>
                <td class="text-secondary small">${row.Institute}</td>
            </tr>`;
    });

    if (practicalStudents.length === 0) {
        tableHTML += `<tr><td colspan="5" class="py-5 text-muted">কোন ব্যবহারিক পরীক্ষার্থী পাওয়া যায়নি</td></tr>`;
    }

    tableHTML += `</tbody></table></div></div>
    <div class="no-print mt-4">
        <button class="btn btn-outline-secondary me-2" onclick="backToMain()">
            <i class="fa-solid fa-arrow-left me-2"></i>Back to Filtered List
        </button>
        <button class="btn btn-dark px-4" onclick="window.print()">
            <i class="fa-solid fa-print me-2"></i>Print Report
        </button>
    </div>
    </div>`;

    displayArea.innerHTML = tableHTML;
}

function calculateInstituteFee() {
    const tableRows = document.querySelectorAll("#tableBody tr");
    const instQ = document.getElementById('instSearch').value;

    if (tableRows.length === 0 || tableRows[0].cells.length < 2) {
        alert("আগে ডাটা ফিল্টার করুন!"); return;
    }

    const visibleRolls = Array.from(tableRows).map(row => row.cells[1].innerText.trim());
    const filteredData = excelData.filter(row => visibleRolls.includes(String(row.Roll).trim()));

    // ১. Regular Students Calculation
    const regStudents = filteredData.filter(r => String(r.Type).toLowerCase() === 'regular');
    const regCount = regStudents.length;
    const regCenterFee = 550;
    const regPracCenterFee = 100;
    const regPracSubCount = 4; // Physics, Chemistry, Trade1, Trade2
    const regPracFeePerSub = 35;
    const regTotal = regCount * (regCenterFee + regPracCenterFee + (regPracSubCount * regPracFeePerSub));

    // ২. Irregular Students (Without Practical)
    const irregNoPrac = filteredData.filter(r => {
        const type = String(r.Type).toLowerCase();
        const hasPrac = ["Physics", "Chemistry", "Trade1", "Trade2"].some(sub => r[sub] == 1);
        return type === 'irregular' && !hasPrac;
    });
    const irregNoPracCount = irregNoPrac.length;
    const irregNoPracTotal = irregNoPracCount * 550;

    // ৩. Irregular Students (With Practical) - Grouped by failed subject count
    const irregWithPrac = filteredData.filter(r => {
        const type = String(r.Type).toLowerCase();
        const hasPrac = ["Physics", "Chemistry", "Trade1", "Trade2"].some(sub => r[sub] == 1);
        return type === 'irregular' && hasPrac;
    });

    let irregPracSummary = { 1: 0, 2: 0, 3: 0, 4: 0 }; // কয়জন কয়টি বিষয়ে ফেল করেছে
    irregWithPrac.forEach(r => {
        let failedCount = 0;
        if (r.Physics == 1) failedCount++;
        if (r.Chemistry == 1) failedCount++;
        if (r.Trade1 == 1) failedCount++;
        if (r.Trade2 == 1) failedCount++;
        if (failedCount > 0) irregPracSummary[failedCount]++;
    });

    // UI জেনারেশন
    document.getElementById('main-ui').style.display = 'none';
    const displayArea = document.getElementById('seat-plan-ui');
    displayArea.style.display = 'block';

    let irregWithPracRows = "";
    let irregPracGrandTotal = 0;

    [1, 2, 3, 4].forEach(num => {
        const count = irregPracSummary[num];
        if (count >= 0) { // আপনার ছবির মতো ০ হলেও দেখাবে
            const centerFee = 450;
            const pracCenterFee = 100;
            const pracFee = num * 35;
            const rowTotal = count * (centerFee + pracCenterFee + pracFee);
            irregPracGrandTotal += rowTotal;
            irregWithPracRows += `
                <tr>
                    <td>${num}</td>
                    <td>${count}</td>
                    <td>${centerFee}</td>
                    <td>${pracCenterFee}</td>
                    <td>${pracFee}</td>
                    <td>${rowTotal}</td>
                </tr>`;
        }
    });

    displayArea.innerHTML = `
        <div class="p-5 bg-white shadow mx-auto" style="max-width: 950px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
            <div class="text-center mb-4">
                <h4 class="fw-bold">${instQ || "Institute Name"}</h4>
                <h5 class="text-decoration-underline text-secondary">Center Fee Statement</h5>
            </div>

            <div class="mb-4">
                <h6 class="fw-bold text-dark">1. Regular Students</h6>
                <table class="table table-bordered text-center align-middle">
                    <thead class="table-light">
                        <tr>
                            <th>Students Nos</th><th>Center Fee</th><th>Practical Center Fee</th>
                            <th>Practical Subject Nos</th><th>Practical Fee (35/-)</th><th>Total Fee</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>${regCount}</td><td>${regCenterFee}</td><td>${regPracCenterFee}</td>
                            <td>${regPracSubCount}</td><td>${regPracSubCount * regPracFeePerSub}</td>
                            <td class="fw-bold">${regTotal}</td>
                        </tr>
                    </tbody>
                </table>
            </div>

            <h6 class="fw-bold text-dark">2. Irregular Students</h6>
            <div class="ms-3 mb-4">
                <p class="mb-1 fw-bold text-muted">2.1 Without Practical Subjects</p>
                <table class="table table-bordered text-center align-middle">
                    <thead class="table-light">
                        <tr><th>Students Nos</th><th>Center Fee</th><th>Practical Center Fee</th><th>Practical Fee</th><th>Total Fee</th></tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>${irregNoPracCount}</td><td>550</td><td>0</td><td>0</td>
                            <td class="fw-bold">${irregNoPracTotal}</td>
                        </tr>
                    </tbody>
                </table>

                <p class="mb-1 fw-bold text-muted mt-4">2.2 With Practical Subjects</p>
                <table class="table table-bordered text-center align-middle">
                    <thead class="table-light">
                        <tr>
                            <th>Practical Subject Nos</th><th>Students Nos</th><th>Center Fee</th>
                            <th>Practical Center Fee</th><th>Practical Fee (35/-)</th><th>Total Fee</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${irregWithPracRows}
                        <tr class="table-secondary fw-bold">
                            <td colspan="5" class="text-end">Total Irregular (With Prac):</td>
                            <td>${irregPracGrandTotal}</td>
                        </tr>
                    </tbody>
                </table>
            </div>

            <div class="mt-5 p-3 border-top border-3">
                <div class="d-flex justify-content-between align-items-center">
                    <h5 class="fw-bold">Total Center Fee = ${regTotal} + ${irregNoPracTotal} + ${irregPracGrandTotal} =</h5>
                    <h4 class="bg-dark text-white px-4 py-2 rounded">${regTotal + irregNoPracTotal + irregPracGrandTotal} /-</h4>
                </div>
            </div>

            <div class="no-print mt-5 d-flex gap-2">
                <button class="btn btn-outline-secondary px-4" onclick="backToMain()">Back</button>
                <button class="btn btn-primary px-4" onclick="window.print()">Print Statement</button>
            </div>
        </div>`;
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

    // ✅ SweetAlert ইনপুট
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
    const container = document.getElementById('seat-plan-ui');
    container.style.display = 'block';

    container.innerHTML = `
        <div class="p-3 text-center">

            <div class="no-print mb-3 d-flex justify-content-center gap-2">
                <button class="btn btn-dark" onclick="backToMain()">⬅ Back</button>
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

function downloadSortedRolls() {
    if (!filteredData.length) {
        alert("No data to download!");
        return;
    }

    const sorted = [...filteredData].sort((a, b) => Number(a.Roll) - Number(b.Roll));

    let html = `
    <html>
    <head>
        <title>Sorted Rolls</title>
        <style>
            @page {
                size: A4;
                margin: 10mm;
            }

            body {
                font-family: Arial, sans-serif;
                margin: 0;
            }

            .container {
                column-count: 6;
                column-gap: 10px;
                column-fill: auto; /* 🔥 FIX */
                height: 100vh;     /* 🔥 FIX */
            }

            .item {
                display: flex;
                justify-content: space-between;
                border-bottom: 1px solid #000;
                font-size: 10px;
                padding: 2px 4px;
                break-inside: avoid;
            }

            .header {
                font-weight: bold;
                border-bottom: 2px solid #000;
                margin-bottom: 3px;
            }
        </style>
    </head>
    <body>

    <div class="container">
        <div class="item header"><span>SL</span><span>Roll</span></div>
    `;

    sorted.forEach((row, index) => {
        html += `
            <div class="item">
                <span>${index + 1}</span>
                <span>${row.Roll}</span>
            </div>
        `;
    });

    html += `
    </div>
    </body>
    </html>
    `;

    const win = window.open('', '_blank');
    win.document.write(html);
    win.document.close();

    win.print();
}