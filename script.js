let studentDatabase = [];

// ১. এক্সেল রিড করা
document.getElementById('excel-file').addEventListener('change', function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = function(evt) {
        try {
            const data = evt.target.result;
            const workbook = XLSX.read(data, {type: 'binary'});
            const sheetName = workbook.SheetNames[0];
            studentDatabase = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            alert("ভীমগঞ্জ মডেল স্কুল: ডাটাবেজ সফলভাবে লোড হয়েছে!");
        } catch (err) {
            alert("ফাইলটি সঠিক নয়!");
        }
    };
    reader.readAsBinaryString(file);
});

// ২. স্টুডেন্ট খোঁজা (স্মার্ট সার্চ)
function findStudent() {
    const rollInput = document.getElementById('in-roll').value.trim();
    const clsInput = document.getElementById('in-class').value.trim();
    
    if(rollInput && clsInput) {
        const student = studentDatabase.find(s => {
            // রোল এবং ক্লাস মিলিয়ে দেখা
            let exRoll = String(s.Roll || s.roll || s['রোল'] || "").trim();
            let exClass = String(s.Class || s.class || s['শ্রেণি'] || "").trim();
            return exRoll === rollInput && exClass === clsInput;
        });
        
        if(student) {
            // নামের কলামটি খুঁজে বের করা (যদি Name না থাকে তবে প্রথম কলামের ভ্যালু নেবে)
            const keys = Object.keys(student);
            const nameKey = keys.find(k => k.toLowerCase().includes('name') || k.includes('নাম')) || keys[0];
            
            document.getElementById('in-student-name').value = student[nameKey] || "নাম পাওয়া যায়নি";
            
            // বকেয়া টাকা বের করা
            const dueKey = keys.find(k => k.toLowerCase().includes('due') || k.includes('বকেয়া')) || keys[3];
            let dueVal = String(student[dueKey] || "0").replace(/[^0-9]/g, '');
            document.getElementById('prev-due').value = dueVal;
            
            calculateBalance();
        } else {
            document.getElementById('in-student-name').value = "পাওয়া যায়নি";
            document.getElementById('prev-due').value = 0;
            document.getElementById('current-due').value = 0;
        }
    }
}

// ৩. হিসাব-নিকাশ
function calculateBalance() {
    const prev = parseFloat(document.getElementById('prev-due').value) || 0;
    const pay = parseFloat(document.getElementById('today-pay').value) || 0;
    document.getElementById('current-due').value = prev - pay;
    updateWords(pay);
}

// ৪. রসিদ তৈরি
function generateReceipt() {
    const name = document.getElementById('in-student-name').value;
    if(!name || name === "পাওয়া যায়নি") { alert("আগে সঠিক ছাত্র খুঁজুন"); return; }

    const clsMap = {"1":"১ম","2":"২য়","3":"৩য়","4":"৪র্থ","5":"৫ম","6":"৬ষ্ঠ","7":"৭ম","8":"৮ম","9":"৯ম","10":"১০ম"};
    const clsBn = clsMap[document.getElementById('in-class').value];

    const html = `
        <div class="receipt-container" style="border: 2px solid #000; padding: 20px; width: 500px; margin: auto; font-family: sans-serif;">
            <div style="text-align:center;">
                <h2 style="margin:0;">ভীমগঞ্জ মডেল স্কুল</h2>
                <p style="margin:2px;">ভীমগঞ্জ বাজার, শেরপুর</p>
                <p><b>বেতন আদায়ের রশিদ</b></p>
            </div>
            <p>রশিদ নং: <b>${document.getElementById('in-receipt-no').value}</b> &nbsp;&nbsp; তারিখ: <b>${new Date().toLocaleDateString('bn-BD')}</b></p>
            <hr>
            <p>নাম: <b>${name}</b></p>
            <p>শ্রেণি: <b>${clsBn}</b> &nbsp;&nbsp; রোল: <b>${document.getElementById('in-roll').value}</b></p>
            <table border="1" style="width:100%; border-collapse:collapse;">
                <tr><th style="padding:5px;">বিবরণ</th><th style="padding:5px; text-align:right;">টাকা</th></tr>
                <tr><td style="padding:5px;">পূর্বের বকেয়া</td><td style="padding:5px; text-align:right;">${document.getElementById('prev-due').value}/-</td></tr>
                <tr style="background:#eee;"><td style="padding:5px;"><b>আজকের জমা</b></td><td style="padding:5px; text-align:right;"><b>${document.getElementById('today-pay').value}/-</b></td></tr>
                <tr><td style="padding:5px;">অবশিষ্ট বকেয়া</td><td style="padding:5px; text-align:right;">${document.getElementById('current-due').value}/-</td></tr>
            </table>
            <p>কথায়: <b>${document.getElementById('in-words-text').value}</b></p>
            <div style="display:flex; justify-content:space-between; margin-top:40px;">
                <p style="border-top:1px solid #000; width:100px; text-align:center;">অভিভাবক</p>
                <p style="border-top:1px solid #000; width:100px; text-align:center;">আদায়কারী</p>
            </div>
        </div>
    `;
    document.getElementById('receipt-area').innerHTML = html;
    document.getElementById('receipt-area').style.display = 'block';
}

// ৫. টাকা থেকে কথায়
function updateWords(n) {
    const units = ['', 'এক', 'দুই', 'তিন', 'চার', 'পাঁচ', 'ছয়', 'সাত', 'আট', 'নয়'];
    const teens = ['দশ', 'এগারো', 'বারো', 'তেরো', 'চৌদ্দ', 'পনেরো', 'ষোলো', 'সতেরো', 'আঠারো', 'উনিশ'];
    const tens = ['', '', 'বিশ', 'ত্রিশ', 'চল্লিশ', 'পঞ্চাশ', 'ষাট', 'সত্তুর', 'আশি', 'নব্বই'];
    function conv(num) {
        if (num == 0) return "";
        let res = "";
        if (num >= 1000) { res += conv(Math.floor(num / 1000)) + " হাজার "; num %= 1000; }
        if (num >= 100) { res += units[Math.floor(num / 100)] + " শত "; num %= 100; }
        if (num > 0) {
            if (num < 10) res += units[num];
            else if (num < 20) res += teens[num - 10];
            else { res += tens[Math.floor(num / 10)]; if (num % 10 > 0) res += " " + units[num % 10]; }
        }
        return res.trim();
    }
    document.getElementById('in-words-text').value = n > 0 ? conv(n) + " টাকা মাত্র" : "";
}

window.onload = function() {
    document.getElementById('in-date').valueAsDate = new Date();
    document.getElementById('in-receipt-no').value = Math.floor(Math.random() * 9000) + 1000;
};