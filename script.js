const EXCEL_URL = "./PRV File.xlsx";   // your 2-column Excel file

let rows = [];


/* ---------- Load Excel ---------- */
async function loadExcel() {

  const res = await fetch(`${EXCEL_URL}?ts=${Date.now()}`);
  const data = await res.arrayBuffer();

  const wb = XLSX.read(data, { type:"array"});
  const sheet = wb.Sheets[wb.SheetNames[0]];

  const all = XLSX.utils.sheet_to_json(sheet,{header:1});

  // skip header automatically
  const start = all[0][0].toString().toUpperCase() === "ID" ? 1 : 0;

  rows = all.slice(start).map(r => [
    String(r[0]||"").trim(),
    String(r[1]||"").trim()
  ]);

  console.log("Loaded rows:", rows.length);
}


/* ---------- Lookup ---------- */
function findValue(id){
  const hit = rows.find(r => r[0] === id);
  return hit ? hit[1] : null;
}


/* ---------- Search ---------- */
document.getElementById("searchForm").addEventListener("submit", e => {

  e.preventDefault();

  const id = document.getElementById("id").value.trim();
  const out = document.getElementById("result");

  out.textContent = "";

  if(!id) return out.textContent = "Enter an ID";

  const value = findValue(id);

  if(!value) out.textContent = "ID not found";
  else out.textContent = value;

});


/* ---------- Scanner ---------- */
let qr = null;

async function startScanner(){

  qr = new Html5Qrcode("qr-reader");

  document.getElementById("scannerWrap").style.display="block";

  await qr.start(
    { facingMode:"environment" },
    { fps:10, qrbox:250 },
    txt=>{
      stopScanner();
      document.getElementById("id").value = txt.trim();
      document.getElementById("searchForm").requestSubmit();
    }
  );

}

async function stopScanner(){
  if(qr) await qr.stop();
  document.getElementById("scannerWrap").style.display="none";
}


/* ---------- Buttons ---------- */
document.getElementById("scanBtn").addEventListener("click", startScanner);
document.getElementById("stopScanBtn").addEventListener("click", stopScanner);


/* ---------- Init ---------- */
loadExcel();
