const supabaseUrl = "https://rxpzwdqyhyudrqxnwukm.supabase.co";
const supabaseKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJ4cHp3ZHF5aHl1ZHJxeG53dWttIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzc4ODcwMTgsImV4cCI6MjA5MzQ2MzAxOH0.3BtreFMQT0opsfVXkZ4HSwwY_6oAI3Wb4XCMxAL1vws";
const supabaseClient = supabase.createClient(supabaseUrl, supabaseKey);

let tumKayitlar = [];

function showPage(id) {
  document.querySelectorAll(".page").forEach(p => p.classList.add("hidden"));
  document.getElementById(id).classList.remove("hidden");

  if (id === "kayitlar") kayitliBiletleriListele();
  if (id === "firmaRaporu") firmaDropdownlariniDoldur();
}

function paraFormat(sayi) {
  return Number(sayi || 0).toFixed(2) + " TL";
}

function durumNormalize(durum) {
  const d = String(durum || "").toLocaleUpperCase("tr-TR");

  if (d.includes("SATISIPTAL") || d.includes("SATIŞIPTAL") || d.includes("SATIŞİPTAL")) return "SATISIPTAL";
  if (d.includes("IADE") || d.includes("İADE")) return "IADE";
  if (d.includes("VOID") || d.includes("REZIPTAL") || d.includes("REZİPTAL") || d.includes("REZERVASYONIPTAL")) return "VOID";
  if (d.includes("SATIS") || d.includes("SATIŞ")) return "SATIS";

  return d;
}

function kdvHesapla(kdvDahil, seyahatTipi) {
  const oran = seyahatTipi === "Yurtiçi" ? 0.20 : 0;
  const kdvHaric = oran > 0 ? kdvDahil / 1.20 : kdvDahil;
  const kdvTutari = kdvDahil - kdvHaric;
  return { kdvHaric, kdvTutari, kdvDahil, oran };
}

async function firmaEkle() {
  const input = document.getElementById("firmaAdi");
  const firmaAdi = input.value.trim();
  if (!firmaAdi) return alert("Firma adı boş olamaz");

  const { error } = await supabaseClient.from("firmalar").insert([{ firma_adi: firmaAdi }]);
  if (error) return alert("Hata: " + error.message);

  input.value = "";
  firmalariListele();
  firmaDropdownlariniDoldur();
}

async function firmaSil(id) {
  if (!confirm("Firmayı silmek istiyor musun?")) return;

  const { error } = await supabaseClient.from("firmalar").delete().eq("id", id);
  if (error) return alert("Silme hatası: " + error.message);

  firmalariListele();
  firmaDropdownlariniDoldur();
}

async function firmalariListele() {
  const ul = document.getElementById("firmaListesi");
  if (!ul) return;

  const { data, error } = await supabaseClient.from("firmalar").select("*").order("id", { ascending: false });
  if (error) return alert("Firma listeleme hatası: " + error.message);

  ul.innerHTML = "";

  (data || []).forEach(f => {
    ul.innerHTML += `
      <li class="bg-white rounded-xl shadow p-4 mb-3 flex justify-between items-center">
        <span class="font-semibold text-gray-800">${f.firma_adi}</span>
        <button onclick="firmaSil(${f.id})" class="bg-red-500 text-white px-3 py-1 rounded-lg">Sil</button>
      </li>
    `;
  });
}

async function firmaDropdownlariniDoldur() {
  const { data } = await supabaseClient.from("firmalar").select("*").order("firma_adi", { ascending: true });
  const firmalar = data || [];

  ["kayitFirmaFilter", "raporFirma"].forEach(id => {
    const select = document.getElementById(id);
    if (!select) return;

    const eski = select.value;
    select.innerHTML = id === "raporFirma"
      ? `<option value="">Firma seç</option>`
      : `<option value="">Tüm Firmalar</option>`;

    firmalar.forEach(f => {
      select.innerHTML += `<option value="${f.id}">${f.firma_adi}</option>`;
    });

    select.value = eski;
  });
}

function hasilatOku() {
  const input = document.getElementById("hasilatInput");
  const sonuc = document.getElementById("hasilatSonuc");
  if (!input.files.length) return alert("Hasılat dosyası seç");

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    window.hasilatData = json;
    sonuc.innerHTML = `<b>${json.length}</b> hasılat satırı okundu.`;
  };

  reader.readAsArrayBuffer(input.files[0]);
}

function excelOku() {
  const input = document.getElementById("excelInput");
  const sonuc = document.getElementById("excelSonuc");
  const pnrOzetDiv = document.getElementById("pnrOzet");
  if (!input.files.length) return alert("Detaylı işlem dosyası seç");

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    sonuc.innerHTML = `<b>${json.length}</b> ham satır okundu.`;

    const gruplar = {};
    json.forEach(satir => {
      const pnr = String(satir["Pnr"] || satir["PNR"] || "").trim();
      if (!pnr) return;
      if (!gruplar[pnr]) gruplar[pnr] = [];
      gruplar[pnr].push(satir);
    });

    const pnrListesi = Object.keys(gruplar);

    let html = `
      <div class="bg-white rounded-xl shadow p-4">
      <h3 class="font-bold mb-3">${pnrListesi.length} farklı PNR bulundu</h3>
      <table class="w-full text-sm border">
        <tr class="bg-gray-100">
          <th class="border p-2">PNR</th>
          <th class="border p-2">Yolcu</th>
          <th class="border p-2">Bilet</th>
          <th class="border p-2">Hareket</th>
          <th class="border p-2">Son Durum</th>
          <th class="border p-2">Net Maliyet</th>
          <th class="border p-2">Detay</th>
        </tr>
    `;

    pnrListesi.forEach(pnr => {
      const hareketler = gruplar[pnr];
      const ozet = pnrOzetHesapla(hareketler);
      const fiyat = pnrFiyatHesapla(pnr);

      html += `
        <tr>
          <td class="border p-2">${pnr}</td>
          <td class="border p-2">${ozet.yolcuSayisi}</td>
          <td class="border p-2">${ozet.biletSayisi}</td>
          <td class="border p-2">${hareketler.length}</td>
          <td class="border p-2">${ozet.sonDurum}</td>
          <td class="border p-2">${paraFormat(fiyat)}</td>
          <td class="border p-2">
            <button onclick='pnrDetayGoster(${JSON.stringify(pnr)}, ${JSON.stringify(hareketler)})' class="bg-blue-600 text-white px-3 py-1 rounded">Aç</button>
          </td>
        </tr>
      `;
    });

    html += `</table></div>`;
    pnrOzetDiv.innerHTML = html;
  };

  reader.readAsArrayBuffer(input.files[0]);
}

function pnrOzetHesapla(hareketler) {
  const yolcular = new Set();
  const biletler = new Set();

  hareketler.forEach(satir => {
    const yolcu = String(satir["Ad ve Soyad"] || "").trim();
    const bilet = String(satir["E-Bilet No"] || "").trim();
    if (yolcu) yolcular.add(yolcu);
    if (bilet) biletler.add(bilet);
  });

  const sirali = [...hareketler].sort((a, b) => tarihCevir(a["İşlem Tarihi"]) - tarihCevir(b["İşlem Tarihi"]));
  const son = sirali[sirali.length - 1] || {};

  return {
    yolcuSayisi: yolcular.size,
    biletSayisi: biletler.size,
    sonDurum: son["Son Durum"] || son["İşlem Tip"] || "-"
  };
}

async function pnrDetayGoster(pnr, hareketler) {
  const firmalar = await firmalariGetir();
  const hafiza = await yolcuFirmaHafizasiGetir();
  const maliyet = pnrFiyatHesapla(pnr);
  const seyahatTipi = seyahatTipiBul(hareketler);

  const biletGruplari = {};
  hareketler.forEach(satir => {
    const biletNo = String(satir["E-Bilet No"] || "").trim();
    const yolcu = String(satir["Ad ve Soyad"] || "").trim();
    const key = biletNo || yolcu;
    if (!key) return;
    if (!biletGruplari[key]) biletGruplari[key] = [];
    biletGruplari[key].push(satir);
  });

  const ozetler = [];

  Object.keys(biletGruplari).forEach(key => {
    const liste = biletGruplari[key].sort((a, b) => tarihCevir(a["İşlem Tarihi"]) - tarihCevir(b["İşlem Tarihi"]));
    const son = liste[liste.length - 1] || {};

    ozetler.push({
      yolcu: son["Ad ve Soyad"],
      bilet: son["E-Bilet No"],
      durum: son["Son Durum"] || son["İşlem Tip"],
      tarih: son["İşlem Tarihi"]
    });
  });

  let html = `
    <div class="bg-white rounded-xl shadow p-4 mt-4">
      <button onclick='dbKaydet("${pnr}", ${JSON.stringify(ozetler)})' class="bg-green-600 text-white px-4 py-2 rounded-lg mb-3">DB’ye Kaydet</button>
      <p><b>Net Maliyet:</b> ${paraFormat(maliyet)}</p>

      <div class="border rounded-xl p-4 my-4 bg-gray-50">
        <h3 class="font-bold mb-3">Satış / KDV Bilgisi</h3>

        <select id="seyahatTipi_${pnr}" class="border p-2 rounded">
          <option value="Yurtiçi" ${seyahatTipi === "Yurtiçi" ? "selected" : ""}>Yurtiçi</option>
          <option value="Yurtdışı" ${seyahatTipi === "Yurtdışı" ? "selected" : ""}>Yurtdışı</option>
        </select>

        <input type="number" id="kdvDahil_${pnr}" placeholder="KDV dahil satış" class="border p-2 rounded ml-2">
        <button onclick="fiyatHesaplaUI('${pnr}')" class="bg-blue-600 text-white px-3 py-2 rounded ml-2">Hesapla</button>
        <div id="fiyatSonuc_${pnr}" class="mt-3"></div>
      </div>

      <h3 class="font-bold mb-3">${pnr} Yolcu/Bilet Özeti</h3>

      <table class="w-full text-sm border">
        <tr class="bg-gray-100">
          <th class="border p-2">Yolcu</th>
          <th class="border p-2">Firma</th>
          <th class="border p-2">E-Bilet No</th>
          <th class="border p-2">Durum</th>
          <th class="border p-2">Tarih</th>
          <th class="border p-2">Kaydet</th>
        </tr>
  `;

  ozetler.forEach((o, i) => {
    const yolcuAdi = String(o.yolcu || "").trim();
    const kayitliFirmaId = hafiza[yolcuAdi] || "";

    let options = `<option value="">Firma seç</option>`;
    firmalar.forEach(f => {
      const selected = String(f.id) === String(kayitliFirmaId) ? "selected" : "";
      options += `<option value="${f.id}" ${selected}>${f.firma_adi}</option>`;
    });

    html += `
      <tr>
        <td class="border p-2">${yolcuAdi}</td>
        <td class="border p-2"><select id="firmaSec_${i}" class="border p-1 rounded">${options}</select></td>
        <td class="border p-2">${o.bilet || ""}</td>
        <td class="border p-2">${o.durum || ""}</td>
        <td class="border p-2">${o.tarih || ""}</td>
        <td class="border p-2">
          <button onclick='yolcuFirmaKaydet(${JSON.stringify(yolcuAdi)}, document.getElementById("firmaSec_${i}").value)' class="bg-purple-600 text-white px-3 py-1 rounded">Kaydet</button>
        </td>
      </tr>
    `;
  });

  html += `</table></div>`;
  document.getElementById("excelSonuc").innerHTML = html;
}

function fiyatHesaplaUI(pnr) {
  const kdvDahil = Number(document.getElementById(`kdvDahil_${pnr}`).value || 0);
  const tip = document.getElementById(`seyahatTipi_${pnr}`).value;
  const maliyet = pnrFiyatHesapla(pnr);

  if (!kdvDahil) return alert("Satış fiyatı gir");

  const h = kdvHesapla(kdvDahil, tip);
  const kar = h.kdvDahil - maliyet;

  document.getElementById(`fiyatSonuc_${pnr}`).innerHTML = `
    <p><b>KDV Hariç:</b> ${paraFormat(h.kdvHaric)}</p>
    <p><b>KDV:</b> ${paraFormat(h.kdvTutari)}</p>
    <p><b>KDV Dahil:</b> ${paraFormat(h.kdvDahil)}</p>
    <p><b>Kar/Zarar:</b> ${paraFormat(kar)}</p>
  `;
}

async function dbKaydet(pnr, ozetler) {
  const hafiza = await yolcuFirmaHafizasiGetir();
  const maliyet = pnrFiyatHesapla(pnr);
  const tip = document.getElementById(`seyahatTipi_${pnr}`)?.value || "Yurtiçi";
  const kdvDahil = Number(document.getElementById(`kdvDahil_${pnr}`)?.value || 0);
  const h = kdvHesapla(kdvDahil, tip);
  const kar = h.kdvDahil - maliyet;

  const kayitlar = ozetler.map(o => {
    const yolcuAdi = String(o.yolcu || "").trim();

    return {
      pnr,
      yolcu_adi: yolcuAdi,
      e_bilet_no: o.bilet,
      son_durum: o.durum,
      son_islem_tarihi: tarihCevir(o.tarih),
      firma_id: hafiza[yolcuAdi] || null,
      net_tutar: maliyet,
      seyahat_tipi: tip,
      satis_kdv_dahil: h.kdvDahil,
      satis_kdv_haric: h.kdvHaric,
      satis_kdv_tutari: h.kdvTutari,
      kdv_orani: h.oran,
      kar_zarar: kar,
      fatura_durumu: "Kesilmedi"
    };
  });

  const { error } = await supabaseClient.from("pnr_kayitlari").insert(kayitlar);
  if (error) return alert("DB hata: " + error.message);

  alert("DB kaydedildi");
  kayitliBiletleriListele();
}

async function firmalariGetir() {
  const { data } = await supabaseClient.from("firmalar").select("*").order("firma_adi");
  return data || [];
}

async function yolcuFirmaHafizasiGetir() {
  const { data } = await supabaseClient.from("yolcu_firma_hafizasi").select("yolcu_adi, firma_id");
  const map = {};
  (data || []).forEach(r => map[String(r.yolcu_adi || "").trim()] = r.firma_id);
  return map;
}

async function yolcuFirmaKaydet(yolcuAdi, firmaId) {
  if (!firmaId) return alert("Firma seç");

  const { error } = await supabaseClient
    .from("yolcu_firma_hafizasi")
    .upsert({ yolcu_adi: yolcuAdi, firma_id: Number(firmaId) }, { onConflict: "yolcu_adi" });

  if (error) return alert("Hafıza hatası: " + error.message);
  alert("Firma hafızaya kaydedildi");
}

function pnrFiyatHesapla(pnr) {
  if (!window.hasilatData) return 0;

  const satirlar = window.hasilatData.filter(x => String(x["Pnr"] || "").trim() === String(pnr).trim());
  let toplam = 0;

  satirlar.forEach(s => {
    toplam += paraCevir(s["Ön Ödeme Tutarı"]);
    toplam += paraCevir(s["Hopi Tutarı"]);
    toplam += paraCevir(s["Kredi Kartı Tutarı"]);
  });

  return toplam;
}

function paraCevir(deger) {
  if (!deger) return 0;
  return parseFloat(String(deger).replace("TL", "").replace(/\./g, "").replace(",", ".").trim()) || 0;
}

function seyahatTipiBul(hareketler) {
  const rota = String((hareketler[0] || {})["Kalkış Varış"] || "").toLocaleLowerCase("tr-TR");

  if (!rota) return "Yurtiçi";

  const yurtdisiUlkeler = [
    "abd", "amerika", "fransa", "kanada", "almanya", "ingiltere", "hollanda",
    "italya", "ispanya", "belçika", "isviçre", "azerbaycan", "bae", "dubai",
    "katar", "mısır", "suudi", "rusya", "gürcistan", "yunanistan", "kore", "çin", "japonya"
  ];

  return yurtdisiUlkeler.some(u => rota.includes(u)) ? "Yurtdışı" : "Yurtiçi";
}

function tarihCevir(tarih) {
  if (!tarih) return new Date(0);

  if (typeof tarih === "number") {
    return new Date((tarih - 25569) * 86400 * 1000);
  }

  const text = String(tarih);
  const tarihKismi = text.split(" ")[0];
  const saat = text.split(" ")[1] || "00:00:00";

  if (tarihKismi.includes(".")) {
    const [gun, ay, yil] = tarihKismi.split(".");
    return new Date(`${yil}-${ay}-${gun}T${saat}`);
  }

  return new Date(text);
}

async function kayitliBiletleriListele() {
  const div = document.getElementById("kayitliBiletler");
  const ozetDiv = document.getElementById("kayitliOzet");
  if (!div) return;

  const { data, error } = await supabaseClient
    .from("pnr_kayitlari")
    .select(`
      id,pnr,yolcu_adi,e_bilet_no,son_durum,son_islem_tarihi,net_tutar,firma_id,
      seyahat_tipi,satis_kdv_dahil,satis_kdv_haric,satis_kdv_tutari,kar_zarar,
      fatura_durumu,fatura_no,fatura_tarihi,
      firmalar(firma_adi)
    `)
    .order("id", { ascending: false });

  if (error) {
    div.innerHTML = "Kayıt alınamadı: " + error.message;
    return;
  }

  tumKayitlar = data || [];

  let filtered = [...tumKayitlar];

  const pnrAra = document.getElementById("aramaPnr")?.value.toLowerCase() || "";
  const yolcuAra = document.getElementById("aramaYolcu")?.value.toLocaleLowerCase("tr-TR") || "";
  const firmaId = document.getElementById("kayitFirmaFilter")?.value || "";
  const durum = document.getElementById("durumFilter")?.value || "";
  const fatura = document.getElementById("faturaFilter")?.value || "";
  const baslangic = document.getElementById("baslangicTarih")?.value || "";
  const bitis = document.getElementById("bitisTarih")?.value || "";

  if (pnrAra) filtered = filtered.filter(x => String(x.pnr || "").toLowerCase().includes(pnrAra));
  if (yolcuAra) filtered = filtered.filter(x => String(x.yolcu_adi || "").toLocaleLowerCase("tr-TR").includes(yolcuAra));
  if (firmaId) filtered = filtered.filter(x => String(x.firma_id) === String(firmaId));
  if (durum) filtered = filtered.filter(x => durumNormalize(x.son_durum) === durum);
  if (fatura) filtered = filtered.filter(x => String(x.fatura_durumu || "Kesilmedi") === fatura);

  if (baslangic) filtered = filtered.filter(x => String(x.son_islem_tarihi || "").slice(0, 10) >= baslangic);
  if (bitis) filtered = filtered.filter(x => String(x.son_islem_tarihi || "").slice(0, 10) <= bitis);

  ozetBas(filtered, ozetDiv);

  let html = `
    <table class="w-full bg-white rounded-xl shadow text-sm">
      <tr class="bg-gray-100">
        <th class="p-2">PNR</th>
        <th>Yolcu</th>
        <th>Firma</th>
        <th>E-Bilet</th>
        <th>Durum</th>
        <th>Seyahat</th>
        <th>Maliyet</th>
        <th>Satış</th>
        <th>KDV</th>
        <th>Kar</th>
        <th>Fatura</th>
        <th>Tarih</th>
      </tr>
  `;

  filtered.forEach(x => {
    html += `
      <tr class="border-t">
        <td class="p-2">${x.pnr || ""}</td>
        <td>${x.yolcu_adi || ""}</td>
        <td>${x.firmalar?.firma_adi || ""}</td>
        <td>${x.e_bilet_no || ""}</td>
        <td>${x.son_durum || ""}</td>
        <td>${x.seyahat_tipi || ""}</td>
        <td>${paraFormat(x.net_tutar)}</td>
        <td>${paraFormat(x.satis_kdv_dahil)}</td>
        <td>${paraFormat(x.satis_kdv_tutari)}</td>
        <td>${paraFormat(x.kar_zarar)}</td>
        <td>${x.fatura_durumu || "Kesilmedi"}</td>
        <td>${String(x.son_islem_tarihi || "").slice(0, 10)}</td>
      </tr>
    `;
  });

  html += `</table>`;
  div.innerHTML = html;

  dashboard(tumKayitlar);
}

function ozetBas(data, div) {
  if (!div) return;

  const toplamMaliyet = data.reduce((t, x) => t + Number(x.net_tutar || 0), 0);
  const toplamSatis = data.reduce((t, x) => t + Number(x.satis_kdv_dahil || 0), 0);
  const toplamKdv = data.reduce((t, x) => t + Number(x.satis_kdv_tutari || 0), 0);
  const toplamKar = data.reduce((t, x) => t + Number(x.kar_zarar || 0), 0);
  const karOrani = toplamSatis ? (toplamKar / toplamSatis) * 100 : 0;

  div.innerHTML = `
    <div class="grid grid-cols-5 gap-3">
      <div class="bg-white p-3 rounded-xl shadow"><p>Kayıt</p><b>${data.length}</b></div>
      <div class="bg-white p-3 rounded-xl shadow"><p>Maliyet</p><b>${paraFormat(toplamMaliyet)}</b></div>
      <div class="bg-white p-3 rounded-xl shadow"><p>Satış</p><b>${paraFormat(toplamSatis)}</b></div>
      <div class="bg-white p-3 rounded-xl shadow"><p>KDV</p><b>${paraFormat(toplamKdv)}</b></div>
      <div class="bg-white p-3 rounded-xl shadow"><p>Kar / Oran</p><b>${paraFormat(toplamKar)} / %${karOrani.toFixed(2)}</b></div>
    </div>
  `;
}

async function firmaRaporuGetir() {
  const firmaId = document.getElementById("raporFirma").value;
  const baslangic = document.getElementById("raporBaslangic").value;
  const bitis = document.getElementById("raporBitis").value;

  if (!firmaId) return alert("Firma seç");

  await kayitliBiletleriListele();

  let data = tumKayitlar.filter(x => String(x.firma_id) === String(firmaId));

  if (baslangic) data = data.filter(x => String(x.son_islem_tarihi || "").slice(0, 10) >= baslangic);
  if (bitis) data = data.filter(x => String(x.son_islem_tarihi || "").slice(0, 10) <= bitis);

  ozetBas(data, document.getElementById("firmaRaporOzet"));

  let html = `
    <table class="w-full bg-white rounded-xl shadow text-sm">
      <tr class="bg-gray-100">
        <th class="p-2">PNR</th><th>Yolcu</th><th>E-Bilet</th><th>Durum</th><th>Maliyet</th><th>Satış</th><th>KDV</th><th>Kar</th><th>Fatura</th><th>Tarih</th>
      </tr>
  `;

  data.forEach(x => {
    html += `
      <tr class="border-t">
        <td class="p-2">${x.pnr || ""}</td>
        <td>${x.yolcu_adi || ""}</td>
        <td>${x.e_bilet_no || ""}</td>
        <td>${x.son_durum || ""}</td>
        <td>${paraFormat(x.net_tutar)}</td>
        <td>${paraFormat(x.satis_kdv_dahil)}</td>
        <td>${paraFormat(x.satis_kdv_tutari)}</td>
        <td>${paraFormat(x.kar_zarar)}</td>
        <td>${x.fatura_durumu || "Kesilmedi"}</td>
        <td>${String(x.son_islem_tarihi || "").slice(0, 10)}</td>
      </tr>
    `;
  });

  html += `</table>`;
  document.getElementById("firmaRaporTablo").innerHTML = html;
}

function dashboard(data) {
  const toplamMaliyet = data.reduce((t, x) => t + Number(x.net_tutar || 0), 0);
  const toplamSatis = data.reduce((t, x) => t + Number(x.satis_kdv_dahil || 0), 0);
  const toplamKar = data.reduce((t, x) => t + Number(x.kar_zarar || 0), 0);

  document.getElementById("dashKayit").innerText = data.length;
  document.getElementById("dashMaliyet").innerText = paraFormat(toplamMaliyet);
  document.getElementById("dashSatis").innerText = paraFormat(toplamSatis);
  document.getElementById("dashKar").innerText = paraFormat(toplamKar);
}

window.onload = function () {
  firmalariListele();
  firmaDropdownlariniDoldur();
  kayitliBiletleriListele();
};
