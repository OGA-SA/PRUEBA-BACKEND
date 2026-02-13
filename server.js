const express = require("express");
const multer = require("multer");
const fetch = require("node-fetch");
const qs = require("querystring");
const cors = require("cors");
const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");

const app = express();

app.use(express.json({ limit: "15mb" }));
app.use(express.urlencoded({ limit: "15mb", extended: true }));

const upload = multer();


// ================= ENV =================

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_ID = process.env.SITE_ID;
const DRIVE_ID = process.env.DRIVE_ID;

const DEFAULT_FOLDER = process.env.FOLDER_PATH || "Extra Seguro";

const allowedOrigins = (process.env.ALLOWED_ORIGIN || "").split(",").filter(Boolean);

app.use(cors({
  origin: allowedOrigins.length > 0 ? allowedOrigins : "*",
  methods: ["POST","OPTIONS","GET"]
}));

app.options("/upload", cors());
app.options("/generate-pdf-editable", cors());

app.get("/", (req,res)=>res.send("✅ Backend funcionando"));


// ================= TOKEN =================

async function getAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  const body = qs.stringify({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const r = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const data = await r.json();
  if (!r.ok) throw new Error(JSON.stringify(data));
  return data.access_token;
}


// ================= SHAREPOINT UPLOAD =================

async function uploadToSharePoint(accessToken, buffer, filename, folder) {
  const safeFolder = encodeURI(folder);
  const safeName = encodeURIComponent(filename);

  const uploadUrl =
    `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:/${safeFolder}/${safeName}:/content`;

  const res = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/pdf"
    },
    body: buffer
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(text);
  }

  return res.json();
}


// ================= PDF EDITABLE =================

app.post("/generate-pdf-editable", async (req,res)=>{
  try{

    const data = req.body;

    const pdfDoc = await PDFDocument.create();
    const page = pdfDoc.addPage([595, 842]);

    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

    const form = pdfDoc.getForm();

    function field(name, x, y, w=200, h=18, value=""){
      const f = form.createTextField(name);
      f.setText(value || "");
      f.addToPage(page,{ x, y, width:w, height:h });
      f.setFontSize(10);
    }

    // ===== CAMPOS PRINCIPALES =====

    field("taller", 50, 790, 200, 18, data.taller);
    field("serieNumero", 300, 790, 200, 18, data.serieNumero);
    field("fecha", 50, 760, 200, 18, data.fecha);
    field("siniestro", 300, 760, 200, 18, data.siniestro1+"-"+data.siniestro2);
    field("dificultadVisual", 50, 730, 450, 18, data.dificultadVisual);
    field("quien", 50, 700, 300, 18, data.quien);

    // ===== TABLAS =====

    let y = 660;

    function drawTabla(tabla){
      tabla.forEach((row,i)=>{
        field(`pieza_${i}`, 50, y, 220, 16, row.pieza);
        field(`chapa_${i}`, 280, y, 100, 16, row.chapa);
        field(`pintura_${i}`, 390, y, 100, 16, row.pintura);
        y -= 20;
        if(y < 80){
          y = 760;
          pdfDoc.addPage();
        }
      });
    }

    drawTabla(data.tabla1 || []);
    drawTabla(data.tabla2 || []);

    // ===== CANVAS IMAGE =====

    if(data.canvasImage){
      const base64 = data.canvasImage.split(",")[1];
      const img = await pdfDoc.embedPng(Buffer.from(base64,"base64"));
      page.drawImage(img,{
        x:50,
        y:500,
        width:500,
        height:200
      });
    }

    const pdfBytes = await pdfDoc.save();

    const filename =
      `${data.siniestro1}_${data.siniestro2}_EXTRA_SEGURO_EDITABLE.pdf`;

    const token = await getAccessToken();

    const result = await uploadToSharePoint(
      token,
      Buffer.from(pdfBytes),
      filename,
      DEFAULT_FOLDER
    );

    res.json({
      ok:true,
      webUrl: result.webUrl,
      name: result.name
    });

  }catch(e){
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});


// ================= ENDPOINT ORIGINAL =================

app.post("/upload", upload.single("pdf"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "Falta pdf" });
    }

    const filename = req.file.originalname;
    const folder = DEFAULT_FOLDER;

    const token = await getAccessToken();
    const result = await uploadToSharePoint(token, req.file.buffer, filename, folder);

    res.json({ ok:true, webUrl: result.webUrl });

  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});


// ================= START =================

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log("🚀 Backend listo puerto", PORT);
});
