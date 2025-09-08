// server.js
require("dotenv").config();
const express = require("express");
const bodyParser = require("body-parser");
const nodemailer = require("nodemailer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = 3000;

// Middleware
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
// Servir arquivos estáticos da pasta frontend
app.use(express.static(path.join(__dirname, "../frontend")));


// Servir o formulário (HTML) automaticamente
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "../frontend/questionario.html"));
});

// Rota para receber respostas
app.post("/submit", async (req, res) => {
  const dados = req.body;

  try {
    // === 1) Gerar Excel bem formatado ===
    const excelPath = path.join(__dirname, "respostas.xlsx");
let wb;

// Se o arquivo já existir, abre ele
if (fs.existsSync(excelPath)) {
  wb = XLSX.readFile(excelPath);
  const ws = wb.Sheets["Respostas"];
  const data = XLSX.utils.sheet_to_json(ws);

  // Adiciona a nova resposta
  data.push(dados);

  // Cria a planilha novamente
  const newWs = XLSX.utils.json_to_sheet(data);
  wb.Sheets["Respostas"] = newWs;
} else {
  // Se não existir, cria um novo
  wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet([dados]);
  XLSX.utils.book_append_sheet(wb, ws, "Respostas");
}

// Salva o Excel corretamente
XLSX.writeFile(wb, excelPath);


    // === 2) Criar corpo do e-mail ===
    let corpoEmail = "📋 Nova resposta recebida:\n\n";
    for (const [campo, valor] of Object.entries(dados)) {
      corpoEmail += `${campo}: ${valor}\n`;
    }

    // === 3) Enviar e-mail ===
    let transporter = nodemailer.createTransport({
      service: "gmail",
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
      },
    });

    await transporter.sendMail({
      from: `"TMCO Questionário" <${process.env.EMAIL_USER}>`,
      to: process.env.EMAIL_DESTINO,
      subject: "📩 Nova resposta do questionário",
      text: corpoEmail,
      attachments: [
        {
          filename: "respostas.xlsx",
          path: excelPath,
        },
      ],
    });

    console.log("✅ E-mail enviado com sucesso!");
    res.json({ message: "Formulário enviado com sucesso!" });
  } catch (err) {
    console.error("❌ Erro ao enviar:", err);
    res.status(500).json({ message: "Erro ao processar o formulário" });
  }
});

// Servir arquivos estáticos da pasta frontend
app.use("/static", express.static(path.join(__dirname, "../frontend")));

// Iniciar servidor
app.listen(PORT, () => {
  console.log(`🚀 Servidor rodando em http://localhost:${PORT}`);
});
// 

