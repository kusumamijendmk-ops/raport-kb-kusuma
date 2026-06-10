import { GoogleGenAI } from "@google/genai";
import { Groq } from "groq-sdk";
import { NextRequest, NextResponse } from "next/server";

// Initialize existing Google Gen AI client
const googleAi = new GoogleGenAI({
  apiKey: process.env.GEMINI_API_KEY || "",
});

// Initialize Groq client if key is available
const groq = process.env.GROQ_API_KEY ? new Groq({
  apiKey: process.env.GROQ_API_KEY,
}) : null;

export async function POST(req: NextRequest) {
  try {
    const { prompt, context, provider = "groq", model: requestedModel } = await req.json();

    if (!context || !context.namaSiswa) {
      return NextResponse.json(
        { error: "Nama siswa diperlukan untuk konteks pembuatan catatan." },
        { status: 400 }
      );
    }

    const systemInstruction = 
      "Anda adalah seorang konsultan psikologi anak dan kepala sekolah PAUD/TK berpengalaman di Indonesia. " +
      "Tugas Anda adalah menolong guru menyusun narasi 'Catatan Perkembangan Anak' yang mendalam, ramah, objektif, dan suportif. " +
      "Gunakan bahasa Indonesia yang baik, santun, dan menyejukkan bagi orang tua siswa. " +
      "Fokus pada hal-hal positif yang telah dikuasai anak, diikuti dengan saran stimulasi yang dapat dilakukan bersama orang tua di rumah untuk aspek yang masih perlu bimbingan. " +
      "Hindari penggunaan kata sandi teknis yang membingungkan. Berikan keluaran berupa teks narasi bersih siap salin (terdiri dari 2-3 paragraf terstruktur).";

    const fullPrompt = `Buatkan narasi perkembangan anak untuk siswa berikut:
Nama Siswa: ${context.namaSiswa}
Kelas: ${context.namaKelas || "PAUD"}
Tingkat Perkembangan Intrakurikuler:
${JSON.stringify(context.intrakurikuler, null, 2)}

Tingkat Perkembangan Kokurikuler (Projek):
${JSON.stringify(context.kokurikuler, null, 2)}

Petunjuk Tambahan dari Guru: "${prompt || "Anak aktif, ceria, dan bersosialisasi dengan baik."}"

Tolong formulasikan narasi raport komprehensif yang rapi dan menginspirasi orang tua.`;

    let resultText = "";

    // Prefer Groq if available and requested/default
    if (groq && (provider === "groq" || !process.env.GEMINI_API_KEY)) {
      const chatCompletion = await groq.chat.completions.create({
        messages: [
          { role: "system", content: systemInstruction },
          { role: "user", content: fullPrompt },
        ],
        model: requestedModel || "llama-3.3-70b-versatile",
        temperature: 0.7,
      });
      resultText = chatCompletion.choices[0]?.message?.content || "";
    } else {
      // Fallback to Gemini
      let modelName = requestedModel || "gemini-3.5-flash";
      if (modelName === "gemini-1.5-flash" || modelName === "gemini-1.5-pro") {
        modelName = "gemini-3.5-flash";
      }
      const response = await googleAi.models.generateContent({
        model: modelName,
        contents: fullPrompt,
        config: {
          systemInstruction,
          temperature: 0.7,
        },
      });
      resultText = response.text || "";
    }

    if (!resultText) {
      throw new Error("Gagal mendapatkan respon dari AI.");
    }

    return NextResponse.json({ text: resultText });
  } catch (error: any) {
    console.error("AI Generation Error:", error);
    return NextResponse.json(
      { error: error?.message || "Terjadi kesalahan internal server dalam memproses AI." },
      { status: 500 }
    );
  }
}
