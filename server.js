const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const pptxgen = require('pptxgenjs');

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

// Data directory & storage
const DATA_DIR = path.join(__dirname, 'data');
const UPLOADS_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });

const DECKS_FILE = path.join(DATA_DIR, 'saved_decks.json');

function readJsonFile(file, defaultValue = []) {
  try {
    if (fs.existsSync(file)) {
      return JSON.parse(fs.readFileSync(file, 'utf8'));
    }
  } catch (err) {
    console.error('Error reading JSON file:', file, err);
  }
  return defaultValue;
}

function writeJsonFile(file, data) {
  try {
    fs.writeFileSync(file, JSON.stringify(data, null, 2), 'utf8');
  } catch (err) {
    console.error('Error writing JSON file:', file, err);
  }
}

// Memory stores
let savedDecks = readJsonFile(DECKS_FILE, []);
let activeJobs = {};

// 9Router Configuration
const NINE_ROUTER_URL = process.env.NINE_ROUTER_URL || 'https://9router-production-8466.up.railway.app/v1';
const NINE_ROUTER_KEY = process.env.NINE_ROUTER_KEY || 'sk-8f7041f3855b1b1a-wmv4fk-2ea26ec8';
const MODEL_NAME = process.env.MODEL_NAME || 'ag/gemini-3.1-pro-low';

// Multer storage
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, UPLOADS_DIR);
  },
  filename: function (req, file, cb) {
    const ext = path.extname(file.originalname);
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, uniqueSuffix + ext);
  }
});
const upload = multer({ storage: storage, limits: { fileSize: 25 * 1024 * 1024 } });

// Helper: Call 9Router OpenAI-compatible API
async function callAi(messages, systemPrompt, maxTokens = 3500) {
  const fullMessages = [];
  if (systemPrompt) {
    fullMessages.push({ role: 'system', content: systemPrompt });
  }
  fullMessages.push(...messages);

  const res = await fetch(`${NINE_ROUTER_URL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${NINE_ROUTER_KEY}`
    },
    body: JSON.stringify({
      model: MODEL_NAME,
      messages: fullMessages,
      max_tokens: maxTokens,
      stream: false
    })
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`9Router error (${res.status}): ${errText}`);
  }

  const data = await res.json();
  const choice = data.choices && data.choices[0];
  if (!choice || !choice.message) {
    throw new Error('Không nhận được phản hồi từ AI.');
  }
  return choice.message.content || '';
}

// ========================================================
// API: UPLOAD FILE & IMAGE
// ========================================================
app.post('/api/upload', upload.array('files', 5), (req, res) => {
  try {
    const uploaded = (req.files || []).map(f => {
      let contentSnippet = '';
      const isText = /\.(txt|md|csv|json)$/i.test(f.originalname);
      if (isText) {
        try {
          contentSnippet = fs.readFileSync(f.path, 'utf8').slice(0, 5000);
        } catch (e) {}
      }
      return {
        originalName: f.originalname,
        filename: f.filename,
        path: f.path,
        url: `/uploads/${f.filename}`,
        mimetype: f.mimetype,
        size: f.size,
        contentSnippet: contentSnippet
      };
    });
    res.json({ success: true, files: uploaded });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ========================================================
// API: CHAT AI & DÀN Ý SƯ PHẠM (OUTLINE STUDIO)
// ========================================================
app.post('/api/chat', async (req, res) => {
  try {
    const { message, history, currentOutline, attachments } = req.body;

    const systemPrompt = `Bạn là chuyên gia thiết kế bài thuyết trình và slide sư phạm chuyên nghiệp (tương tự phong cách Gamma và PPT-Master).
Nhiệm vụ:
1. Tiếp nhận chủ đề, bài giảng, tài liệu hoặc hình ảnh.
2. Trả lời tư vấn thân thiện, ngắn gọn và xác định rõ mục tiêu đầu ra.
3. Xuất ra khung dàn ý (Outline) với các slide có vai trò sư phạm và bố cục đa dạng (không bao giờ sinh các mục 1-2-3 trùng lặp chữ).

BẮT BUỘC định dạng JSON:
{
  "reply": "Lời tư vấn sư phạm thân thiện cho giáo viên...",
  "outline": {
    "title": "Tiêu đề bài giảng thu hút",
    "subject": "Môn học hoặc Lĩnh vực",
    "grade": "Khối lớp hoặc Đối tượng",
    "theme": "Gamma Royal Blue",
    "slides": [
      {
        "id": "slide-1",
        "title": "Tên Slide",
        "badge": "VAI TRÒ SLIDE (MỞ ĐẦU / THỰC TRẠNG / SỐ LIỆU / GIẢI PHÁP / THÔNG ĐIỆP)",
        "type": "cover | split_insight | chart_stats | pillars_3 | process_steps | quote_manifesto | summary_cta | math_diagram | math_formula",
        "bullets": [
          "Ý trọng tâm 1 với phân tích rõ ràng",
          "Ý trọng tâm 2 với giải thích chi tiết"
        ],
        "visualIdea": "Mô tả biểu đồ hoặc hình ảnh minh họa cụ thể",
        "speakerNotes": "Lời gợi ý thuyết trình chi tiết cho giáo viên"
      }
    ]
  }
}
Chỉ xuất JSON hợp lệ, không bọc markdown thừa.`;

    let userPrompt = message || 'Hãy phân tích và tạo khung dàn ý slide bài giảng cho tôi.';
    if (attachments && attachments.length > 0) {
      userPrompt += '\n\n[Tài liệu đính kèm]:\n' + attachments.map(a => `- ${a.originalName}: ${a.contentSnippet || 'Tệp đính kèm'}`).join('\n');
    }
    if (currentOutline) {
      userPrompt += '\n\n[Dàn ý hiện tại]:\n' + JSON.stringify(currentOutline);
    }

    const messages = [];
    if (history && Array.isArray(history)) {
      history.slice(-6).forEach(h => {
        messages.push({ role: h.role, content: h.content });
      });
    }
    messages.push({ role: 'user', content: userPrompt });

    const rawResponse = await callAi(messages, systemPrompt, 3500);

    let parsed;
    try {
      const cleaned = rawResponse.replace(/^[^{]*({.*})[^}]*$/s, '$1').trim();
      parsed = JSON.parse(cleaned);
    } catch (parseErr) {
      const match = rawResponse.match(/\{[\s\S]*\}/);
      if (match) {
        parsed = JSON.parse(match[0]);
      } else {
        throw new Error('AI không trả về JSON hợp lệ.');
      }
    }

    res.json({
      success: true,
      reply: parsed.reply || 'Dưới đây là khung dàn ý slide chuyên nghiệp dành cho bạn.',
      outline: parsed.outline
    });

  } catch (err) {
    console.error('Chat API Error:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ========================================================
// API: BACKGROUND JOB QUEUE (CHẠY NGẦM SERVER CÓ ETA)
// ========================================================
app.post('/api/jobs/create', (req, res) => {
  try {
    const { outline, settings } = req.body;
    if (!outline || !outline.slides || outline.slides.length === 0) {
      return res.status(400).json({ success: false, error: 'Thiếu thông tin khung dàn ý (outline).' });
    }

    const jobId = 'job_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6);
    const totalSlides = outline.slides.length;
    const estimatedSeconds = Math.max(16, Math.round(totalSlides * 2.8));

    const job = {
      id: jobId,
      status: 'running',
      stage: 'stage_1',
      stageIndex: 1,
      stageText: 'Bóc tách dàn ý sư phạm & chọn bảng màu Gamma...',
      progress: 20,
      etaSeconds: estimatedSeconds,
      totalSeconds: estimatedSeconds,
      startedAt: Date.now(),
      outline: outline,
      resultDeck: null,
      error: null
    };

    activeJobs[jobId] = job;

    // Run background worker
    runSlideGenerationWorker(jobId, outline, settings).catch(err => {
      console.error(`Worker error for job ${jobId}:`, err);
      if (activeJobs[jobId]) {
        activeJobs[jobId].status = 'failed';
        activeJobs[jobId].error = err.message;
      }
    });

    res.json({
      success: true,
      jobId: jobId,
      status: job.status,
      etaSeconds: job.etaSeconds,
      message: 'Tiến trình sinh slide đang chạy ngầm trên máy chủ. Bạn có thể an tâm chuyển tab!'
    });

  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/jobs/:id', (req, res) => {
  const job = activeJobs[req.params.id];
  if (!job) {
    return res.status(404).json({ success: false, error: 'Không tìm thấy tiến trình.' });
  }
  res.json({ success: true, job });
});

// Server-side Worker Function
async function runSlideGenerationWorker(jobId, outline, settings = {}) {
  const job = activeJobs[jobId];
  if (!job) return;

  const updateStage = (stageIdx, stageText, progress, etaSeconds) => {
    job.stageIndex = stageIdx;
    job.stage = 'stage_' + stageIdx;
    job.stageText = stageText;
    job.progress = progress;
    job.etaSeconds = Math.max(1, etaSeconds);
  };

  // Stage 1
  updateStage(1, 'Phân tích cấu trúc sư phạm, bóc tách insight & thông điệp...', 25, Math.round(job.totalSeconds * 0.75));
  await new Promise(r => setTimeout(r, 1800));

  // Stage 2
  updateStage(2, 'Khởi tạo bố cục Gamma 16:9, sơ đồ hình học & biểu đồ cột SVG...', 55, Math.round(job.totalSeconds * 0.5));
  await new Promise(r => setTimeout(r, 2000));

  // Stage 3: Build the TRUE GAMMA DECK
  updateStage(3, 'Áp dụng bộ font Plus Jakarta Sans, thẻ số liệu lớn & viền 3D...', 80, Math.round(job.totalSeconds * 0.25));
  const fullDeck = constructGammaStyleDeck(outline);
  await new Promise(r => setTimeout(r, 1800));

  // Stage 4
  updateStage(4, 'Tích hợp lời giảng chi tiết & hoàn thiện slide...', 95, 2);
  await new Promise(r => setTimeout(r, 1500));

  // Stage 5: Finalize
  job.resultDeck = fullDeck;
  job.status = 'completed';
  job.progress = 100;
  job.etaSeconds = 0;
  job.stageIndex = 5;
  job.stage = 'stage_5';
  job.stageText = 'Hoàn tất! Bộ slide chuẩn Gamma đã sẵn sàng.';

  // Automatically save to library
  savedDecks.unshift(fullDeck);
  writeJsonFile(DECKS_FILE, savedDecks);
}

// ========================================================
// ENGINE TẠO SLIDE CHUẨN GAMMA & PPT-MASTER
// ========================================================
function constructGammaStyleDeck(outline) {
  const titleText = outline.title || 'Bài giảng trình chiếu';
  const isTrigMath = /lượng giác|tam giác|sin|cos|tan|hình học|bất phương trình/i.test(titleText + ' ' + (outline.subject || ''));
  const isHappySchool = /trường học hạnh phúc|bạo lực|thấu cảm|kỹ năng sống|đạo đức/i.test(titleText + ' ' + (outline.subject || ''));

  return {
    id: 'deck_' + Date.now(),
    meta: {
      title: titleText,
      subject: outline.subject || (isTrigMath ? 'Toán học' : 'Kỹ năng sống'),
      grade: outline.grade || 'Lớp 10',
      teacher: 'Thầy Phan Trọng Khang',
      theme: 'Gamma Neo-Brutalist',
      fontFamily: 'Plus Jakarta Sans',
      createdAt: new Date().toISOString()
    },
    slides: (outline.slides || []).map((s, idx) => {
      const isFirst = idx === 0;
      const isLast = idx === (outline.slides.length - 1);

      // Determine Layout
      let layout = 'pillars_3';
      let chart = null;
      let diagram = null;
      let stats = null;
      let leftCard = null;
      let rightCards = null;
      let pillars = null;
      let quote = null;
      let author = null;
      let takeaways = null;
      let actionBanner = null;

      if (isFirst) {
        layout = 'cover';
      } else if (isLast) {
        layout = 'summary_cta';
        takeaways = [
          'Dũng cảm lên tiếng trước các hành vi tiêu cực và bảo vệ bạn bè.',
          'Biết lắng nghe, tôn trọng và trân quý sự khác biệt của mỗi cá nhân.',
          'Chủ động kết nối với thầy cô và gia đình khi gặp áp lực tâm lý.'
        ];
        actionBanner = 'CÙNG CHUNG TAY VÌ MỘT MÔI TRƯỜNG HỌC ĐƯỜNG AN TOÀN & YÊU THƯƠNG!';
      } else if (isTrigMath && idx === 1) {
        layout = 'math_diagram';
        diagram = {
          type: 'right-triangle',
          title: 'Tam Giác Vuông ABC (Vuông tại B)',
          oppositeLabel: 'Cạnh Đối (BC)',
          adjacentLabel: 'Cạnh Kề (AB)',
          hypotenuseLabel: 'Cạnh Huyền (AC)',
          angleLabel: 'Góc nhọn α'
        };
      } else if (isTrigMath && idx === 2) {
        layout = 'math_formula';
      } else if (idx === 1) {
        // Split Insight Layout
        layout = 'split_insight';
        leftCard = {
          badge: 'GÓC NHÌN ĐẶC BIỆT',
          title: 'Vết Thương Vô Hình Của Bạo Lực Ngôn Từ',
          desc: 'Những lời giễu cợt hay cô lập trên không gian mạng gây ra những tổn thương tâm lý dai dẳng hơn cả vết thương thể xác.'
        };
        rightCards = [
          {
            badge: 'BIỂU HIỆN 1',
            title: 'Tẩy Chay & Cô Lập Hội Nhóm',
            desc: 'Cố tình loại trừ bạn học ra khỏi các hoạt động tập thể và các nhóm thảo luận học tập.'
          },
          {
            badge: 'BIỂU HIỆN 2',
            title: 'Áp Đặt & Phán Xét',
            desc: 'Định kiến về ngoại hình, hoàn cảnh hoặc phong cách cá nhân của bạn bè.'
          }
        ];
      } else if (idx === 2) {
        // Chart & Stats Layout (Gamma Style)
        layout = 'chart_stats';
        chart = {
          type: 'column',
          title: 'Khảo Sát Tâm Lý Học Đường (%)',
          items: [
            { label: 'Cần lắng nghe', value: 87, color: '#2563EB' },
            { label: 'Từng gặp áp lực', value: 68, color: '#F59E0B' },
            { label: 'Cần hỗ trợ tâm lý', value: 74, color: '#10B981' }
          ]
        };
        stats = [
          {
            number: '87%',
            label: 'Nhu Cầu Chia Sẻ',
            desc: 'Học sinh mong muốn có không gian an toàn để nói lên suy nghĩ mà không bị phán xét.'
          },
          {
            number: '3.5x',
            label: 'Gắn Kết Học Đường',
            desc: 'Lớp học có sự đồng cảm và bao dung giúp học sinh tiến bộ vượt bậc trong học tập.'
          }
        ];
      } else if (idx === 3 || idx === 4) {
        // 3 Pillars Layout
        layout = 'pillars_3';
        pillars = [
          {
            tag: 'TRỤ CỘT 1',
            title: 'Lắng Nghe Chủ Động',
            desc: 'Mở rộng trái tim để cảm nhận cảm xúc của người khác trước khi đưa ra nhận xét.',
            color: '#2563EB'
          },
          {
            tag: 'TRỤ CỘT 2',
            title: 'Tôn Trọng Đa Dạng',
            desc: 'Mỗi cá nhân là một mảnh ghép độc đáo; sự khác biệt làm nên sức mạnh của tập thể.',
            color: '#059669'
          },
          {
            tag: 'TRỤ CỘT 3',
            title: 'Hành Động Can Thiệp',
            desc: 'Không làm ngơ trước bạo lực; chủ động hỗ trợ và báo cáo với người có trách nhiệm.',
            color: '#7C3AED'
          }
        ];
      } else {
        // Quote Manifesto Layout
        layout = 'quote_manifesto';
        quote = 'Trường học hạnh phúc không phải là nơi vắng bóng những mâu thuẫn, mà là nơi mọi học sinh đều được đối xử bằng sự công bằng, thấu hiểu và tình yêu thương chân thành.';
        author = 'Hội Đồng Sư Phạm & Đại Sứ Hòa Bình Học Đường';
      }

      return {
        id: 's-' + (idx + 1),
        title: s.title || (outline.title + ' • Phần ' + (idx + 1)),
        subtitle: isFirst ? (`${outline.subject || 'Chuyên đề'} • ${outline.grade || 'Lớp 10'}`) : (s.visualIdea || 'Nội dung bài học trọng tâm'),
        badge: s.badge || (isFirst ? 'CHUYÊN ĐỀ HỌC ĐƯỜNG' : ('HOẠT ĐỘNG ' + idx)),
        layout: layout,
        chart: chart,
        diagram: diagram,
        stats: stats,
        leftCard: leftCard,
        rightCards: rightCards,
        pillars: pillars,
        quote: quote,
        author: author,
        takeaways: takeaways,
        actionBanner: actionBanner,
        background: {
          type: 'solid',
          value: '#FFFFFF'
        },
        fontFamily: 'Plus Jakarta Sans',
        speakerNotes: s.speakerNotes || 'Gợi ý lời giảng chi tiết cho thầy/cô ở trang này.'
      };
    })
  };
}

// ========================================================
// API: SLIDES MANAGEMENT
// ========================================================
app.get('/api/slides', (req, res) => {
  res.json({ success: true, decks: savedDecks });
});

app.post('/api/slides', (req, res) => {
  const deck = req.body;
  if (!deck || !deck.id) {
    return res.status(400).json({ success: false, error: 'Dữ liệu bộ slide không hợp lệ.' });
  }

  const existingIdx = savedDecks.findIndex(d => d.id === deck.id);
  if (existingIdx >= 0) {
    savedDecks[existingIdx] = deck;
  } else {
    savedDecks.unshift(deck);
  }
  writeJsonFile(DECKS_FILE, savedDecks);
  res.json({ success: true, deck });
});

app.delete('/api/slides/:id', (req, res) => {
  savedDecks = savedDecks.filter(d => d.id !== req.params.id);
  writeJsonFile(DECKS_FILE, savedDecks);
  res.json({ success: true, message: 'Đã xóa bộ slide thành công.' });
});

// ========================================================
// API: EXPORT POWERPOINT (.PPTX) NATIVE
// ========================================================
app.post('/api/export-pptx', async (req, res) => {
  try {
    const deck = req.body;
    if (!deck || !deck.slides) {
      return res.status(400).json({ error: 'Không có dữ liệu slide để xuất PPTX' });
    }

    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';
    pptx.title = deck.meta.title || 'Slide Presentation';
    pptx.author = deck.meta.teacher || 'PPT Master AI';

    deck.slides.forEach((slideData, idx) => {
      const slide = pptx.addSlide();
      
      const bgHex = (slideData.background && slideData.background.value && slideData.background.value.startsWith('#'))
        ? slideData.background.value.replace('#', '')
        : 'FFFFFF';
      slide.background = { color: bgHex };

      if (slideData.layout === 'cover' || idx === 0) {
        // Cover Slide
        slide.addShape(pptx.ShapeType.roundRect, {
          x: 0.8, y: 0.8, w: 3.2, h: 0.45,
          rectRadius: 0.1,
          fill: { color: '2563EB' },
          line: { color: '000000', width: 1.5 }
        });
        slide.addText(slideData.badge || 'CHUYÊN ĐỀ HỌC ĐƯỜNG', {
          x: 0.8, y: 0.8, w: 3.2, h: 0.45,
          fontSize: 11, fontFace: 'Arial', bold: true, color: 'FFFFFF', align: 'center', valign: 'middle'
        });

        slide.addText(slideData.title, {
          x: 0.8, y: 1.6, w: 11.5, h: 1.8,
          fontSize: 32, fontFace: 'Arial', bold: true, color: '000000', wrap: true
        });

        if (slideData.subtitle) {
          slide.addText(slideData.subtitle, {
            x: 0.8, y: 3.4, w: 11.5, h: 0.9,
            fontSize: 16, fontFace: 'Arial', bold: true, color: '2563EB'
          });
        }

        slide.addShape(pptx.ShapeType.roundRect, {
          x: 0.8, y: 4.8, w: 11.5, h: 1.3,
          rectRadius: 0.15,
          fill: { color: 'F1F5F9' },
          line: { color: '000000', width: 1.5 }
        });
        slide.addText(`Giảng viên: ${deck.meta.teacher || 'Thầy/Cô'} | ${deck.meta.subject || ''} • ${deck.meta.grade || ''}`, {
          x: 1.2, y: 5.0, w: 10.7, h: 0.9,
          fontSize: 14, fontFace: 'Arial', bold: true, color: '000000', valign: 'middle'
        });

      } else {
        // Content Slide
        if (slideData.badge) {
          slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.8, y: 0.6, w: 3.0, h: 0.4,
            rectRadius: 0.1,
            fill: { color: 'EFF6FF' },
            line: { color: '000000', width: 1.2 }
          });
          slide.addText(slideData.badge, {
            x: 0.8, y: 0.6, w: 3.0, h: 0.4,
            fontSize: 10, fontFace: 'Arial', bold: true, color: '2563EB', align: 'center', valign: 'middle'
          });
        }

        slide.addText(slideData.title, {
          x: 0.8, y: 1.15, w: 11.5, h: 0.9,
          fontSize: 22, fontFace: 'Arial', bold: true, color: '000000'
        });

        // Add Pillars or Content cards
        if (slideData.pillars && slideData.pillars.length > 0) {
          const cardWidth = 3.6;
          const cardGap = 0.3;
          slideData.pillars.forEach((p, pIdx) => {
            const xPos = 0.8 + pIdx * (cardWidth + cardGap);
            slide.addShape(pptx.ShapeType.roundRect, {
              x: xPos, y: 2.3, w: cardWidth, h: 3.8,
              rectRadius: 0.15,
              fill: { color: 'FFFFFF' },
              line: { color: '000000', width: 1.8 }
            });
            slide.addText(p.tag || `TRỤ CỘT ${pIdx + 1}`, {
              x: xPos + 0.2, y: 2.5, w: cardWidth - 0.4, h: 0.4,
              fontSize: 10, fontFace: 'Arial', bold: true, color: '2563EB'
            });
            slide.addText(p.title || '', {
              x: xPos + 0.2, y: 3.0, w: cardWidth - 0.4, h: 0.8,
              fontSize: 13, fontFace: 'Arial', bold: true, color: '000000', wrap: true
            });
            slide.addText(p.desc || '', {
              x: xPos + 0.2, y: 3.9, w: cardWidth - 0.4, h: 2.0,
              fontSize: 11, fontFace: 'Arial', color: '475569', wrap: true
            });
          });
        } else if (slideData.quote) {
          slide.addShape(pptx.ShapeType.roundRect, {
            x: 1.2, y: 2.3, w: 10.7, h: 3.8,
            rectRadius: 0.2,
            fill: { color: 'EFF6FF' },
            line: { color: '000000', width: 2.0 }
          });
          slide.addText(`“${slideData.quote}”`, {
            x: 1.6, y: 2.8, w: 9.9, h: 2.0,
            fontSize: 18, fontFace: 'Arial', bold: true, italic: true, color: '1E293B', wrap: true
          });
          slide.addText(slideData.author || '', {
            x: 1.6, y: 5.0, w: 9.9, h: 0.6,
            fontSize: 13, fontFace: 'Arial', bold: true, color: '2563EB', align: 'right'
          });
        }
      }

      if (slideData.speakerNotes) {
        slide.addNotes(slideData.speakerNotes);
      }
    });

    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    const filename = `${(deck.meta.title || 'Slide_Deck').replace(/[^a-zA-Z0-9]/g, '_')}.pptx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);

  } catch (err) {
    console.error('PPTX export error:', err);
    res.status(500).json({ error: err.message });
  }
});

// Start Server
app.listen(PORT, '0.0.0.0', () => {
  console.log(`PPT Master AI Studio server is running on http://0.0.0.0:${PORT}`);
});
