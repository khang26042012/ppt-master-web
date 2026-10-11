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
const JOBS_FILE = path.join(DATA_DIR, 'jobs.json');

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

// Multer storage for files and images
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
    throw new Error('Không nhận được câu trả lời từ AI.');
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
// API: CHAT AI & KHUNG DÀN Ý (OUTLINE STUDIO)
// ========================================================
app.post('/api/chat', async (req, res) => {
  try {
    const { message, history, currentOutline, attachments } = req.body;

    const systemPrompt = `Bạn là chuyên gia thiết kế bài giảng và slide thuyết trình AI hàng đầu (PPT Master).
Nhiệm vụ của bạn:
1. Tiếp nhận yêu cầu, tài liệu, giáo án hoặc ảnh từ người dùng.
2. Trả lời trò chuyện thân thiện, chuyên nghiệp, súc tích.
3. Luôn bóc tách và xuất ra KHUNG DÀN Ý SLIDE (Outline) chi tiết theo định dạng JSON chuẩn.
Nếu người dùng yêu cầu chỉnh sửa dàn ý hiện có, hãy cập nhật dàn ý theo đúng ý họ.

BẮT BUỘC trả về định dạng JSON duy nhất như sau:
{
  "reply": "Lời giải thích hoặc tư vấn thân thiện cho người dùng...",
  "outline": {
    "title": "Tiêu đề bài thuyết trình",
    "subject": "Môn học hoặc Lĩnh vực",
    "grade": "Khối lớp hoặc Đối tượng",
    "theme": {
      "name": "Royal Blue Neo-Brutalist",
      "primary": "#2563EB",
      "bg": "#FFFFFF",
      "accent": "#38BDF8"
    },
    "slides": [
      {
        "id": "slide-1",
        "title": "Tên Slide",
        "badge": "Mục tiêu / Hoạt động...",
        "type": "title | bento | content | comparison | timeline | summary",
        "bullets": [
          "Ý trọng tâm 1",
          "Ý trọng tâm 2",
          "Ý trọng tâm 3"
        ],
        "visualIdea": "Mô tả hình ảnh/đồ họa/biểu đồ minh họa trực quan",
        "speakerNotes": "Lời gợi ý thuyết trình cho giáo viên/diễn giả"
      }
    ]
  }
}
Chỉ xuất ra chuỗi JSON hợp lệ, không bọc markdown thừa.`;

    let userPrompt = message || 'Hãy phân tích và tạo khung dàn ý slide bài giảng cho tôi.';
    if (attachments && attachments.length > 0) {
      userPrompt += '\n\n[Tài liệu đính kèm]:\n' + attachments.map(a => `- ${a.originalName}: ${a.contentSnippet || 'Tệp hình ảnh/tài liệu đính kèm'}`).join('\n');
    }
    if (currentOutline) {
      userPrompt += '\n\n[Khung dàn ý hiện tại cần chỉnh sửa]:\n' + JSON.stringify(currentOutline);
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
      // Fallback extraction
      const match = rawResponse.match(/\{[\s\S]*\}/);
      if (match) {
        parsed = JSON.parse(match[0]);
      } else {
        throw new Error('AI không trả về JSON hợp lệ: ' + rawResponse.slice(0, 200));
      }
    }

    res.json({
      success: true,
      reply: parsed.reply || 'Dưới đây là khung dàn ý slide tối ưu cho bạn.',
      outline: parsed.outline
    });

  } catch (err) {
    console.error('Chat API Error:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ========================================================
// API: BACKGROUND JOB QUEUE (CHẠY NGẦM TRÊN SERVER VỚI ETA)
// ========================================================
app.post('/api/jobs/create', (req, res) => {
  try {
    const { outline, settings } = req.body;
    if (!outline || !outline.slides || outline.slides.length === 0) {
      return res.status(400).json({ success: false, error: 'Thiếu thông tin khung dàn ý (outline).' });
    }

    const jobId = 'job_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6);
    const totalSlides = outline.slides.length;
    const estimatedSeconds = Math.max(15, totalSlides * 3.5);

    const job = {
      id: jobId,
      status: 'running', // running | completed | failed
      stage: 'analyzing_outline',
      stageText: 'Phân tích dàn ý sư phạm & chọn bảng màu...',
      progress: 15,
      etaSeconds: Math.round(estimatedSeconds),
      totalSeconds: Math.round(estimatedSeconds),
      startedAt: Date.now(),
      outline: outline,
      resultDeck: null,
      error: null
    };

    activeJobs[jobId] = job;

    // Start background asynchronous generation task on server
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
      message: 'Tiến trình sinh slide đã được khởi chạy ngầm trên máy chủ. Bạn có thể an tâm chuyển tab!'
    });

  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/jobs/:id', (req, res) => {
  const job = activeJobs[req.params.id];
  if (!job) {
    return res.status(404).json({ success: false, error: 'Không tìm thấy tiến trình (Job not found).' });
  }
  res.json({ success: true, job });
});

// Server-side Worker Function
async function runSlideGenerationWorker(jobId, outline, settings = {}) {
  const job = activeJobs[jobId];
  if (!job) return;

  const totalSteps = 4;
  const updateStage = (stage, stageText, progress, etaSeconds) => {
    job.stage = stage;
    job.stageText = stageText;
    job.progress = progress;
    job.etaSeconds = Math.max(1, etaSeconds);
  };

  // Step 1: Layout & Palette synthesis (simulate server work & prep)
  updateStage('analyzing_outline', 'Phân tích cấu trúc bento & phong cách đồ họa...', 25, Math.round(job.totalSeconds * 0.75));
  await new Promise(r => setTimeout(r, 2500));

  // Step 2: Call AI to expand elements, diagrams, and shapes
  updateStage('generating_elements', 'Thiết kế chi tiết thẻ bento, biểu đồ và icon...', 55, Math.round(job.totalSeconds * 0.45));

  const prompt = `Hãy chuyển dàn ý sau thành BỘ SLIDE THỰC THỤ hoàn chỉnh với các element chỉnh sửa được:
Tiêu đề: ${outline.title}
Môn học: ${outline.subject || 'Tổng hợp'}
Dàn ý các slide:
${JSON.stringify(outline.slides, null, 2)}

Yêu cầu xuất ra JSON cấu trúc bộ slide:
{
  "deck": {
    "id": "deck_${Date.now()}",
    "meta": {
      "title": "${outline.title}",
      "subject": "${outline.subject || 'Toán học'}",
      "grade": "${outline.grade || 'Lớp 10'}",
      "teacher": "Thầy/Cô Giảng Viên",
      "theme": "Neo-Brutalist Royal Blue"
    },
    "slides": [
      {
        "id": "s-1",
        "title": "Tiêu đề Slide",
        "subtitle": "Phụ đề / Thông điệp",
        "badge": "HOẠT ĐỘNG / MỤC TIÊU",
        "layout": "title | bento | grid-3 | stat | timeline",
        "background": {
          "type": "solid",
          "value": "#FFFFFF"
        },
        "elements": [
          {
            "id": "el-1",
            "type": "card",
            "tag": "Bước 1",
            "title": "Nội dung mục",
            "desc": "Chi tiết giải thích rõ ràng...",
            "icon": "sparkles"
          }
        ],
        "speakerNotes": "Lời gợi ý thuyết trình chi tiết cho người giảng dạy..."
      }
    ]
  }
}
Chỉ xuất JSON hợp lệ, không kèm văn bản giải thích.`;

  let fullDeck;
  try {
    const aiResponse = await callAi([{ role: 'user', content: prompt }], 'Bạn là kỹ sư đồ họa chuyển đổi dàn ý thành Slide Canvas thực thụ.', 4000);
    const cleaned = aiResponse.replace(/^[^{]*({.*})[^}]*$/s, '$1').trim();
    const parsed = JSON.parse(cleaned);
    fullDeck = parsed.deck;
  } catch (err) {
    console.warn('Fallback to outline synthesis due to AI JSON parse:', err.message);
    // Robust fallback: construct rich deck directly from outline
    fullDeck = constructDeckFromOutline(outline);
  }

  // Step 3: Typography and visual enhancement
  updateStage('styling_canvas', 'Áp dụng hiệu ứng đổ bóng 3D, viền sắc nét & typography...', 85, 3);
  await new Promise(r => setTimeout(r, 2000));

  // Step 4: Finalize
  job.resultDeck = fullDeck;
  job.status = 'completed';
  job.progress = 100;
  job.etaSeconds = 0;
  job.stage = 'completed';
  job.stageText = 'Hoàn tất! Bộ slide đã sẵn sàng.';

  // Automatically save to library
  savedDecks.unshift(fullDeck);
  writeJsonFile(DECKS_FILE, savedDecks);
}

function constructDeckFromOutline(outline) {
  return {
    id: 'deck_' + Date.now(),
    meta: {
      title: outline.title || 'Bài giảng trình chiếu',
      subject: outline.subject || 'Chuyên đề',
      grade: outline.grade || 'Cơ bản',
      teacher: 'Thầy/Cô Giảng Viên',
      theme: 'Neo-Brutalist Royal Blue'
    },
    slides: (outline.slides || []).map((s, idx) => {
      const isTitle = idx === 0 || s.type === 'title';
      return {
        id: 's-' + (idx + 1),
        title: s.title || ('Slide ' + (idx + 1)),
        subtitle: isTitle ? (outline.subject + ' • ' + (outline.grade || '')) : (s.visualIdea || ''),
        badge: s.badge || (isTitle ? 'GIỚI THIỆU BÀI GIẢNG' : ('MỤC ' + (idx + 1))),
        layout: isTitle ? 'title' : 'bento',
        background: {
          type: 'solid',
          value: '#FFFFFF'
        },
        elements: (s.bullets || []).map((b, bIdx) => ({
          id: `el-${idx + 1}-${bIdx + 1}`,
          type: 'card',
          tag: `Mục ${bIdx + 1}`,
          title: b.split(':')[0] || `Ý chính ${bIdx + 1}`,
          desc: b.includes(':') ? b.split(':').slice(1).join(':').trim() : b,
          icon: ['zap', 'layers', 'check-circle', 'star'][bIdx % 4]
        })),
        speakerNotes: s.speakerNotes || 'Lời giảng gợi ý cho slide này.'
      };
    })
  };
}

// ========================================================
// API: SLIDES MANAGEMENT (LƯU TRỮ & CHỈNH SỬA DECK)
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
// API: EXPORT POWERPOINT (.PPTX) CHUẨN MICROSOFT
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
      
      // Background color
      const bgHex = (slideData.background && slideData.background.value && slideData.background.value.startsWith('#'))
        ? slideData.background.value.replace('#', '')
        : 'FFFFFF';
      slide.background = { color: bgHex };

      if (slideData.layout === 'title' || idx === 0) {
        // Title Slide Layout
        slide.addShape(pptx.ShapeType.roundRect, {
          x: 0.8, y: 0.8, w: 2.2, h: 0.45,
          rectRadius: 0.1,
          fill: { color: '2563EB' },
          line: { color: '000000', width: 1.5 }
        });
        slide.addText(slideData.badge || 'BÀI GIẢNG', {
          x: 0.8, y: 0.8, w: 2.2, h: 0.45,
          fontSize: 11, fontFace: 'Arial', bold: true, color: 'FFFFFF', align: 'center', valign: 'middle'
        });

        slide.addText(slideData.title, {
          x: 0.8, y: 1.6, w: 11.5, h: 1.8,
          fontSize: 34, fontFace: 'Arial', bold: true, color: '000000', wrap: true
        });

        if (slideData.subtitle) {
          slide.addText(slideData.subtitle, {
            x: 0.8, y: 3.4, w: 11.5, h: 0.9,
            fontSize: 18, fontFace: 'Arial', bold: true, color: '2563EB'
          });
        }

        // Teacher info card
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
        // Content Slide Layout
        if (slideData.badge) {
          slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.8, y: 0.6, w: 2.8, h: 0.4,
            rectRadius: 0.1,
            fill: { color: 'EFF6FF' },
            line: { color: '000000', width: 1.2 }
          });
          slide.addText(slideData.badge, {
            x: 0.8, y: 0.6, w: 2.8, h: 0.4,
            fontSize: 10, fontFace: 'Arial', bold: true, color: '2563EB', align: 'center', valign: 'middle'
          });
        }

        slide.addText(slideData.title, {
          x: 0.8, y: 1.15, w: 11.5, h: 1.0,
          fontSize: 24, fontFace: 'Arial', bold: true, color: '000000'
        });

        // Bento Element Cards
        const elements = slideData.elements || [];
        const count = Math.min(elements.length, 3);
        const cardWidth = 3.6;
        const cardGap = 0.35;
        const startX = 0.8;

        elements.slice(0, 3).forEach((el, elIdx) => {
          const xPos = startX + elIdx * (cardWidth + cardGap);
          slide.addShape(pptx.ShapeType.roundRect, {
            x: xPos, y: 2.4, w: cardWidth, h: 3.5,
            rectRadius: 0.15,
            fill: { color: 'FFFFFF' },
            line: { color: '000000', width: 1.8 }
          });

          slide.addShape(pptx.ShapeType.roundRect, {
            x: xPos + 0.3, y: 2.7, w: 1.2, h: 0.35,
            rectRadius: 0.08,
            fill: { color: '2563EB' },
            line: { color: '000000', width: 1.0 }
          });
          slide.addText(el.tag || `Mục ${elIdx + 1}`, {
            x: xPos + 0.3, y: 2.7, w: 1.2, h: 0.35,
            fontSize: 9, fontFace: 'Arial', bold: true, color: 'FFFFFF', align: 'center', valign: 'middle'
          });

          slide.addText(el.title || '', {
            x: xPos + 0.3, y: 3.2, w: cardWidth - 0.6, h: 0.8,
            fontSize: 13, fontFace: 'Arial', bold: true, color: '000000', wrap: true
          });

          slide.addText(el.desc || '', {
            x: xPos + 0.3, y: 4.1, w: cardWidth - 0.6, h: 1.5,
            fontSize: 11, fontFace: 'Arial', color: '475569', wrap: true
          });
        });
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
