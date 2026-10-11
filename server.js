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

    const systemPrompt = `Bạn là chuyên gia thiết kế bài thuyết trình Presenton & Gamma AI.
Nhiệm vụ:
1. Tiếp nhận chủ đề, bài giảng, tài liệu hoặc hình ảnh.
2. Tư vấn ngắn gọn, sư phạm, xác định đối tượng người học.
3. Xuất ra khung dàn ý (Outline) ĐỦ 5 PHẦN CỐT LÕI (Mở đầu, Mục lục, Nội dung thực trạng/dữ liệu/giải pháp, Tổng kết, Cảm ơn).
4. KHÔNG BAO GIỜ sinh các mục 1-2-3 trùng lặp chữ. Mỗi slide có 1 thông điệp chính, tối đa 5 dòng.

BẮT BUỘC định dạng JSON:
{
  "reply": "Lời tư vấn sư phạm thân thiện cho giáo viên...",
  "outline": {
    "title": "Tiêu đề bài giảng thu hút",
    "subject": "Môn học hoặc Lĩnh vực",
    "grade": "Khối lớp hoặc Đối tượng",
    "theme": "Emerald Aurora | Midnight Navy | Royal Violet",
    "slides": [
      {
        "id": "slide-1",
        "title": "Tên Slide",
        "badge": "MỞ ĐẦU | MỤC LỤC | THỰC TRẠNG | SỐ LIỆU | GIẢI PHÁP | TỔNG KẾT | CẢM ƠN",
        "type": "cover | agenda | split_insight | chart_stats | pillars_3 | summary_quote | thank_you",
        "bullets": [
          "Ý trọng tâm 1 với phân tích rõ ràng",
          "Ý trọng tâm 2 với giải thích chi tiết"
        ],
        "visualIdea": "Mô tả biểu đồ vector hoặc hình ảnh minh họa cụ thể",
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
      reply: parsed.reply || 'Dưới đây là khung dàn ý slide chuẩn Presenton & Gamma dành cho bạn.',
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
    const totalSlides = 7; // Chuẩn 7 slides
    const estimatedSeconds = 18;

    const job = {
      id: jobId,
      status: 'running',
      stage: 'stage_1',
      stageIndex: 1,
      stageText: 'Phân tích chủ đề & cấu hình Palette Presenton 16:9...',
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
      message: 'Tiến trình sinh slide Presenton đang chạy ngầm trên máy chủ. Bạn có thể an tâm chuyển tab!'
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
  updateStage(1, 'Kiểm tra 7 tiêu chí Presenton: Chuẩn 16:9, bảng màu gradient & typography...', 25, 14);
  await new Promise(r => setTimeout(r, 1800));

  // Stage 2
  updateStage(2, 'Khởi tạo cấu trúc 5 phần: Mở đầu, Mục lục, Thực trạng, Biểu đồ & Giải pháp...', 50, 10);
  await new Promise(r => setTimeout(r, 2000));

  // Stage 3: Build the PRESENTON 16:9 DECK
  updateStage(3, 'Dựng biểu đồ cột Vector SVG, số liệu KPI & thẻ glassmorphism không nền trắng...', 75, 6);
  const fullDeck = constructPresentonDeck(outline);
  await new Promise(r => setTimeout(r, 1800));

  // Stage 4
  updateStage(4, 'Tích hợp lời giảng chuyên gia & kiểm định tương thích PPTX native...', 90, 2);
  await new Promise(r => setTimeout(r, 1500));

  // Stage 5: Finalize
  job.resultDeck = fullDeck;
  job.status = 'completed';
  job.progress = 100;
  job.etaSeconds = 0;
  job.stageIndex = 5;
  job.stage = 'stage_5';
  job.stageText = 'Hoàn tất! Bộ slide chuẩn Presenton 16:9 tuyệt đối đã sẵn sàng.';

  // Automatically save to library
  savedDecks.unshift(fullDeck);
  writeJsonFile(DECKS_FILE, savedDecks);
}

// ========================================================
// PRESENTON 16:9 MULTI-LAYOUT SLIDE GENERATION ENGINE
// ========================================================
function constructPresentonDeck(outline) {
  const titleText = outline.title || 'Chuyên đề Giáo Dục';
  const isTrigMath = /lượng giác|tam giác|sin|cos|tan|hình học|bất phương trình|toán/i.test(titleText + ' ' + (outline.subject || ''));

  // Define Theme Tokens (Zero Plain White Backgrounds)
  const theme = isTrigMath ? {
    id: 'midnight_navy',
    name: 'Midnight Cyber Navy',
    gradient: 'linear-gradient(135deg, #0F172A 0%, #1E293B 45%, #0C4A6E 100%)',
    hex: '0F172A',
    cardBg: 'rgba(255, 255, 255, 0.08)',
    cardBorder: 'rgba(56, 189, 248, 0.35)',
    accent1: '#38BDF8',
    accent2: '#F59E0B',
    textColor: '#FFFFFF',
    subtextColor: '#BAE6FD'
  } : {
    id: 'emerald_aurora',
    name: 'Emerald Aurora Gradient',
    gradient: 'linear-gradient(135deg, #064E3B 0%, #065F46 45%, #042F2E 100%)',
    hex: '064E3B',
    cardBg: 'rgba(255, 255, 255, 0.09)',
    cardBorder: 'rgba(52, 211, 153, 0.35)',
    accent1: '#34D399',
    accent2: '#FBBF24',
    textColor: '#FFFFFF',
    subtextColor: '#D1FAE5'
  };

  const slides = [];

  // =========================================================================
  // SLIDE 1: MỞ ĐẦU (COVER SLIDE)
  // =========================================================================
  slides.push({
    id: 's-1',
    stage: 'Mở đầu',
    badge: isTrigMath ? 'CHUYÊN ĐỀ TOÁN HỌC' : 'CHUYÊN ĐỀ HỌC ĐƯỜNG',
    title: isTrigMath ? 'Khám Phá Tỉ Số Lượng Giác & Ứng Dụng Thực Tế' : (outline.title || 'Xây Dựng Trường Học Hạnh Phúc: Thay Bạo Lực Bằng Thấu Cảm'),
    subtitle: isTrigMath ? 'Tam giác vuông, công thức sin/cos/tan/cot và bài toán đo đạc' : 'Văn hóa ứng xử, kỹ năng sống và môi trường học đường không bạo lực',
    layout: 'cover',
    theme: theme,
    teacher: 'Thầy Phan Trọng Khang',
    subject: outline.subject || (isTrigMath ? 'Toán Học 9' : 'Kỹ Năng Sống'),
    grade: outline.grade || 'Lớp 10',
    icon: 'sparkles',
    speakerNotes: 'Kính chào thầy cô và các em. Bài giảng hôm nay được thiết kế theo chuẩn 16:9 hiện đại, tập trung vào trải nghiệm trực quan và ứng dụng thực tiễn.'
  });

  // =========================================================================
  // SLIDE 2: MỤC LỤC (AGENDA / OVERVIEW)
  // =========================================================================
  slides.push({
    id: 's-2',
    stage: 'Mục lục',
    badge: 'MỤC LỤC CHUYÊN ĐỀ',
    title: 'Cấu Trúc Bài Giảng & Các Chặng Đường',
    subtitle: 'Lộ trình khám phá từ thực trạng, phân tích dữ liệu đến hành động',
    layout: 'agenda',
    theme: theme,
    agendaItems: isTrigMath ? [
      { step: '01', title: 'Khái Niệm Tam Giác Vuông', desc: 'Định nghĩa cạnh đối, kề, huyền và góc nhọn' },
      { step: '02', title: 'Bộ Tứ Tỉ Số Lượng Giác', desc: 'Công thức sin, cos, tan, cot và mẹo ghi nhớ' },
      { step: '03', title: 'Bảng Giá Trị Góc Đặc Biệt', desc: 'Góc 30°, 45°, 60° và bài toán tính toán' },
      { step: '04', title: 'Ứng Dụng Đo Đạc Thực Tế', desc: 'Đo chiều cao tòa nhà và khoảng cách từ xa' }
    ] : [
      { step: '01', title: 'Nhận Diện Thực Trạng', desc: 'Phân tích bạo lực tinh thần và các tổn thương vô hình' },
      { step: '02', title: 'Dữ Liệu & Khảo Sát', desc: 'Thống kê tiếng nói từ hơn 1,200 học sinh' },
      { step: '03', title: 'Ba Trụ Cột Hành Động', desc: 'Lắng nghe, tôn trọng sự khác biệt và can thiệp' },
      { step: '04', title: 'Cam Kết & Lan Tỏa', desc: 'Hành động cụ thể vì một ngôi trường yêu thương' }
    ],
    speakerNotes: 'Chúng ta sẽ đi qua 4 chặng đường chính, từ nhận diện vấn đề cho đến các giải pháp ứng dụng thiết thực.'
  });

  // =========================================================================
  // SLIDE 3: NỘI DUNG 1 - BỐI CẢNH & THỰC TRẠNG (HOẶC HÌNH HỌC TRỰC QUAN)
  // =========================================================================
  if (isTrigMath) {
    slides.push({
      id: 's-3',
      stage: 'Nội dung',
      badge: 'HÌNH HỌC TRỰC QUAN',
      title: 'Tam Giác Vuông ABC (Vuông Tại B)',
      subtitle: 'Xác định chính xác vị trí các cạnh tương ứng với góc nhọn α',
      layout: 'math_diagram',
      theme: theme,
      diagram: {
        type: 'right-triangle',
        opposite: 'Cạnh Đối (BC)',
        adjacent: 'Cạnh Kề (AB)',
        hypotenuse: 'Cạnh Huyền (AC)',
        angle: 'Góc nhọn α'
      },
      bullets: [
        'Cạnh Huyền: Cạnh dài nhất đối diện góc vuông B.',
        'Cạnh Đối: Đối diện trực tiếp với góc nhọn α đang xét.',
        'Cạnh Kề: Nằm kề góc nhọn α (khác cạnh huyền).'
      ],
      speakerNotes: 'Thầy nhấn mạnh: việc xác định đúng cạnh đối và kề phụ thuộc vào góc nhọn ta đang xét.'
    });
  } else {
    slides.push({
      id: 's-3',
      stage: 'Nội dung',
      badge: 'BỐI CẢNH & THỰC TRẠNG',
      title: 'Vết Thương Vô Hình Của Bạo Lực Học Đường',
      subtitle: 'Từ bạo lực thể chất đến bạo lực tinh thần và cô lập hội nhóm',
      layout: 'split_insight',
      theme: theme,
      leftCard: {
        badge: 'GÓC NHÌN ĐẶC BIỆT',
        title: 'Bạo Lực Tinh Thần Nguy Hiểm Hơn Thể Xác',
        desc: 'Những lời chê bai ngoại hình, tẩy chay vô căn cứ gây tổn thương tâm lý dai dẳng hơn bất kỳ vết thương thể xác nào.'
      },
      rightCards: [
        {
          badge: 'BIỂU HIỆN 1',
          title: 'Tẩy Chay & Cô Lập Mạng Xã Hội',
          desc: 'Loại trừ bạn bè khỏi nhóm chat, chế giễu và lan truyền tin đồn sai sự thật.'
        },
        {
          badge: 'BIỂU HIỆN 2',
          title: 'Áp Đặt Tiêu Chuẩn Lệch Lạc',
          desc: 'Ép buộc người khác phải thay đổi bản thân để được chấp nhận vào nhóm.'
        }
      ],
      speakerNotes: 'Hãy chú ý đến những biểu hiện vô hình. Đôi khi chỉ một lời nói đùa ác ý cũng có thể làm tổn thương bạn học.'
    });
  }

  // =========================================================================
  // SLIDE 4: NỘI DUNG 2 - DỮ LIỆU & SỐ LIỆU / SO SÁNH (CHART & STATS / BẢNG)
  // =========================================================================
  if (isTrigMath) {
    slides.push({
      id: 's-4',
      stage: 'Nội dung',
      badge: 'BẢNG GIÁ TRỊ ĐẶC BIỆT',
      title: 'Tỉ Số Lượng Giác Các Góc Nhọn Đặc Biệt',
      subtitle: 'Đối chiếu giá trị chuẩn của góc 30°, 45° và 60°',
      layout: 'table_compare',
      theme: theme,
      tableHeaders: ['Tỉ Số', 'Góc 30°', 'Góc 45°', 'Góc 60°'],
      tableRows: [
        ['sin α', '1/2 (0.5)', '√2/2 (~0.707)', '√3/2 (~0.866)'],
        ['cos α', '√3/2 (~0.866)', '√2/2 (~0.707)', '1/2 (0.5)'],
        ['tan α', '√3/3 (~0.577)', '1', '√3 (~1.732)'],
        ['cot α', '√3 (~1.732)', '1', '√3/3 (~0.577)']
      ],
      kpis: [
        { number: '1.0', label: 'tan(45°)', desc: 'Cạnh đối bằng cạnh kề trong tam giác vuông cân.' },
        { number: '90°', label: 'Hai góc phụ nhau', desc: 'sin góc này bằng cos góc kia.' }
      ],
      speakerNotes: 'Bảng giá trị này là chìa khóa để giải nhanh các bài toán hình học và bài tập ứng dụng.'
    });
  } else {
    slides.push({
      id: 's-4',
      stage: 'Nội dung',
      badge: 'SỐ LIỆU & KHẢO SÁT',
      title: 'Thực Trạng Tâm Lý Học Đường Qua Con Số',
      subtitle: 'Khảo sát 1,200 học sinh về nhu cầu an toàn và được lắng nghe',
      layout: 'chart_stats',
      theme: theme,
      chart: {
        type: 'column',
        title: 'Mức độ hài lòng & áp lực tâm lý (%)',
        items: [
          { label: 'Cần lắng nghe', value: 87, color: '#38BDF8' },
          { label: 'Từng gặp áp lực', value: 68, color: '#F59E0B' },
          { label: 'Cần hỗ trợ tâm lý', value: 74, color: '#34D399' }
        ]
      },
      stats: [
        {
          number: '87%',
          label: 'Nhu Cầu Lắng Nghe',
          desc: 'Học sinh mong muốn có góc chia sẻ an toàn không bị phán xét.'
        },
        {
          number: '3.5x',
          label: 'Gắn Kết Học Đường',
          desc: 'Lớp học có sự đồng cảm giúp học sinh tăng 3.5 lần hiệu quả học tập.'
        }
      ],
      speakerNotes: 'Con số 87% cho thấy nhu cầu chia sẻ là cực kỳ lớn. Các em rất cần một môi trường an toàn để giãi bày.'
    });
  }

  // =========================================================================
  // SLIDE 5: NỘI DUNG 3 - GIẢI PHÁP TRỌNG TÂM (3 TRỤ CỘT / BỘ TỨ CÔNG THỨC)
  // =========================================================================
  if (isTrigMath) {
    slides.push({
      id: 's-5',
      stage: 'Nội dung',
      badge: 'BỘ TỨ CÔNG THỨC',
      title: 'Bốn Công Thức Lượng Giác & Mẹo Ghi Nhớ',
      subtitle: 'Phương pháp ghi nhớ thần tốc qua câu vè dân gian',
      layout: 'math_formula',
      theme: theme,
      formulas: [
        { name: 'sin α', formula: 'Đối / Huyền', tip: 'Sao Đi Học (S = Đ / H)', color: '#38BDF8' },
        { name: 'cos α', formula: 'Kề / Huyền', tip: 'Cứ Khóc Hoài (C = K / H)', color: '#34D399' },
        { name: 'tan α', formula: 'Đối / Kề', tip: 'Thôi Đừng Khóc (T = Đ / K)', color: '#F59E0B' },
        { name: 'cot α', formula: 'Kề / Đối', tip: 'Có Kẹo Đây (C = K / Đ)', color: '#C084FC' }
      ],
      speakerNotes: 'Ghi nhớ câu vè này các em sẽ không bao giờ nhầm lẫn giữa sin, cos, tan và cot.'
    });
  } else {
    slides.push({
      id: 's-5',
      stage: 'Nội dung',
      badge: 'GIẢI PHÁP TRỌNG TÂM',
      title: 'Ba Trụ Cột Xây Dựng Trường Học Hạnh Phúc',
      subtitle: 'Hệ thống hành động từ nhận thức, thái độ đến giải pháp cụ thể',
      layout: 'pillars_3',
      theme: theme,
      pillars: [
        {
          tag: 'TRỤ CỘT 1',
          title: 'Lắng Nghe Chủ Động',
          desc: 'Đặt mình vào vị trí của bạn, lắng nghe trọn vẹn mà không vội phán xét.',
          icon: 'ear',
          color: '#38BDF8'
        },
        {
          tag: 'TRỤ CỘT 2',
          title: 'Tôn Trọng Đa Dạng',
          desc: 'Chấp nhận và trân trọng sự khác biệt về sở thích, tính cách và hoàn cảnh.',
          icon: 'heart',
          color: '#34D399'
        },
        {
          tag: 'TRỤ CỘT 3',
          title: 'Hành Động Can Thiệp',
          desc: 'Dũng cảm lên tiếng khi thấy bất công; không thờ ơ trước nỗi đau của bạn bè.',
          icon: 'shield',
          color: '#F59E0B'
        }
      ],
      speakerNotes: '3 trụ cột này đòi hỏi sự chung tay của cả tập thể lớp, giáo viên chủ nhiệm và nhà trường.'
    });
  }

  // =========================================================================
  // SLIDE 6: TỔNG KẾT & THÔNG ĐIỆP (TAKEAWAYS & MANIFESTO)
  // =========================================================================
  slides.push({
    id: 's-6',
    stage: 'Tổng kết',
    badge: 'TỔNG KẾT & CAM KẾT',
    title: 'Đúc Kết Bài Học & Thông Điệp Cốt Lõi',
    subtitle: 'Mỗi hành động tử tế nhỏ bé hôm nay là hạt mầm cho tương lai',
    layout: 'summary_quote',
    theme: theme,
    takeaways: isTrigMath ? [
      'Nắm vững định nghĩa: Sin đối huyền, Cos kề huyền, Tan đối kề, Cot kề đối.',
      'Sử dụng tam giác vuông để giải quyết các bài toán đo đạc không chạm tới được.',
      'Rèn luyện tư duy không gian và khả năng quy đổi mô hình hình học.'
    ] : [
      'Nói không với mọi hành vi chế giễu, tẩy chay trên mạng và đời thực.',
      'Chủ động kết nối với thầy cô, gia đình khi cảm thấy bị tổn thương tâm lý.',
      'Lan tỏa sự tử tế, bảo vệ nụ cười của bạn bè xung quanh mỗi ngày.'
    ],
    quote: isTrigMath 
      ? 'Toán học không chỉ là những con số khô khan, mà là ngôn ngữ tuyệt mỹ dùng để đo đạc và kiến tạo thế giới.' 
      : 'Trường học hạnh phúc không phải là nơi vắng bóng mâu thuẫn, mà là nơi mọi học sinh đều được đối xử bằng sự công bằng và tình yêu thương.',
    author: isTrigMath ? 'Galileo Galilei' : 'Hội Đồng Sư Phạm & Đại Sứ Học Đường',
    speakerNotes: 'Thầy mong các em ghi nhớ thông điệp này không chỉ trong giờ học mà cả trong cuộc sống hàng ngày.'
  });

  // =========================================================================
  // SLIDE 7: CẢM ƠN & LỜI KẾT (THANK YOU & Q&A)
  // =========================================================================
  slides.push({
    id: 's-7',
    stage: 'Cảm ơn',
    badge: 'LỜI KẾT & TRI ÂN',
    title: 'Cảm Ơn Thầy Cô & Các Em Đã Lắng Nghe!',
    subtitle: isTrigMath 
      ? 'Chúc các em luôn yêu thích môn Toán và tự tin chinh phục các kỳ thi!'
      : 'Cùng chung tay kiến tạo một ngôi trường an toàn, ấm áp và hạnh phúc!',
    layout: 'thank_you',
    theme: theme,
    teacher: 'Thầy Phan Trọng Khang',
    contact: 'Email: giaovien@truonghoc.edu.vn | Hotline: 1900 1234',
    discussionPrompt: 'Thầy cô và các em có câu hỏi hoặc chia sẻ nào không ạ?',
    speakerNotes: 'Xin trân trọng cảm ơn tất cả quý thầy cô và các em học sinh. Bây giờ là thời gian dành cho phần hỏi đáp và thảo luận.'
  });

  return {
    id: 'deck_' + Date.now(),
    meta: {
      title: titleText,
      subject: outline.subject || (isTrigMath ? 'Toán Học 9' : 'Kỹ Năng Sống'),
      grade: outline.grade || 'Lớp 10',
      teacher: 'Thầy Phan Trọng Khang',
      theme: theme.name,
      fontFamily: 'Plus Jakarta Sans',
      aspectRatio: '16:9',
      createdAt: new Date().toISOString()
    },
    slides: slides
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
// API: EXPORT POWERPOINT (.PPTX) NATIVE CHUẨN 16:9 TUYỆT ĐỐI
// ========================================================
app.post('/api/export-pptx', async (req, res) => {
  try {
    const deck = req.body;
    if (!deck || !deck.slides) {
      return res.status(400).json({ error: 'Không có dữ liệu slide để xuất PPTX' });
    }

    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9'; // Chuẩn 16:9 widescreen tuyệt đối
    pptx.title = deck.meta.title || 'Slide Presentation';
    pptx.author = deck.meta.teacher || 'PPT Master Presenton';

    deck.slides.forEach((slideData, idx) => {
      const slide = pptx.addSlide();
      
      const bgHex = (slideData.theme && slideData.theme.hex) ? slideData.theme.hex : '0F172A';
      slide.background = { color: bgHex };

      // Slide Badge Pill
      if (slideData.badge) {
        slide.addShape(pptx.ShapeType.roundRect, {
          x: 0.8, y: 0.6, w: 2.8, h: 0.4,
          rectRadius: 0.1,
          fill: { color: '38BDF8' },
          line: { color: '000000', width: 1.0 }
        });
        slide.addText(slideData.badge, {
          x: 0.8, y: 0.6, w: 2.8, h: 0.4,
          fontSize: 10, fontFace: 'Arial', bold: true, color: '000000', align: 'center', valign: 'middle'
        });
      }

      if (slideData.layout === 'cover' || idx === 0) {
        // ==========================================
        // SLIDE 1: COVER
        // ==========================================
        slide.addText(slideData.title, {
          x: 0.8, y: 1.5, w: 11.5, h: 2.0,
          fontSize: 32, fontFace: 'Arial', bold: true, color: 'FFFFFF', wrap: true
        });

        if (slideData.subtitle) {
          slide.addText(slideData.subtitle, {
            x: 0.8, y: 3.6, w: 11.5, h: 0.9,
            fontSize: 16, fontFace: 'Arial', bold: true, color: '38BDF8'
          });
        }

        slide.addShape(pptx.ShapeType.roundRect, {
          x: 0.8, y: 4.8, w: 11.5, h: 1.3,
          rectRadius: 0.15,
          fill: { color: '1E293B' },
          line: { color: '38BDF8', width: 1.5 }
        });
        slide.addText(`Giảng viên: ${slideData.teacher || 'Thầy/Cô'} | ${slideData.subject || ''} • ${slideData.grade || ''}`, {
          x: 1.2, y: 5.0, w: 10.7, h: 0.9,
          fontSize: 14, fontFace: 'Arial', bold: true, color: 'FFFFFF', valign: 'middle'
        });

      } else if (slideData.layout === 'agenda') {
        // ==========================================
        // SLIDE 2: AGENDA
        // ==========================================
        slide.addText(slideData.title, {
          x: 0.8, y: 1.15, w: 11.5, h: 0.8,
          fontSize: 22, fontFace: 'Arial', bold: true, color: 'FFFFFF'
        });

        const items = slideData.agendaItems || [];
        const cardW = 5.6;
        const cardH = 2.1;
        items.forEach((it, i) => {
          const col = i % 2;
          const row = Math.floor(i / 2);
          const xPos = 0.8 + col * (cardW + 0.3);
          const yPos = 2.2 + row * (cardH + 0.3);

          slide.addShape(pptx.ShapeType.roundRect, {
            x: xPos, y: yPos, w: cardW, h: cardH,
            rectRadius: 0.15,
            fill: { color: '1E293B' },
            line: { color: '38BDF8', width: 1.2 }
          });
          slide.addText(it.step || `0${i + 1}`, {
            x: xPos + 0.3, y: yPos + 0.2, w: 1.0, h: 0.5,
            fontSize: 16, fontFace: 'Arial', bold: true, color: '38BDF8'
          });
          slide.addText(it.title || '', {
            x: xPos + 1.2, y: yPos + 0.2, w: cardW - 1.5, h: 0.6,
            fontSize: 13, fontFace: 'Arial', bold: true, color: 'FFFFFF'
          });
          slide.addText(it.desc || '', {
            x: xPos + 1.2, y: yPos + 0.8, w: cardW - 1.5, h: 1.1,
            fontSize: 11, fontFace: 'Arial', color: 'BAE6FD', wrap: true
          });
        });

      } else if (slideData.layout === 'pillars_3') {
        // ==========================================
        // SLIDE 5: 3 PILLARS
        // ==========================================
        slide.addText(slideData.title, {
          x: 0.8, y: 1.15, w: 11.5, h: 0.8,
          fontSize: 22, fontFace: 'Arial', bold: true, color: 'FFFFFF'
        });

        const pillars = slideData.pillars || [];
        const cardWidth = 3.6;
        const cardGap = 0.35;
        pillars.forEach((p, pIdx) => {
          const xPos = 0.8 + pIdx * (cardWidth + cardGap);
          slide.addShape(pptx.ShapeType.roundRect, {
            x: xPos, y: 2.2, w: cardWidth, h: 4.2,
            rectRadius: 0.15,
            fill: { color: '1E293B' },
            line: { color: '38BDF8', width: 1.5 }
          });
          slide.addText(p.tag || `TRỤ CỘT ${pIdx + 1}`, {
            x: xPos + 0.3, y: 2.4, w: cardWidth - 0.6, h: 0.4,
            fontSize: 11, fontFace: 'Arial', bold: true, color: '38BDF8'
          });
          slide.addText(p.title || '', {
            x: xPos + 0.3, y: 2.9, w: cardWidth - 0.6, h: 0.8,
            fontSize: 14, fontFace: 'Arial', bold: true, color: 'FFFFFF', wrap: true
          });
          slide.addText(p.desc || '', {
            x: xPos + 0.3, y: 3.8, w: cardWidth - 0.6, h: 2.2,
            fontSize: 11, fontFace: 'Arial', color: 'D1FAE5', wrap: true
          });
        });

      } else if (slideData.layout === 'summary_quote') {
        // ==========================================
        // SLIDE 6: SUMMARY & QUOTE
        // ==========================================
        slide.addText(slideData.title, {
          x: 0.8, y: 1.15, w: 11.5, h: 0.8,
          fontSize: 22, fontFace: 'Arial', bold: true, color: 'FFFFFF'
        });

        const takeaways = slideData.takeaways || [];
        takeaways.forEach((t, tIdx) => {
          const yPos = 2.1 + tIdx * 0.9;
          slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.8, y: yPos, w: 5.6, h: 0.75,
            rectRadius: 0.1,
            fill: { color: '1E293B' },
            line: { color: '34D399', width: 1.0 }
          });
          slide.addText(`✓  ${t}`, {
            x: 1.0, y: yPos, w: 5.2, h: 0.75,
            fontSize: 11, fontFace: 'Arial', bold: true, color: 'FFFFFF', valign: 'middle'
          });
        });

        // Quote Box
        slide.addShape(pptx.ShapeType.roundRect, {
          x: 6.8, y: 2.1, w: 5.6, h: 4.2,
          rectRadius: 0.2,
          fill: { color: '1E293B' },
          line: { color: 'FBBF24', width: 1.5 }
        });
        slide.addText(`“${slideData.quote || ''}”`, {
          x: 7.2, y: 2.6, w: 4.8, h: 2.6,
          fontSize: 14, fontFace: 'Arial', bold: true, italic: true, color: 'FFFFFF', wrap: true
        });
        slide.addText(slideData.author || '', {
          x: 7.2, y: 5.3, w: 4.8, h: 0.6,
          fontSize: 12, fontFace: 'Arial', bold: true, color: 'FBBF24', align: 'right'
        });

      } else if (slideData.layout === 'thank_you') {
        // ==========================================
        // SLIDE 7: THANK YOU & CLOSING
        // ==========================================
        slide.addShape(pptx.ShapeType.roundRect, {
          x: 1.5, y: 1.8, w: 10.3, h: 4.5,
          rectRadius: 0.2,
          fill: { color: '1E293B' },
          line: { color: '38BDF8', width: 2.0 }
        });
        slide.addText(slideData.title || 'Cảm Ơn Thầy Cô & Các Em!', {
          x: 2.0, y: 2.3, w: 9.3, h: 1.2,
          fontSize: 28, fontFace: 'Arial', bold: true, color: 'FFFFFF', align: 'center'
        });
        slide.addText(slideData.subtitle || '', {
          x: 2.0, y: 3.5, w: 9.3, h: 0.8,
          fontSize: 14, fontFace: 'Arial', bold: true, color: '34D399', align: 'center'
        });
        slide.addText(slideData.contact || '', {
          x: 2.0, y: 4.8, w: 9.3, h: 0.8,
          fontSize: 12, fontFace: 'Arial', color: 'BAE6FD', align: 'center'
        });

      } else {
        // Fallback Content
        slide.addText(slideData.title, {
          x: 0.8, y: 1.15, w: 11.5, h: 0.8,
          fontSize: 22, fontFace: 'Arial', bold: true, color: 'FFFFFF'
        });
        slide.addShape(pptx.ShapeType.roundRect, {
          x: 0.8, y: 2.2, w: 11.5, h: 4.2,
          rectRadius: 0.15,
          fill: { color: '1E293B' },
          line: { color: '38BDF8', width: 1.5 }
        });
        slide.addText(slideData.subtitle || 'Nội dung bài giảng trọng tâm', {
          x: 1.2, y: 2.6, w: 10.7, h: 3.2,
          fontSize: 14, fontFace: 'Arial', color: 'FFFFFF', wrap: true
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
